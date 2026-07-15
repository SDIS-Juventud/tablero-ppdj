# -*- coding: utf-8 -*-
"""
Genera data/resultados.json para el tablero HTML.

Igual que generar_productos.py pero para la vista de Resultados:
- Datos trimestrales 2018-2024; metas 2020-2025 (el año 2025 entra a la
  serie solo con meta, sin valor, para que la línea de meta continúe).
- Regla "Creciente" corregida (Q4 menos Q4 del año anterior): 17 casos.
- "Decreciente" (10 casos) se mantiene como Q4, igual que Constante,
  porque no hay regla documentada — pendiente de definición del equipo.

Ejecución: python programas/generar_resultados.py  (desde la raíz del repo)
"""

import json
import os
from datetime import date

import pandas as pd

from comun_pipeline import (DIMENSIONES, DIR_INPUTS, anualizar_tablero,
                            cargar_objetivos, clasificar_indicador_tipo,
                            extraer_anio_y_trimestre, normalizar_texto,
                            parsear_fecha_mixta)
from generar_productos import (a_fecha_iso, a_numero, porcentaje_avance,
                               redondear)

# Insumo del ciclo 2025 (generado por convertir_formato_resultados.py a
# partir del formato crudo de la SDP; trae los datos de 2025)
RUTA_INPUT = os.path.join(DIR_INPUTS, 'Seguimiento_Resultados_PPDJ_2025_excel.xlsx')
RUTA_JSON = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'data', 'resultados.json')

# Los resultados como tal arrancan en 2020 (año de implementación de la
# política — aclaración de Carolina, 2026-07-08, igual que el Power BI).
# El insumo trae trimestres de 2018-2019 pero no entran a la vista;
# 2025 solo tiene meta.
ANIOS = list(range(2020, 2026))


def construir():
    bd = pd.read_excel(RUTA_INPUT, sheet_name='Cuanti', engine='openpyxl')
    # Solo los resultados vigentes entran al tablero. En el ciclo 2025 salió
    # el 2.7 (No Vigente) y entró el 2.8; OJO: eso corrió la numeración un
    # puesto desde el antiguo R12 — "Resultado No." no es un identificador
    # estable entre ciclos (comparar por el código del texto, ej. "3.1").
    if 'Estado del Indicador' in bd.columns:
        bd = bd[bd['Estado del Indicador'] == 'Vigente']
    columnas_trimestrales = [c for c in bd.columns if extraer_anio_y_trimestre(c)[0] is not None]

    objetivos = cargar_objetivos(pd)
    objetivos_por_key = {f'{int(k)}.': normalizar_texto(o)
                         for k, o in zip(objetivos['Key'], objetivos['Objetivo'])}

    items = []
    sectores = set()
    for _, fila in bd.iterrows():
        numero = int(fila['Resultado No.'])
        key_dim = f"{int(fila['Key'])}."
        tipo = fila['Tipo de anualización']
        tipo_limpio = normalizar_texto(tipo)
        linea_base = a_numero(fila['Valor Linea Base'])
        sector = normalizar_texto(fila['Sector Líder'])
        if sector:
            sectores.add(sector)

        trimestres = {}
        for col in columnas_trimestrales:
            anio, tri = extraer_anio_y_trimestre(col)
            trimestres.setdefault(anio, {})[tri] = fila[col]

        # Ventana de reporte: años sin reporte se muestran en 0 desde 2020
        # (inicio de la implementación, para TODOS los resultados) hasta su
        # último año cerrado; después queda null (sin dato).
        anio_inicio = 2020
        anio_ultimo = a_numero(fila.get('Año del último reporte'))
        corte = str(fila.get('Corte del último reporte') or '').strip()
        if anio_ultimo is None:
            anios_con_dato = [a for a, ts in trimestres.items()
                              if any(a_numero(v) is not None for v in ts.values())]
            anio_ultimo = max(anios_con_dato) if anios_con_dato else None
        elif corte and corte != 'Q4':
            anio_ultimo = int(anio_ultimo) - 1

        serie = []
        for anio in ANIOS:
            # Los resultados son indicadores de NIVEL (tasas, porcentajes):
            # el valor del año es el último reporte (Q4), sin descontar el
            # año anterior. La corrección de "Creciente" (Q4 menos Q4
            # anterior) aplica solo a productos, que son conteos reportados
            # de forma acumulada. Diferenciar una tasa producía valores
            # negativos sin sentido (corrección de Carolina, 2026-07-08).
            valor = anualizar_tablero(
                {t: a_numero(v) for t, v in trimestres.get(anio, {}).items()},
                tipo)
            if valor is None and anio_ultimo is not None and anio_inicio <= anio <= int(anio_ultimo):
                valor = 0.0
            meta = a_numero(fila.get(f'Meta_{anio}'))
            # 4 decimales: los resultados son proporciones y tasas — con 2
            # decimales se pierde precisión real (0,3626 se mostraría 36%
            # en vez de 36,3%)
            diff = redondear(valor - meta, 4) if valor is not None and meta is not None else None
            # % oficial: en Creciente descuenta la línea base; el valor de
            # los resultados ya es el nivel reportado (no se diferencia)
            porcentaje = porcentaje_avance(valor, meta, tipo_limpio, linea_base)
            serie.append({'anio': anio, 'valor': redondear(valor, 4), 'meta': redondear(meta, 4),
                          'porcentaje': porcentaje, 'diff': diff})

        items.append({
            'llave': f'R{numero}',
            'numero': numero,
            'key_dimension': key_dim,
            'dimension': DIMENSIONES[key_dim]['nombre'],
            'esperado': normalizar_texto(fila['Resultado esperado']),
            'indicador': normalizar_texto(fila['Nombre indicador de Resultado']),
            'sector_lider': sector,
            'indicador_tipo': clasificar_indicador_tipo(fila['Nombre indicador de Resultado']),
            'ponderacion': a_numero(fila['Ponderación relativa del Resultado (%)']),
            'valor_linea_base': (a_numero(fila['Valor Linea Base'])
                                 if a_numero(fila['Valor Linea Base']) is not None
                                 else normalizar_texto(fila['Valor Linea Base'])),
            'tipo_anualizacion': normalizar_texto(tipo),
            'fecha_inicio': a_fecha_iso(fila['Fecha de Inicio']),
            'fecha_fin': a_fecha_iso(fila['Fecha de Finalización']),
            'serie': serie,
        })

    datos = {
        'generado': date.today().isoformat(),
        'vista': 'resultados',
        'anios': ANIOS,
        'dimensiones': [{'key': k, 'nombre': v['nombre'], 'color': v['color'],
                         'objetivo': objetivos_por_key.get(k)}
                        for k, v in DIMENSIONES.items()],
        'sectores': sorted(sectores),
        'items': items,
    }
    return datos


def main():
    from generar_productos import escribir_json_y_js, exportar_base_excel
    datos = construir()
    escribir_json_y_js(datos, RUTA_JSON, 'DATOS_RESULTADOS')
    ruta_base = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                             'salidas', 'Base_resultados_PPDJ_python.xlsx')
    exportar_base_excel(datos, ruta_base)
    kb = os.path.getsize(RUTA_JSON) / 1024
    print(f'Generado: {RUTA_JSON} (+ .js) ({kb:.0f} KB)')
    print(f'Base Excel: {ruta_base}')
    print(f'Items: {len(datos["items"])} | Dimensiones: {len(datos["dimensiones"])} | Sectores: {len(datos["sectores"])}')


if __name__ == '__main__':
    main()
