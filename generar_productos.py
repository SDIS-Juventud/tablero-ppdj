# -*- coding: utf-8 -*-
"""
Genera data/productos.json para el tablero HTML.

A diferencia de los scripts de paridad (que replican el R con sus bugs),
este generador aplica las reglas CORREGIDAS:
- "Creciente": Q4 del año menos Q4 del año anterior (regla real de
  Reglas de validación.xlsx). Afecta 15 productos.
- Años sin ningún dato trimestral quedan como null (sin dato), no como 0.

El JSON alimenta productos.html: filtros por dimensión y sector, tabla
maestro, ficha del producto y gráfica Valor vs Meta por año.

Ejecución: python generar_productos.py  (desde tablero-ppdj/)
"""

import json
import math
import os
from datetime import date, datetime

import pandas as pd

from comun_pipeline import (DIMENSIONES, DIR_INPUTS, anualizar_tablero,
                            cargar_objetivos, extraer_anio_y_trimestre,
                            normalizar_texto, parsear_fecha_mixta)

RUTA_INPUT = os.path.join(DIR_INPUTS, 'Seguimiento_Productos_PPDJ_Q1_2025_excel.xlsx')
RUTA_JSON = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'data', 'productos.json')

ANIOS = list(range(2019, 2026))


def a_numero(valor):
    """Convierte a float nativo de Python; None si es nulo o no numérico."""
    if valor is None or (isinstance(valor, float) and math.isnan(valor)):
        return None
    try:
        return float(valor)
    except (TypeError, ValueError):
        return None


def a_fecha_iso(valor):
    """Fecha en formato ISO (YYYY-MM-DD) o None."""
    fecha = parsear_fecha_mixta(valor)
    return fecha.strftime('%Y-%m-%d') if fecha else None


def redondear(valor, decimales=2):
    return None if valor is None else round(valor, decimales)


def construir():
    bd = pd.read_excel(RUTA_INPUT, sheet_name='Cuanti', skiprows=2, engine='openpyxl')

    # Valores trimestrales por producto: {numero: {(año, trimestre): valor}}
    columnas_trimestrales = [c for c in bd.columns if extraer_anio_y_trimestre(c)[0] is not None]

    objetivos = cargar_objetivos(pd)
    objetivos_por_key = {f'{int(k)}.': normalizar_texto(o)
                         for k, o in zip(objetivos['Key'], objetivos['Objetivo'])}

    items = []
    sectores = set()
    for _, fila in bd.iterrows():
        numero = int(fila['Producto No.'])
        key_dim = f"{int(fila['Key'])}."
        tipo = fila['Tipo de anualización']
        sector = normalizar_texto(fila['Sector Líder'])
        if sector:
            sectores.add(sector)

        trimestres = {}
        for col in columnas_trimestrales:
            anio, tri = extraer_anio_y_trimestre(col)
            trimestres.setdefault(anio, {})[tri] = fila[col]

        # Ventana de reporte: los años sin ningún trimestre reportado se
        # muestran en 0 (decisión de Carolina 2026-07-07, igual que el Power
        # BI publicado) pero solo entre el año de inicio del producto y el
        # último año CERRADO (si el corte va en Q1-Q3, ese año queda "en
        # curso" y no se rellena). Fuera de la ventana queda null (sin dato).
        fecha_ini = parsear_fecha_mixta(fila['Fecha de Inicio'])
        anio_inicio = fecha_ini.year if fecha_ini else ANIOS[0]
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
            q4_anterior = a_numero(trimestres.get(anio - 1, {}).get(4))
            valor = anualizar_tablero(
                {t: a_numero(v) for t, v in trimestres.get(anio, {}).items()},
                tipo, q4_anio_anterior=q4_anterior)
            if valor is None and anio_ultimo is not None and anio_inicio <= anio <= int(anio_ultimo):
                valor = 0.0
            meta = a_numero(fila.get(f'Meta_programada_{anio}'))
            diff = redondear(valor - meta) if valor is not None and meta is not None else None
            porcentaje = (redondear(valor / meta * 100, 1)
                          if valor is not None and meta not in (None, 0) else None)
            # Acumulado y meta acumulada: vienen tal cual del insumo (el pbix
            # los grafica directo; el R los descartaba). La gráfica del
            # tablero usa estas dos series para coincidir con el Power BI.
            acumulado = a_numero(fila.get(f'Acumulado {anio}'))
            meta_acum = a_numero(fila.get(f'Meta_programada_acum_{anio}'))
            if acumulado is None and anio_ultimo is not None and anio_inicio <= anio <= int(anio_ultimo):
                acumulado = 0.0
            porcentaje_acum = (redondear(acumulado / meta_acum * 100, 1)
                               if acumulado is not None and meta_acum not in (None, 0) else None)
            serie.append({'anio': anio, 'valor': redondear(valor), 'meta': redondear(meta),
                          'porcentaje': porcentaje, 'diff': diff,
                          'acumulado': redondear(acumulado), 'meta_acum': redondear(meta_acum),
                          'porcentaje_acum': porcentaje_acum})

        items.append({
            'llave': f'P{numero}',
            'numero': numero,
            'key_dimension': key_dim,
            'dimension': DIMENSIONES[key_dim]['nombre'],
            'esperado': normalizar_texto(fila['Producto esperado']),
            'indicador': normalizar_texto(fila['Nombre indicador de Producto']),
            'sector_lider': sector,
            'ponderacion': a_numero(fila['Ponderación relativa del Producto (%)']),
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
        'vista': 'productos',
        'anios': ANIOS,
        'dimensiones': [{'key': k, 'nombre': v['nombre'], 'color': v['color'],
                         'objetivo': objetivos_por_key.get(k)}
                        for k, v in DIMENSIONES.items()],
        'sectores': sorted(sectores),
        'items': items,
    }
    return datos


def escribir_json_y_js(datos, ruta_json, nombre_variable):
    """Escribe el JSON y un compañero .js (window.<VARIABLE> = ...).
    El .js es lo que cargan las páginas HTML con <script src>: funciona
    al abrir el archivo con doble clic (file://), donde fetch() falla por
    CORS. El JSON queda para inspección y otros consumidores."""
    os.makedirs(os.path.dirname(ruta_json), exist_ok=True)
    with open(ruta_json, 'w', encoding='utf-8') as f:
        json.dump(datos, f, ensure_ascii=False, indent=1)
    ruta_js = ruta_json[:-5] + '.js'
    with open(ruta_js, 'w', encoding='utf-8') as f:
        f.write(f'window.{nombre_variable} = ')
        json.dump(datos, f, ensure_ascii=False, separators=(',', ':'))
        f.write(';\n')
    return ruta_js


def main():
    datos = construir()
    escribir_json_y_js(datos, RUTA_JSON, 'DATOS_PRODUCTOS')
    kb = os.path.getsize(RUTA_JSON) / 1024
    print(f'Generado: {RUTA_JSON} (+ .js) ({kb:.0f} KB)')
    print(f'Items: {len(datos["items"])} | Dimensiones: {len(datos["dimensiones"])} | Sectores: {len(datos["sectores"])}')


if __name__ == '__main__':
    main()
