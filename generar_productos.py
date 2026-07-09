# -*- coding: utf-8 -*-
"""
Genera data/productos.json para el tablero HTML.

Desde 2026-07-08 el insumo es el formato SDP nuevo (convertido por
convertir_formato_productos.py), que reemplaza al corte Q1 2025:
- Solo entran los productos VIGENTES (99 de 123). Decisión de Carolina.
- 71 productos cambiaron de tipo de anualización frente al corte viejo
  (ej. 1.1.1 pasó de Constante a Suma); el tipo manda sobre el cálculo.
- La serie llega hasta 2026, que es un año PARCIAL (corte a 1er trimestre):
  se marca con anio_parcial en el JSON y el HTML lo rotula "2026p".
- El acumulado ya NO viene del insumo (el bloque del formato trae textos
  de ERROR en varios Creciente): se calcula aquí según el tipo:
    Suma      -> suma corrida de los valores anuales
    Creciente -> nivel del último trimestre reportado del año
    Constante / Decreciente -> igual al valor anual (no hay acumulación
                real; el HTML no muestra columnas de acumulado para estos)

Reglas de anualización por tipo (valor del año):
- Suma: suma de los trimestres reportados del año.
- Constante / Creciente / Decreciente reportan nivel año corrido: el valor
  del año es el último trimestre reportado (Q4 en años cerrados; en el año
  parcial, el último corte disponible). Para Creciente el valor del año es
  el nivel menos el nivel del año anterior (regla de Reglas de
  validación.xlsx); si no hay año anterior, el nivel completo.
- Años sin ningún dato dentro de la ventana de reporte se muestran en 0.

El JSON alimenta productos.html: filtros por dimensión y sector, tabla
maestro, ficha del producto y gráfica Valor vs Meta por año.

Ejecución: python generar_productos.py  (desde tablero-ppdj/)
"""

import json
import math
import os
from datetime import date

import pandas as pd

from comun_pipeline import (DIMENSIONES, DIR_INPUTS, cargar_objetivos,
                            clasificar_indicador_tipo,
                            extraer_anio_y_trimestre, normalizar_texto,
                            parsear_fecha_mixta)

RUTA_INPUT = os.path.join(DIR_INPUTS, 'Seguimiento_Productos_PPDJ_2026_excel.xlsx')
RUTA_JSON = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'data', 'productos.json')

ANIOS = list(range(2019, 2027))
ANIO_PARCIAL = 2026  # en este corte solo trae el 1er trimestre
NOTA_PARCIAL = '2026p: dato parcial de 2026, con corte al primer trimestre.'


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


def clave_dimension(valor):
    """Normaliza la Key de dimensión a la forma "1.": el insumo viejo la
    traía numérica (1.0) y el formato SDP la trae como texto ("1.")."""
    texto = str(valor).strip()
    if texto.endswith('.'):
        return texto
    return f'{int(float(texto))}.'


def porcentaje_avance(reportado, meta, tipo, linea_base):
    """Porcentaje de avance según la metodología oficial de anualización:
    - Creciente: la línea base es el punto de partida del esfuerzo y se
      descuenta de ambos lados: (reportado - LB) / (meta - LB). Es la misma
      fórmula del bloque "Avance de la Vigencia" del formato SDP. Si el
      reporte es menor a la línea base el resultado queda negativo (el
      formato lo marca como ERROR de reporte); si no hay línea base
      numérica se usa 0.
    - Los demás tipos: reportado / meta (la línea base no impacta).
    Devuelve None si falta algún dato o el denominador es 0."""
    if reportado is None or meta is None:
        return None
    if tipo == 'Creciente':
        lb = linea_base if linea_base is not None else 0.0
        denominador = meta - lb
        if denominador == 0:
            return None
        return redondear((reportado - lb) / denominador * 100, 1)
    if meta == 0:
        return None
    return redondear(reportado / meta * 100, 1)


def construir():
    bd = pd.read_excel(RUTA_INPUT, sheet_name='Cuanti', skiprows=0, engine='openpyxl')
    # Solo los productos vigentes entran al tablero (el formato trae
    # "No vigente" y "No Vigente"; se filtra por el valor exacto "Vigente")
    if 'Estado del Indicador' in bd.columns:
        bd = bd[bd['Estado del Indicador'] == 'Vigente']

    columnas_trimestrales = [c for c in bd.columns if extraer_anio_y_trimestre(c)[0] is not None]

    objetivos = cargar_objetivos(pd)
    objetivos_por_key = {f'{int(k)}.': normalizar_texto(o)
                         for k, o in zip(objetivos['Key'], objetivos['Objetivo'])}

    items = []
    sectores = set()
    for _, fila in bd.iterrows():
        numero = int(fila['Producto No.'])
        key_dim = clave_dimension(fila['Key'])
        tipo = normalizar_texto(fila['Tipo de anualización'])
        linea_base = a_numero(fila['Valor Linea Base'])
        sector = normalizar_texto(fila['Sector Líder'])
        if sector:
            sectores.add(sector)

        # Valores trimestrales por año, solo los numéricos presentes
        trimestres = {}
        for col in columnas_trimestrales:
            anio, tri = extraer_anio_y_trimestre(col)
            v = a_numero(fila[col])
            if v is not None:
                trimestres.setdefault(anio, {})[tri] = v

        # Ventana de reporte: los años sin ningún trimestre reportado se
        # muestran en 0 (decisión de Carolina 2026-07-07, igual que el Power
        # BI publicado) pero solo entre el año de inicio del producto y el
        # último año CERRADO. El año parcial (2026) nunca se rellena.
        fecha_ini = parsear_fecha_mixta(fila['Fecha de Inicio'])
        anio_inicio = fecha_ini.year if fecha_ini else ANIOS[0]
        anio_ultimo = a_numero(fila.get('Año del último reporte'))
        corte = str(fila.get('Corte del último reporte') or '').strip()
        if anio_ultimo is None:
            anios_con_dato = [a for a, ts in trimestres.items() if ts]
            anio_ultimo = max(anios_con_dato) if anios_con_dato else None
        elif corte and corte != 'Q4':
            anio_ultimo = int(anio_ultimo) - 1
        if anio_ultimo is not None:
            anio_ultimo = min(int(anio_ultimo), ANIO_PARCIAL - 1)

        # Meta acumulada COHERENTE con el tipo de anualización, calculada a
        # partir de las metas anuales (la columna acumulada del insumo trae
        # vacíos y ceros inconsistentes).
        # - Suma: suma corrida de las metas anuales (arranca en 2018, primer
        #   año con columna de meta en el insumo).
        # - Constante / Creciente / Decreciente: la meta anual, que ya
        #   expresa el nivel objetivo del año.
        metas_anuales = {a: a_numero(fila.get(f'Meta_programada_{a}')) for a in ANIOS}
        metas_acum = {}
        meta_2018 = a_numero(fila.get('Meta_programada_2018'))
        suma_corrida = meta_2018 if meta_2018 is not None else 0.0
        hay_meta = meta_2018 is not None
        for a in ANIOS:
            m = metas_anuales[a]
            if tipo == 'Suma':
                if m is not None:
                    suma_corrida += m
                    hay_meta = True
                metas_acum[a] = suma_corrida if hay_meta else None
            else:
                metas_acum[a] = m

        serie = []
        acum_suma = None      # acumulado corrido (solo tipo Suma)
        nivel_previo = None   # último nivel conocido (solo tipo Creciente)
        for anio in ANIOS:
            validos = trimestres.get(anio, {})
            en_ventana = anio_ultimo is not None and anio_inicio <= anio <= anio_ultimo

            if tipo == 'Suma':
                valor = sum(validos.values()) if validos else None
                if valor is None and en_ventana:
                    valor = 0.0
                if valor is not None:
                    acum_suma = (acum_suma or 0.0) + valor
                acumulado = acum_suma if valor is not None else None
            elif tipo == 'Creciente':
                nivel = validos[max(validos)] if validos else None
                if nivel is None and en_ventana:
                    # sin reporte dentro de la ventana el nivel persiste
                    nivel = nivel_previo if nivel_previo is not None else 0.0
                if nivel is not None:
                    valor = nivel - (nivel_previo if nivel_previo is not None else 0.0)
                    nivel_previo = nivel
                    acumulado = nivel
                else:
                    valor = None
                    acumulado = None
            else:
                # Constante / Decreciente: nivel del último trimestre; no hay
                # acumulación real, el acumulado se iguala al valor anual
                valor = validos[max(validos)] if validos else None
                if valor is None and en_ventana:
                    valor = 0.0
                acumulado = valor

            meta = metas_anuales[anio]
            # 4 decimales: en conteos no cambia nada, pero los indicadores de
            # proporción (0 a 1) pierden precisión real con 2 (0,3626 -> 0,36
            # se mostraría 36% en vez de 36,3%)
            diff = redondear(valor - meta, 4) if valor is not None and meta is not None else None
            meta_acum = metas_acum[anio]
            # En Creciente el % oficial del año se mide sobre el NIVEL
            # reportado (acumulado), no sobre el incremento anual
            reportado_anio = acumulado if tipo == 'Creciente' else valor
            porcentaje = porcentaje_avance(reportado_anio, meta, tipo, linea_base)
            porcentaje_acum = porcentaje_avance(acumulado, meta_acum, tipo, linea_base)
            serie.append({'anio': anio, 'valor': redondear(valor, 4), 'meta': redondear(meta, 4),
                          'porcentaje': porcentaje, 'diff': diff,
                          'acumulado': redondear(acumulado, 4), 'meta_acum': redondear(meta_acum, 4),
                          'porcentaje_acum': porcentaje_acum})

        items.append({
            'llave': f'P{numero}',
            'numero': numero,
            'key_dimension': key_dim,
            'dimension': DIMENSIONES[key_dim]['nombre'],
            'esperado': normalizar_texto(fila['Producto esperado']),
            'indicador': normalizar_texto(fila['Nombre indicador de Producto']),
            'sector_lider': sector,
            'indicador_tipo': clasificar_indicador_tipo(fila['Nombre indicador de Producto']),
            'ponderacion': a_numero(fila['Ponderación relativa del Producto (%)']),
            'valor_linea_base': (a_numero(fila['Valor Linea Base'])
                                 if a_numero(fila['Valor Linea Base']) is not None
                                 else normalizar_texto(fila['Valor Linea Base'])),
            'tipo_anualizacion': tipo,
            'fecha_inicio': a_fecha_iso(fila['Fecha de Inicio']),
            'fecha_fin': a_fecha_iso(fila['Fecha de Finalización']),
            'serie': serie,
        })

    datos = {
        'generado': date.today().isoformat(),
        'vista': 'productos',
        'anios': ANIOS,
        'anio_parcial': ANIO_PARCIAL,
        'nota_parcial': NOTA_PARCIAL,
        'dimensiones': [{'key': k, 'nombre': v['nombre'], 'color': v['color'],
                         'objetivo': objetivos_por_key.get(k)}
                        for k, v in DIMENSIONES.items()],
        'sectores': sorted(sectores),
        'items': items,
    }
    return datos


def exportar_base_excel(datos, ruta_excel):
    """Exporta la base procesada a Excel (una fila por producto/resultado,
    con Indicador tipo como columna y los valores anuales corregidos).
    Es la versión presentable de "cómo quedó la base en Python"."""
    anio_parcial = datos.get('anio_parcial')
    filas = []
    for it in datos['items']:
        fila = {
            'Key': it['key_dimension'],
            'Dimensión': it['dimension'],
            'No.': it['numero'],
            'Esperado': it['esperado'],
            'Nombre indicador': it['indicador'],
            'Indicador tipo': it['indicador_tipo'],
            'Sector líder': it['sector_lider'],
            'Tipo de anualización': it['tipo_anualizacion'],
            'Fecha inicio': it['fecha_inicio'],
            'Fecha fin': it['fecha_fin'],
        }
        tiene_acumulado = any(p.get('acumulado') is not None for p in it['serie'])
        for p in it['serie']:
            a = f"{p['anio']}p" if p['anio'] == anio_parcial else p['anio']
            fila[f'Valor_{a}'] = p['valor']
            fila[f'Meta_{a}'] = p['meta']
            fila[f'Diff_{a}'] = p['diff']
            if tiene_acumulado:
                fila[f'Acumulado_{a}'] = p.get('acumulado')
                fila[f'Meta_acum_{a}'] = p.get('meta_acum')
        filas.append(fila)
    os.makedirs(os.path.dirname(ruta_excel), exist_ok=True)
    pd.DataFrame(filas).to_excel(ruta_excel, index=False, engine='openpyxl')
    return ruta_excel


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
    ruta_base = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                             'salidas', 'Base_productos_PPDJ_python.xlsx')
    exportar_base_excel(datos, ruta_base)
    kb = os.path.getsize(RUTA_JSON) / 1024
    print(f'Generado: {RUTA_JSON} (+ .js) ({kb:.0f} KB)')
    print(f'Base Excel: {ruta_base}')
    print(f'Items: {len(datos["items"])} | Dimensiones: {len(datos["dimensiones"])} | Sectores: {len(datos["sectores"])}')


if __name__ == '__main__':
    main()
