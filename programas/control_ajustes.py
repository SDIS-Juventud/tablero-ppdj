# -*- coding: utf-8 -*-
"""
Carga el Control de ajustes de la PPDJ y lo organiza por código de
producto/resultado, para que los generadores del tablero anexen a cada
ítem la historia de sus ajustes (qué se ajustó, cuándo y en qué estado).

Insumo: el archivo más reciente que coincida con el patrón
    Insumos/Control de ajustes PPDJ*.xlsx
La Subdirección versiona el archivo en el nombre (ej. "V11_05_2026" es la
fecha de la versión). Cuando llegue una versión nueva basta con dejarla en
la carpeta Insumos/ y volver a correr los generadores: este módulo toma el
archivo de modificación más reciente, sin tocar código.

Estructura esperada del Excel (dos hojas):
- "Hist. Aprobados": historial de solicitudes tramitadas. Encabezados en
  la fila 3 (índice 2): Sector PPJ | Entidad Solicitante | Objetivo |
  Producto / Resultado | Ajuste realizado | Tipo de ajuste | Fecha de
  aprobación | Trámite. La columna de fecha trae fechas reales o el texto
  "No aprobado".
- "Importantes pend. No Aprobados": solicitudes importantes denegadas o en
  trámite, con Fecha de evaluación e Instancia en vez de fecha y trámite
  de aprobación.

El código del producto ("1.1.2", tres partes) o del resultado ("6.4", dos
partes, a veces escrito "R 6.4.") se extrae del inicio de la columna
"Producto / Resultado". Las filas sin código reconocible (ej. ajustes
generales de encuestas) se devuelven aparte como "sin asignar".

Al tablero solo pasan los ajustes APROBADOS (decisión de Carolina,
2026-07-28): el sitio es público y las solicitudes denegadas o en trámite
son información interna del proceso. Para inspeccionarlas todas, correr
este módulo directo o llamar cargar_ajustes(solo_aprobados=False).
"""

import glob
import os
import re
from datetime import datetime

import pandas as pd

from comun_pipeline import RAIZ_PROYECTO, normalizar_texto, parsear_fecha_mixta

# El control vive en la raíz de Insumos/ (junto a la carpeta Datos tablero)
DIR_INSUMOS_RAIZ = os.path.normpath(os.path.join(RAIZ_PROYECTO, '..', 'Insumos'))
PATRON_CONTROL = os.path.join(DIR_INSUMOS_RAIZ, 'Control de ajustes PPDJ*.xlsx')

# Columnas por posición (el encabezado real está en la fila 3 y algunas
# celdas traen saltos de línea, por eso se leen por índice):
COL_PRODUCTO, COL_AJUSTE, COL_TIPO, COL_FECHA, COL_TRAMITE = 3, 4, 5, 6, 7


def ruta_control_mas_reciente():
    """Ruta del Control de ajustes más reciente en Insumos/, o None."""
    candidatos = glob.glob(PATRON_CONTROL)
    if not candidatos:
        return None
    return max(candidatos, key=os.path.getmtime)


def _extraer_codigo(texto):
    """Extrae el código del texto "Producto / Resultado".
    Devuelve (codigo, es_producto): "1.1.2" con es_producto=True si el
    código tiene tres partes; "6.4" con es_producto=False si tiene dos.
    (None, None) si no hay código reconocible."""
    m = re.match(r'\s*(?:R\s*)?(\d+)\.\s*(\d+)(?:\.\s*(\d+))?', str(texto or ''))
    if not m:
        return None, None
    if m.group(3):
        return f'{m.group(1)}.{m.group(2)}.{m.group(3)}', True
    return f'{m.group(1)}.{m.group(2)}', False


def _estado_y_fecha(celda_fecha, tramite, hoja_aprobados):
    """Deriva el estado del ajuste y la fecha en ISO (o None).
    - Hoja de aprobados: fecha real -> Aprobado; texto "No aprobado" -> No
      aprobado (en trámite del 2do semestre, según el propio control).
    - Hoja de pendientes: "Denegado" si el texto del trámite lo dice;
      "En trámite" en los demás casos."""
    fecha = None
    try:
        parseada = parsear_fecha_mixta(celda_fecha)
        if isinstance(parseada, datetime):
            fecha = parseada.strftime('%Y-%m-%d')
    except (ValueError, TypeError):
        pass
    if hoja_aprobados:
        return ('Aprobado', fecha) if fecha else ('No aprobado', None)
    texto_tramite = str(tramite or '').lower()
    if 'denegado' in texto_tramite or 'denegada' in texto_tramite:
        return 'Denegado', fecha
    return 'En trámite', fecha


def cargar_ajustes(solo_aprobados=True):
    """Lee el Control de ajustes y devuelve:
    (ajustes_productos, ajustes_resultados, sin_asignar, ruta_usada)
    - ajustes_productos: {"1.1.2": [ajuste, ...]} con códigos de 3 partes
    - ajustes_resultados: {"6.4": [ajuste, ...]} con códigos de 2 partes
    - sin_asignar: lista de descripciones de filas sin código
    - ruta_usada: ruta del Excel leído (None si no se encontró)
    Cada ajuste es un dict {fecha, estado, ajuste}, ordenado por fecha
    (los sin fecha al final).
    Con solo_aprobados=True (el valor con el que corren los generadores)
    se omiten las solicitudes no aprobadas, denegadas o en trámite."""
    ruta = ruta_control_mas_reciente()
    if ruta is None:
        return {}, {}, [], None

    productos, resultados, sin_asignar = {}, {}, []
    hojas = pd.read_excel(ruta, sheet_name=None, header=None, engine='openpyxl')
    for nombre_hoja, df in hojas.items():
        hoja_aprobados = 'aprobado' in nombre_hoja.lower() and 'no aprobado' not in nombre_hoja.lower()
        # las filas de datos empiezan después del encabezado (fila índice 2)
        for i in range(3, len(df)):
            fila = df.iloc[i]
            texto_producto = normalizar_texto(fila[COL_PRODUCTO])
            texto_ajuste = normalizar_texto(fila[COL_AJUSTE])
            if not texto_producto and not texto_ajuste:
                continue
            estado, fecha = _estado_y_fecha(fila[COL_FECHA], fila[COL_TRAMITE], hoja_aprobados)
            if solo_aprobados and estado != 'Aprobado':
                continue
            # A las fichas del tablero solo van fecha y ajuste (decisión de
            # Carolina); el estado se conserva para la inspección de los no
            # aprobados. El "tipo de ajuste según criterios SDP" no se usa.
            registro = {
                'fecha': fecha,
                'estado': estado,
                'ajuste': texto_ajuste,
            }
            codigo, es_producto = _extraer_codigo(texto_producto)
            if codigo is None:
                sin_asignar.append(f'{texto_producto or "(sin producto)"} — {texto_ajuste or "(sin detalle)"}')
            elif es_producto:
                productos.setdefault(codigo, []).append(registro)
            else:
                resultados.setdefault(codigo, []).append(registro)

    # orden cronológico dentro de cada indicador; los sin fecha al final
    for diccionario in (productos, resultados):
        for lista in diccionario.values():
            lista.sort(key=lambda r: (r['fecha'] is None, r['fecha'] or ''))
    return productos, resultados, sin_asignar, ruta


if __name__ == '__main__':
    # Corrida de inspección: muestra el resumen completo, incluyendo las
    # solicitudes no aprobadas que NO pasan al tablero
    prods, res, sueltos, ruta = cargar_ajustes(solo_aprobados=False)
    print(f'Archivo: {ruta}')
    total_p = sum(len(v) for v in prods.values())
    total_r = sum(len(v) for v in res.values())
    print(f'Productos con ajustes: {len(prods)} ({total_p} ajustes)')
    print(f'Resultados con ajustes: {len(res)} ({total_r} ajustes)')
    no_aprobados = [(c, a) for d in (prods, res) for c, lista in d.items()
                    for a in lista if a['estado'] != 'Aprobado']
    print(f'No aprobados / denegados / en trámite (no pasan al tablero): {len(no_aprobados)}')
    for c, a in no_aprobados:
        print(f'  - [{c}] {a["estado"]}: {(a["ajuste"] or "")[:90]}')
    print(f'Filas sin código asignable: {len(sueltos)}')
    for s in sueltos:
        print(f'  - {s[:120]}')
