# -*- coding: utf-8 -*-
"""
Comparador de paridad: verifica celda a celda que las salidas generadas
por los scripts de Python coincidan con los archivos de referencia
producidos por el pipeline R original.

Para cada par (generado, referencia):
  1. Compara el conjunto y el orden de columnas.
  2. Alinea las filas por columna(s) clave (no depende del orden de filas).
  3. Compara numéricas con tolerancia (diferencias de punto flotante) y
     texto tras normalizar NBSP y _x000D_ en ambos lados (defectos del
     Excel fuente, no diferencias reales).
  4. Imprime PASS/FAIL por archivo y guarda un reporte de discrepancias
     en salidas_paridad/ cuando hay diferencias.

Ejecución (desde la carpeta paridad/):
  python comparar_excel.py            # compara todos los pares disponibles
  python comparar_excel.py productos  # compara solo los pares cuyo nombre contenga "productos"
"""

import os
import sys

import numpy as np
import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from comun_pipeline import DIR_SALIDAS_R, normalizar_texto

DIR_PARIDAD = os.path.dirname(os.path.abspath(__file__))
DIR_GENERADOS = os.path.join(DIR_PARIDAD, 'salidas_paridad')

PARES = [
    dict(nombre='productos_cuanti',
         generado='Seguimiento_Productos_PPDJ_anual_PY.xlsx',
         referencia='Seguimiento_Productos_PPDJ_anual_20082025.xlsx',
         clave=['Producto No.'], tolerancia=0.011),
    # drift_filas_conocidas: la referencia se generó el 26/06/2025 y el insumo
    # cambió después — 26 celdas de descripción fueron vaciadas. Se verificó
    # una a una que las 26 filas faltantes corresponden a celdas hoy vacías
    # (verificación del 2026-07-07, ver Notes/). No es un error del pipeline.
    dict(nombre='productos_cuali',
         generado='Seguimiento_cuali_Productos_PPDJ_anual_PY.xlsx',
         referencia='Seguimiento_cuali_Productos_PPDJ_anual_26062025.xlsx',
         clave=['llave', 'Año', 'Trimestre', 'Tipo'], tolerancia=None,
         drift_filas_conocidas=26),
    # La referencia de lógica es la versión _26032025: es la que el R actual
    # sí reproduce con el insumo actual. El archivo "limpio" sin sufijo viene
    # de una corrida ANTERIOR del pipeline (verificado el 2026-07-07: sus
    # Total_2024 en 0 y Meta_2024 viejas corresponden a un insumo previo).
    # El esquema de columnas del generado sigue siendo el del archivo limpio;
    # por eso se comparan solo las columnas comunes, alineadas por indicador.
    dict(nombre='resultados_cuanti',
         generado='Seguimiento_Resultados_PPDJ_anual_PY.xlsx',
         referencia='Seguimiento_Resultados_PPDJ_anual_26032025.xlsx',
         clave=['Nombre indicador de Resultado'], tolerancia=0.011,
         solo_columnas_comunes=True),
    # Igual que en resultados_cuanti: la referencia de lógica es la versión
    # _26032025 (la que produce el R actual); la sin sufijo es de una corrida
    # anterior con menos datos (558 vs 623 filas).
    # drift_filas_conocidas: verificado el 2026-07-07 que las 41 filas de la
    # referencia ausentes en el generado son celdas hoy vacías en el insumo.
    dict(nombre='resultados_cuali',
         generado='Seguimiento_cuali_resultados_PPDJ_anual_PY.xlsx',
         referencia='Seguimiento_cuali_resultados_PPDJ_anual_26032025.xlsx',
         clave=['llave', 'Año', 'Trimestre', 'Tipo'], tolerancia=None,
         drift_filas_conocidas=41),
]


def normalizar_para_comparar(valor):
    """Normalización estricta solo para la comparación: además de la limpieza
    estándar, colapsa TODO espacio en blanco (incluidos saltos de línea) a un
    solo espacio. Las diferencias de espacios/saltos entre la referencia y el
    insumo actual son ediciones cosméticas del archivo fuente, no contenido."""
    import re
    texto = normalizar_texto(valor)
    if texto is None:
        return None
    return re.sub(r'\s+', ' ', texto)


def preparar(df, clave):
    """Alinea un DataFrame por la(s) columna(s) clave y normaliza texto."""
    df = df.copy()
    for col in df.columns:
        if df[col].dtype == object:
            df[col] = df[col].map(normalizar_para_comparar)
    df = df.sort_values(by=clave, kind='mergesort').set_index(clave)
    return df


def comparar_par(par):
    ruta_gen = os.path.join(DIR_GENERADOS, par['generado'])
    ruta_ref = os.path.join(DIR_SALIDAS_R, par['referencia'])
    if not os.path.exists(ruta_gen):
        return None, f"pendiente (no existe {par['generado']})"
    df_gen = pd.read_excel(ruta_gen, engine='openpyxl')
    df_ref = pd.read_excel(ruta_ref, engine='openpyxl')

    problemas = []

    # 1. Columnas: mismo conjunto y mismo orden (salvo pares marcados para
    # comparar solo columnas comunes, cuando la referencia tiene otro esquema)
    if not par.get('solo_columnas_comunes') and list(df_gen.columns) != list(df_ref.columns):
        solo_gen = [c for c in df_gen.columns if c not in df_ref.columns]
        solo_ref = [c for c in df_ref.columns if c not in df_gen.columns]
        if solo_gen or solo_ref:
            problemas.append(f'columnas solo en generado: {solo_gen} | solo en referencia: {solo_ref}')
        else:
            problemas.append('mismas columnas pero en distinto orden')

    # 2. Conteo de filas
    if len(df_gen) != len(df_ref):
        problemas.append(f'filas: generado {len(df_gen)} vs referencia {len(df_ref)}')

    # 3. Comparación celda a celda sobre columnas comunes, alineadas por clave
    comunes = [c for c in df_ref.columns if c in df_gen.columns]
    gen = preparar(df_gen[comunes], par['clave'])
    ref = preparar(df_ref[comunes], par['clave'])

    indice_comun = gen.index.intersection(ref.index)
    faltan = ref.index.difference(gen.index)
    sobran = gen.index.difference(ref.index)
    drift = par.get('drift_filas_conocidas', 0)
    nota_drift = ''
    if len(faltan):
        if drift and len(faltan) <= drift and not len(sobran):
            # Drift documentado del insumo: no cuenta como fallo de paridad
            nota_drift = f' [drift conocido: {len(faltan)} filas de la referencia ya no existen en el insumo]'
            problemas = [p for p in problemas if not p.startswith('filas:')]
        else:
            problemas.append(f'{len(faltan)} filas de la referencia no están en el generado (ej: {list(faltan[:3])})')
    if len(sobran):
        problemas.append(f'{len(sobran)} filas del generado no están en la referencia (ej: {list(sobran[:3])})')

    gen, ref = gen.loc[indice_comun], ref.loc[indice_comun]
    discrepancias = []

    def celdas_iguales(vg, vr, tol):
        """Compara un par de celdas: numérico con tolerancia si ambos lados
        son convertibles a número (las columnas mixtas texto/número llegan
        como object y el round-trip de Excel mete ruido de punto flotante);
        si no, como texto."""
        nulo_g = pd.isna(vg) if not isinstance(vg, (list, tuple)) else False
        nulo_r = pd.isna(vr) if not isinstance(vr, (list, tuple)) else False
        if nulo_g or nulo_r:
            return bool(nulo_g and nulo_r)
        try:
            return bool(np.isclose(float(str(vg).replace(',', '.')),
                                   float(str(vr).replace(',', '.')), atol=tol))
        except (TypeError, ValueError):
            return str(vg) == str(vr)

    for col in gen.columns:
        g, r = gen[col], ref[col]
        tol = par['tolerancia'] or 1e-9
        if pd.api.types.is_numeric_dtype(r) and pd.api.types.is_numeric_dtype(g):
            iguales = np.isclose(g.astype(float), r.astype(float), atol=tol, equal_nan=True)
        elif pd.api.types.is_datetime64_any_dtype(r) or pd.api.types.is_datetime64_any_dtype(g):
            iguales = (pd.to_datetime(g, errors='coerce') == pd.to_datetime(r, errors='coerce')) \
                      | (g.isna() & r.isna())
        else:
            iguales = np.array([celdas_iguales(vg, vr, tol) for vg, vr in zip(g, r)])
        for idx in g.index[~iguales]:
            discrepancias.append({'fila': idx, 'columna': col,
                                  'generado': g.loc[idx], 'referencia': r.loc[idx]})

    if discrepancias:
        problemas.append(f'{len(discrepancias)} celdas distintas')
        reporte = pd.DataFrame(discrepancias)
        ruta_reporte = os.path.join(DIR_GENERADOS, f"reporte_{par['nombre']}.xlsx")
        reporte.to_excel(ruta_reporte, index=False, engine='openpyxl')
        problemas.append(f'reporte: {ruta_reporte}')

    total_celdas = gen.size or 1
    pct = 100 * (1 - len(discrepancias) / total_celdas)
    return (not problemas), (f'{pct:.2f}% de celdas coinciden'
                             + ('; ' + ' | '.join(problemas) if problemas else '')
                             + nota_drift)


def main():
    filtro = sys.argv[1].lower() if len(sys.argv) > 1 else ''
    for par in PARES:
        if filtro and filtro not in par['nombre']:
            continue
        ok, detalle = comparar_par(par)
        estado = 'PASS' if ok else ('----' if ok is None else 'FAIL')
        print(f"[{estado}] {par['nombre']}: {detalle}")
        if discrepante := (ok is False):
            _ = discrepante  # el detalle ya incluye la ruta del reporte


if __name__ == '__main__':
    main()
