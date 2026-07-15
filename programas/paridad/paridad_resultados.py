# -*- coding: utf-8 -*-
"""
Fase 3 de paridad: replica la sección RESULTADOS (cuantitativo) de
Procesamiento.R (líneas 147-211).

Diferencias con la sección Productos, tal como están en el R:
- El parseo de fechas SÍ está activo: mezcla de "dd/mm/yyyy" (texto) y
  seriales de Excel en la misma columna; el resultado son strings
  "YYYY-MM-DD" (as.character(as.Date(...)) en R).
- La anualización agrega "Decreciente" al grupo que toma solo Q4.
- No hay redondeo.
- Las metas van con guion bajo (Meta_2020) y Meta_2025 se descarta.

Target de paridad (decisión de Carolina): el archivo "limpio" sin sufijo
Salidas/Seguimiento_Resultados_PPDJ_anual.xlsx (36 filas, 28 columnas,
Diff solo 2020-2023). OJO: el Procesamiento.R actual NO reproduce ese
archivo tal cual (produce la versión _26032025 con columnas desordenadas);
aquí se reconstruye el esquema limpio por nombre de columna.

PREGUNTAS ABIERTAS para quien mantiene el pipeline (ver Notes/):
- ¿Debería existir Diff_2024? Total_2024 y Meta_2024 existen, pero el
  archivo limpio termina en Diff_2023.
- ¿"Resultado No." se omitió a propósito de la salida limpia?
"""

import os
import sys

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from comun_pipeline import (DIR_INPUTS, anualizar, cargar_objetivos,
                            extraer_anio_y_trimestre, parsear_fecha_mixta)

RUTA_INPUT = os.path.join(DIR_INPUTS, 'Seguimiento_Resultados_PPDJ_2024_excel.xlsx')
DIR_SALIDA = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'salidas_paridad')
RUTA_SALIDA = os.path.join(DIR_SALIDA, 'Seguimiento_Resultados_PPDJ_anual_PY.xlsx')

ANIOS_TOTAL = list(range(2018, 2025))   # trimestres del insumo: 2018-2024
ANIOS_META = list(range(2020, 2025))    # Meta_2020..Meta_2024 (Meta_2025 se descarta)
ANIOS_DIFF = list(range(2020, 2024))    # el archivo limpio solo trae Diff_2020..Diff_2023


def construir_base_resultados():
    bd = pd.read_excel(RUTA_INPUT, sheet_name='Cuanti', engine='openpyxl')

    # Fechas mixtas -> string "YYYY-MM-DD" (líneas 152-156 del R)
    for original, nueva in [('Fecha de Inicio', 'fecha_inicio'), ('Fecha de Finalización', 'fecha_fin')]:
        bd[nueva] = bd[original].map(
            lambda v: f.strftime('%Y-%m-%d') if (f := parsear_fecha_mixta(v)) else None)

    bd['llave'] = (bd['Resultado No.'].astype(str) + '/' + bd['Key'].astype(str)
                   + '/' + bd['Resultado esperado'].astype(str))

    # Pivot largo de las 28 columnas trimestrales (2018-2024)
    columnas_trimestrales = [c for c in bd.columns if extraer_anio_y_trimestre(c)[0] is not None]
    largo = bd.melt(id_vars=['llave', 'Tipo de anualización'],
                    value_vars=columnas_trimestrales,
                    var_name='Columna', value_name='Valor')
    largo[['Anio', 'Trimestre_Num']] = largo['Columna'].apply(
        lambda c: pd.Series(extraer_anio_y_trimestre(c)))

    def total_del_grupo(grupo):
        tipo = grupo['Tipo de anualización'].iloc[0]
        valores = dict(zip(grupo['Trimestre_Num'], grupo['Valor']))
        return anualizar(valores, tipo)

    totales = (largo.groupby(['llave', 'Anio'], sort=False)
               .apply(total_del_grupo, include_groups=False)
               .reset_index(name='Total'))
    duplicados = totales.duplicated(subset=['llave', 'Anio']).sum()
    if duplicados:
        raise ValueError(f'Hay {duplicados} combinaciones llave+año duplicadas: revisar insumo.')
    ancho = totales.pivot(index='llave', columns='Anio', values='Total')
    ancho.columns = [f'Total_{a}' for a in ancho.columns]
    bd = bd.merge(ancho.reset_index(), on='llave', how='left')

    # Join con Objetivos: ambas Keys son numéricas (1.0 a 7.0)
    objetivos = cargar_objetivos(pd)
    bd = bd.merge(objetivos, on='Key', how='left')

    # Esquema limpio del archivo de referencia, construido por nombre
    columnas_finales = (
        ['Key', 'Objetivo', 'Resultado esperado', 'Nombre indicador de Resultado',
         'Sector Líder', 'Ponderación relativa del Resultado (%)', 'Valor Linea Base',
         'Tipo de anualización', 'Periodicidad de medición del indicador',
         'fecha_inicio', 'fecha_fin', 'Año del último reporte']
        + [f'Total_{a}' for a in ANIOS_TOTAL]
        + [f'Meta_{a}' for a in ANIOS_META]
    )
    bd = bd[columnas_finales]

    # Diff solo 2020-2023, replicando el artefacto del archivo limpio
    for a in ANIOS_DIFF:
        bd[f'Diff_{a}'] = bd[f'Total_{a}'] - bd[f'Meta_{a}']

    return bd


def main():
    bd = construir_base_resultados()
    os.makedirs(DIR_SALIDA, exist_ok=True)
    bd.to_excel(RUTA_SALIDA, index=False, engine='openpyxl')
    print(f'Salida: {RUTA_SALIDA}')
    print(f'Filas: {len(bd)} | Columnas: {len(bd.columns)}')


if __name__ == '__main__':
    main()
