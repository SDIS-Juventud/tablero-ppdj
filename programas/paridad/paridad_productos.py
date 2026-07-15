# -*- coding: utf-8 -*-
"""
Fase 1 de paridad: replica en Python la sección PRODUCTOS (cuantitativo)
de Procesamiento.R, tal como está hoy, para verificar que el pipeline
portado produce exactamente la misma salida que el R.

IMPORTANTE: este script replica a propósito el comportamiento ACTUAL del R,
incluido el bug de "Creciente" (toma solo el Q4, sin restar el Q4 del año
anterior). La corrección de esa regla se hace en la fase 5, después de
verificada la paridad.

Referencia de paridad:
  Salidas/Seguimiento_Productos_PPDJ_anual_20082025.xlsx (106 filas, 32 columnas)

Ejecución (desde la carpeta paridad/):
  python paridad_productos.py
"""

import os
import sys

import pandas as pd

# Permite importar comun_pipeline desde la carpeta padre
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from comun_pipeline import DIR_INPUTS, anualizar, cargar_objetivos, extraer_anio_y_trimestre

RUTA_INPUT = os.path.join(DIR_INPUTS, 'Seguimiento_Productos_PPDJ_Q1_2025_excel.xlsx')
DIR_SALIDA = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'salidas_paridad')
RUTA_SALIDA = os.path.join(DIR_SALIDA, 'Seguimiento_Productos_PPDJ_anual_PY.xlsx')

# Años del loop de Diff, hardcodeados igual que en el R (línea 89)
ANIOS = list(range(2019, 2026))


def construir_base_productos():
    """Replica la sección Productos de Procesamiento.R paso a paso."""
    # Lectura: hoja Cuanti, saltando las 2 filas de título fusionado (skip=2 en R)
    bd = pd.read_excel(RUTA_INPUT, sheet_name='Cuanti', skiprows=2, engine='openpyxl')

    # Llave compuesta, igual que la línea 27 del R
    bd['llave'] = (bd['Producto No.'].astype(str) + '/' + bd['Key'].astype(str)
                   + '/' + bd['Producto esperado'].astype(str))

    # Pivot largo de las 28 columnas trimestrales (2019-2025)
    columnas_trimestrales = [c for c in bd.columns if extraer_anio_y_trimestre(c)[0] is not None]
    largo = bd.melt(id_vars=['llave', 'Tipo de anualización'],
                    value_vars=columnas_trimestrales,
                    var_name='Columna', value_name='Valor')
    largo[['Anio', 'Trimestre_Num']] = largo['Columna'].apply(
        lambda c: pd.Series(extraer_anio_y_trimestre(c)))

    # Anualización por llave y año (replica el case_when del R, bug incluido)
    def total_del_grupo(grupo):
        tipo = grupo['Tipo de anualización'].iloc[0]
        valores = dict(zip(grupo['Trimestre_Num'], grupo['Valor']))
        return anualizar(valores, tipo)

    totales = (largo.groupby(['llave', 'Anio'], sort=False)
               .apply(total_del_grupo, include_groups=False)
               .reset_index(name='Total'))

    # Pivot ancho a Total_2019..Total_2025 (values_fn=sum del R: aquí validamos
    # explícitamente que no haya llaves duplicadas en vez de sumar en silencio)
    duplicados = totales.duplicated(subset=['llave', 'Anio']).sum()
    if duplicados:
        raise ValueError(f'Hay {duplicados} combinaciones llave+año duplicadas: revisar insumo.')
    ancho = totales.pivot(index='llave', columns='Anio', values='Total')
    ancho.columns = [f'Total_{a}' for a in ancho.columns]
    ancho = ancho.reset_index()

    # Joins: totales por llave, objetivos por Key (líneas 57-59 del R)
    bd = bd.merge(ancho, on='llave', how='left')
    objetivos = cargar_objetivos(pd)
    bd = bd.merge(objetivos, on='Key', how='left')

    # Renombres: fechas (línea 65) y Meta_programada_YYYY -> "Meta YYYY" (líneas 82-86)
    bd = bd.rename(columns={'Fecha de Inicio': 'fecha_inicio', 'Fecha de Finalización': 'fecha_fin'})
    bd = bd.rename(columns={f'Meta_programada_{a}': f'Meta {a}' for a in ANIOS})

    # Selección y orden final por nombre (equivale al select posicional de la línea 62)
    columnas_finales = (
        ['Key', 'Objetivo', 'Producto No.', 'Producto esperado', 'Nombre indicador de Producto',
         'Sector Líder', 'Ponderación relativa del Producto (%)', 'Valor Linea Base',
         'Tipo de anualización', 'fecha_inicio', 'fecha_fin']
        + [f'Total_{a}' for a in ANIOS]
        + [f'Meta {a}' for a in ANIOS]
    )
    bd = bd[columnas_finales]

    # Redondeo a 2 decimales de todas las columnas numéricas (línea 69 del R).
    # Nota: el redondeo a 1 decimal de las líneas 106-111 del R es código
    # muerto (solo se imprime, nunca se exporta) — no se replica.
    numericas = bd.select_dtypes(include='number').columns
    bd[numericas] = bd[numericas].round(2)

    # Diff_YYYY = Total_YYYY - Meta YYYY (loop de las líneas 91-104, después del redondeo)
    for a in ANIOS:
        bd[f'Diff_{a}'] = bd[f'Total_{a}'] - bd[f'Meta {a}']

    return bd


def main():
    bd = construir_base_productos()
    os.makedirs(DIR_SALIDA, exist_ok=True)
    bd.to_excel(RUTA_SALIDA, index=False, engine='openpyxl')
    print(f'Salida: {RUTA_SALIDA}')
    print(f'Filas: {len(bd)} | Columnas: {len(bd.columns)}')


if __name__ == '__main__':
    main()
