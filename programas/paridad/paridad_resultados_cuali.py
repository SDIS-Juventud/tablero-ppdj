# -*- coding: utf-8 -*-
"""
Fase 4 de paridad: replica la sección cualitativa de RESULTADOS de
Procesamiento.R (líneas 214-243) tal como está hoy.

Comportamiento replicado a propósito:
- El filtro de la línea 229 del R tiene un bug de operador (usa | en vez
  de &), por lo que las exclusiones de "N.A."/"N/A"/"No aplica" NUNCA se
  aplican: en la práctica equivale a filtrar solo descripciones no nulas.
  Aquí se replica ese comportamiento (solo dropna). La corrección del
  filtro se hace en la fase 5, después de verificada la paridad.
- Se exporta el detalle largo; la concatenación anual del R es código
  muerto que nunca se exporta.

Referencia de paridad: Salidas/Seguimiento_cuali_resultados_PPDJ_anual_26032025.xlsx
(la que produce el R actual; la versión sin sufijo es de una corrida anterior).
"""

import os
import sys

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from comun_pipeline import DIR_INPUTS
from paridad_productos_cuali import formatear_como_r, formatear_key

RUTA_INPUT = os.path.join(DIR_INPUTS, 'Seguimiento_Resultados_PPDJ_2024_excel.xlsx')
DIR_SALIDA = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'salidas_paridad')
RUTA_SALIDA = os.path.join(DIR_SALIDA, 'Seguimiento_cuali_resultados_PPDJ_anual_PY.xlsx')


def construir_cuali_resultados():
    bd = pd.read_excel(RUTA_INPUT, sheet_name='Cuali', engine='openpyxl')

    # Llave compuesta (línea 216 del R). La Key numérica se normaliza al
    # formato "N." con el que se generaron las referencias.
    bd['llave'] = (bd['Resultado No.'].map(formatear_como_r) + '/'
                   + bd['Key'].map(formatear_key) + '/'
                   + bd['Resultado esperado'].map(formatear_como_r))

    columnas_datos = [c for c in bd.columns if isinstance(c, str) and '__Q' in c]
    largo = bd.melt(id_vars=['llave'], value_vars=columnas_datos,
                    var_name='Categoria', value_name='Descripción')

    partes = largo['Categoria'].str.split('__', n=1, expand=True)
    largo['Año'] = partes[0].astype(int)
    pregunta = partes[1].str.split('_', n=1, expand=True)
    largo['Trimestre'] = pregunta[0]
    largo['Tipo'] = pregunta[1]

    # Réplica del filtro con bug (equivale a solo !is.na, ver docstring)
    largo = largo.dropna(subset=['Descripción'])

    return largo[['llave', 'Año', 'Trimestre', 'Tipo', 'Descripción']]


def main():
    datos = construir_cuali_resultados()
    os.makedirs(DIR_SALIDA, exist_ok=True)
    datos.to_excel(RUTA_SALIDA, index=False, engine='openpyxl')
    print(f'Salida: {RUTA_SALIDA}')
    print(f'Filas: {len(datos)} | Columnas: {len(datos.columns)}')


if __name__ == '__main__':
    main()
