# -*- coding: utf-8 -*-
"""
Fase 2 de paridad: replica la sección cualitativa de PRODUCTOS de
Procesamiento.R (líneas 116-145) tal como está hoy.

Comportamiento replicado a propósito:
- Se exporta el detalle largo (datos_largo): una fila por llave x trimestre
  x tipo. El R calcula además una concatenación anual (datos_anual) pero
  NUNCA la exporta — ese código muerto no se replica.
- El único filtro real es Descripción no nula (!is.na en R).

Referencia de paridad:
  Salidas/Seguimiento_cuali_Productos_PPDJ_anual_26062025.xlsx (2740 filas, 5 columnas)
"""

import math
import os
import sys

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from comun_pipeline import DIR_INPUTS

RUTA_INPUT = os.path.join(DIR_INPUTS, 'Seguimiento_Productos_PPDJ_Q1_2025_excel.xlsx')
DIR_SALIDA = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'salidas_paridad')
RUTA_SALIDA = os.path.join(DIR_SALIDA, 'Seguimiento_cuali_Productos_PPDJ_anual_PY.xlsx')


def formatear_como_r(valor):
    """Convierte un valor a texto igual que paste() de R: los números
    enteros van sin decimales (1, no 1.0) y los nulos quedan como 'NA'."""
    if valor is None or (isinstance(valor, float) and math.isnan(valor)):
        return 'NA'
    if isinstance(valor, float) and valor.is_integer():
        return str(int(valor))
    return str(valor)


def formatear_key(valor):
    """La Key de las dimensiones es "1." a "7.". En la hoja Cuanti viene como
    texto, pero en la hoja Cuali del insumo actual viene como número (1.0).
    Cuando el R generó la referencia, la columna era texto — el insumo cambió
    después. Se normaliza al formato "N." para reproducir la llave original."""
    if isinstance(valor, float) and valor.is_integer():
        return f'{int(valor)}.'
    if isinstance(valor, int):
        return f'{valor}.'
    return formatear_como_r(valor)


def construir_cuali_productos():
    bd = pd.read_excel(RUTA_INPUT, sheet_name='Cuali', engine='openpyxl')

    # Llave compuesta idéntica a la del R (línea 118)
    bd['llave'] = (bd['Producto No.'].map(formatear_como_r) + '/'
                   + bd['Key'].map(formatear_key) + '/'
                   + bd['Producto esperado'].map(formatear_como_r))

    # Columnas de datos trimestre x tipo: formato AAAA__Q#_Cualitativo/Enfoques
    columnas_datos = [c for c in bd.columns
                      if isinstance(c, str) and '__Q' in c]

    largo = bd.melt(id_vars=['llave'], value_vars=columnas_datos,
                    var_name='Categoria', value_name='Descripción')

    # Separación en dos pasos, igual que el R: "__" -> Año / Pregunta,
    # luego "_" -> Trimestre / Tipo
    partes = largo['Categoria'].str.split('__', n=1, expand=True)
    largo['Año'] = partes[0].astype(int)  # la referencia trae el año numérico
    pregunta = partes[1].str.split('_', n=1, expand=True)
    largo['Trimestre'] = pregunta[0]
    largo['Tipo'] = pregunta[1]

    # Filtro: solo descripciones no nulas (línea 131 del R)
    largo = largo.dropna(subset=['Descripción'])

    return largo[['llave', 'Año', 'Trimestre', 'Tipo', 'Descripción']]


def main():
    datos = construir_cuali_productos()
    os.makedirs(DIR_SALIDA, exist_ok=True)
    datos.to_excel(RUTA_SALIDA, index=False, engine='openpyxl')
    print(f'Salida: {RUTA_SALIDA}')
    print(f'Filas: {len(datos)} | Columnas: {len(datos.columns)}')


if __name__ == '__main__':
    main()
