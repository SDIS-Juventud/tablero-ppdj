# -*- coding: utf-8 -*-
"""
Convierte el formato crudo de la SDP para PRODUCTOS (xlsb, hojas
"Avance Cuantitativo" y "Avance Cualitativo") al input procesado que
consume el pipeline (xlsx con hojas "Cuanti" y "Cuali").

Es el gemelo de convertir_formato_resultados.py. Diferencias del formato
de productos frente al insumo viejo (corte Q1 2025):
- Trae la columna "Estado del Indicador" (99 vigentes de 123).
- 71 productos cambiaron de tipo de anualización.
- Trae datos hasta el 4o trimestre de 2025 y 1er trimestre de 2026.
- Las metas anuales llegan hasta 2026 (más la Meta Final, que no se usa).
- El bloque "Acumulado" del formato trae textos de ERROR para varios
  productos Creciente, por eso NO se copia: el acumulado lo calcula
  generar_productos.py según el tipo de anualización.

Entrada:  Insumos/1. Formato de Seguimiento a Productos 3.1 PP Juventud Prelim SDP.xlsb
Salida:   Insumos/Datos tablero/Inputs/Seguimiento_Productos_PPDJ_2026_excel.xlsx

Ejecución: python programas/convertir_formato_productos.py  (desde la raíz del repo)
"""

import os

import pandas as pd

# Raíz del repo: este script vive en programas/, un nivel abajo de la raíz
RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
RUTA_XLSB = os.path.normpath(os.path.join(
    RAIZ, '..', 'Insumos', '1. Formato de Seguimiento a Productos 3.1 PP Juventud Prelim SDP.xlsb'))
RUTA_SALIDA = os.path.normpath(os.path.join(
    RAIZ, '..', 'Insumos', 'Datos tablero', 'Inputs', 'Seguimiento_Productos_PPDJ_2026_excel.xlsx'))

# Renombres del formato crudo al esquema del input procesado
RENOMBRES = {
    'Indicador No.': 'Producto No.',
    'Indicador esperado': 'Producto esperado',
    'Nombre del Indicador': 'Nombre indicador de Producto',
    'Sector Responsable': 'Sector Líder',
    'Ponderación relativa (%)': 'Ponderación relativa del Producto (%)',
    'Periodicidad de medición': 'Periodicidad de medición del indicador',
}
METADATOS = ['Producto No.', 'Key', 'Estado del Indicador', 'Producto esperado',
             'Nombre indicador de Producto', 'Sector Líder',
             'Ponderación relativa del Producto (%)', 'Valor Linea Base',
             'Tipo de anualización', 'Periodicidad de medición del indicador',
             'Fecha de Inicio', 'Fecha de Finalización',
             'Corte del último reporte', 'Año del último reporte']


def convertir_cuanti():
    df = pd.read_excel(RUTA_XLSB, sheet_name='Avance Cuantitativo',
                       engine='pyxlsb', skiprows=4)
    df = df.dropna(subset=['Indicador No.']).rename(columns=RENOMBRES)

    # trimestres 2018-2026 (el 1er trimestre 2026 ya trae datos)
    trimestres = [c for c in df.columns if isinstance(c, str) and 'Trimestre' in c]

    # metas anuales: la PRIMERA columna nombrada con el año es la del bloque
    # "METAS PROGRAMADAS" (los bloques de porcentaje vienen después y pandas
    # los renombra 2020.1, 2020.2, ...). Mismo nombre del insumo viejo.
    metas = {}
    for anio in range(2018, 2027):
        for candidato in (str(anio), anio, float(anio)):
            if candidato in df.columns:
                metas[f'Meta_programada_{anio}'] = df[candidato]
                break
        else:
            raise ValueError(f'No se encontró la columna de meta anual {anio}')

    cuanti = df[METADATOS + trimestres].copy()
    for nombre, serie in metas.items():
        cuanti[nombre] = serie
    return cuanti


def convertir_cuali():
    # el encabezado real está repartido en 3 filas: año (2), trimestre (3)
    # y tipo Cualitativo/Enfoques (4); los datos empiezan en la fila 5
    crudo = pd.read_excel(RUTA_XLSB, sheet_name='Avance Cualitativo',
                          engine='pyxlsb', header=None)
    anios = crudo.iloc[2].ffill()
    trims = crudo.iloc[3].ffill()
    tipos = crudo.iloc[4]

    columnas = []
    for i in range(crudo.shape[1]):
        tipo = tipos.iloc[i]
        if tipo in ('Cualitativo', 'Enfoques') and pd.notna(anios.iloc[i]):
            anio = int(float(anios.iloc[i]))
            columnas.append((i, f'{anio}__{trims.iloc[i]}_{tipo}'))
        elif pd.notna(tipo) and tipo not in ('Cualitativo', 'Enfoques'):
            columnas.append((i, str(tipo).strip()))

    datos = crudo.iloc[5:].reset_index(drop=True)
    cuali = pd.DataFrame({nombre: datos[i] for i, nombre in columnas})
    cuali = cuali.rename(columns=RENOMBRES)
    cuali = cuali.dropna(subset=['Producto No.'])
    if 'Enfoque a Reportar' in cuali.columns:
        cuali = cuali.drop(columns=['Enfoque a Reportar'])
    return cuali


def main():
    cuanti = convertir_cuanti()
    cuali = convertir_cuali()
    with pd.ExcelWriter(RUTA_SALIDA, engine='openpyxl') as escritor:
        cuanti.to_excel(escritor, sheet_name='Cuanti', index=False)
        cuali.to_excel(escritor, sheet_name='Cuali', index=False)
    print(f'Salida: {RUTA_SALIDA}')
    print(f'Cuanti: {cuanti.shape[0]} filas x {cuanti.shape[1]} columnas')
    print(f'Cuali:  {cuali.shape[0]} filas x {cuali.shape[1]} columnas')
    if 'Estado del Indicador' in cuanti.columns:
        print('Estado del Indicador:', cuanti['Estado del Indicador'].value_counts(dropna=False).to_dict())


if __name__ == '__main__':
    main()
