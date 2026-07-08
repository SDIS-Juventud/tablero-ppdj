# -*- coding: utf-8 -*-
"""
Convierte el formato crudo de la SDP (xlsb, hojas "Avance Cuantitativo" y
"Avance Cualitativo") al input procesado que consume el pipeline (xlsx con
hojas "Cuanti" y "Cuali", mismo esquema del insumo 2024).

Antes esta conversión la hacía el equipo a mano; este script la vuelve
repetible. Requiere pyxlsb (está en requirements.txt).

Entrada:  Insumos/2. Formato de Seguimiento a Resultados 3.1 PP Juventud Prelim SDP.xlsb
Salida:   Insumos/Datos tablero/Inputs/Seguimiento_Resultados_PPDJ_2025_excel.xlsx

Ejecución: python convertir_formato_resultados.py  (desde tablero-ppdj/)
"""

import os

import pandas as pd

RAIZ = os.path.dirname(os.path.abspath(__file__))
RUTA_XLSB = os.path.normpath(os.path.join(
    RAIZ, '..', 'Insumos', '2. Formato de Seguimiento a Resultados 3.1 PP Juventud Prelim SDP.xlsb'))
RUTA_SALIDA = os.path.normpath(os.path.join(
    RAIZ, '..', 'Insumos', 'Datos tablero', 'Inputs', 'Seguimiento_Resultados_PPDJ_2025_excel.xlsx'))

# Renombres del formato crudo al esquema del input procesado
RENOMBRES = {
    'Indicador No.': 'Resultado No.',
    'Indicador esperado': 'Resultado esperado',
    'Nombre del Indicador': 'Nombre indicador de Resultado',
    'Sector Responsable': 'Sector Líder',
    'Ponderación relativa (%)': 'Ponderación relativa del Resultado (%)',
    'Periodicidad de medición': 'Periodicidad de medición del indicador',
}
METADATOS = ['Resultado No.', 'Key', 'Estado del Indicador', 'Resultado esperado',
             'Nombre indicador de Resultado', 'Sector Líder',
             'Ponderación relativa del Resultado (%)', 'Valor Linea Base',
             'Tipo de anualización', 'Periodicidad de medición del indicador',
             'Fecha de Inicio', 'Fecha de Finalización',
             'Corte del último reporte', 'Año del último reporte']


def convertir_cuanti():
    df = pd.read_excel(RUTA_XLSB, sheet_name='Avance Cuantitativo',
                       engine='pyxlsb', skiprows=4)
    df = df.dropna(subset=['Indicador No.']).rename(columns=RENOMBRES)

    # trimestres 2018-2025 (2026 viene vacío y no entra)
    trimestres = [c for c in df.columns if isinstance(c, str)
                  and 'Trimestre' in c and not c.endswith('2026')]

    # metas anuales: columnas nombradas con el año, entre "Acumulado" y "Meta Final"
    metas = {}
    for anio in range(2020, 2026):
        for candidato in (str(anio), anio, float(anio)):
            if candidato in df.columns:
                metas[f'Meta_{anio}'] = df[candidato]
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
    cuali = cuali.dropna(subset=['Resultado No.'])
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
