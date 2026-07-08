# -*- coding: utf-8 -*-
"""
Genera data/resultados-cualitativo.json para el tablero HTML.

Reutiliza la lógica de generar_productos_cualitativo.py cambiando el
insumo y la columna de número. Mismo filtro corregido (excluye
"N.A."/"N/A"/"No aplica") y misma estructura por_llave ("R<numero>").

Ejecución: python generar_resultados_cualitativo.py  (desde tablero-ppdj/)
"""

import json
import os

from comun_pipeline import DIR_INPUTS
from generar_productos_cualitativo import construir

# Insumo del ciclo 2025 (generado por convertir_formato_resultados.py)
RUTA_INPUT = os.path.join(DIR_INPUTS, 'Seguimiento_Resultados_PPDJ_2025_excel.xlsx')
RUTA_JSON = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'data', 'resultados-cualitativo.json')


def main():
    from generar_productos import escribir_json_y_js
    datos = construir(ruta_input=RUTA_INPUT, columna_numero='Resultado No.', prefijo='R')
    escribir_json_y_js(datos, RUTA_JSON, 'DATOS_RESULTADOS_CUALI')
    kb = os.path.getsize(RUTA_JSON) / 1024
    total = sum(len(v) for v in datos['por_llave'].values())
    print(f'Generado: {RUTA_JSON} (+ .js) ({kb:.0f} KB)')
    print(f'Resultados con registros: {len(datos["por_llave"])} | Registros: {total}')


if __name__ == '__main__':
    main()
