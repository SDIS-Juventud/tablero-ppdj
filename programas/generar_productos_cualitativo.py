# -*- coding: utf-8 -*-
"""
Genera data/productos-cualitativo.json para el tablero HTML.

Estructura: {"por_llave": {"P<numero>": [{anio, trimestre, tipo, descripcion}]}}
agrupado por producto para que la página de análisis cualitativo haga el
lookup directo al llegar por querystring (?llave=P12).

Aplica el filtro CORREGIDO (el R tenía un bug de operador que nunca
excluía los literales "N.A."/"N/A"/"No aplica"): aquí sí se excluyen.

Este JSON pesa varios MB, así que no se embebe en el HTML: la página lo
carga desde el .js compañero (<script src>), que funciona también al
abrir el archivo con doble clic (file://), sin necesidad de servidor.

Ejecución: python programas/generar_productos_cualitativo.py  (desde la raíz del repo)
"""

import json
import os
from datetime import date

import pandas as pd

from comun_pipeline import DIR_INPUTS, normalizar_texto

# Desde 2026-07-08: formato SDP nuevo convertido (solo vigentes, textos
# hasta el corte 2026). Generado por convertir_formato_productos.py.
RUTA_INPUT = os.path.join(DIR_INPUTS, 'Seguimiento_Productos_PPDJ_2026_excel.xlsx')
RUTA_JSON = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'data', 'productos-cualitativo.json')

LITERALES_SIN_CONTENIDO = {'N.A.', 'N/A', 'No aplica', 'N.A', 'NA'}


def construir(ruta_input=RUTA_INPUT, columna_numero='Producto No.', prefijo='P',
              anio_maximo=2025):
    # anio_maximo: el tablero muestra hasta 2025 cerrado (decisión de
    # Carolina 2026-07-14); los textos del corte 2026 no entran.
    bd = pd.read_excel(ruta_input, sheet_name='Cuali', engine='openpyxl')
    # Si el insumo trae estado (formato 2025 de resultados), solo los vigentes
    if 'Estado del Indicador' in bd.columns:
        bd = bd[bd['Estado del Indicador'] == 'Vigente']
    columnas_datos = [c for c in bd.columns if isinstance(c, str) and '__Q' in c]

    por_llave = {}
    for _, fila in bd.iterrows():
        if pd.isna(fila[columna_numero]):
            continue
        llave = f'{prefijo}{int(fila[columna_numero])}'
        registros = []
        for col in columnas_datos:
            anio_txt, pregunta = col.split('__', 1)
            if int(anio_txt) > anio_maximo:
                continue
            descripcion = normalizar_texto(fila[col])
            if not descripcion or descripcion.strip() in LITERALES_SIN_CONTENIDO:
                continue
            trimestre, tipo = pregunta.split('_', 1)
            registros.append({'anio': int(anio_txt), 'trimestre': trimestre,
                              'tipo': tipo, 'descripcion': descripcion})
        if registros:
            registros.sort(key=lambda r: (r['anio'], r['trimestre'], r['tipo']))
            # Si el insumo trae filas duplicadas del mismo número, se acumulan
            por_llave.setdefault(llave, []).extend(registros)

    return {'generado': date.today().isoformat(), 'por_llave': por_llave}


def main():
    from generar_productos import escribir_json_y_js
    datos = construir()
    escribir_json_y_js(datos, RUTA_JSON, 'DATOS_PRODUCTOS_CUALI')
    kb = os.path.getsize(RUTA_JSON) / 1024
    total = sum(len(v) for v in datos['por_llave'].values())
    print(f'Generado: {RUTA_JSON} (+ .js) ({kb:.0f} KB)')
    print(f'Productos con registros: {len(datos["por_llave"])} | Registros: {total}')


if __name__ == '__main__':
    main()
