# -*- coding: utf-8 -*-
"""
Módulo común del pipeline de Productos y Resultados PPDJ.

Contiene las constantes y funciones compartidas por los scripts de paridad
(carpeta paridad/) y los generadores de JSON del tablero (generar_*.py).

Este módulo replica la lógica del pipeline original en R (Procesamiento.R).
La función anualizar() reproduce por defecto el comportamiento ACTUAL del R,
incluido el bug conocido de "Creciente" (se trata igual que "Constante",
sin restar el Q4 del año anterior). La corrección se activa solo con
corregir_creciente=True, para la fase posterior a la verificación de paridad.
"""

import math
import os
import re
from datetime import datetime, timedelta

# Ruta base de los insumos: la carpeta Insumos/ vive junto a tablero-ppdj/,
# fuera del repo (no se versiona). Se resuelve relativa a este archivo.
RAIZ_PROYECTO = os.path.dirname(os.path.abspath(__file__))
DIR_INSUMOS = os.path.normpath(os.path.join(RAIZ_PROYECTO, '..', 'Insumos', 'Datos tablero'))
DIR_INPUTS = os.path.join(DIR_INSUMOS, 'Inputs')
DIR_SALIDAS_R = os.path.join(DIR_INSUMOS, 'Salidas')

# Las 7 dimensiones de la PPDJ. La Key ("1." a "7.") es la misma que usa
# Objetivos.xlsx. Los colores son los mismos del tablero-cij para mantener
# consistencia visual entre los dos tableros hermanos.
DIMENSIONES = {
    '1.': {'nombre': 'Ser Joven', 'color': '#f4676e'},
    '2.': {'nombre': 'Educación', 'color': '#1eaf76'},
    '3.': {'nombre': 'Inclusión Productiva', 'color': '#663a93'},
    '4.': {'nombre': 'Salud Integral y Autocuidado', 'color': '#1e9da3'},
    '5.': {'nombre': 'Cultura, Recreación y Deporte', 'color': '#f58b53'},
    '6.': {'nombre': 'Paz, Convivencia y Justicia', 'color': '#2fa4d4'},
    '7.': {'nombre': 'Hábitat', 'color': '#1e7895'},
}

# Mapeo del prefijo de columna trimestral ("1er Trimestre 2019") a número.
TRIMESTRE_A_NUMERO = {'1er': 1, '2o': 2, '3er': 3, '4o': 4}


def normalizar_texto(valor):
    """Limpia los defectos de texto que traen los Excel fuente:
    NBSP (\\xa0), saltos de línea legacy de Excel (_x000D_) y espacios
    repetidos. Devuelve None si el valor es nulo."""
    if valor is None or (isinstance(valor, float) and math.isnan(valor)):
        return None
    texto = str(valor)
    texto = texto.replace('\xa0', ' ')
    # Canonicaliza saltos de línea: el escape legacy de Excel (_x000D_) y el
    # retorno de carro (\r) equivalen a un salto; "\r\n" debe quedar como un
    # solo "\n", no como dos.
    texto = texto.replace('_x000D_', '\r')
    texto = texto.replace('\r\n', '\n').replace('\r', '\n')
    # Más de dos saltos seguidos no aportan contenido: se colapsan a párrafo
    texto = re.sub(r'\n{3,}', '\n\n', texto)
    texto = re.sub(r'[ \t]+', ' ', texto)
    return texto.strip()


def parsear_fecha_mixta(valor):
    """Replica el parseo de fechas de la sección Resultados del R:
    si el valor contiene "/", se interpreta como día/mes/año (dmy);
    si es numérico, se interpreta como serial de Excel (origen 1899-12-30).
    Los datos reales mezclan ambos formatos en la misma columna."""
    if valor is None or (isinstance(valor, float) and math.isnan(valor)):
        return None
    if isinstance(valor, datetime):
        return valor
    texto = str(valor).strip()
    if '/' in texto:
        return datetime.strptime(texto, '%d/%m/%Y')
    # Serial de Excel: días desde 1899-12-30 (mismo origen que usa R)
    return datetime(1899, 12, 30) + timedelta(days=float(texto))


def anualizar(valores_trimestre, tipo, corregir_creciente=False, q4_anio_anterior=None):
    """Anualiza los 4 valores trimestrales de un año según el tipo de
    anualización, replicando el case_when del R.

    valores_trimestre: dict {1: v1, 2: v2, 3: v3, 4: v4} (valores pueden ser None/NaN)
    tipo: "Constante" | "Creciente" | "Decreciente" | "Suma"

    Comportamiento por defecto (paridad con el R actual):
    - Constante, Creciente y Decreciente toman solo el Q4. El R usa
      sum(Valor[Trimestre_Num == 4], na.rm = TRUE), que devuelve 0 (no NA)
      cuando el Q4 es NA — ese detalle se replica aquí tal cual.
    - Suma suma los 4 trimestres ignorando NA (también 0 si todos son NA).
    - Cualquier otro tipo devuelve None (NA_real_ en R).

    Con corregir_creciente=True (fase de corrección, según la regla de
    Reglas de validación.xlsx): Creciente = Q4 del año menos Q4 del año
    anterior (q4_anio_anterior; si no hay año anterior se usa 0).
    """
    validos = {q: v for q, v in valores_trimestre.items()
               if v is not None and not (isinstance(v, float) and math.isnan(v))}

    if tipo in ('Constante', 'Creciente', 'Decreciente'):
        valor_q4 = validos.get(4, 0.0)  # replica sum(..., na.rm=TRUE) sobre vacío = 0
        if corregir_creciente and tipo == 'Creciente':
            anterior = q4_anio_anterior if q4_anio_anterior is not None else 0.0
            return valor_q4 - anterior
        return valor_q4
    if tipo == 'Suma':
        return sum(validos.values()) if validos else 0.0
    return None


def anualizar_tablero(valores_trimestre, tipo, q4_anio_anterior=None):
    """Anualización CORREGIDA para el tablero (fase posterior a la paridad).

    Diferencias con anualizar() (la réplica fiel del R):
    - "Creciente" aplica la regla real de Reglas de validación.xlsx:
      Q4 del año menos Q4 del año anterior (si no hay año anterior, solo Q4).
    - Un año sin ningún dato trimestral devuelve None (sin dato), no 0:
      el 0 del R era un artefacto de sum(..., na.rm=TRUE) sobre vacío.
    - "Decreciente" se mantiene como Q4 (igual que Constante) porque no hay
      regla documentada — pendiente de definición del equipo. Hay 10 casos
      reales en Resultados.
    """
    validos = {q: v for q, v in valores_trimestre.items()
               if v is not None and not (isinstance(v, float) and math.isnan(v))}
    if not validos:
        return None
    if tipo == 'Suma':
        return sum(validos.values())
    if tipo in ('Constante', 'Decreciente'):
        return validos.get(4)
    if tipo == 'Creciente':
        valor_q4 = validos.get(4)
        if valor_q4 is None:
            return None
        if q4_anio_anterior is None or (isinstance(q4_anio_anterior, float) and math.isnan(q4_anio_anterior)):
            return valor_q4
        return valor_q4 - q4_anio_anterior
    return None


def clasificar_indicador_tipo(nombre_indicador):
    """Deriva el campo "Indicador Tipo" de la ficha (Valor / Porcentaje /
    Tasa) a partir del nombre del indicador. El campo original vive solo en
    el modelo binario del Power BI (no está en ningún Excel), así que se
    reconstruye con esta regla transparente basada en cómo empieza el nombre.
    Verificado contra los 106 productos y 36 resultados del insumo actual."""
    if not nombre_indicador:
        return None
    inicio = str(nombre_indicador).strip().lower()
    if inicio.startswith(('porcentaje', '%', 'proporción', 'proporcion', 'deserción', 'desercion')):
        return 'Porcentaje'
    if inicio.startswith('tasa'):
        return 'Tasa'
    return 'Valor'


def cargar_objetivos(pandas_modulo):
    """Lee Objetivos.xlsx (columnas Key y Objetivo, Keys "1." a "7.")
    y lo devuelve como DataFrame para el join por Key."""
    ruta = os.path.join(DIR_INSUMOS, 'Objetivos.xlsx')
    return pandas_modulo.read_excel(ruta, engine='openpyxl')


def extraer_anio_y_trimestre(nombre_columna):
    """De un nombre de columna trimestral tipo "1er Trimestre 2019"
    extrae (2019, 1). Devuelve (None, None) si el nombre no corresponde
    a una columna trimestral."""
    coincidencia = re.match(r'^(1er|2o|3er|4o)\s+Trimestre\s+(\d{4})$', str(nombre_columna).strip())
    if not coincidencia:
        return None, None
    return int(coincidencia.group(2)), TRIMESTRE_A_NUMERO[coincidencia.group(1)]
