# Tablero de Productos y Resultados PPDJ — instrucciones de actualización

Este tablero muestra el seguimiento a los Productos y Resultados del Plan de
Acción Indicativo de la Política Pública Distrital de Juventud 2019-2030.
Son páginas HTML estáticas que leen datos generados por scripts de Python.

## Qué archivos usa

| Insumo | Dónde va |
|---|---|
| Seguimiento a Productos (Excel, hojas Cuanti y Cuali) | `..\Insumos\Datos tablero\Inputs\` |
| Seguimiento a Resultados (Excel, hojas Cuanti y Cuali) | `..\Insumos\Datos tablero\Inputs\` |
| Objetivos por dimensión | `..\Insumos\Datos tablero\Objetivos.xlsx` |

La carpeta `Insumos\` está fuera del repositorio a propósito: contiene los
formatos de trabajo del equipo y no se publica.

## Cómo actualizar los datos

1. Guardar el nuevo Excel de seguimiento en `..\Insumos\Datos tablero\Inputs\`
   **sin borrar el anterior**.
2. Si cambió el nombre del archivo, actualizar la constante `RUTA_INPUT` en el
   script correspondiente (`generar_productos.py`, `generar_resultados.py`,
   `generar_productos_cualitativo.py`, `generar_resultados_cualitativo.py`).
3. Correr los cuatro generadores desde esta carpeta:

```
python generar_productos.py
python generar_resultados.py
python generar_productos_cualitativo.py
python generar_resultados_cualitativo.py
```

4. Abrir `index.html` con doble clic y revisar que los datos nuevos aparezcan.
5. Subir los cambios (los `data/*.json` y `data/*.js` regenerados) al repositorio.

Requisitos de Python: `pip install -r requirements.txt` (pandas y openpyxl).

## Reglas de cálculo (anualización)

El valor anual de cada indicador se calcula desde los reportes trimestrales
según su tipo de anualización (ver `indicadores.html`):

- **Suma**: suma de los cuatro trimestres.
- **Constante**: dato del Q4.
- **Creciente**: dato del Q4 menos el Q4 del año anterior. *(El pipeline viejo
  en R no restaba el año anterior; aquí ya está corregido.)*
- **Decreciente**: dato del Q4, provisional, mientras el equipo define la regla.

Un año sin ningún reporte trimestral queda sin dato (no en cero).

## Carpeta paridad/

Contiene los scripts que replican el pipeline original de R **tal cual, con
sus bugs**, y el comparador que verifica que Python reproduce las salidas ya
publicadas. Solo se usan para auditoría; el tablero se alimenta de los
`generar_*.py`. Ver `python paridad\comparar_excel.py`.

## Estructura de páginas

| Página | Contenido |
|---|---|
| `index.html` | Portada: 7 dimensiones con conteo de productos y resultados |
| `productos.html` / `resultados.html` | Detalle: filtros, tabla, ficha y gráfica Valor vs Meta |
| `productos-cualitativo.html` / `resultados-cualitativo.html` | Bitácora cualitativa trimestral del ítem seleccionado |
| `indicadores.html` | Explicación de los tipos de anualización |

Los datos llegan a las páginas por `<script src="data/*.js">`, que funciona
al abrir los archivos directamente (file://), sin servidor.
