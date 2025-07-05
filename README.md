# Conversión de ecuaciones a Word

Este repositorio contiene un pequeño script para generar un archivo `.docx` a partir de un cuaderno Jupyter u otro archivo compatible con pandoc, manteniendo las ecuaciones de LaTeX como ecuaciones editables en Microsoft Word.

## Requisitos
- Python 3
- [pandoc](https://pandoc.org/) disponible en el sistema
- El paquete `pypandoc`

Puedes instalar `pypandoc` ejecutando:

```bash
pip install pypandoc
```

## Uso

Para convertir `Laura.ipynb` a `output.docx`:

```bash
python convert_to_docx.py
```

También puedes indicar archivos de entrada y salida:

```bash
python convert_to_docx.py entrada.ipynb salida.docx
```

El proceso conserva las ecuaciones escritas en LaTeX en el cuaderno para que aparezcan como objetos de ecuación en Word.
