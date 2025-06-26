# Procesador de Subitems

This repository contains a small GUI utility that cleans Monday.com Excel exports.
The script loads an `.xlsx` file, fills empty "subitem" rows using the previous
valid row, removes helper lines, and highlights the processed rows. Original
rows are marked in blue and generated rows in yellow. A new workbook with
`_procesado` appended to the filename is created in the same folder.

## Requirements

Install the following packages in your Python environment:

```
pip install pandas openpyxl Pillow
```

`tkinter` is included with standard Python on most platforms.

## Running

Execute the script directly with Python:

```
python procesador_subitems.py
```

A window will ask for the Excel file and will open the resulting
`*_procesado.xlsx` when finished.

## Building an executable

A PyInstaller specification file (`procesador_subitems.spec`) is provided to
create a standalone executable:

```
pyinstaller procesador_subitems.spec
```

This will place the build artifacts in the `build/` directory.
