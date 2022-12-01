# Korespondencja seryjna

## Środowisko

### Utwórz jeśli nie utworzone

```bash
python3 -m venv .
```

### Aktywuj

```bash
source bin/activate
```

## Zainstaluj wymagane

```bash
bin/pip install --no-cache-dir -r requirements.txt
```

## Sposób użycia

```bash
$ bin/python main.py -h
usage: main.py [-h] -w WORKBOOK [-s SHEET_NAME] [-t TEMPLATE] [-o OUTFILE]

Based on the document template (docx) and data specified in sheet (xlsx) produces document of series.

optional arguments:
  -h, --help            show this help message and exit
  -w WORKBOOK, --workbook WORKBOOK
                        Name of workbook file with data. [xlsx file]
  -s SHEET_NAME, --sheet-name SHEET_NAME
                        Name of sheet. Optional. Default: active is taken.
  -t TEMPLATE, --template TEMPLATE
                        Name of template file. [docx file]. If not specified json output to stdout.
  -o OUTFILE, --outfile OUTFILE
                        Name of document file to be generated. [docx file] Optional. Default: out.docx
```

Skrypt main.py czyta pierwszy arkusz lub wskazany opcją -s ze wskazanego dokumentu opcją -w. W pierwszym wierszu są nagłówki, pozostałe zawierają dane. Jeżeli wskazany jest dokument z szablonem opcją -t to wygenerowany jest dokument. Nazwę dokumentu do wygenerowania podajemy opcją -o. Jeśli nie podano out.docx jest domyślna.
Przykład użycia

```bash
bin/python main.py -w workbook.xlsx -s Sheet1 -t template.docx -o out.docx
```

## Dokumentacja

Jinja2

- https://jinja.palletsprojects.com/en/3.1.x/

python-docx-template

- https://github.com/elapouya/python-docx-template 

- https://docxtpl.readthedocs.io/en/latest/#

OpenPyXL

- https://openpyxl.readthedocs.io/en/stable/index.html#
