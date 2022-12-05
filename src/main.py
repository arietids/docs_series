#!/usr/bin/env python3
import sys
import argparse
import openpyxl
from pathlib import Path
import json
from docxtpl import DocxTemplate
import jinja2

__VERSION__ = '0.1'

def usage_args():
    parser = argparse.ArgumentParser(
                        description='Based on the document template (docx) and data specified in sheet (xlsx) produces document of series.'
                        )
    parser.add_argument('-w', '--workbook',
                        dest='workbook',
                        required=True,
                        type=Path,
                        help='Name of workbook file with data. [xlsx file]')
    parser.add_argument('-t', '--template',
                        dest='template', 
                        action='store',
                        help='Name of template file. [docx file]. If not specified json output to stdout.')
    parser.add_argument('-o', '--outfile',
                        dest='outfile',
                        action='store',
                        default='out.docx',
                        help='Name of document file to be generated. [docx file] Optional. Default: out.docx')
    parser.add_argument('--version',
                        action='version',
                        version='%(prog)s {version}'.format(version=__VERSION__))

    return parser.parse_args()

if __name__ == "__main__":

    args = usage_args()

    xlsx_file = Path(args.workbook)
    try:
        workbook = openpyxl.load_workbook(filename=xlsx_file, read_only=True)
    except FileNotFoundError:
        print(f"File {xlsx_file} not found.")
        sys.exit(1)
    except Exception as err:
        print(f"Unexpected {err}, {type(err)}")
        sys.exit(2)

    j = {}
    
    for sheet in workbook:
        
        data = []

        for i, row in enumerate(sheet.iter_rows(values_only=True)):
            if i == 0:
                header = row
            else:
                rowdata = {header[i] : row[i] for i, _ in enumerate(row)}
                data.append(rowdata)

        j[sheet.title] = data

    j["page_break"] = "\f"

    if args.template != None and args.template != '':
        try:
            docx_tpl = DocxTemplate(args.template)
            docx_tpl.render(j)
            docx_tpl.save(args.outfile)
        except jinja2.exceptions.TemplateSyntaxError as err:
            print(f"{err}")
            sys.exit(2)
        except Exception as err:
            print(f"{err}, {type(err)}")
            sys.exit(2)

    else:
        print(json.dumps(j))
