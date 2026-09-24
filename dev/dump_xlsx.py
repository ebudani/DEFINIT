"""Convierte una copia .xlsx de la planilla en dev/raw.json (mismo formato que getValues()).

Solo para desarrollo local: permite probar el parser y previsualizar el tablero
sin desplegar en Apps Script. raw.json contiene datos de ventas y NO se versiona.

Uso: python dev/dump_xlsx.py ruta/a/Facturacion.xlsx
"""
import datetime
import json
import pathlib
import sys

import openpyxl


def valor(v):
    if isinstance(v, (datetime.datetime, datetime.date)):
        return v.strftime("%Y-%m-%d")
    return "" if v is None else v


def main(ruta):
    wb = openpyxl.load_workbook(ruta, data_only=True, read_only=True)
    salida = {}
    for ws in wb.worksheets:
        # Igual que en Apps Script: solo el resumen y las pestañas mensuales.
        if ws.title == "Ventas Mensuales" or ws.title[:-2].lower() in {
            "enero", "febrero", "marzo", "abril", "mayo", "junio", "julio",
            "agosto", "septiembre", "setiembre", "octubre", "noviembre", "diciembre",
        }:
            salida[ws.title] = [[valor(v) for v in fila] for fila in ws.iter_rows(values_only=True)]
    destino = pathlib.Path(__file__).with_name("raw.json")
    destino.write_text(json.dumps(salida, ensure_ascii=False), encoding="utf-8")
    print(f"{destino} ({len(salida)} pestañas)")


if __name__ == "__main__":
    main(sys.argv[1])
