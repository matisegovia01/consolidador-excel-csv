"""Prueba reproducible del ejecutable; solo usa datos ficticios."""
import json
from pathlib import Path

from openpyxl import load_workbook
import xlrd

from .shell import ConsolidatorApp
from modules.afip_consolidator.models import ModoSalida


def verificar(carpeta):
    destino = Path(carpeta).resolve()
    destino.mkdir(parents=True, exist_ok=True)
    origen = destino / "ventas_prueba.csv"
    # No sobrescribir datos existentes al ejecutar la prueba.
    with origen.open("x", encoding="utf-8-sig") as archivo:
        archivo.write("Fecha de Emisión;Tipo;Número;Importe\n01/04/2026;001;000123;1.234,56\n")
    app = ConsolidatorApp()
    app.withdraw()
    try:
        app.update_idletasks()
        app.controller.seleccionar([origen])
        resultado = app.controller.preparar(periodo=("04/2026", "04/2026"))
        assert len(resultado.comprobantes) == 1
        assert "Falta Compras 04/2026" in resultado.avisos
        for modo in ModoSalida:
            salida = destino / (modo.name + ".xlsx")
            app.controller.exportar(salida, modo)
            libro = load_workbook(salida, read_only=True)
            try:
                assert len(libro.sheetnames) == 1
                assert libro.active["C2"].value == "000123"
                assert libro.active["D2"].value == "1.234,56"
            finally:
                libro.close()
        (destino / "resultado.json").write_text(json.dumps({"estado": "OK", "modalidades": 2,
                                                           "interfaz": "OK", "xlrd": xlrd.__version__}), encoding="utf-8")
    finally:
        app.destroy()
