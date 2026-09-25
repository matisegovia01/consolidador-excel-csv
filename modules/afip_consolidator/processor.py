from pathlib import Path
from datetime import datetime
import re

import pandas as pd

from .importers import leer_archivo
from .models import Resultado, ErrorConsolidacion, COLUMNAS_ORIGEN, COLUMNA_PERIODO, COLUMNA_ARCHIVO
from .importers import _normalizar_texto


def tipo_archivo(ruta, columnas):
    nombre = _normalizar_texto(Path(ruta).stem)
    compra = bool(re.search(r"(?:^|[^a-z])compras?(?:[^a-z]|$)", nombre))
    venta = bool(re.search(r"(?:^|[^a-z])ventas?(?:[^a-z]|$)", nombre))
    if compra != venta:
        return "Compras" if compra else "Ventas"
    nombres = " ".join(_normalizar_texto(c) for c in columnas)
    if ("vendedor" in nombres) != ("comprador" in nombres):
        return "Compras" if "vendedor" in nombres else "Ventas"
    return "Sin identificar"


def meses_periodo(desde, hasta):
    try:
        inicio = datetime.strptime(desde.strip(), "%m/%Y")
        fin = datetime.strptime(hasta.strip(), "%m/%Y")
    except ValueError as exc:
        raise ErrorConsolidacion("Use MM/AAAA para el período; por ejemplo 02/2024.") from exc
    if inicio > fin:
        raise ErrorConsolidacion("El inicio del período no puede ser posterior al final.")
    return [str(m) for m in pd.period_range(inicio, fin, freq="M")]


def consolidar(rutas: list[str | Path], periodo: tuple[str, str] | None = None) -> Resultado:
    """Conserva filas y valores; si un archivo falla, no produce salida parcial."""
    if not rutas:
        raise ErrorConsolidacion("Seleccione al menos un archivo.")
    tablas, archivos, avisos, errores = [], [], [], []
    vistas = set()
    nombres = {}
    meses = meses_periodo(*periodo) if periodo is not None else []
    tipos = {}
    cobertura = set()
    for seleccion in rutas:
        ruta = Path(seleccion).resolve()
        if ruta in vistas:
            avisos.append(f"{ruta.name}: archivo seleccionado más de una vez; se leyó una sola vez.")
            continue
        vistas.add(ruta)
        if ruta.name.casefold() in nombres:
            errores.append(f"{ruta.name}: hay archivos de carpetas distintas con el mismo nombre. "
                           "Renombre una copia para conservar un origen inequívoco.")
            continue
        nombres[ruta.name.casefold()] = ruta
        try:
            tabla, origen, advertencias = leer_archivo(ruta)
            tipo = tipo_archivo(ruta, tabla.columns)
            tipos[ruta.name] = tipo
            if tipo == "Sin identificar":
                advertencias.append(f"{ruta.name}: no se identificó Compras/Ventas. Renombre el archivo para identificarlo. "
                                    "Se ordenará al final de su mes y no acreditará cobertura por tipo.")
            for mes in tabla[COLUMNA_PERIODO].unique():
                cobertura.add((mes, tipo))
            tablas.append(tabla)
            archivos.append(origen)
            avisos.extend(advertencias)
        except Exception as exc:
            errores.append(f"{ruta.name}: {exc}")
    if errores:
        raise ErrorConsolidacion("No se preparó la consolidación. Revise estos archivos:\n\n" + "\n".join(errores))
    esquemas = [tuple(c for c in t.columns if c not in COLUMNAS_ORIGEN) for t in tablas]
    if len(set(esquemas)) > 1:
        avisos.append("Los encabezados difieren entre archivos. Se conserva la unión de columnas; "
                      "los campos ausentes quedan vacíos. No se unifican nombres distintos automáticamente.")
    datos = pd.concat(tablas, ignore_index=True, sort=False)
    orden = datos[COLUMNA_ARCHIVO].map(tipos).map({"Compras": 0, "Ventas": 1, "Sin identificar": 2})
    indices = pd.DataFrame({"mes": datos[COLUMNA_PERIODO], "tipo": orden}).sort_values(
        ["mes", "tipo"], kind="stable").index
    datos = datos.loc[indices].reset_index(drop=True)
    if periodo is not None:
        faltantes = [f"Falta {tipo} {mes[5:]}/{mes[:4]}" for mes in meses
                     for tipo in ("Compras", "Ventas") if (mes, tipo) not in cobertura]
        avisos.extend(faltantes or ["Cobertura completa: se encontraron Compras y Ventas en cada mes solicitado."])
        avisos.append("La cobertura verifica presencia de datos por mes y tipo; no garantiza que estén todos los comprobantes. "
                      "El período no filtra ni elimina filas.")
    columnas = [c for c in datos.columns if c not in COLUMNAS_ORIGEN] + list(COLUMNAS_ORIGEN)
    return Resultado(datos[columnas], archivos, avisos)
