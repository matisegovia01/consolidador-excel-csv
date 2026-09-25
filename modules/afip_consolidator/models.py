from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from pathlib import Path

import pandas as pd


class ModoSalida(str, Enum):
    SABANA = "Vista sábana"
    POR_PERIODO = "Hojas por período"


@dataclass
class ArchivoOrigen:
    ruta: Path
    filas: int
    filas_vacias: int
    hoja: str


@dataclass
class Resultado:
    comprobantes: pd.DataFrame
    archivos: list[ArchivoOrigen]
    avisos: list[str] = field(default_factory=list)


class ErrorConsolidacion(ValueError):
    """Entrada no compatible; no debe generarse una salida parcial."""


# Columnas añadidas: nunca reemplazan campos del usuario.
COLUMNA_ARCHIVO = "Origen - archivo"
COLUMNA_HOJA = "Origen - hoja"
COLUMNA_FILA = "Origen - fila"
COLUMNA_PERIODO = "Origen - período"
COLUMNAS_ORIGEN = (COLUMNA_ARCHIVO, COLUMNA_HOJA, COLUMNA_FILA, COLUMNA_PERIODO)
