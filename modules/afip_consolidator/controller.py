from pathlib import Path

from .processor import consolidar
from .exporter import exportar
from .models import Resultado, ModoSalida


class ConsolidatorController:
    """Coordina selección y resultados, sin depender de la interfaz gráfica."""

    def __init__(self):
        self.rutas: list[Path] = []
        self.resultado: Resultado | None = None

    def seleccionar(self, rutas):
        self.resultado = None
        self.rutas = list(dict.fromkeys(Path(r).resolve() for r in rutas))

    def preparar(self, periodo=None):
        self.resultado = None
        self.resultado = consolidar(self.rutas.copy(), periodo=periodo)
        return self.resultado

    def exportar(self, ruta, modo: ModoSalida, *, reemplazar=False):
        if self.resultado is None:
            raise ValueError("Primero prepare la selección actual.")
        return exportar(self.resultado, ruta, modo, reemplazar=reemplazar)
