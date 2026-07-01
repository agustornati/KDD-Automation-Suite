"""Interfaz común de los exportadores de resultados."""

from __future__ import annotations

from abc import ABC, abstractmethod
from pathlib import Path
from typing import Any


class ResultExporter(ABC):
    """Contrato para exportar un resultado de dominio a archivos.

    Cada implementación (Excel, PDF, etc.) recibe un objeto de resultado y una
    carpeta destino, y devuelve las rutas de los archivos generados. El motor
    que produce el resultado no conoce a los exportadores.
    """

    @abstractmethod
    def export(self, result: Any, output_dir: Path) -> list[Path]:
        """Genera los archivos de salida y devuelve sus rutas."""
        raise NotImplementedError
