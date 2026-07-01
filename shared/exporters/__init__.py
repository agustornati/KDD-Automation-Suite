"""Exportadores reutilizables por todos los módulos.

Los exportadores están **desacoplados** de los motores: reciben un resultado de
dominio (o un conjunto de archivos) y producen una salida en un formato
concreto. Así se pueden agregar nuevos formatos (PDF, API, e-mail, dashboard)
sin tocar la lógica de negocio.
"""

from .base import ResultExporter
from .zip_exporter import ZipExporter

__all__ = ["ResultExporter", "ZipExporter"]
