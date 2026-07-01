"""Configuración de logging para toda la plataforma.

Provee una base simple y sólida para diagnóstico: registra en un archivo dentro
de ``logs/`` y también en consola. Los módulos obtienen su logger con
:func:`get_logger`.
"""

from __future__ import annotations

import logging
from logging.handlers import RotatingFileHandler

from config import paths, settings

_LOG_FORMAT = "%(asctime)s | %(levelname)-8s | %(name)s | %(message)s"
_DATE_FORMAT = "%Y-%m-%d %H:%M:%S"

_configured = False


def setup_logging(level: int | None = None) -> None:
    """Configura el logger raíz una única vez (idempotente).

    Args:
        level: Nivel de logging. Por defecto usa ``settings.LOG_LEVEL``.
    """
    global _configured
    if _configured:
        return

    log_level = level if level is not None else settings.LOG_LEVEL
    formatter = logging.Formatter(_LOG_FORMAT, datefmt=_DATE_FORMAT)

    file_handler = RotatingFileHandler(
        paths.LOGS_DIR / settings.LOG_FILE_NAME,
        maxBytes=1_000_000,
        backupCount=5,
        encoding="utf-8",
    )
    file_handler.setFormatter(formatter)

    console_handler = logging.StreamHandler()
    console_handler.setFormatter(formatter)

    root = logging.getLogger()
    root.setLevel(log_level)
    root.addHandler(file_handler)
    root.addHandler(console_handler)

    _configured = True


def get_logger(name: str) -> logging.Logger:
    """Devuelve un logger con la configuración de la plataforma aplicada.

    Args:
        name: Nombre del logger (habitualmente ``__name__``).
    """
    setup_logging()
    return logging.getLogger(name)
