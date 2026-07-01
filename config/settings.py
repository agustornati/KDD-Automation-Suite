"""Ajustes de la aplicación.

Parámetros de presentación y comportamiento general de la plataforma. La
configuración de rutas vive en :mod:`config.paths`.
"""

from __future__ import annotations

import logging

# Identidad de la aplicación (usada por la interfaz).
APP_NAME: str = "KDD Automation Suite"
APP_TAGLINE: str = "Powered by KDD Consulting"
APP_VERSION: str = "0.1.1"

# Logging.
LOG_LEVEL: int = logging.INFO
LOG_FILE_NAME: str = "kdd_automation.log"

# Módulo de conciliación bancaria.
BANK_STATEMENT_EXTENSIONS: tuple[str, ...] = (".pdf",)
SAP_LEDGER_EXTENSIONS: tuple[str, ...] = (".xlsx", ".xlsm")

# Formato esperado del período ingresado por el usuario.
PERIOD_FORMAT_HINT: str = "YYYY-MM"
