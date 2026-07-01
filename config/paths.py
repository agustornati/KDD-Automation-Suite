"""Rutas del proyecto.

Centraliza todas las ubicaciones del sistema de archivos para que ningún
módulo tenga rutas hardcodeadas. Al importarse, garantiza que las carpetas
de trabajo existan.
"""

from __future__ import annotations

from pathlib import Path

# Raíz del repositorio (config/ está un nivel por debajo).
BASE_DIR: Path = Path(__file__).resolve().parent.parent

# Código y recursos.
MODULES_DIR: Path = BASE_DIR / "modules"
SHARED_DIR: Path = BASE_DIR / "shared"
ASSETS_DIR: Path = BASE_DIR / "assets"
LEGACY_DIR: Path = BASE_DIR / "legacy"

# Datos de trabajo.
DATA_DIR: Path = BASE_DIR / "data"
UPLOADS_DIR: Path = DATA_DIR / "uploads"
OUTPUTS_DIR: Path = DATA_DIR / "outputs"
TEMP_DIR: Path = DATA_DIR / "temp"

# Diagnóstico.
LOGS_DIR: Path = BASE_DIR / "logs"

# Carpetas que deben existir en tiempo de ejecución.
RUNTIME_DIRS: tuple[Path, ...] = (UPLOADS_DIR, OUTPUTS_DIR, TEMP_DIR, LOGS_DIR)


def ensure_runtime_dirs() -> None:
    """Crea las carpetas de trabajo si no existen (idempotente)."""
    for directory in RUNTIME_DIRS:
        directory.mkdir(parents=True, exist_ok=True)


# Se ejecuta al importar para que el resto del código pueda asumir su existencia.
ensure_runtime_dirs()
