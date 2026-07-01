"""Exportador que empaqueta archivos en un ZIP."""

from __future__ import annotations

import zipfile
from pathlib import Path
from typing import Sequence


class ZipExporter:
    """Empaqueta un conjunto de archivos en un único ``.zip``.

    Es genérico: no depende de ningún módulo ni formato de origen.
    """

    def bundle(self, files: Sequence[Path], zip_path: Path) -> Path:
        """Comprime ``files`` dentro de ``zip_path``.

        Args:
            files: Archivos a incluir (cada uno se guarda por su nombre).
            zip_path: Ruta del archivo ZIP a crear.

        Returns:
            La ruta del ZIP generado.
        """
        zip_path = Path(zip_path)
        zip_path.parent.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
            for file in files:
                file = Path(file)
                if file.exists():
                    zf.write(file, arcname=file.name)
        return zip_path
