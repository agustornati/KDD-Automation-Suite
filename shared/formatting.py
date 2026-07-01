"""Helpers de formato para presentación (números y tamaños).

Transversales y reutilizables por la interfaz y por exportadores futuros. No
contienen lógica de negocio.
"""

from __future__ import annotations


def format_ars(value: float, decimals: int = 2) -> str:
    """Formatea un número al estilo argentino (miles con ``.``, decimales ``,``).

    Ejemplo: ``239716.74`` → ``"239.716,74"``.
    """
    formatted = f"{value:,.{decimals}f}"
    return formatted.replace(",", "\x00").replace(".", ",").replace("\x00", ".")


def human_size(num_bytes: int) -> str:
    """Convierte una cantidad de bytes a un tamaño legible (KB, MB, …)."""
    size = float(num_bytes)
    for unit in ("B", "KB", "MB", "GB"):
        if size < 1024 or unit == "GB":
            return f"{size:.0f} {unit}" if unit == "B" else f"{size:.1f} {unit}"
        size /= 1024
    return f"{size:.1f} GB"
