"""Módulo de conciliación bancaria de KDD Automation Suite.

API pública:
    - :func:`reconcile`: ejecuta la conciliación y devuelve un
      :class:`ReconciliationResult` (sin generar archivos).
"""

from .models import (
    PeriodInfo,
    ReconciliationData,
    ReconciliationResult,
    ReconciliationStats,
)
from .service import reconcile
from .validators import ValidationError

__all__ = [
    "reconcile",
    "ReconciliationResult",
    "ReconciliationStats",
    "ReconciliationData",
    "PeriodInfo",
    "ValidationError",
]
