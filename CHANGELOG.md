# Changelog

Todas las versiones relevantes del proyecto se documentan acá.
El formato sigue, de manera simplificada,
[Keep a Changelog](https://keepachangelog.com/es-ES/1.0.0/), y el versionado
[SemVer](https://semver.org/lang/es/).

## [No publicado]

### En desarrollo
- Sprint 1: primer módulo de conciliación bancaria (ver más abajo v0.1.0).

## [0.1.0] — Sprint 1

Primera versión funcional de KDD Automation Suite.

### Agregado
- **Módulo de conciliación bancaria** (`modules/bank_reconciliation/`):
  - Servicio `reconcile()` que ejecuta la conciliación y devuelve un objeto de
    dominio (`ReconciliationResult`) con estadísticas, DataFrames y payload
    crudo, **sin generar archivos**.
  - Motor (`engine.py`) migrado del script legacy con comportamiento idéntico:
    parseo de extracto (PDF) y libro SAP (Excel), anulación de
    autocancelatorios, matching 1:1 y agrupado N:1 por suma.
- **Exportadores desacoplados**: `ExcelExporter` (3 libros Excel) y
  `ZipExporter` (empaquetado), listos para sumar PDF/e-mail/API en el futuro.
- **Interfaz Streamlit** (`app.py`): período, extracto PDF, libro SAP, saldo
  contable, ejecución, estadísticas y descarga de resultados (ZIP).
- **Base de plataforma**: configuración (`config/paths.py`, `config/settings.py`),
  logging (`shared/logging_config.py`), estructura de datos (`data/`),
  documentación (`docs/`) y tests (`tests/`).
- **Legacy congelado** (`legacy/conciliacion_mvp.py`) como referencia funcional.
