# Changelog

Todas las versiones relevantes del proyecto se documentan acá.
El formato sigue, de manera simplificada,
[Keep a Changelog](https://keepachangelog.com/es-ES/1.0.0/), y el versionado
[SemVer](https://semver.org/lang/es/).

## [No publicado]

## [0.1.1] — Sprint 1.5 (mejoras de UX)

Mejoras de experiencia de usuario para uso diario por personal administrativo.
**Sin cambios en el motor, el servicio ni las reglas de negocio.**

### Agregado
- Encabezado con logo opcional, número de versión y fecha/hora de la última
  ejecución.
- Layout organizado en secciones (Nueva Conciliación, Resultado, Descarga, Log
  de ejecución).
- Cargas de archivo más grandes y con borde visible, y resumen del archivo
  cargado (nombre, tamaño y estado).
- Validación de campos obligatorios antes de ejecutar, con un mensaje claro que
  indica exactamente qué falta.
- Progreso por etapas durante el proceso (`st.status`).
- Estadísticas presentadas como tarjetas con iconos y mensaje de éxito con
  tiempo de ejecución, registros procesados y diferencias.
- Sección de descarga independiente con cantidad de archivos y tamaño total.
- Log de ejecución con los mensajes importantes del proceso.
- La app recuerda el último período y saldo contable usados (`session_state`).
- Helpers de formato `format_ars` (formato argentino) y `human_size` en
  `shared/formatting.py`.
- Smoke tests de la interfaz con Streamlit `AppTest`.

### Cambiado
- Los errores técnicos ya no se muestran al usuario: se traducen a mensajes
  amigables y el detalle completo queda registrado únicamente en `logs/`.

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
