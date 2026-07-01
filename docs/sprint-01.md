# Sprint 1 — Módulo de conciliación bancaria

## Objetivo

Construir la **primera versión completamente funcional** de KDD Automation Suite:
un producto usable que permita ejecutar la conciliación bancaria de punta a
punta, reutilizando el motor existente y dejando una base de software
profesional y escalable.

**No** se agregan funcionalidades nuevas ni se modifican las reglas de negocio.

## Alcance

El usuario puede:

1. Abrir la aplicación.
2. Ingresar el período (`YYYY-MM`).
3. Subir el extracto bancario (PDF).
4. Subir el libro contable exportado de SAP (Excel).
5. Ingresar el saldo contable de cierre.
6. Ejecutar la conciliación.
7. Ver las estadísticas del proceso.
8. Descargar los resultados (ZIP con los 3 Excel).

## Fuera de alcance (Sprints futuros)

IA, login, usuarios, base de datos, dashboard, OCR, reportes, otras
automatizaciones e integraciones con ARCA.

## Decisiones de diseño

- **Entrada del banco = PDF**: el motor existente parsea el extracto desde PDF
  (con `pdfplumber`). Se mantiene ese formato para no alterar la lógica.
- **Salida = ZIP** con los 3 Excel que ya produce el motor (Conciliación,
  No Coincidentes, Resumen de saldos).
- **Saldo contable obligatorio**: necesario para el Resumen de saldos.
- **Motor desacoplado de exportadores**: `reconcile()` devuelve un objeto de
  dominio; los exportadores generan los archivos (ver
  [`architecture.md`](architecture.md)).

## Reutilización del motor existente

El script original (`legacy/conciliacion_mvp.py`) se migró a
`modules/bank_reconciliation/` **sin cambiar el comportamiento**:

- `engine.py` conserva los algoritmos (parseo, anulados, matching 1:1 y N:1).
- Regla contable intacta: *Conta DEBE ↔ Banco CRÉDITO | Conta HABER ↔ Banco DÉBITO*.
- Refactor limitado a: separar responsabilidades, type hints, docstrings,
  nombres y organización.

## Criterio de validación (Definition of Done)

El módulo se considera validado cuando, con los datos reales de un período
(p. ej. `2025-07`), los tres Excel generados por el nuevo motor son
**equivalentes** a los del script legacy. Hasta entonces, `legacy/` permanece
como referencia funcional y no se elimina.
