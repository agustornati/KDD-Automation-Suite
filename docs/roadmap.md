# Roadmap — NAT Automation Suite

El producto crece **una automatización a la vez**, cada una como módulo
independiente dentro de `modules/`. Las versiones son orientativas y se ajustan
según las prioridades del estudio.

## v0.1 — Conciliación bancaria  *(Sprint 1, en curso)*
- Primer módulo funcional: conciliación bancaria (Banco Credicoop).
- Reutiliza el motor existente encapsulado como servicio (`reconcile()`).
- Interfaz Streamlit: subir extracto (PDF) + libro SAP (Excel), ejecutar,
  ver estadísticas y descargar resultados (ZIP con 3 Excel).
- Base profesional: arquitectura modular, logging, tests, documentación.

## v0.2 — Gestión de clientes
- Alta y administración de clientes del estudio.
- Asociar conciliaciones y procesos a cada cliente.

## v0.3 — OCR de PDFs
- Lectura automática de comprobantes y documentos escaneados.

## v0.4 — Gestión documental
- Repositorio ordenado de documentos por cliente y período.

## v0.5 — Reportes
- Generación de reportes consolidados y exportables.

## v1.0 — Plataforma completa
- Asistente con IA.
- Dashboard integral.
- Automatización de procesos de punta a punta.
- Integraciones con ARCA.

## Ideas / backlog
- Procesamiento de facturas.
- Agenda de vencimientos impositivos.
- Nuevos exportadores (PDF, e-mail, API).
- Persistencia con SQLite.
