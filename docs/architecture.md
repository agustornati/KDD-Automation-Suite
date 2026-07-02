# Arquitectura — NAT Automation Suite

## Visión general

La plataforma sigue una arquitectura **modular por capas**. La regla central es:
la lógica de negocio vive en los módulos, nunca en la interfaz, y **el motor de
cada proceso no depende del formato de salida**.

```
┌──────────────────────────────────────────────┐
│                  app.py (UI)                   │  Streamlit — solo interacción
└───────────────┬───────────────────────────────┘
                │ llama
┌───────────────▼───────────────────────────────┐
│          modules/<automatización>/             │  lógica de negocio
│   service.py  → punto de entrada del módulo     │
│   engine.py   → algoritmos puros (sin I/O)      │
│   models.py   → objetos de dominio              │
│   validators.py, utils.py, exporters.py         │
└───────────────┬───────────────────────────────┘
                │ produce (objeto de dominio)      │ usa
┌───────────────▼──────────────┐   ┌──────────────▼─────────────┐
│   ReconciliationResult        │   │   shared/ (transversal)     │
│   (datos, stats, DataFrames)  │   │   exporters/, logging       │
└──────────────────────────────┘   └────────────────────────────┘
                ▲
                │ consume
        exporters (Excel, Zip, …futuros PDF/Email/API)
```

## Estructura de carpetas

```
NAT-Automation-Suite/
├── app.py                  # Interfaz Streamlit (sin lógica de negocio)
├── config/
│   ├── paths.py            # Rutas del proyecto
│   └── settings.py         # Parámetros de la aplicación
├── modules/
│   └── bank_reconciliation/
│       ├── service.py      # reconcile(): punto de entrada
│       ├── engine.py       # parseo + matching (lógica pura)
│       ├── models.py       # dataclasses / TypedDict del dominio
│       ├── validators.py   # validación de entradas
│       ├── utils.py        # helpers de dominio y DataFrames
│       └── exporters.py    # ExcelExporter (específico del módulo)
├── shared/
│   ├── logging_config.py   # logging para toda la plataforma
│   └── exporters/          # exportadores genéricos (base, ZipExporter)
├── data/{uploads,outputs,temp}/
├── logs/
├── tests/
├── docs/
├── assets/
└── legacy/                 # motor original congelado (referencia funcional)
```

## Principio clave: motor desacoplado de exportadores

El motor (`reconcile()`) **no genera archivos**. Devuelve un objeto de dominio
(`ReconciliationResult`) con:

- estadísticas (`ReconciliationStats`),
- DataFrames listos para consumir,
- el payload crudo (`ReconciliationData`) para reconstruir cualquier salida,
- advertencias y errores.

Los **exportadores** son componentes independientes que reciben ese resultado y
producen un formato concreto:

- `ExcelExporter` → los 3 libros Excel (Conciliación, No Coincidentes, Resumen).
- `ZipExporter` → empaqueta archivos en un ZIP.

En el futuro se pueden agregar `PdfExporter`, `EmailExporter`, `ApiExporter`,
`DashboardExporter` **sin modificar el motor**.

## Convenciones por módulo

Cada automatización futura (`invoices/`, `taxes/`, `reports/`, …) replica esta
estructura: un `service.py` como punto de entrada, un `engine.py` con la lógica
pura, `models.py` de dominio y exportadores desacoplados.

## Configuración, logging y datos

- **Configuración** separada en `config/paths.py` (sistema de archivos) y
  `config/settings.py` (parámetros de la app).
- **Logging** centralizado en `shared/logging_config.py`: archivo rotativo en
  `logs/` + consola.
- **Datos** de trabajo en `data/` (`uploads/`, `outputs/`, `temp/`), fuera del
  control de versiones.

## Motor legacy

`legacy/conciliacion_mvp.py` es el script original congelado. El motor migrado
debe producir **resultados idénticos**; el legacy permanece como referencia
funcional hasta validar la equivalencia (ver [`sprint-01.md`](sprint-01.md)).
