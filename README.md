# KDD Automation Suite

Plataforma de automatización del estudio contable **KDD Consulting**.

Un único lugar desde donde ejecutar todas las automatizaciones del estudio. El
primer módulo es la **conciliación bancaria**; la plataforma está diseñada para
crecer incorporando nuevas automatizaciones como módulos independientes
(facturas, OCR, reportes, gestión documental, IA, etc.).

> Visión completa y plan de evolución en [`docs/vision.md`](docs/vision.md) y
> [`docs/roadmap.md`](docs/roadmap.md).

## Objetivo

Reemplazar procesos manuales y scripts aislados por módulos mantenibles,
reutilizables y con calidad de software empresarial.

## Requisitos

- Python 3.10 o superior.
- Dependencias en [`requirements.txt`](requirements.txt): Streamlit, Pandas,
  OpenPyXL, pdfplumber (y pytest para los tests).

## Instalación

```bash
git clone https://github.com/agustornati/KDD-Automation-Suite.git
cd KDD-Automation-Suite

python -m venv .venv
# Windows
.venv\Scripts\activate
# Linux / macOS
source .venv/bin/activate

pip install -r requirements.txt
```

## Ejecución

```bash
streamlit run app.py
```

Se abre la aplicación en el navegador. Para la conciliación bancaria:

1. Ingresá el **período** (`YYYY-MM`, ej. `2025-07`).
2. Subí el **extracto bancario** (PDF).
3. Subí el **libro SAP** (Excel).
4. Ingresá el **saldo contable** de cierre.
5. Presioná **Conciliar**.
6. Revisá las estadísticas y **descargá los resultados** (ZIP con 3 Excel).

## Estructura del proyecto

```
KDD-Automation-Suite/
├── app.py                  # Interfaz Streamlit (sin lógica de negocio)
├── config/                 # Configuración (paths.py, settings.py)
├── modules/                # Automatizaciones (una por carpeta)
│   └── bank_reconciliation/#   → módulo de conciliación bancaria
├── shared/                 # Utilidades transversales (logging, exporters)
├── data/                   # uploads/ · outputs/ · temp/ (no versionado)
├── logs/                   # Logs de ejecución (no versionado)
├── tests/                  # Pruebas automatizadas
├── docs/                   # Documentación (visión, arquitectura, roadmap)
├── assets/                 # Recursos estáticos
└── legacy/                 # Motor original congelado (referencia)
```

Detalle de la arquitectura en [`docs/architecture.md`](docs/architecture.md).

## Tests

```bash
pytest
```

## Versionado

Los cambios se documentan en [`CHANGELOG.md`](CHANGELOG.md).
