# Visión — NAT Automation Suite

## Qué es

NAT Automation Suite es la **plataforma única de automatización** del estudio
contable **NAT Consulting**.

El objetivo es contar con un solo lugar desde donde el estudio ejecute todas sus
automatizaciones, reemplazando scripts sueltos y tareas manuales repetitivas por
módulos mantenibles y reutilizables.

## Problema que resuelve

El trabajo contable involucra procesos repetitivos, propensos a error y que
consumen mucho tiempo: conciliaciones, carga de facturas, lectura de PDFs,
generación de reportes, seguimiento de vencimientos, etc. Hoy suelen resolverse
con planillas y scripts aislados, difíciles de mantener y escalar.

## Visión de largo plazo

Una plataforma que **crezca durante años**, incorporando una automatización a la
vez, cada una como un módulo independiente:

- Conciliación bancaria
- Procesamiento de facturas
- OCR de PDFs
- Reportes
- Gestión documental
- Agenda de vencimientos
- Integraciones con ARCA
- Automatización de procesos
- Dashboard
- Asistente con IA

## Principios rectores

1. **Escalabilidad**: cada automatización es un módulo dentro de `modules/`.
   Agregar una nueva no debe requerir tocar las existentes.
2. **Mantenibilidad**: código modular, con responsabilidades separadas, type
   hints, docstrings y funciones pequeñas.
3. **Reutilización**: la lógica que ya funciona se encapsula, no se reescribe.
4. **Desacoplamiento**: el motor de cada proceso no depende del formato de
   salida ni de la interfaz. Los exportadores y la UI son intercambiables.
5. **Calidad empresarial**: documentación, tests y logging desde el primer día.

## Estado actual

Primer módulo: **conciliación bancaria** (Banco Credicoop), reutilizando el
motor ya validado en producción. Ver [`roadmap.md`](roadmap.md) para el plan de
evolución.
