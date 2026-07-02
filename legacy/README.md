# Legacy — motor de conciliación original

Esta carpeta contiene el **script original** de conciliación bancaria del Banco
Credicoop (`conciliacion_mvp.py`), tal como funcionaba antes de la migración a la
arquitectura de NAT Automation Suite.

## Propósito

Es la **referencia funcional congelada**. El motor migrado en
`modules/bank_reconciliation/` debe producir **exactamente los mismos
resultados** que este script.

## Reglas

- **No modificar** este código.
- **No eliminarlo** hasta validar que el nuevo motor genera resultados idénticos
  (ver `tests/` y `docs/sprint-01.md`).
- Cualquier cambio en las reglas de negocio se decide y documenta explícitamente;
  este archivo queda como testigo del comportamiento original.

## Uso original (CLI)

```bash
python conciliacion_mvp.py --periodo 2025-07 --saldo-contable 21667987.68
```
