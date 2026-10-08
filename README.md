# Calculadora de IVA

App de escritorio ligera (Python + Tkinter + SQLite) para registrar ventas y compras diarias,
calcular el IVA y exportar planillas Excel por mes o por rango de fechas.

## Uso

Doble clic en `iniciar.bat` (la primera vez instala `openpyxl`), o bien:

```
python -m pip install -r requirements.txt
python main.py
```

## Secciones

- **Ventas** y **Compras**: cada una tiene dos subpestañas.
  - **Registro diario**: elige la fecha (◀ Hoy ▶) e ingresa los registros. El monto puede ser total
    (con IVA, se desglosa el neto) o neto (se le suma el IVA). Doble clic para editar, Supr para eliminar.
    Muestra los totales del día y el acumulado del mes.
  - **Libro de ventas / compras**: lista del mes o rango de fechas (cada registro o totales por día)
    y botón para exportar solo ventas o solo compras a Excel.
- **Resumen IVA (ventas − compras)**: totales diarios combinados, IVA débito, IVA crédito e IVA a pagar
  o saldo a favor, con exportación a Excel.
- **Configuración**: tasa de IVA, si los montos incluyen IVA por defecto, decimales, moneda y datos de la empresa.
  Cada registro guarda su propia tasa, así que cambiarla no altera lo ya ingresado.

## Planillas Excel

- **Libro de ventas / compras**: hojas **Resumen**, **Por día** y **Detalle** (con totales y filtros).
- **Resumen IVA**: hojas **Resumen** (IVA a pagar o saldo a favor), **Resumen diario** (con fórmulas)
  y **Detalle** (ventas y compras juntas, con filtros).

Los datos quedan en `datos_iva.db`, junto al programa. Desde Configuración puedes crear una copia de seguridad.
