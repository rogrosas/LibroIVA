# Calculadora de IVA

App de escritorio ligera (Python + Tkinter + SQLite) para registrar ventas y compras diarias,
calcular el IVA y exportar planillas Excel por mes o por rango de fechas.

## Uso

Doble clic en `iniciar.bat` (la primera vez instala `openpyxl`), o bien:

```
python -m pip install -r requirements.txt
python main.py
```

## Pestañas

- **Registro diario**: elige la fecha (◀ Hoy ▶), ingresa ventas o compras. El monto puede ser
  total (con IVA, se desglosa el neto) o neto (se le suma el IVA). Doble clic para editar, Supr para eliminar.
  Abajo se ve el IVA débito, el IVA crédito y el IVA a pagar del día.
- **Resumen y exportación**: totales diarios del mes o del rango de fechas, y botón **Exportar a Excel…**.
- **Configuración**: tasa de IVA, si los montos incluyen IVA por defecto, decimales, moneda y datos de la empresa.
  Cada movimiento guarda su propia tasa, así que cambiarla no altera lo ya registrado.

## Planilla Excel

Tiene tres hojas: **Resumen** (totales e IVA a pagar o saldo a favor), **Resumen diario**
(una fila por día, con fórmulas) y **Detalle** (cada movimiento, con filtros).

Los datos quedan en `datos_iva.db`, junto al programa. Desde Configuración puedes crear una copia de seguridad.
