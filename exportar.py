"""Exportación del libro de IVA a Excel (.xlsx)."""
from datetime import datetime

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from datos import fmt_fecha

AZUL = "1F4E78"
VERDE = "E2EFDA"
NARANJA = "FCE4D6"
GRIS = "F2F2F2"

F_TITULO = Font(size=14, bold=True, color=AZUL)
F_ENCABEZADO = Font(bold=True, color="FFFFFF")
R_ENCABEZADO = PatternFill("solid", fgColor=AZUL)
F_NEGRITA = Font(bold=True)
BORDE_SUP = Border(top=Side(style="thin"), bottom=Side(style="double"))
CENTRO = Alignment(horizontal="center", vertical="center", wrap_text=True)


def _formato_moneda(simbolo, decimales):
    simbolo = simbolo.replace('"', "")
    base = "#,##0" + ("." + "0" * decimales if decimales else "")
    return f'"{simbolo} "{base};-"{simbolo} "{base}' if simbolo else f"{base};-{base}"


def _encabezados(ws, fila, titulos, anchos):
    for col, (titulo, ancho) in enumerate(zip(titulos, anchos), start=1):
        c = ws.cell(row=fila, column=col, value=titulo)
        c.font, c.fill, c.alignment = F_ENCABEZADO, R_ENCABEZADO, CENTRO
        ws.column_dimensions[get_column_letter(col)].width = ancho
    ws.row_dimensions[fila].height = 30


def exportar_excel(ruta, db, desde, hasta, etiqueta, cfg):
    dias = db.resumen_diario(desde, hasta)
    movimientos = db.listar(desde, hasta)
    if not dias:
        raise ValueError("No hay movimientos en el período seleccionado.")

    fmt = _formato_moneda(cfg.get("simbolo_moneda", "$"), int(cfg.get("decimales", 0)))
    wb = Workbook()

    # ------------------------------------------------------------ Resumen diario
    wd = wb.active
    wd.title = "Resumen diario"
    _encabezados(
        wd, 1,
        ["Fecha", "N° ventas", "Ventas neto", "IVA débito\n(ventas)", "Ventas total",
         "N° compras", "Compras neto", "IVA crédito\n(compras)", "Compras total", "IVA a pagar\n(débito − crédito)"],
        [12, 10, 16, 16, 16, 11, 16, 16, 16, 18],
    )
    for i, d in enumerate(dias, start=2):
        wd.append([
            d["fecha"], d["n_ventas"], d["ventas_neto"], d["ventas_iva"], d["ventas_total"],
            d["n_compras"], d["compras_neto"], d["compras_iva"], d["compras_total"], f"=D{i}-H{i}",
        ])
        wd.cell(row=i, column=1).number_format = "DD/MM/YYYY"
        for col in (3, 4, 5, 7, 8, 9, 10):
            wd.cell(row=i, column=col).number_format = fmt
        relleno = PatternFill("solid", fgColor=GRIS) if i % 2 else None
        if relleno:
            for col in range(1, 11):
                wd.cell(row=i, column=col).fill = relleno
    ultima = len(dias) + 1
    tot = ultima + 1
    wd.cell(row=tot, column=1, value="TOTAL")
    for col in range(2, 11):
        letra = get_column_letter(col)
        c = wd.cell(row=tot, column=col, value=f"=SUM({letra}2:{letra}{ultima})")
        if col not in (2, 6):
            c.number_format = fmt
    for col in range(1, 11):
        c = wd.cell(row=tot, column=col)
        c.font, c.border = F_NEGRITA, BORDE_SUP
    wd.freeze_panes = "A2"
    wd.auto_filter.ref = f"A1:J{ultima}"

    # ------------------------------------------------------------ Detalle
    wm = wb.create_sheet("Detalle")
    _encabezados(
        wm, 1,
        ["Fecha", "Tipo", "N° documento", "Descripción", "Tasa IVA %", "Neto", "IVA", "Total"],
        [12, 10, 16, 40, 11, 16, 16, 16],
    )
    for i, m in enumerate(movimientos, start=2):
        wm.append([
            datetime.fromisoformat(m["fecha"]).date(), m["tipo"].capitalize(), m["documento"],
            m["descripcion"], m["tasa"] / 100, m["neto"], m["iva"], m["total"],
        ])
        wm.cell(row=i, column=1).number_format = "DD/MM/YYYY"
        wm.cell(row=i, column=5).number_format = "0.##%"
        for col in (6, 7, 8):
            wm.cell(row=i, column=col).number_format = fmt
        color = VERDE if m["tipo"] == "venta" else NARANJA
        wm.cell(row=i, column=2).fill = PatternFill("solid", fgColor=color)
    wm.freeze_panes = "A2"
    wm.auto_filter.ref = f"A1:H{len(movimientos) + 1}"

    # ------------------------------------------------------------ Resumen (primera hoja)
    wr = wb.create_sheet("Resumen", 0)
    wr.column_dimensions["A"].width = 34
    for col in "BCDE":
        wr.column_dimensions[col].width = 18

    wr["A1"] = f"Libro de IVA — {etiqueta}"
    wr["A1"].font = F_TITULO
    empresa = " · ".join(x for x in (cfg.get("nombre_empresa", ""), cfg.get("id_fiscal", "")) if x)
    if empresa:
        wr["A2"] = empresa
        wr["A2"].font = F_NEGRITA
    wr["A3"] = f"Período: {fmt_fecha(desde)} al {fmt_fecha(hasta)}"
    wr["A4"] = f"Generado: {datetime.now():%d/%m/%Y %H:%M}   ·   Tasa de IVA configurada: {cfg.get('tasa_iva')}%"
    wr["A4"].font = Font(italic=True, color="666666")

    _encabezados(wr, 6, ["Concepto", "Cantidad", "Neto", "IVA", "Total"], [34, 12, 18, 18, 18])
    ref = f"'Resumen diario'!"
    filas = [
        ("Ventas (IVA débito)", "B", "C", "D", "E", VERDE),
        ("Compras (IVA crédito)", "F", "G", "H", "I", NARANJA),
    ]
    for r, (nombre, n, neto, iva, total, color) in enumerate(filas, start=7):
        wr.cell(row=r, column=1, value=nombre).fill = PatternFill("solid", fgColor=color)
        wr.cell(row=r, column=2, value=f"={ref}{n}{tot}")
        for col, letra in ((3, neto), (4, iva), (5, total)):
            wr.cell(row=r, column=col, value=f"={ref}{letra}{tot}").number_format = fmt

    wr["A10"] = "IVA a pagar (débito − crédito)"
    wr["D10"] = "=D7-D8"
    wr["D10"].number_format = fmt
    for c in (wr["A10"], wr["D10"]):
        c.font = Font(bold=True, size=12)
        c.border = BORDE_SUP
    wr["A11"] = '=IF(D10>=0,"Resultado: IVA a pagar","Resultado: saldo a favor (remanente de crédito)")'
    wr["A11"].font = Font(italic=True, color=AZUL)
    wr["A12"] = "Días con movimientos"
    wr["B12"] = len(dias)

    wb.save(ruta)
