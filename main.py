"""Calculadora de IVA — registro diario de ventas y compras con exportación a Excel."""
import calendar
import os
import subprocess
import sys
import tkinter as tk
from datetime import date, timedelta
from tkinter import filedialog, font as tkfont, messagebox, ttk

from datos import BaseDatos, calcular_iva, fmt_fecha, fmt_monto, fmt_numero, parse_fecha, parse_monto

MESES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio",
         "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
DIAS = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]
COLOR_VENTA = "#1b6e2e"
COLOR_COMPRA = "#a04a00"


def parse_tasa(texto):
    try:
        tasa = float(texto.strip().replace("%", "").replace(",", "."))
    except ValueError:
        raise ValueError(f"Tasa de IVA inválida: «{texto}».")
    if not 0 <= tasa <= 100:
        raise ValueError("La tasa de IVA debe estar entre 0 y 100.")
    return tasa


def abrir_archivo(ruta):
    if sys.platform == "win32":
        os.startfile(ruta)
    else:
        subprocess.Popen(["open" if sys.platform == "darwin" else "xdg-open", ruta])


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Calculadora de IVA")
        escala = self.winfo_fpixels("1i") / 96
        self.geometry(f"{int(1080 * escala)}x{int(720 * escala)}")
        self.minsize(int(900 * escala), int(600 * escala))

        self.db = BaseDatos()
        self.cfg = self.db.obtener_config()
        self._estilos()

        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=8, pady=8)
        self.tab_registro = RegistroTab(nb, self)
        self.tab_resumen = ResumenTab(nb, self)
        self.tab_config = ConfigTab(nb, self)
        nb.add(self.tab_registro, text="  Registro diario  ")
        nb.add(self.tab_resumen, text="  Resumen y exportación  ")
        nb.add(self.tab_config, text="  Configuración  ")
        nb.bind("<<NotebookTabChanged>>", lambda e: self.tab_resumen.refrescar()
                if nb.select() == str(self.tab_resumen) else None)

    def _estilos(self):
        estilo = ttk.Style(self)
        if "vista" in estilo.theme_names():
            estilo.theme_use("vista")
        base = tkfont.nametofont("TkDefaultFont")
        base.configure(size=10)
        tkfont.nametofont("TkTextFont").configure(size=10)
        alto = base.metrics("linespace")
        estilo.configure("Treeview", rowheight=int(alto * 1.5))
        estilo.configure("Treeview.Heading", font=(base.actual("family"), 10, "bold"))
        estilo.configure("Bold.TLabel", font=(base.actual("family"), 10, "bold"))
        estilo.configure("Grande.TLabel", font=(base.actual("family"), 14, "bold"))
        estilo.configure("Sub.TLabel", foreground="#666666")
        estilo.configure("Accent.TButton", font=(base.actual("family"), 10, "bold"))

    # ---- helpers de configuración
    @property
    def decimales(self):
        return int(self.cfg.get("decimales", 0))

    @property
    def tasa(self):
        return float(self.cfg.get("tasa_iva", 19))

    def fmt(self, valor):
        return fmt_monto(valor, self.cfg.get("simbolo_moneda", "$"), self.decimales)

    def config_actualizada(self):
        self.cfg = self.db.obtener_config()
        self.tab_registro.aplicar_config()
        self.tab_resumen.refrescar()


# =========================================================================== Registro diario

class RegistroTab(ttk.Frame):
    def __init__(self, master, app):
        super().__init__(master, padding=12)
        self.app = app
        self.editando_id = None

        # ---- barra de fecha
        barra = ttk.Frame(self)
        barra.pack(fill="x")
        ttk.Label(barra, text="Fecha:", style="Bold.TLabel").pack(side="left")
        self.var_fecha = tk.StringVar(value=fmt_fecha(date.today()))
        ent_fecha = ttk.Entry(barra, textvariable=self.var_fecha, width=12, justify="center")
        ent_fecha.pack(side="left", padx=(6, 4))
        ent_fecha.bind("<Return>", lambda e: self.cargar_dia())
        ent_fecha.bind("<FocusOut>", lambda e: self.cargar_dia(silencioso=True))
        ttk.Button(barra, text="◀", width=3, command=lambda: self.mover_dia(-1)).pack(side="left")
        ttk.Button(barra, text="Hoy", width=6, command=self.ir_hoy).pack(side="left", padx=2)
        ttk.Button(barra, text="▶", width=3, command=lambda: self.mover_dia(1)).pack(side="left")
        self.lbl_dia = ttk.Label(barra, style="Sub.TLabel")
        self.lbl_dia.pack(side="left", padx=12)

        cuerpo = ttk.Frame(self)
        cuerpo.pack(fill="both", expand=True, pady=(10, 0))

        # ---- formulario
        self.form = ttk.LabelFrame(cuerpo, text="Nuevo movimiento", padding=12)
        self.form.pack(side="left", fill="y")
        f = self.form

        self.var_tipo = tk.StringVar(value="venta")
        fila = ttk.Frame(f)
        fila.grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 8))
        ttk.Radiobutton(fila, text="Venta", value="venta", variable=self.var_tipo).pack(side="left")
        ttk.Radiobutton(fila, text="Compra", value="compra", variable=self.var_tipo).pack(side="left", padx=14)

        self.var_monto = tk.StringVar()
        self.var_incluye = tk.BooleanVar()
        self.var_tasa = tk.StringVar()
        self.var_doc = tk.StringVar()
        self.var_desc = tk.StringVar()

        ttk.Label(f, text="Monto:").grid(row=1, column=0, sticky="w", pady=3)
        self.ent_monto = ttk.Entry(f, textvariable=self.var_monto, width=22, justify="right")
        self.ent_monto.grid(row=1, column=1, sticky="ew", pady=3)
        ttk.Checkbutton(f, text="El monto incluye IVA (total)", variable=self.var_incluye) \
            .grid(row=2, column=1, sticky="w", pady=(0, 3))
        ttk.Label(f, text="Tasa IVA %:").grid(row=3, column=0, sticky="w", pady=3)
        ttk.Entry(f, textvariable=self.var_tasa, width=8, justify="right").grid(row=3, column=1, sticky="w", pady=3)
        ttk.Label(f, text="N° documento:").grid(row=4, column=0, sticky="w", pady=3)
        ttk.Entry(f, textvariable=self.var_doc, width=22).grid(row=4, column=1, sticky="ew", pady=3)
        ttk.Label(f, text="Descripción:").grid(row=5, column=0, sticky="w", pady=3)
        ttk.Entry(f, textvariable=self.var_desc, width=22).grid(row=5, column=1, sticky="ew", pady=3)

        prev = ttk.Frame(f, padding=(0, 10))
        prev.grid(row=6, column=0, columnspan=2, sticky="ew")
        self.lbl_prev = {}
        for i, (clave, texto) in enumerate((("neto", "Neto"), ("iva", "IVA"), ("total", "Total"))):
            ttk.Label(prev, text=texto + ":", style="Sub.TLabel").grid(row=i, column=0, sticky="w")
            self.lbl_prev[clave] = ttk.Label(prev, text="—", style="Bold.TLabel")
            self.lbl_prev[clave].grid(row=i, column=1, sticky="e", padx=(10, 0))
        prev.columnconfigure(1, weight=1)

        botones = ttk.Frame(f)
        botones.grid(row=7, column=0, columnspan=2, sticky="ew")
        self.btn_guardar = ttk.Button(botones, text="Agregar", style="Accent.TButton", command=self.guardar)
        self.btn_guardar.pack(side="left", fill="x", expand=True)
        self.btn_cancelar = ttk.Button(botones, text="Cancelar", command=self.limpiar_form)

        for var in (self.var_monto, self.var_incluye, self.var_tasa):
            var.trace_add("write", lambda *a: self.actualizar_preview())
        for widget in f.winfo_children():
            if isinstance(widget, ttk.Entry):
                widget.bind("<Return>", lambda e: self.guardar())

        # ---- lista del día
        derecha = ttk.Frame(cuerpo)
        derecha.pack(side="left", fill="both", expand=True, padx=(12, 0))
        cols = (("tipo", "Tipo", 80, "w"), ("doc", "N° doc.", 90, "w"), ("desc", "Descripción", 200, "w"),
                ("neto", "Neto", 110, "e"), ("iva", "IVA", 100, "e"), ("total", "Total", 110, "e"))
        contenedor = ttk.Frame(derecha)
        contenedor.pack(fill="both", expand=True)
        self.tree = ttk.Treeview(contenedor, columns=[c[0] for c in cols], show="headings", selectmode="browse")
        for clave, titulo, ancho, anc in cols:
            self.tree.heading(clave, text=titulo, anchor=anc)
            self.tree.column(clave, width=ancho, anchor=anc, stretch=clave == "desc")
        sb = ttk.Scrollbar(contenedor, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="left", fill="y")
        self.tree.tag_configure("venta", foreground=COLOR_VENTA)
        self.tree.tag_configure("compra", foreground=COLOR_COMPRA)
        self.tree.bind("<Double-1>", lambda e: self.editar())
        self.tree.bind("<Delete>", lambda e: self.eliminar())

        acciones = ttk.Frame(derecha)
        acciones.pack(fill="x", pady=(6, 0))
        ttk.Button(acciones, text="Editar", command=self.editar).pack(side="left")
        ttk.Button(acciones, text="Eliminar", command=self.eliminar).pack(side="left", padx=6)
        ttk.Label(acciones, text="Doble clic para editar · Supr para eliminar", style="Sub.TLabel").pack(side="right")

        # ---- resumen del día
        res = ttk.LabelFrame(self, text="Resumen del día", padding=12)
        res.pack(fill="x", pady=(10, 0))
        self.lbl_res = {}
        bloques = (
            ("Ventas", COLOR_VENTA, (("ventas_neto", "Neto"), ("ventas_iva", "IVA débito"), ("ventas_total", "Total"))),
            ("Compras", COLOR_COMPRA, (("compras_neto", "Neto"), ("compras_iva", "IVA crédito"), ("compras_total", "Total"))),
        )
        for col, (titulo, color, items) in enumerate(bloques):
            bloque = ttk.Frame(res)
            bloque.grid(row=0, column=col, sticky="nw", padx=(0, 40))
            self.lbl_res["n_" + titulo.lower()] = ttk.Label(bloque, style="Bold.TLabel", foreground=color)
            self.lbl_res["n_" + titulo.lower()].grid(row=0, column=0, columnspan=2, sticky="w")
            for i, (clave, texto) in enumerate(items, start=1):
                ttk.Label(bloque, text=texto + ":", style="Sub.TLabel").grid(row=i, column=0, sticky="w")
                self.lbl_res[clave] = ttk.Label(bloque)
                self.lbl_res[clave].grid(row=i, column=1, sticky="e", padx=(12, 0))
        final = ttk.Frame(res)
        final.grid(row=0, column=2, sticky="ne")
        res.columnconfigure(2, weight=1)
        self.lbl_res["titulo_pagar"] = ttk.Label(final, style="Sub.TLabel")
        self.lbl_res["titulo_pagar"].pack(anchor="e")
        self.lbl_res["iva_pagar"] = ttk.Label(final, style="Grande.TLabel")
        self.lbl_res["iva_pagar"].pack(anchor="e")

        self.aplicar_config()
        self.ent_monto.focus_set()

    # ---- fecha
    def fecha_actual(self, silencioso=False):
        try:
            return parse_fecha(self.var_fecha.get())
        except ValueError as e:
            if not silencioso:
                messagebox.showerror("Fecha inválida", str(e), parent=self)
            return None

    def mover_dia(self, delta):
        f = self.fecha_actual()
        if f:
            self.var_fecha.set(fmt_fecha(f + timedelta(days=delta)))
            self.cargar_dia()

    def ir_hoy(self):
        self.var_fecha.set(fmt_fecha(date.today()))
        self.cargar_dia()

    def cargar_dia(self, silencioso=False):
        f = self.fecha_actual(silencioso)
        if f is None:
            return
        self.var_fecha.set(fmt_fecha(f))
        self.lbl_dia.config(text=f"{DIAS[f.weekday()]} {f.day} de {MESES[f.month - 1].lower()} de {f.year}")

        self.tree.delete(*self.tree.get_children())
        fmt = self.app.fmt
        for m in self.app.db.listar(f, f):
            self.tree.insert("", "end", iid=str(m["id"]), tags=(m["tipo"],), values=(
                m["tipo"].capitalize(), m["documento"], m["descripcion"],
                fmt(m["neto"]), fmt(m["iva"]), fmt(m["total"])))

        t = self.app.db.totales(f, f)
        self.lbl_res["n_ventas"].config(text=f"Ventas ({t['n_ventas']})")
        self.lbl_res["n_compras"].config(text=f"Compras ({t['n_compras']})")
        for clave in ("ventas_neto", "ventas_iva", "ventas_total", "compras_neto", "compras_iva", "compras_total"):
            self.lbl_res[clave].config(text=fmt(t[clave]))
        pagar = t["iva_pagar"]
        self.lbl_res["titulo_pagar"].config(
            text="IVA a pagar del día (débito − crédito)" if pagar >= 0 else "Saldo a favor del día (crédito > débito)")
        self.lbl_res["iva_pagar"].config(text=fmt(pagar), foreground="#1F4E78" if pagar >= 0 else COLOR_VENTA)

    # ---- formulario
    def aplicar_config(self):
        if self.editando_id is None:
            self.var_tasa.set(fmt_numero(self.app.tasa, 2).rstrip("0").rstrip(","))
            self.var_incluye.set(self.app.cfg.get("montos_incluyen_iva") == "1")
        self.actualizar_preview()
        self.cargar_dia(silencioso=True)

    def _leer_form(self):
        monto = parse_monto(self.var_monto.get())
        if monto <= 0:
            raise ValueError("El monto debe ser mayor que cero.")
        return monto, parse_tasa(self.var_tasa.get())

    def actualizar_preview(self):
        try:
            monto, tasa = self._leer_form()
            valores = calcular_iva(monto, tasa, self.var_incluye.get(), self.app.decimales)
        except ValueError:
            valores = None
        for clave, valor in zip(("neto", "iva", "total"), valores or (None,) * 3):
            self.lbl_prev[clave].config(text="—" if valor is None else self.app.fmt(valor))

    def guardar(self):
        f = self.fecha_actual()
        if f is None:
            return
        try:
            monto, tasa = self._leer_form()
        except ValueError as e:
            messagebox.showerror("Dato inválido", str(e), parent=self)
            return
        args = dict(fecha=f, tipo=self.var_tipo.get(), monto=monto, incluye_iva=self.var_incluye.get(),
                    tasa=tasa, decimales=self.app.decimales,
                    documento=self.var_doc.get().strip(), descripcion=self.var_desc.get().strip())
        if self.editando_id is None:
            self.app.db.agregar(**args)
        else:
            self.app.db.actualizar(self.editando_id, **args)
        self.limpiar_form()
        self.cargar_dia()

    def limpiar_form(self):
        self.editando_id = None
        self.form.config(text="Nuevo movimiento")
        self.btn_guardar.config(text="Agregar")
        self.btn_cancelar.pack_forget()
        self.var_monto.set("")
        self.var_doc.set("")
        self.var_desc.set("")
        self.aplicar_config()
        self.ent_monto.focus_set()

    def _seleccion(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Sin selección", "Selecciona un movimiento de la lista.", parent=self)
            return None
        return int(sel[0])

    def editar(self):
        id_ = self._seleccion()
        if id_ is None:
            return
        m = self.app.db.obtener(id_)
        self.editando_id = id_
        self.form.config(text=f"Editando movimiento #{id_}")
        self.btn_guardar.config(text="Guardar cambios")
        self.btn_cancelar.pack(side="left", padx=(6, 0))
        self.var_tipo.set(m["tipo"])
        self.var_incluye.set(bool(m["incluye_iva"]))
        self.var_tasa.set(fmt_numero(m["tasa"], 2).rstrip("0").rstrip(","))
        self.var_monto.set(fmt_numero(m["monto"], self.app.decimales))
        self.var_doc.set(m["documento"])
        self.var_desc.set(m["descripcion"])
        self.ent_monto.focus_set()
        self.ent_monto.select_range(0, "end")

    def eliminar(self):
        id_ = self._seleccion()
        if id_ is None:
            return
        m = self.app.db.obtener(id_)
        if messagebox.askyesno("Eliminar", f"¿Eliminar la {m['tipo']} por {self.app.fmt(m['total'])}?", parent=self):
            self.app.db.eliminar(id_)
            if self.editando_id == id_:
                self.limpiar_form()
            self.cargar_dia()


# =========================================================================== Resumen

class ResumenTab(ttk.Frame):
    def __init__(self, master, app):
        super().__init__(master, padding=12)
        self.app = app
        hoy = date.today()

        ctrl = ttk.LabelFrame(self, text="Período", padding=10)
        ctrl.pack(fill="x")
        self.var_modo = tk.StringVar(value="mes")
        ttk.Radiobutton(ctrl, text="Por mes", value="mes", variable=self.var_modo, command=self._modo) \
            .grid(row=0, column=0, sticky="w", pady=3)
        self.cmb_mes = ttk.Combobox(ctrl, values=MESES, state="readonly", width=12)
        self.cmb_mes.current(hoy.month - 1)
        self.cmb_mes.grid(row=0, column=1, padx=6, sticky="w")
        self.var_anio = tk.StringVar(value=str(hoy.year))
        self.spn_anio = ttk.Spinbox(ctrl, from_=2000, to=2100, textvariable=self.var_anio, width=6)
        self.spn_anio.grid(row=0, column=2, sticky="w")

        ttk.Radiobutton(ctrl, text="Rango de fechas", value="rango", variable=self.var_modo, command=self._modo) \
            .grid(row=1, column=0, sticky="w", pady=3)
        self.var_desde = tk.StringVar(value=fmt_fecha(hoy.replace(day=1)))
        self.var_hasta = tk.StringVar(value=fmt_fecha(hoy))
        rango = ttk.Frame(ctrl)
        rango.grid(row=1, column=1, columnspan=3, sticky="w", padx=6)
        ttk.Label(rango, text="Desde").pack(side="left")
        self.ent_desde = ttk.Entry(rango, textvariable=self.var_desde, width=12, justify="center")
        self.ent_desde.pack(side="left", padx=(4, 10))
        ttk.Label(rango, text="Hasta").pack(side="left")
        self.ent_hasta = ttk.Entry(rango, textvariable=self.var_hasta, width=12, justify="center")
        self.ent_hasta.pack(side="left", padx=4)

        ctrl.columnconfigure(4, weight=1)
        ttk.Button(ctrl, text="Consultar", command=lambda: self.refrescar(mostrar_errores=True)) \
            .grid(row=0, column=5, rowspan=2, padx=4, ipady=6)
        ttk.Button(ctrl, text="Exportar a Excel…", style="Accent.TButton", command=self.exportar) \
            .grid(row=0, column=6, rowspan=2, padx=4, ipady=6)

        self.cmb_mes.bind("<<ComboboxSelected>>", lambda e: self.refrescar())
        self.spn_anio.configure(command=self.refrescar)
        for w in (self.spn_anio, self.ent_desde, self.ent_hasta):
            w.bind("<Return>", lambda e: self.refrescar(mostrar_errores=True))

        self.lbl_periodo = ttk.Label(self, style="Bold.TLabel")
        self.lbl_periodo.pack(anchor="w", pady=(10, 4))

        cols = (("fecha", "Fecha", 90), ("vn", "Ventas neto", 110), ("vi", "IVA débito", 105),
                ("vt", "Ventas total", 110), ("cn", "Compras neto", 110), ("ci", "IVA crédito", 105),
                ("ct", "Compras total", 110), ("pagar", "IVA a pagar", 115))
        contenedor = ttk.Frame(self)
        contenedor.pack(fill="both", expand=True)
        self.tree = ttk.Treeview(contenedor, columns=[c[0] for c in cols], show="headings")
        for clave, titulo, ancho in cols:
            anc = "w" if clave == "fecha" else "e"
            self.tree.heading(clave, text=titulo, anchor=anc)
            self.tree.column(clave, width=ancho, anchor=anc)
        sb = ttk.Scrollbar(contenedor, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="left", fill="y")
        self.tree.tag_configure("total", font=(tkfont.nametofont("TkDefaultFont").actual("family"), 10, "bold"),
                                background="#E8EEF6")
        self.tree.bind("<Double-1>", self._ir_al_dia)

        self.lbl_total = ttk.Label(self, style="Grande.TLabel")
        self.lbl_total.pack(anchor="e", pady=(10, 0))

        self._modo()

    def _modo(self):
        mes = self.var_modo.get() == "mes"
        self.cmb_mes.configure(state="readonly" if mes else "disabled")
        self.spn_anio.configure(state="normal" if mes else "disabled")
        for w in (self.ent_desde, self.ent_hasta):
            w.configure(state="disabled" if mes else "normal")
        self.refrescar()

    def periodo(self):
        if self.var_modo.get() == "mes":
            mes = self.cmb_mes.current() + 1
            try:
                anio = int(self.var_anio.get())
            except ValueError:
                raise ValueError("Año inválido.")
            desde = date(anio, mes, 1)
            hasta = date(anio, mes, calendar.monthrange(anio, mes)[1])
            return desde, hasta, f"{MESES[mes - 1]} {anio}", f"IVA_{anio}-{mes:02d}"
        desde = parse_fecha(self.var_desde.get())
        hasta = parse_fecha(self.var_hasta.get())
        if desde > hasta:
            raise ValueError("La fecha «Desde» es posterior a «Hasta».")
        return (desde, hasta, f"{fmt_fecha(desde)} al {fmt_fecha(hasta)}",
                f"IVA_{desde.isoformat()}_a_{hasta.isoformat()}")

    def refrescar(self, mostrar_errores=False):
        try:
            desde, hasta, etiqueta, _ = self.periodo()
        except ValueError as e:
            if mostrar_errores:
                messagebox.showerror("Período inválido", str(e), parent=self)
            return
        fmt = self.app.fmt
        self.tree.delete(*self.tree.get_children())
        dias = self.app.db.resumen_diario(desde, hasta)
        for d in dias:
            self.tree.insert("", "end", iid=d["fecha"].isoformat(), values=(
                fmt_fecha(d["fecha"]), fmt(d["ventas_neto"]), fmt(d["ventas_iva"]), fmt(d["ventas_total"]),
                fmt(d["compras_neto"]), fmt(d["compras_iva"]), fmt(d["compras_total"]), fmt(d["iva_pagar"])))
        t = self.app.db.totales(desde, hasta)
        if dias:
            self.tree.insert("", "end", iid="total", tags=("total",), values=(
                "TOTAL", fmt(t["ventas_neto"]), fmt(t["ventas_iva"]), fmt(t["ventas_total"]),
                fmt(t["compras_neto"]), fmt(t["compras_iva"]), fmt(t["compras_total"]), fmt(t["iva_pagar"])))
        self.lbl_periodo.config(
            text=f"{etiqueta} — {len(dias)} día(s) con movimientos · {t['n_ventas']} venta(s) · {t['n_compras']} compra(s)")
        pagar = t["iva_pagar"]
        self.lbl_total.config(
            text=f"IVA a pagar del período: {fmt(pagar)}" if pagar >= 0
            else f"Saldo a favor del período: {fmt(-pagar)}")

    def _ir_al_dia(self, _evento):
        sel = self.tree.selection()
        if sel and sel[0] != "total":
            reg = self.app.tab_registro
            reg.var_fecha.set(fmt_fecha(date.fromisoformat(sel[0])))
            reg.cargar_dia()
            self.master.select(reg)

    def exportar(self):
        try:
            desde, hasta, etiqueta, nombre = self.periodo()
        except ValueError as e:
            messagebox.showerror("Período inválido", str(e), parent=self)
            return
        if not self.app.db.resumen_diario(desde, hasta):
            messagebox.showinfo("Sin datos", "No hay movimientos en el período seleccionado.", parent=self)
            return
        try:
            from exportar import exportar_excel
        except ImportError:
            messagebox.showerror("Falta una dependencia",
                                 "Para exportar a Excel instala openpyxl:\n\npython -m pip install openpyxl",
                                 parent=self)
            return
        ruta = filedialog.asksaveasfilename(
            parent=self, title="Guardar planilla Excel", defaultextension=".xlsx",
            initialfile=nombre + ".xlsx", filetypes=[("Libro de Excel", "*.xlsx")])
        if not ruta:
            return
        try:
            exportar_excel(ruta, self.app.db, desde, hasta, etiqueta, self.app.cfg)
        except PermissionError:
            messagebox.showerror("No se pudo guardar",
                                 "No se pudo escribir el archivo. ¿Está abierto en Excel?", parent=self)
            return
        if messagebox.askyesno("Exportado", f"Planilla guardada en:\n{ruta}\n\n¿Abrirla ahora?", parent=self):
            abrir_archivo(ruta)


# =========================================================================== Configuración

class ConfigTab(ttk.Frame):
    def __init__(self, master, app):
        super().__init__(master, padding=16)
        self.app = app
        cfg = app.cfg

        caja = ttk.LabelFrame(self, text="Cálculo de IVA", padding=12)
        caja.pack(fill="x")
        self.var_tasa = tk.StringVar(value=cfg["tasa_iva"])
        self.var_incluye = tk.BooleanVar(value=cfg["montos_incluyen_iva"] == "1")
        self.var_decimales = tk.StringVar(value=cfg["decimales"])
        ttk.Label(caja, text="Tasa de IVA (%):").grid(row=0, column=0, sticky="w", pady=4)
        ttk.Entry(caja, textvariable=self.var_tasa, width=8, justify="right").grid(row=0, column=1, sticky="w")
        ttk.Checkbutton(caja, text="Por defecto, los montos ingresados incluyen IVA (se desglosa el neto)",
                        variable=self.var_incluye).grid(row=1, column=0, columnspan=3, sticky="w", pady=4)
        ttk.Label(caja, text="Decimales:").grid(row=2, column=0, sticky="w", pady=4)
        ttk.Combobox(caja, textvariable=self.var_decimales, values=("0", "2"), state="readonly", width=5) \
            .grid(row=2, column=1, sticky="w")
        ttk.Label(caja, style="Sub.TLabel",
                  text="Cambiar la tasa no altera los movimientos ya registrados: cada uno guarda la tasa con que se ingresó.") \
            .grid(row=3, column=0, columnspan=3, sticky="w", pady=(8, 0))

        caja2 = ttk.LabelFrame(self, text="Datos para la planilla", padding=12)
        caja2.pack(fill="x", pady=12)
        self.var_simbolo = tk.StringVar(value=cfg["simbolo_moneda"])
        self.var_empresa = tk.StringVar(value=cfg["nombre_empresa"])
        self.var_idf = tk.StringVar(value=cfg["id_fiscal"])
        for i, (texto, var, ancho) in enumerate((
                ("Símbolo de moneda:", self.var_simbolo, 6),
                ("Nombre / razón social:", self.var_empresa, 40),
                ("ID fiscal (RUT, NIT, RFC, CIF…):", self.var_idf, 20))):
            ttk.Label(caja2, text=texto).grid(row=i, column=0, sticky="w", pady=4)
            ttk.Entry(caja2, textvariable=var, width=ancho).grid(row=i, column=1, sticky="w", padx=6)

        ttk.Button(self, text="Guardar configuración", style="Accent.TButton", command=self.guardar) \
            .pack(anchor="w", ipady=4)

        caja3 = ttk.LabelFrame(self, text="Datos", padding=12)
        caja3.pack(fill="x", pady=(16, 0))
        ttk.Label(caja3, text=f"Base de datos: {app.db.ruta}", style="Sub.TLabel").pack(anchor="w")
        ttk.Button(caja3, text="Crear copia de seguridad…", command=self.respaldar).pack(anchor="w", pady=(8, 0))

    def guardar(self):
        try:
            tasa = parse_tasa(self.var_tasa.get())
        except ValueError as e:
            messagebox.showerror("Dato inválido", str(e), parent=self)
            return
        self.app.db.guardar_config({
            "tasa_iva": f"{tasa:g}",
            "montos_incluyen_iva": "1" if self.var_incluye.get() else "0",
            "decimales": self.var_decimales.get(),
            "simbolo_moneda": self.var_simbolo.get().strip(),
            "nombre_empresa": self.var_empresa.get().strip(),
            "id_fiscal": self.var_idf.get().strip(),
        })
        self.var_tasa.set(f"{tasa:g}")
        self.app.config_actualizada()
        messagebox.showinfo("Configuración", "Configuración guardada.", parent=self)

    def respaldar(self):
        ruta = filedialog.asksaveasfilename(
            parent=self, title="Guardar copia de seguridad", defaultextension=".db",
            initialfile=f"respaldo_iva_{date.today().isoformat()}.db", filetypes=[("Base de datos", "*.db")])
        if ruta:
            self.app.db.respaldar(ruta)
            messagebox.showinfo("Copia de seguridad", f"Copia guardada en:\n{ruta}", parent=self)


if __name__ == "__main__":
    if sys.platform == "win32":
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass
    App().mainloop()
