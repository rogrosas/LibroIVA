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
AZUL = "#1F4E78"

TIPOS = {
    "venta": {"singular": "venta", "plural": "ventas", "titulo": "Ventas", "iva": "IVA débito",
              "color": COLOR_VENTA, "libro": "Libro de ventas"},
    "compra": {"singular": "compra", "plural": "compras", "titulo": "Compras", "iva": "IVA crédito",
               "color": COLOR_COMPRA, "libro": "Libro de compras"},
}


def parse_tasa(texto):
    try:
        tasa = float(texto.strip().replace("%", "").replace(",", "."))
    except ValueError:
        raise ValueError(f"Tasa de IVA inválida: «{texto}».")
    if not 0 <= tasa <= 100:
        raise ValueError("La tasa de IVA debe estar entre 0 y 100.")
    return tasa


def fmt_tasa(tasa):
    return fmt_numero(tasa, 2).rstrip("0").rstrip(",")


def abrir_archivo(ruta):
    if sys.platform == "win32":
        os.startfile(ruta)
    else:
        subprocess.Popen(["open" if sys.platform == "darwin" else "xdg-open", ruta])


def crear_tabla(master, columnas):
    """Treeview con scrollbar. `columnas` = [(clave, título, ancho, ancla, estirar)]."""
    contenedor = ttk.Frame(master)
    tree = ttk.Treeview(contenedor, columns=[c[0] for c in columnas], show="headings", selectmode="browse")
    for clave, titulo, ancho, ancla, estirar in columnas:
        tree.heading(clave, text=titulo, anchor=ancla)
        tree.column(clave, width=ancho, anchor=ancla, stretch=estirar)
    sb = ttk.Scrollbar(contenedor, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=sb.set)
    tree.pack(side="left", fill="both", expand=True)
    sb.pack(side="left", fill="y")
    familia = tkfont.nametofont("TkDefaultFont").actual("family")
    tree.tag_configure("total", font=(familia, 10, "bold"), background="#E8EEF6")
    tree.tag_configure("venta", foreground=COLOR_VENTA)
    tree.tag_configure("compra", foreground=COLOR_COMPRA)
    return contenedor, tree


def exportar_con_dialogo(parent, nombre, funcion):
    """Pide la ruta, ejecuta `funcion(ruta)` y ofrece abrir el archivo."""
    ruta = filedialog.asksaveasfilename(
        parent=parent, title="Guardar planilla Excel", defaultextension=".xlsx",
        initialfile=nombre + ".xlsx", filetypes=[("Libro de Excel", "*.xlsx")])
    if not ruta:
        return
    try:
        funcion(ruta)
    except PermissionError:
        messagebox.showerror("No se pudo guardar",
                             "No se pudo escribir el archivo. ¿Está abierto en Excel?", parent=parent)
        return
    except ValueError as e:
        messagebox.showinfo("Sin datos", str(e), parent=parent)
        return
    if messagebox.askyesno("Exportado", f"Planilla guardada en:\n{ruta}\n\n¿Abrirla ahora?", parent=parent):
        abrir_archivo(ruta)


def cargar_exportador():
    try:
        import exportar
        return exportar
    except ImportError:
        messagebox.showerror("Falta una dependencia",
                             "Para exportar a Excel instala openpyxl:\n\npython -m pip install openpyxl")
        return None


# =========================================================================== App

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Calculadora de IVA")
        escala = self.winfo_fpixels("1i") / 96
        self.geometry(f"{int(1100 * escala)}x{int(740 * escala)}")
        self.minsize(int(920 * escala), int(620 * escala))

        self.db = BaseDatos()
        self.cfg = self.db.obtener_config()
        self._estilos()

        self.nb = ttk.Notebook(self)
        self.nb.pack(fill="both", expand=True, padx=8, pady=8)
        self.secciones = {tipo: SeccionTipo(self.nb, self, tipo) for tipo in TIPOS}
        self.tab_resumen = ResumenTab(self.nb, self)
        self.tab_config = ConfigTab(self.nb, self)
        self.nb.add(self.secciones["venta"], text="  Ventas  ")
        self.nb.add(self.secciones["compra"], text="  Compras  ")
        self.nb.add(self.tab_resumen, text="  Resumen IVA (ventas − compras)  ")
        self.nb.add(self.tab_config, text="  Configuración  ")

    def _estilos(self):
        estilo = ttk.Style(self)
        if "vista" in estilo.theme_names():
            estilo.theme_use("vista")
        base = tkfont.nametofont("TkDefaultFont")
        base.configure(size=10)
        tkfont.nametofont("TkTextFont").configure(size=10)
        familia = base.actual("family")
        estilo.configure("Treeview", rowheight=int(base.metrics("linespace") * 1.5))
        estilo.configure("Treeview.Heading", font=(familia, 10, "bold"))
        estilo.configure("Bold.TLabel", font=(familia, 10, "bold"))
        estilo.configure("Grande.TLabel", font=(familia, 14, "bold"))
        estilo.configure("Sub.TLabel", foreground="#666666")
        estilo.configure("Accent.TButton", font=(familia, 10, "bold"))

    @property
    def decimales(self):
        return int(self.cfg.get("decimales", 0))

    @property
    def tasa(self):
        return float(self.cfg.get("tasa_iva", 19))

    def fmt(self, valor):
        return fmt_monto(valor, self.cfg.get("simbolo_moneda", "$"), self.decimales)

    def datos_cambiados(self):
        for s in self.secciones.values():
            s.libro.refrescar()
        self.tab_resumen.refrescar()

    def config_actualizada(self):
        self.cfg = self.db.obtener_config()
        for s in self.secciones.values():
            s.registro.aplicar_config()
        self.datos_cambiados()

    def ir_a_dia(self, tipo, fecha, editar_id=None):
        seccion = self.secciones[tipo]
        self.nb.select(seccion)
        seccion.nb.select(seccion.registro)
        seccion.registro.ir_a(fecha, editar_id)


# =========================================================================== Selector de período

class SelectorPeriodo(ttk.LabelFrame):
    """Elige un mes o un rango de fechas. Llama a `al_cambiar()` cuando cambia."""

    def __init__(self, master, al_cambiar, botones=()):
        super().__init__(master, text="Período", padding=10)
        self.al_cambiar = al_cambiar
        hoy = date.today()

        self.var_modo = tk.StringVar(value="mes")
        ttk.Radiobutton(self, text="Por mes", value="mes", variable=self.var_modo, command=self._modo) \
            .grid(row=0, column=0, sticky="w", pady=3)
        self.cmb_mes = ttk.Combobox(self, values=MESES, state="readonly", width=12)
        self.cmb_mes.current(hoy.month - 1)
        self.cmb_mes.grid(row=0, column=1, padx=6, sticky="w")
        self.var_anio = tk.StringVar(value=str(hoy.year))
        self.spn_anio = ttk.Spinbox(self, from_=2000, to=2100, textvariable=self.var_anio, width=6,
                                    command=self.al_cambiar)
        self.spn_anio.grid(row=0, column=2, sticky="w")

        ttk.Radiobutton(self, text="Rango de fechas", value="rango", variable=self.var_modo, command=self._modo) \
            .grid(row=1, column=0, sticky="w", pady=3)
        self.var_desde = tk.StringVar(value=fmt_fecha(hoy.replace(day=1)))
        self.var_hasta = tk.StringVar(value=fmt_fecha(hoy))
        rango = ttk.Frame(self)
        rango.grid(row=1, column=1, columnspan=3, sticky="w", padx=6)
        ttk.Label(rango, text="Desde").pack(side="left")
        self.ent_desde = ttk.Entry(rango, textvariable=self.var_desde, width=12, justify="center")
        self.ent_desde.pack(side="left", padx=(4, 10))
        ttk.Label(rango, text="Hasta").pack(side="left")
        self.ent_hasta = ttk.Entry(rango, textvariable=self.var_hasta, width=12, justify="center")
        self.ent_hasta.pack(side="left", padx=4)

        self.columnconfigure(4, weight=1)
        botones = [("Consultar", lambda: self.al_cambiar(mostrar_errores=True), False), *botones]
        for i, (texto, comando, destacado) in enumerate(botones, start=5):
            ttk.Button(self, text=texto, command=comando, style="Accent.TButton" if destacado else "TButton") \
                .grid(row=0, column=i, rowspan=2, padx=4, ipady=6)

        self.cmb_mes.bind("<<ComboboxSelected>>", lambda e: self.al_cambiar())
        for w in (self.spn_anio, self.ent_desde, self.ent_hasta):
            w.bind("<Return>", lambda e: self.al_cambiar(mostrar_errores=True))
        self._modo(notificar=False)

    def _modo(self, notificar=True):
        mes = self.var_modo.get() == "mes"
        self.cmb_mes.configure(state="readonly" if mes else "disabled")
        self.spn_anio.configure(state="normal" if mes else "disabled")
        for w in (self.ent_desde, self.ent_hasta):
            w.configure(state="disabled" if mes else "normal")
        if notificar:
            self.al_cambiar()

    def periodo(self):
        """Devuelve (desde, hasta, etiqueta, sufijo_archivo). Lanza ValueError si es inválido."""
        if self.var_modo.get() == "mes":
            mes = self.cmb_mes.current() + 1
            try:
                anio = int(self.var_anio.get())
            except ValueError:
                raise ValueError("Año inválido.")
            desde = date(anio, mes, 1)
            hasta = date(anio, mes, calendar.monthrange(anio, mes)[1])
            return desde, hasta, f"{MESES[mes - 1]} {anio}", f"{anio}-{mes:02d}"
        desde = parse_fecha(self.var_desde.get())
        hasta = parse_fecha(self.var_hasta.get())
        if desde > hasta:
            raise ValueError("La fecha «Desde» es posterior a «Hasta».")
        return (desde, hasta, f"{fmt_fecha(desde)} al {fmt_fecha(hasta)}",
                f"{desde.isoformat()}_a_{hasta.isoformat()}")


# =========================================================================== Sección Ventas / Compras

class SeccionTipo(ttk.Frame):
    def __init__(self, master, app, tipo):
        super().__init__(master, padding=(4, 6, 4, 4))
        self.nb = ttk.Notebook(self)
        self.nb.pack(fill="both", expand=True)
        self.registro = RegistroTab(self.nb, app, tipo)
        self.libro = LibroTab(self.nb, app, tipo)
        self.nb.add(self.registro, text="  Registro diario  ")
        self.nb.add(self.libro, text=f"  {TIPOS[tipo]['libro']} y exportación  ")


class RegistroTab(ttk.Frame):
    def __init__(self, master, app, tipo):
        super().__init__(master, padding=12)
        self.app = app
        self.tipo = tipo
        self.t = TIPOS[tipo]
        self.editando_id = None

        # ---- barra de fecha
        barra = ttk.Frame(self)
        barra.pack(fill="x")
        ttk.Label(barra, text=self.t["titulo"], style="Grande.TLabel", foreground=self.t["color"]) \
            .pack(side="left", padx=(0, 20))
        ttk.Label(barra, text="Fecha:", style="Bold.TLabel").pack(side="left")
        self.var_fecha = tk.StringVar(value=fmt_fecha(date.today()))
        ent_fecha = ttk.Entry(barra, textvariable=self.var_fecha, width=12, justify="center")
        ent_fecha.pack(side="left", padx=(6, 4))
        ent_fecha.bind("<Return>", lambda e: self.cargar_dia())
        ent_fecha.bind("<FocusOut>", lambda e: self.cargar_dia(silencioso=True))
        ttk.Button(barra, text="◀", width=3, command=lambda: self.mover_dia(-1)).pack(side="left")
        ttk.Button(barra, text="Hoy", width=6, command=lambda: self.ir_a(date.today())).pack(side="left", padx=2)
        ttk.Button(barra, text="▶", width=3, command=lambda: self.mover_dia(1)).pack(side="left")
        self.lbl_dia = ttk.Label(barra, style="Sub.TLabel")
        self.lbl_dia.pack(side="left", padx=12)

        cuerpo = ttk.Frame(self)
        cuerpo.pack(fill="both", expand=True, pady=(10, 0))

        # ---- formulario
        self.form = ttk.LabelFrame(cuerpo, padding=12)
        self.form.pack(side="left", fill="y")
        f = self.form
        self.var_monto = tk.StringVar()
        self.var_incluye = tk.BooleanVar()
        self.var_tasa = tk.StringVar()
        self.var_doc = tk.StringVar()
        self.var_desc = tk.StringVar()

        ttk.Label(f, text="Monto:").grid(row=0, column=0, sticky="w", pady=3)
        self.ent_monto = ttk.Entry(f, textvariable=self.var_monto, width=22, justify="right")
        self.ent_monto.grid(row=0, column=1, sticky="ew", pady=3)
        ttk.Checkbutton(f, text="El monto incluye IVA (total)", variable=self.var_incluye) \
            .grid(row=1, column=1, sticky="w", pady=(0, 3))
        ttk.Label(f, text="Tasa IVA %:").grid(row=2, column=0, sticky="w", pady=3)
        ttk.Entry(f, textvariable=self.var_tasa, width=8, justify="right").grid(row=2, column=1, sticky="w", pady=3)
        ttk.Label(f, text="N° documento:").grid(row=3, column=0, sticky="w", pady=3)
        ttk.Entry(f, textvariable=self.var_doc, width=22).grid(row=3, column=1, sticky="ew", pady=3)
        ttk.Label(f, text="Descripción:").grid(row=4, column=0, sticky="w", pady=3)
        ttk.Entry(f, textvariable=self.var_desc, width=22).grid(row=4, column=1, sticky="ew", pady=3)

        prev = ttk.Frame(f, padding=(0, 10))
        prev.grid(row=5, column=0, columnspan=2, sticky="ew")
        self.lbl_prev = {}
        for i, (clave, texto) in enumerate((("neto", "Neto"), ("iva", self.t["iva"]), ("total", "Total"))):
            ttk.Label(prev, text=texto + ":", style="Sub.TLabel").grid(row=i, column=0, sticky="w")
            self.lbl_prev[clave] = ttk.Label(prev, text="—", style="Bold.TLabel")
            self.lbl_prev[clave].grid(row=i, column=1, sticky="e", padx=(10, 0))
        prev.columnconfigure(1, weight=1)

        botones = ttk.Frame(f)
        botones.grid(row=6, column=0, columnspan=2, sticky="ew")
        self.btn_guardar = ttk.Button(botones, style="Accent.TButton", command=self.guardar)
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
        tabla, self.tree = crear_tabla(derecha, [
            ("doc", "N° doc.", 100, "w", False), ("desc", "Descripción", 220, "w", True),
            ("tasa", "Tasa", 60, "e", False), ("neto", "Neto", 115, "e", False),
            ("iva", self.t["iva"], 110, "e", False), ("total", "Total", 115, "e", False)])
        tabla.pack(fill="both", expand=True)
        self.tree.bind("<Double-1>", lambda e: self.editar())
        self.tree.bind("<Delete>", lambda e: self.eliminar())

        acciones = ttk.Frame(derecha)
        acciones.pack(fill="x", pady=(6, 0))
        ttk.Button(acciones, text="Editar", command=self.editar).pack(side="left")
        ttk.Button(acciones, text="Eliminar", command=self.eliminar).pack(side="left", padx=6)
        ttk.Label(acciones, text="Doble clic para editar · Supr para eliminar", style="Sub.TLabel").pack(side="right")

        # ---- resumen del día y acumulado del mes
        res = ttk.Frame(self)
        res.pack(fill="x", pady=(10, 0))
        self.lbl_res = {}
        for col, (bloque, titulo) in enumerate((("dia", "Totales del día"), ("mes", "Acumulado del mes"))):
            caja = ttk.LabelFrame(res, text=titulo, padding=12)
            caja.grid(row=0, column=col, sticky="nsew", padx=(0, 10) if col == 0 else 0)
            res.columnconfigure(col, weight=1)
            for i, (clave, texto) in enumerate((("n", f"N° de {self.t['plural']}"), ("neto", "Neto"),
                                                ("iva", self.t["iva"]), ("total", "Total"))):
                ttk.Label(caja, text=texto + ":", style="Sub.TLabel").grid(row=i, column=0, sticky="w")
                lbl = ttk.Label(caja, style="Bold.TLabel" if clave == "iva" else "TLabel",
                                foreground=self.t["color"] if clave == "iva" else "")
                lbl.grid(row=i, column=1, sticky="e", padx=(16, 0))
                self.lbl_res[(bloque, clave)] = lbl

        self.limpiar_form()

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
            self.ir_a(f + timedelta(days=delta))

    def ir_a(self, fecha, editar_id=None):
        self.var_fecha.set(fmt_fecha(fecha))
        self.cargar_dia()
        if editar_id is not None and self.tree.exists(str(editar_id)):
            self.tree.selection_set(str(editar_id))
            self.tree.see(str(editar_id))
            self.editar()

    def cargar_dia(self, silencioso=False):
        f = self.fecha_actual(silencioso)
        if f is None:
            return
        self.var_fecha.set(fmt_fecha(f))
        self.lbl_dia.config(text=f"{DIAS[f.weekday()]} {f.day} de {MESES[f.month - 1].lower()} de {f.year}")

        fmt = self.app.fmt
        self.tree.delete(*self.tree.get_children())
        for m in self.app.db.listar(f, f, self.tipo):
            self.tree.insert("", "end", iid=str(m["id"]), values=(
                m["documento"], m["descripcion"], fmt_tasa(m["tasa"]) + "%",
                fmt(m["neto"]), fmt(m["iva"]), fmt(m["total"])))

        inicio_mes = f.replace(day=1)
        fin_mes = f.replace(day=calendar.monthrange(f.year, f.month)[1])
        pref = self.t["plural"]
        for bloque, (desde, hasta) in (("dia", (f, f)), ("mes", (inicio_mes, fin_mes))):
            t = self.app.db.totales(desde, hasta)
            self.lbl_res[(bloque, "n")].config(text=str(t[f"n_{pref}"]))
            for clave in ("neto", "iva", "total"):
                self.lbl_res[(bloque, clave)].config(text=fmt(t[f"{pref}_{clave}"]))

    # ---- formulario
    def aplicar_config(self):
        if self.editando_id is None:
            self.var_tasa.set(fmt_tasa(self.app.tasa))
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
            valores = (None,) * 3
        for clave, valor in zip(("neto", "iva", "total"), valores):
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
        args = dict(fecha=f, tipo=self.tipo, monto=monto, incluye_iva=self.var_incluye.get(),
                    tasa=tasa, decimales=self.app.decimales,
                    documento=self.var_doc.get().strip(), descripcion=self.var_desc.get().strip())
        if self.editando_id is None:
            self.app.db.agregar(**args)
        else:
            self.app.db.actualizar(self.editando_id, **args)
        self.limpiar_form()
        self.app.datos_cambiados()

    def limpiar_form(self):
        self.editando_id = None
        self.form.config(text=f"Nueva {self.t['singular']}")
        self.btn_guardar.config(text=f"Agregar {self.t['singular']}")
        self.btn_cancelar.pack_forget()
        self.var_monto.set("")
        self.var_doc.set("")
        self.var_desc.set("")
        self.aplicar_config()
        self.ent_monto.focus_set()

    def _seleccion(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Sin selección", "Selecciona un registro de la lista.", parent=self)
            return None
        return int(sel[0])

    def editar(self):
        id_ = self._seleccion()
        if id_ is None:
            return
        m = self.app.db.obtener(id_)
        self.editando_id = id_
        self.form.config(text=f"Editando {self.t['singular']} #{id_}")
        self.btn_guardar.config(text="Guardar cambios")
        self.btn_cancelar.pack(side="left", padx=(6, 0))
        self.var_incluye.set(bool(m["incluye_iva"]))
        self.var_tasa.set(fmt_tasa(m["tasa"]))
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
            self.app.datos_cambiados()


class LibroTab(ttk.Frame):
    """Listado de ventas o compras de un período, con exportación exclusiva."""

    def __init__(self, master, app, tipo):
        super().__init__(master, padding=12)
        self.app = app
        self.tipo = tipo
        self.t = TIPOS[tipo]

        self.selector = SelectorPeriodo(self, self.refrescar, botones=[
            (f"Exportar {self.t['plural']} a Excel…", self.exportar, True)])
        self.selector.pack(fill="x")

        self.lbl_periodo = ttk.Label(self, style="Bold.TLabel")
        self.lbl_periodo.pack(anchor="w", pady=(10, 4))

        self.var_vista = tk.StringVar(value="detalle")
        vista = ttk.Frame(self)
        vista.pack(fill="x", pady=(0, 4))
        ttk.Label(vista, text="Ver:").pack(side="left")
        ttk.Radiobutton(vista, text=f"Cada {self.t['singular']}", value="detalle", variable=self.var_vista,
                        command=self.refrescar).pack(side="left", padx=8)
        ttk.Radiobutton(vista, text="Totales por día", value="dia", variable=self.var_vista,
                        command=self.refrescar).pack(side="left")
        ttk.Label(vista, text="Doble clic para ir al registro", style="Sub.TLabel").pack(side="right")

        tabla, self.tree = crear_tabla(self, [
            ("fecha", "Fecha", 95, "w", False), ("doc", "N° doc. / cantidad", 130, "w", False),
            ("desc", "Descripción", 220, "w", True), ("neto", "Neto", 120, "e", False),
            ("iva", self.t["iva"], 115, "e", False), ("total", "Total", 120, "e", False)])
        tabla.pack(fill="both", expand=True)
        self.tree.bind("<Double-1>", self._ir_al_registro)

        self.lbl_total = ttk.Label(self, style="Grande.TLabel", foreground=self.t["color"])
        self.lbl_total.pack(anchor="e", pady=(10, 0))
        self.refrescar()

    def refrescar(self, mostrar_errores=False):
        try:
            desde, hasta, etiqueta, _ = self.selector.periodo()
        except ValueError as e:
            if mostrar_errores:
                messagebox.showerror("Período inválido", str(e), parent=self)
            return
        fmt = self.app.fmt
        pref = self.t["plural"]
        self.tree.delete(*self.tree.get_children())
        if self.var_vista.get() == "detalle":
            for m in self.app.db.listar(desde, hasta, self.tipo):
                fecha = date.fromisoformat(m["fecha"])
                self.tree.insert("", "end", iid=f"m{m['id']}", values=(
                    fmt_fecha(fecha), m["documento"], m["descripcion"],
                    fmt(m["neto"]), fmt(m["iva"]), fmt(m["total"])))
        else:
            for d in self.app.db.resumen_diario(desde, hasta):
                if d[f"n_{pref}"]:
                    self.tree.insert("", "end", iid=f"d{d['fecha'].isoformat()}", values=(
                        fmt_fecha(d["fecha"]), f"{d[f'n_{pref}']} {pref}", "",
                        fmt(d[f"{pref}_neto"]), fmt(d[f"{pref}_iva"]), fmt(d[f"{pref}_total"])))
        t = self.app.db.totales(desde, hasta)
        if t[f"n_{pref}"]:
            self.tree.insert("", "end", iid="total", tags=("total",), values=(
                "TOTAL", f"{t[f'n_{pref}']} {pref}", "",
                fmt(t[f"{pref}_neto"]), fmt(t[f"{pref}_iva"]), fmt(t[f"{pref}_total"])))
        self.lbl_periodo.config(text=f"{self.t['libro']} — {etiqueta}")
        self.lbl_total.config(text=f"{self.t['iva']} del período: {fmt(t[f'{pref}_iva'])}")

    def _ir_al_registro(self, _evento):
        sel = self.tree.selection()
        if not sel or sel[0] == "total":
            return
        if sel[0].startswith("m"):
            m = self.app.db.obtener(int(sel[0][1:]))
            self.app.ir_a_dia(self.tipo, date.fromisoformat(m["fecha"]), m["id"])
        else:
            self.app.ir_a_dia(self.tipo, date.fromisoformat(sel[0][1:]))

    def exportar(self):
        try:
            desde, hasta, etiqueta, sufijo = self.selector.periodo()
        except ValueError as e:
            messagebox.showerror("Período inválido", str(e), parent=self)
            return
        exp = cargar_exportador()
        if exp:
            exportar_con_dialogo(
                self, f"{self.t['titulo']}_{sufijo}",
                lambda ruta: exp.exportar_libro(ruta, self.app.db, self.tipo, desde, hasta, etiqueta, self.app.cfg))


# =========================================================================== Resumen combinado

class ResumenTab(ttk.Frame):
    def __init__(self, master, app):
        super().__init__(master, padding=12)
        self.app = app

        self.selector = SelectorPeriodo(self, self.refrescar, botones=[
            ("Exportar resumen a Excel…", self.exportar, True)])
        self.selector.pack(fill="x")

        self.lbl_periodo = ttk.Label(self, style="Bold.TLabel")
        self.lbl_periodo.pack(anchor="w", pady=(10, 4))

        tabla, self.tree = crear_tabla(self, [
            ("fecha", "Fecha", 90, "w", False), ("vn", "Ventas neto", 110, "e", True),
            ("vi", "IVA débito", 105, "e", True), ("vt", "Ventas total", 110, "e", True),
            ("cn", "Compras neto", 110, "e", True), ("ci", "IVA crédito", 105, "e", True),
            ("ct", "Compras total", 110, "e", True), ("pagar", "IVA a pagar", 115, "e", True)])
        tabla.pack(fill="both", expand=True)
        self.tree.bind("<Double-1>", self._ir_al_dia)

        pie = ttk.Frame(self)
        pie.pack(fill="x", pady=(10, 0))
        self.lbl_debito = ttk.Label(pie, foreground=COLOR_VENTA, style="Bold.TLabel")
        self.lbl_debito.pack(side="left")
        self.lbl_credito = ttk.Label(pie, foreground=COLOR_COMPRA, style="Bold.TLabel")
        self.lbl_credito.pack(side="left", padx=20)
        self.lbl_total = ttk.Label(pie, style="Grande.TLabel", foreground=AZUL)
        self.lbl_total.pack(side="right")
        self.refrescar()

    def refrescar(self, mostrar_errores=False):
        try:
            desde, hasta, etiqueta, _ = self.selector.periodo()
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
        self.lbl_debito.config(text=f"IVA débito: {fmt(t['ventas_iva'])}")
        self.lbl_credito.config(text=f"IVA crédito: {fmt(t['compras_iva'])}")
        pagar = t["iva_pagar"]
        self.lbl_total.config(
            text=f"IVA a pagar: {fmt(pagar)}" if pagar >= 0 else f"Saldo a favor: {fmt(-pagar)}")

    def _ir_al_dia(self, _evento):
        sel = self.tree.selection()
        if not sel or sel[0] == "total":
            return
        fecha = date.fromisoformat(sel[0])
        hay_ventas = self.app.db.totales(fecha, fecha)["n_ventas"] > 0
        self.app.ir_a_dia("venta" if hay_ventas else "compra", fecha)

    def exportar(self):
        try:
            desde, hasta, etiqueta, sufijo = self.selector.periodo()
        except ValueError as e:
            messagebox.showerror("Período inválido", str(e), parent=self)
            return
        exp = cargar_exportador()
        if exp:
            exportar_con_dialogo(
                self, f"IVA_{sufijo}",
                lambda ruta: exp.exportar_excel(ruta, self.app.db, desde, hasta, etiqueta, self.app.cfg))


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
                  text="Cambiar la tasa no altera los registros ya ingresados: cada uno guarda la tasa con que se ingresó.") \
            .grid(row=3, column=0, columnspan=3, sticky="w", pady=(8, 0))

        caja2 = ttk.LabelFrame(self, text="Datos para las planillas", padding=12)
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
