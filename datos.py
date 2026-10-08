"""Persistencia (SQLite), cálculo de IVA y utilidades de formato."""
import re
import sqlite3
from datetime import date, datetime
from decimal import Decimal, ROUND_HALF_UP, InvalidOperation
from pathlib import Path

DB_PATH = Path(__file__).resolve().with_name("datos_iva.db")

CONFIG_POR_DEFECTO = {
    "tasa_iva": "19",
    "montos_incluyen_iva": "1",
    "simbolo_moneda": "$",
    "decimales": "0",
    "nombre_empresa": "",
    "id_fiscal": "",
}

TIPOS = ("venta", "compra")


# --------------------------------------------------------------------------- cálculo

def calcular_iva(monto, tasa, incluye_iva, decimales=0):
    """Devuelve (neto, iva, total) redondeados a `decimales`.

    Si `incluye_iva` es True, `monto` es el total (bruto) y se desglosa;
    si no, `monto` es el neto y se le agrega el IVA.
    """
    q = Decimal(1).scaleb(-decimales)
    m = Decimal(str(monto))
    t = Decimal(str(tasa)) / 100
    if incluye_iva:
        total = m.quantize(q, ROUND_HALF_UP)
        neto = (total / (1 + t)).quantize(q, ROUND_HALF_UP)
        iva = total - neto
    else:
        neto = m.quantize(q, ROUND_HALF_UP)
        iva = (neto * t).quantize(q, ROUND_HALF_UP)
        total = neto + iva
    return float(neto), float(iva), float(total)


# --------------------------------------------------------------------------- formato

def parse_fecha(texto):
    """Acepta dd/mm/aaaa, dd-mm-aaaa o aaaa-mm-dd."""
    texto = texto.strip()
    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%d/%m/%y"):
        try:
            return datetime.strptime(texto, fmt).date()
        except ValueError:
            pass
    raise ValueError(f"Fecha inválida: «{texto}». Usa el formato dd/mm/aaaa.")


def fmt_fecha(f):
    return f.strftime("%d/%m/%Y")


def parse_monto(texto):
    """Interpreta montos como '1.190.000', '1190000', '1.234,56' o '1234.56'."""
    s = re.sub(r"[^\d.,\-]", "", texto.strip())
    if not s:
        raise ValueError("Ingresa un monto.")
    if "." in s and "," in s:
        # el último separador que aparece es el decimal
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        partes = s.split(",")
        if len(partes) > 2 or len(partes[1]) == 3:
            s = s.replace(",", "")  # separador de miles
        else:
            s = s.replace(",", ".")
    elif s.count(".") > 1 or (s.count(".") == 1 and len(s.split(".")[1]) == 3):
        s = s.replace(".", "")  # separador de miles
    try:
        return float(Decimal(s))
    except InvalidOperation:
        raise ValueError(f"Monto inválido: «{texto}».")


def fmt_numero(valor, decimales=0):
    """Formato 1.234.567,89 (punto de miles, coma decimal)."""
    s = f"{valor:,.{decimales}f}"
    return s.replace(",", "_").replace(".", ",").replace("_", ".")


def fmt_monto(valor, simbolo="$", decimales=0):
    signo = "-" if valor < 0 else ""
    return f"{signo}{simbolo}{' ' if simbolo else ''}{fmt_numero(abs(valor), decimales)}"


# --------------------------------------------------------------------------- base de datos

class BaseDatos:
    def __init__(self, ruta=DB_PATH):
        self.ruta = Path(ruta)
        self.con = sqlite3.connect(self.ruta)
        self.con.row_factory = sqlite3.Row
        self._crear_tablas()

    def _crear_tablas(self):
        self.con.executescript("""
            CREATE TABLE IF NOT EXISTS config (
                clave TEXT PRIMARY KEY,
                valor TEXT NOT NULL
            );
            CREATE TABLE IF NOT EXISTS movimientos (
                id          INTEGER PRIMARY KEY AUTOINCREMENT,
                fecha       TEXT NOT NULL,
                tipo        TEXT NOT NULL CHECK (tipo IN ('venta', 'compra')),
                documento   TEXT NOT NULL DEFAULT '',
                descripcion TEXT NOT NULL DEFAULT '',
                monto       REAL NOT NULL,
                incluye_iva INTEGER NOT NULL,
                tasa        REAL NOT NULL,
                neto        REAL NOT NULL,
                iva         REAL NOT NULL,
                total       REAL NOT NULL
            );
            CREATE INDEX IF NOT EXISTS idx_mov_fecha ON movimientos (fecha);
        """)
        self.con.executemany(
            "INSERT OR IGNORE INTO config (clave, valor) VALUES (?, ?)",
            CONFIG_POR_DEFECTO.items(),
        )
        self.con.commit()

    # ---- configuración
    def obtener_config(self):
        return {r["clave"]: r["valor"] for r in self.con.execute("SELECT * FROM config")}

    def guardar_config(self, valores):
        self.con.executemany(
            "INSERT OR REPLACE INTO config (clave, valor) VALUES (?, ?)",
            [(k, str(v)) for k, v in valores.items()],
        )
        self.con.commit()

    # ---- movimientos
    def agregar(self, fecha, tipo, monto, incluye_iva, tasa, decimales, documento="", descripcion=""):
        neto, iva, total = calcular_iva(monto, tasa, incluye_iva, decimales)
        cur = self.con.execute(
            """INSERT INTO movimientos
               (fecha, tipo, documento, descripcion, monto, incluye_iva, tasa, neto, iva, total)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
            (fecha.isoformat(), tipo, documento, descripcion, monto, int(incluye_iva), tasa, neto, iva, total),
        )
        self.con.commit()
        return cur.lastrowid

    def actualizar(self, id_, fecha, tipo, monto, incluye_iva, tasa, decimales, documento="", descripcion=""):
        neto, iva, total = calcular_iva(monto, tasa, incluye_iva, decimales)
        self.con.execute(
            """UPDATE movimientos SET fecha=?, tipo=?, documento=?, descripcion=?, monto=?,
               incluye_iva=?, tasa=?, neto=?, iva=?, total=? WHERE id=?""",
            (fecha.isoformat(), tipo, documento, descripcion, monto, int(incluye_iva), tasa, neto, iva, total, id_),
        )
        self.con.commit()

    def eliminar(self, id_):
        self.con.execute("DELETE FROM movimientos WHERE id=?", (id_,))
        self.con.commit()

    def obtener(self, id_):
        return self.con.execute("SELECT * FROM movimientos WHERE id=?", (id_,)).fetchone()

    def listar(self, desde, hasta, tipo=None):
        sql = "SELECT * FROM movimientos WHERE fecha BETWEEN ? AND ?"
        params = [desde.isoformat(), hasta.isoformat()]
        if tipo:
            sql += " AND tipo = ?"
            params.append(tipo)
        return self.con.execute(sql + " ORDER BY fecha, tipo DESC, id", params).fetchall()

    _SUMAS = """
        COALESCE(SUM(CASE WHEN tipo='venta'  THEN 1     END), 0) AS n_ventas,
        COALESCE(SUM(CASE WHEN tipo='venta'  THEN neto  END), 0) AS ventas_neto,
        COALESCE(SUM(CASE WHEN tipo='venta'  THEN iva   END), 0) AS ventas_iva,
        COALESCE(SUM(CASE WHEN tipo='venta'  THEN total END), 0) AS ventas_total,
        COALESCE(SUM(CASE WHEN tipo='compra' THEN 1     END), 0) AS n_compras,
        COALESCE(SUM(CASE WHEN tipo='compra' THEN neto  END), 0) AS compras_neto,
        COALESCE(SUM(CASE WHEN tipo='compra' THEN iva   END), 0) AS compras_iva,
        COALESCE(SUM(CASE WHEN tipo='compra' THEN total END), 0) AS compras_total
    """

    def resumen_diario(self, desde, hasta):
        """Una fila por día con movimientos dentro del período."""
        filas = self.con.execute(
            f"SELECT fecha, {self._SUMAS} FROM movimientos "
            "WHERE fecha BETWEEN ? AND ? GROUP BY fecha ORDER BY fecha",
            (desde.isoformat(), hasta.isoformat()),
        ).fetchall()
        resultado = []
        for f in filas:
            d = dict(f)
            d["fecha"] = date.fromisoformat(d["fecha"])
            d["iva_pagar"] = d["ventas_iva"] - d["compras_iva"]
            resultado.append(d)
        return resultado

    def totales(self, desde, hasta):
        d = dict(self.con.execute(
            f"SELECT {self._SUMAS} FROM movimientos WHERE fecha BETWEEN ? AND ?",
            (desde.isoformat(), hasta.isoformat()),
        ).fetchone())
        d["iva_pagar"] = d["ventas_iva"] - d["compras_iva"]
        return d

    def respaldar(self, destino):
        copia = sqlite3.connect(destino)
        try:
            self.con.backup(copia)
        finally:
            copia.close()
