#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
CRM INTEGRAL | GROUP EJ MORALES S.U.R.L. — Aplicación de escritorio
===================================================================
Roles:
  - ESPECIALISTA (comercial): trabaja su libro individual, registra clientes,
    oportunidades, ofertas, etc. y genera un PAQUETE DIARIO para enviar a jefatura.
  - JEFATURA: importa los paquetes de los especialistas, consolida en el libro
    maestro, resuelve conflictos y revisa/aprueba ofertas.

Funciona sin internet y directamente sobre los libros Excel existentes,
manteniendo nombres de hojas y encabezados (compatible con Power Query).
"""

import os
import sys
import json
import copy
import shutil
import datetime
import traceback

# --- openpyxl incluida localmente (carpeta libs) --------------------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
LIBS = os.path.join(BASE_DIR, "libs")
if os.path.isdir(LIBS) and LIBS not in sys.path:
    sys.path.insert(0, LIBS)

import openpyxl
from openpyxl.styles import Font, PatternFill

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

APP_NAME = "CRM Group EJ Morales"
APP_TITLE = "CRM INTEGRAL | GROUP EJ MORALES S.U.R.L."

# ---------------------------------------------------------------------------
# Configuración del modelo
# ---------------------------------------------------------------------------

# Hojas de datos (idénticas en libro individual y maestro)
DATA_SHEETS = [
    "Equipo", "Clientes", "Proveedores", "Prospectos", "Actividades", "Ofertas",
    "Compras", "Importaciones", "Participantes_importacion", "Embarques",
    "Contenedores", "Contenido_contenedores", "Documentos_importacion",
    "Costos_importacion", "Contratos", "Ventas", "Logistica", "Facturas_Cobros",
    "Incidencias", "Rendimiento",
]

# Hoja y columna que contiene el identificador único de cada fila
ID_FIELD = {
    "Equipo": "ID_Usuario", "Clientes": "ID_Cliente", "Proveedores": "ID_Proveedor",
    "Prospectos": "ID_Oportunidad", "Actividades": "ID_Actividad", "Ofertas": "ID_Oferta",
    "Compras": "ID_Compra", "Importaciones": "ID_Importacion",
    "Participantes_importacion": "ID_Participante", "Embarques": "ID_Embarque",
    "Contenedores": "ID_Contenedor", "Contenido_contenedores": "ID_Linea",
    "Documentos_importacion": "ID_Documento", "Costos_importacion": "ID_Costo",
    "Contratos": "ID_Contrato", "Ventas": "ID_Operacion", "Logistica": "ID_Entrega",
    "Facturas_Cobros": "ID_Registro", "Incidencias": "ID_Incidencia",
    "Rendimiento": None,
}

# Prefijo para generar IDs automáticos
ID_PREFIX = {
    "Equipo": "USR", "Clientes": "CLI", "Proveedores": "PRO", "Prospectos": "OPO",
    "Actividades": "ACT", "Ofertas": "OFE", "Compras": "CMP", "Importaciones": "IMP",
    "Participantes_importacion": "PAR", "Embarques": "EMB", "Contenedores": "CNT",
    "Contenido_contenedores": "LIN", "Documentos_importacion": "DOC",
    "Costos_importacion": "CST", "Contratos": "CTR", "Ventas": "VTA",
    "Logistica": "ENT", "Facturas_Cobros": "REG", "Incidencias": "INC",
}

# Columnas con lista de valores (tomada de la hoja CATALOGOS)
CATALOG_MAP = {
    ("Prospectos", "Etapa"): "Etapa",
    ("Ofertas", "Estado_revision"): "Estado_revision",
    ("Importaciones", "Estado"): "Estado_importacion",
    ("Embarques", "Estado"): "Estado_embarque",
    ("Contenedores", "Estado"): "Estado_contenedor",
    ("Incidencias", "Prioridad"): "Prioridad",
}

# Estados de oferta que requieren acción de jefatura
OFERTAS_PENDIENTES = [
    "Borrador", "Pendiente de revisión", "En revisión",
    "Modificaciones solicitadas", "Pendiente de autorización de Dirección",
]

TODAY = datetime.date.today

# ---------------------------------------------------------------------------
# Utilidades
# ---------------------------------------------------------------------------

def now_iso():
    return datetime.datetime.now().strftime("%Y-%m-%dT%H:%M:%S")

def today_str():
    return datetime.date.today().strftime("%Y-%m-%d")

def to_text(v):
    if v is None:
        return ""
    if isinstance(v, datetime.datetime):
        return v.strftime("%Y-%m-%d %H:%M")
    if isinstance(v, datetime.date):
        return v.strftime("%Y-%m-%d")
    return str(v)

def coerce_cell(sheet_headers, col, value):
    """Convierte texto del formulario/JSON a un tipo adecuado para Excel."""
    if value is None:
        return None
    if not isinstance(value, str):
        return value
    s = value.strip()
    if s == "":
        return None
    is_date_col = col.startswith("Fecha") or col in ("ETA", "ATA")
    if is_date_col:
        for fmt in ("%Y-%m-%d %H:%M", "%Y-%m-%d", "%d/%m/%Y", "%d/%m/%Y %H:%M"):
            try:
                return datetime.datetime.strptime(s, fmt)
            except ValueError:
                continue
        return s
    if col.endswith("_pct") or col in ("Probabilidad_pct", "Cumplimiento_pct"):
        try:
            return float(s)
        except ValueError:
            return s
    try:
        f = float(s)
        return int(f) if f.is_integer() and "." not in s else f
    except ValueError:
        return s

def rows_differ(a, b, ignore=("Fecha_modificacion",)):
    """Campos que difieren entre dos filas (dicts)."""
    diffs = {}
    for k in set(a) | set(b):
        if k in ignore:
            continue
        va, vb = to_text(a.get(k, "")), to_text(b.get(k, ""))
        if va != vb:
            diffs[k] = (va, vb)
    return diffs

def json_default(o):
    if isinstance(o, (datetime.datetime, datetime.date)):
        return o.isoformat()
    return str(o)

# ---------------------------------------------------------------------------
# Modelo: libro CRM (envoltorio del Excel)
# ---------------------------------------------------------------------------

class CRMBook:
    def __init__(self, path):
        self.path = path
        self.hojas = {}          # hoja -> {"headers": [...], "rows": [dict,...]}
        self.catalogos = {}      # tipo -> [valores]
        self.load()

    # -- carga ---------------------------------------------------------------
    def load(self):
        self.hojas = {}
        self.catalogos = {}
        wb = openpyxl.load_workbook(self.path, data_only=False)
        for name in DATA_SHEETS + ["LOG_CONSOLIDACION"]:
            if name not in wb.sheetnames:
                continue
            ws = wb[name]
            rows = list(ws.iter_rows(values_only=True))
            if not rows:
                continue
            headers = [str(h) if h is not None else f"Col{i+1}" for i, h in enumerate(rows[0])]
            data = []
            for r in rows[1:]:
                if all(v is None or str(v).strip() == "" for v in r):
                    continue
                row = {}
                for i, h in enumerate(headers):
                    v = r[i] if i < len(r) else None
                    row[h] = to_text(v) if v is not None else ""
                data.append(row)
            self.hojas[name] = {"headers": headers, "rows": data}
        if "CATALOGOS" in wb.sheetnames:
            ws = wb["CATALOGOS"]
            for t, v in ws.iter_rows(min_row=2, values_only=True):
                if t and v:
                    self.catalogos.setdefault(str(t), []).append(str(v))
        wb.close()

    # -- guardado --------------------------------------------------------------
    def save(self):
        wb = openpyxl.load_workbook(self.path, data_only=False)
        for name, info in self.hojas.items():
            if name not in wb.sheetnames:
                ws = wb.create_sheet(name)
                for c, h in enumerate(info["headers"], 1):
                    ws.cell(1, c, h)
                    ws.cell(1, c).font = Font(bold=True, color="FFFFFF")
                    ws.cell(1, c).fill = PatternFill("solid", fgColor="1F3A5F")
            else:
                ws = wb[name]
                if ws.max_row > 1:
                    ws.delete_rows(2, ws.max_row - 1)
                if ws.max_row == 0 or ws.cell(1, 1).value is None:
                    for c, h in enumerate(info["headers"], 1):
                        ws.cell(1, c, h)
                        ws.cell(1, c).font = Font(bold=True, color="FFFFFF")
                        ws.cell(1, c).fill = PatternFill("solid", fgColor="1F3A5F")
            for row in info["rows"]:
                vals = []
                for h in info["headers"]:
                    v = row.get(h, "")
                    if isinstance(v, str) and v.strip() == "":
                        vals.append(None)
                    else:
                        vals.append(v)
                ws.append(vals)
        wb.save(self.path)
        wb.close()

    # -- operaciones ------------------------------------------------------------
    def sheet(self, name):
        if name not in self.hojas:
            self.hojas[name] = {"headers": [], "rows": []}
        return self.hojas[name]

    def next_id(self, sheet_name):
        info = self.sheet(sheet_name)
        idf = ID_FIELD.get(sheet_name)
        prefix = ID_PREFIX.get(sheet_name)
        if not idf or not prefix:
            return ""
        max_n = 0
        for r in info["rows"]:
            v = str(r.get(idf, ""))
            if v.startswith(prefix + "-"):
                try:
                    max_n = max(max_n, int(v.split("-")[1]))
                except (ValueError, IndexError):
                    pass
        return f"{prefix}-{max_n + 1:04d}"

    def find_row(self, sheet_name, id_value):
        idf = ID_FIELD.get(sheet_name)
        if not idf:
            return None, -1
        for i, r in enumerate(self.sheet(sheet_name)["rows"]):
            if str(r.get(idf, "")) == str(id_value):
                return r, i
        return None, -1

    def upsert(self, sheet_name, row, user=None):
        """Inserta o actualiza (por ID). Marca Fecha_modificacion."""
        info = self.sheet(sheet_name)
        if "Fecha_modificacion" in info["headers"] and user is not False:
            row["Fecha_modificacion"] = now_iso()[:16].replace("T", " ")
        idf = ID_FIELD.get(sheet_name)
        if idf:
            existing, idx = self.find_row(sheet_name, row.get(idf, ""))
            if existing is not None:
                info["rows"][idx] = row
                return "actualizada"
            info["rows"].append(row)
            return "nueva"
        info["rows"].append(row)
        return "nueva"

    def delete_row(self, sheet_name, id_value):
        _, idx = self.find_row(sheet_name, id_value)
        if idx >= 0:
            self.sheet(sheet_name)["rows"].pop(idx)
            return True
        return False


def center_over(parent, dlg, w, h):
    parent.update_idletasks()
    px, py = parent.winfo_rootx(), parent.winfo_rooty()
    pw, ph = parent.winfo_width(), parent.winfo_height()
    x = px + max(0, (pw - w) // 2)
    y = py + max(0, (ph - h) // 2)
    dlg.geometry(f"{w}x{h}+{x}+{y}")


# ---------------------------------------------------------------------------
# Exportación a Excel
# ---------------------------------------------------------------------------

def export_sheet_to_xlsx(path, sheet_name, headers, rows, title=None):
    """Exporta una hoja del CRM a un archivo Excel independiente y formateado."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet_name[:31]
    if title:
        ws.cell(1, 1, title)
        ws.cell(1, 1).font = Font(bold=True, size=13, color="1F3A5F")
        ws.cell(2, 1, f"Exportado: {now_iso()[:16].replace('T', ' ')}")
        ws.cell(2, 1).font = Font(italic=True, size=9, color="777777")
        head_row = 4
    else:
        head_row = 1
    for c, h in enumerate(headers, 1):
        cell = ws.cell(head_row, c, h)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F3A5F")
    for r, row in enumerate(rows, head_row + 1):
        for c, h in enumerate(headers, 1):
            v = row.get(h, "")
            ws.cell(r, c, v if v != "" else None)
    # anchos, filtro y panel congelado
    for c, h in enumerate(headers, 1):
        ancho = max(12, min(38, len(h) + 4))
        for r in range(head_row + 1, head_row + 1 + len(rows)):
            v = to_text(ws.cell(r, c).value)
            if v:
                ancho = max(ancho, min(42, len(v) + 2))
        ws.column_dimensions[openpyxl.utils.get_column_letter(c)].width = ancho
    if rows:
        ws.auto_filter.ref = (f"A{head_row}:"
                              f"{openpyxl.utils.get_column_letter(len(headers))}{head_row + len(rows)}")
    ws.freeze_panes = ws.cell(head_row + 1, 1)
    wb.save(path)
    wb.close()

# ---------------------------------------------------------------------------
# Paquetes (especialista -> jefatura y retroalimentación jefatura -> especialista)
# ---------------------------------------------------------------------------

def snapshot_path(book_path):
    return book_path + ".snapshot.json"

def save_snapshot(book_path, hojas):
    snap = {"generado": now_iso(), "hojas": {k: v["rows"] for k, v in hojas.items()}}
    with open(snapshot_path(book_path), "w", encoding="utf-8") as f:
        json.dump(snap, f, ensure_ascii=False, default=json_default)

def load_snapshot(book_path):
    p = snapshot_path(book_path)
    if not os.path.exists(p):
        return None
    with open(p, encoding="utf-8") as f:
        return json.load(f)

def build_package(book, especialista):
    """Paquete completo del especialista + resumen de cambios desde el último envío."""
    snap = load_snapshot(book.path)
    cambios = {}
    hojas_out = {}
    for name, info in book.hojas.items():
        rows = copy.deepcopy(info["rows"])
        hojas_out[name] = rows
        if snap:
            prev = {str(r.get(ID_FIELD[name], "")): r for r in snap.get("hojas", {}).get(name, [])
                    if ID_FIELD.get(name)}
            if ID_FIELD.get(name):
                nuevas, modificadas = [], []
                for r in rows:
                    rid = str(r.get(ID_FIELD[name], ""))
                    old = prev.get(rid)
                    if old is None:
                        nuevas.append(rid)
                    elif rows_differ(old, r):
                        modificadas.append(rid)
                if nuevas or modificadas:
                    cambios[name] = {"nuevas": nuevas, "modificadas": modificadas}
    pkg = {
        "tipo": "PAQUETE_ESPECIALISTA",
        "app": APP_NAME,
        "especialista": especialista,
        "libro": os.path.basename(book.path),
        "generado": now_iso(),
        "cambios": cambios,
        "hojas": hojas_out,
    }
    return pkg

def write_package(pkg, dest_dir, prefix="Paquete"):
    fn = f"{prefix}_{pkg['especialista'].replace(' ', '_')}_{datetime.date.today():%Y%m%d_%H%M}.json"
    dest = os.path.join(dest_dir, fn)
    with open(dest, "w", encoding="utf-8") as f:
        json.dump(pkg, f, ensure_ascii=False, default=json_default)
    return dest

def load_package(path):
    with open(path, encoding="utf-8") as f:
        return json.load(f)

def analyze_import(master, pkg):
    """Compara un paquete contra el maestro. Devuelve informe de consolidación."""
    reporte = {}
    for name, rows in pkg.get("hojas", {}).items():
        if name not in DATA_SHEETS or not rows:
            continue
        idf = ID_FIELD.get(name)
        if not idf:
            continue
        m_info = master.sheet(name)
        m_headers = set(m_info["headers"])
        rep = {"nuevas": [], "actualizadas": [], "conflictos": {}}
        for r in rows:
            rid = str(r.get(idf, ""))
            if not rid:
                continue
            existing, _ = master.find_row(name, rid)
            if existing is None:
                rep["nuevas"].append(rid)
            else:
                # solo campos que existen en el maestro
                inc = {k: v for k, v in r.items() if k in m_headers}
                diffs = rows_differ(existing, inc)
                # rellenos: maestro vacío, especialista con dato -> se completan
                conflict = {}
                relleno = False
                for k, (mv, iv) in diffs.items():
                    if mv == "":
                        relleno = True
                    else:
                        conflict[k] = (mv, iv)
                if conflict:
                    rep["conflictos"][rid] = conflict
                elif diffs:
                    rep["actualizadas"].append(rid)
        if rep["nuevas"] or rep["actualizadas"] or rep["conflictos"]:
            reporte[name] = rep
    return reporte

def apply_import(master, pkg, reporte, conflict_choices):
    """Aplica el paquete al maestro según el informe y las decisiones de conflictos.
    conflict_choices: {hoja: {id: True(tomar especialista)|False(conservar maestro)}}"""
    stats = {"nuevas": 0, "actualizadas": 0, "conflictos_especialista": 0, "conflictos_maestro": 0}
    for name, rep in reporte.items():
        idf = ID_FIELD[name]
        m_info = master.sheet(name)
        headers = m_info["headers"]
        rows_in = {str(r.get(idf, "")): r for r in pkg["hojas"].get(name, [])}
        for rid in rep["nuevas"]:
            r = rows_in.get(rid)
            if not r:
                continue
            row = {h: r.get(h, "") for h in headers}
            master.upsert(name, row, user=False)
            stats["nuevas"] += 1
        for rid in rep["actualizadas"]:
            r = rows_in.get(rid)
            if not r:
                continue
            existing, _ = master.find_row(name, rid)
            for h in headers:
                if h == "Fecha_modificacion":
                    continue
                if to_text(existing.get(h, "")) == "" and to_text(r.get(h, "")) != "":
                    existing[h] = r.get(h, "")
            if "Fecha_modificacion" in existing:
                existing["Fecha_modificacion"] = now_iso()[:16].replace("T", " ")
            stats["actualizadas"] += 1
        for rid, fields in rep["conflictos"].items():
            choice = conflict_choices.get(name, {}).get(rid, False)
            r = rows_in.get(rid)
            existing, _ = master.find_row(name, rid)
            if choice and r is not None and existing is not None:
                for h in headers:
                    if h == "Fecha_modificacion":
                        continue
                    if to_text(r.get(h, "")) != "":
                        existing[h] = r.get(h, "")
                stats["conflictos_especialista"] += 1
            else:
                stats["conflictos_maestro"] += 1
            if existing is not None and "Fecha_modificacion" in existing:
                existing["Fecha_modificacion"] = now_iso()[:16].replace("T", " ")
    return stats

def registrar_log(master, pkg, stats, notas=""):
    headers = ["Fecha", "Especialista", "Libro", "Paquete", "Nuevas", "Actualizadas",
               "Conflictos_a_favor_especialista", "Conflictos_conservados_maestro", "Notas"]
    if "LOG_CONSOLIDACION" not in master.hojas:
        master.hojas["LOG_CONSOLIDACION"] = {"headers": headers, "rows": []}
    log = master.hojas["LOG_CONSOLIDACION"]
    log["rows"].append({
        "Fecha": now_iso()[:16].replace("T", " "),
        "Especialista": pkg.get("especialista", ""),
        "Libro": pkg.get("libro", ""),
        "Paquete": pkg.get("generado", ""),
        "Nuevas": stats["nuevas"], "Actualizadas": stats["actualizadas"],
        "Conflictos_a_favor_especialista": stats["conflictos_especialista"],
        "Conflictos_conservados_maestro": stats["conflictos_maestro"],
        "Notas": notas,
    })
    # actualizar Ultima_consolidacion en Equipo
    if "Equipo" in master.hojas:
        for r in master.hojas["Equipo"]["rows"]:
            libro = str(r.get("Libro_asignado", ""))
            nombre = str(r.get("Nombre", ""))
            if libro and libro == pkg.get("libro", ""):
                r["Ultima_consolidacion"] = today_str()

# ---------------------------------------------------------------------------
# Interfaz
# ---------------------------------------------------------------------------

COLOR_BG = "#f4f6f9"
COLOR_HEADER = "#1f3a5f"
COLOR_ACCENT = "#2e86ab"
COLOR_SIDEBAR = "#16293f"

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"{APP_TITLE} — {APP_NAME}")
        self.geometry("1280x800")
        self.minsize(1000, 640)
        self.role = None
        self.book = None
        self.book_path = None
        self.especialista = ""
        self.current_sheet = None
        self._build_style()
        self._build_ui()
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    # ---------- estilo ----------
    def _build_style(self):
        self.configure(bg=COLOR_BG)
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("TFrame", background=COLOR_BG)
        style.configure("White.TFrame", background="white")
        style.configure("Top.TFrame", background=COLOR_HEADER)
        style.configure("Side.TFrame", background=COLOR_SIDEBAR)
        style.configure("Top.TLabel", background=COLOR_HEADER, foreground="white",
                        font=("Segoe UI", 11, "bold"))
        style.configure("TLabel", background=COLOR_BG, font=("Segoe UI", 10))
        style.configure("TButton", font=("Segoe UI", 10), padding=6)
        style.configure("Accent.TButton", background=COLOR_ACCENT, foreground="white",
                        font=("Segoe UI", 10, "bold"))
        style.map("Accent.TButton", background=[("active", "#22577a")])
        style.configure("TNotebook", background=COLOR_BG)
        style.configure("TNotebook.Tab", font=("Segoe UI", 10, "bold"), padding=10)
        style.configure("Treeview", font=("Segoe UI", 9), rowheight=24)
        style.configure("Treeview.Heading", font=("Segoe UI", 9, "bold"))
        style.configure("Side.TButton", background=COLOR_SIDEBAR, foreground="white",
                        anchor="w", font=("Segoe UI", 10), padding=8, borderwidth=0)
        style.map("Side.TButton", background=[("active", "#2a4a6b"), ("selected", COLOR_ACCENT)])

    # ---------- construcción ----------
    def _build_ui(self):
        # barra superior
        top = ttk.Frame(self, style="Top.TFrame", padding=(10, 8))
        top.pack(fill="x", side="top")
        ttk.Label(top, text=f"{APP_TITLE}", style="Top.TLabel").pack(side="left")
        self.lbl_role = ttk.Label(top, text="— sin rol —", style="Top.TLabel",
                                  font=("Segoe UI", 10, "italic"))
        self.lbl_role.pack(side="left", padx=14)
        self.lbl_file = ttk.Label(top, text="", style="Top.TLabel",
                                  font=("Segoe UI", 9, "italic"))
        self.lbl_file.pack(side="right")

        # cuerpo: barra lateral + área principal
        body = ttk.Frame(self)
        body.pack(fill="both", expand=True)

        self.sidebar = ttk.Frame(body, style="Side.TFrame", width=230, padding=6)
        self.sidebar.pack(side="left", fill="y")
        self.sidebar.pack_propagate(False)

        self.main = ttk.Frame(self)
        self.main.pack(side="left", fill="both", expand=True, padx=8, pady=8)

        self.btn_open = ttk.Button(self.sidebar, text="Abrir libro...", style="Side.TButton",
                                   command=self.open_book, state="disabled")
        self.btn_open.pack(fill="x", pady=2)
        self.btn_new = ttk.Button(self.sidebar, text="Nuevo libro...", style="Side.TButton",
                                  command=self.new_book, state="disabled")
        self.btn_new.pack(fill="x", pady=2)
        self.btn_save = ttk.Button(self.sidebar, text="Guardar", style="Side.TButton",
                                   command=self.save_book, state="disabled")
        self.btn_save.pack(fill="x", pady=2)
        ttk.Separator(self.sidebar).pack(fill="x", pady=6)

        self.btn_action1 = ttk.Button(self.sidebar, text="", style="Side.TButton",
                                      command=self.role_action_1, state="disabled")
        self.btn_action1.pack(fill="x", pady=2)
        self.btn_action2 = ttk.Button(self.sidebar, text="", style="Side.TButton",
                                      command=self.role_action_2, state="disabled")
        self.btn_action2.pack(fill="x", pady=2)

        ttk.Separator(self.sidebar).pack(fill="x", pady=6)
        self.nav = ttk.Treeview(self.sidebar, show="tree", selectmode="browse")
        self.nav.pack(fill="both", expand=True)
        self.nav.tag_configure("panel", font=("Segoe UI", 10, "bold"))
        self.nav.bind("<<TreeviewSelect>>", self._on_nav)

        # barra de estado
        self.status = ttk.Label(self, text="Listo.", anchor="w", relief="sunken", padding=6)
        self.status.pack(fill="x", side="bottom")

        self.frames = {}
        self._show_welcome()

    def _show_welcome(self):
        self._clear_main()
        f = ttk.Frame(self.main)
        f.pack(expand=True)
        ttk.Label(f, text="CRM INTEGRAL", font=("Segoe UI", 26, "bold"),
                  foreground=COLOR_HEADER).pack(pady=(30, 4))
        ttk.Label(f, text="GROUP EJ MORALES S.U.R.L.", font=("Segoe UI", 13),
                  foreground=COLOR_ACCENT).pack(pady=(0, 26))
        ttk.Label(f, text="Seleccione su rol para comenzar:", font=("Segoe UI", 11)).pack(pady=6)
        bf = ttk.Frame(f)
        bf.pack(pady=14)
        ttk.Button(bf, text="Especialista (comercial)", style="Accent.TButton",
                   command=lambda: self._set_role("especialista")).grid(row=0, column=0, padx=8, ipadx=14, ipady=8)
        ttk.Button(bf, text="Jefatura", style="Accent.TButton",
                   command=lambda: self._set_role("jefatura")).grid(row=0, column=1, padx=8, ipadx=14, ipady=8)
        ttk.Label(f, text="Especialista: trabaja su libro individual y envía paquetes diarios a jefatura.\n"
                          "Jefatura: importa los paquetes, consolida el maestro y aprueba ofertas.",
                  justify="center", font=("Segoe UI", 9, "italic")).pack(pady=16)

    def _clear_main(self):
        for w in self.main.winfo_children():
            w.destroy()
        self.frames = {}

    def _set_role(self, role):
        self.role = role
        self.lbl_role.config(text=("Especialista" if role == "especialista" else "Jefatura"))
        for b in (self.btn_open, self.btn_new, self.btn_save, self.btn_action1, self.btn_action2):
            b.config(state="normal")
        if role == "especialista":
            self.btn_action1.config(text="Enviar a jefatura")
            self.btn_action2.config(text="Importar retroalimentación")
        else:
            self.btn_action1.config(text="Importar paquete…")
            self.btn_action2.config(text="Exportar retroalimentación")
        self._prompt_book()
        self._refresh_nav()

    # ---------- libros ----------
    def _prompt_book(self):
        role_txt = "individual" if self.role == "especialista" else "maestro de jefatura"
        if messagebox.askyesno(APP_NAME, f"¿Desea abrir ahora su libro {role_txt}?\n\n"
                                          "Sí = abrir un libro existente\n"
                                          "No = crear uno nuevo desde la plantilla"):
            self.open_book()
        else:
            self.new_book()

    def open_book(self):
        path = filedialog.askopenfilename(
            title="Abrir libro CRM",
            filetypes=[("Libros Excel", "*.xlsx"), ("Todos", "*.*")],
            parent=self)
        if not path:
            return
        self._load_book_file(path)

    def new_book(self):
        plantilla = "Plantilla_Libro_Individual.xlsx" if self.role == "especialista" else "Plantilla_Maestro_Jefatura.xlsx"
        src = os.path.join(BASE_DIR, plantilla)
        if not os.path.exists(src):
            messagebox.showerror(APP_NAME, f"No se encontró la plantilla:\n{src}")
            return
        path = filedialog.asksaveasfilename(
            title="Guardar nuevo libro como…", defaultextension=".xlsx",
            filetypes=[("Libros Excel", "*.xlsx")],
            initialfile=f"CRM_Comercial_XX_GroupEJ_Morales_{today_str()}.xlsx"
            if self.role == "especialista" else f"CRM_Maestro_Jefatura_GroupEJ_Morales_{today_str()}.xlsx",
            parent=self)
        if not path:
            return
        shutil.copyfile(src, path)
        self._load_book_file(path)

    def _load_book_file(self, path):
        try:
            self.book = CRMBook(path)
        except Exception as e:
            messagebox.showerror(APP_NAME, f"No se pudo abrir el libro:\n{e}")
            return
        self.book_path = path
        self.lbl_file.config(text=os.path.basename(path))
        self.especialista = os.path.basename(path).replace("CRM_", "").split("_GroupEJ")[0]
        self._refresh_nav()
        self._select_first_sheet()
        self._set_status(f"Libro abierto: {os.path.basename(path)}")

    def save_book(self):
        if not self.book:
            return
        try:
            self.book.save()
            self._set_status(f"Guardado: {os.path.basename(self.book_path)} — {now_iso()[:16].replace('T', ' ')}")
        except Exception as e:
            messagebox.showerror(APP_NAME, f"Error al guardar:\n{e}")

    def _on_close(self):
        if self.book and messagebox.askyesno(APP_NAME, "¿Guardar cambios antes de salir?"):
            self.save_book()
        self.destroy()

    # ---------- navegación ----------
    def _refresh_nav(self):
        self.nav.delete(*self.nav.get_children())
        if not self.book:
            return
        self.nav.insert("", "end", iid="__panel__", text="PANEL", tags=("panel",))
        for name in DATA_SHEETS:
            if name in self.book.hojas:
                n = len(self.book.hojas[name]["rows"])
                self.nav.insert("", "end", iid=name, text=f"  {name}  ({n})")
        if "LOG_CONSOLIDACION" in self.book.hojas:
            n = len(self.book.hojas["LOG_CONSOLIDACION"]["rows"])
            self.nav.insert("", "end", iid="LOG_CONSOLIDACION",
                            text=f"  LOG_CONSOLIDACION  ({n})", tags=("panel",))

    def _select_first_sheet(self):
        if self.book:
            self.nav.selection_set("__panel__")
            self._on_nav(None)

    def _on_nav(self, event):
        sel = self.nav.selection()
        if not sel:
            return
        name = sel[0]
        self.current_sheet = None if name == "__panel__" else name
        if name == "__panel__":
            self._show_panel()
        else:
            self._show_table(name)

    def _set_status(self, msg):
        self.status.config(text=msg)

    # ---------- panel ----------
    def _show_panel(self):
        self._clear_main()
        f = ttk.Frame(self.main)
        f.pack(fill="both", expand=True)
        title = "PANEL DE CONTROL" + (" — JEFATURA" if self.role == "jefatura" else " — ESPECIALISTA")
        head = ttk.Frame(f)
        head.pack(fill="x", pady=(4, 12))
        ttk.Label(head, text=title, font=("Segoe UI", 16, "bold"),
                  foreground=COLOR_HEADER).pack(side="left")
        ttk.Button(head, text="Exportar informe a Excel",
                   command=self._exportar_informe).pack(side="right")

        if not self.book:
            ttk.Label(f, text="Abra o cree un libro para ver el panel.").pack()
            return

        # tarjetas de resumen
        cards = ttk.Frame(f)
        cards.pack(fill="x", pady=6)
        self._panel_cards(cards)

        # zona inferior en dos columnas
        bottom = ttk.Frame(f)
        bottom.pack(fill="both", expand=True, pady=8)
        left = ttk.Frame(bottom)
        left.pack(side="left", fill="both", expand=True, padx=(0, 6))
        right = ttk.Frame(bottom)
        right.pack(side="left", fill="both", expand=True, padx=(6, 0))

        self._panel_vencidos(left)
        if self.role == "jefatura":
            self._panel_ofertas_pendientes(right)
            self._panel_log(right)
        else:
            self._panel_ofertas_mias(right)

    def _card(self, parent, text, value, color=COLOR_ACCENT):
        c = tk.Frame(parent, bg="white", bd=1, relief="solid")
        c.pack(side="left", padx=6, ipadx=14, ipady=10, expand=True)
        tk.Label(c, text=str(value), font=("Segoe UI", 20, "bold"), bg="white",
                 fg=color).pack()
        tk.Label(c, text=text, font=("Segoe UI", 9), bg="white", fg="#555").pack()

    def _panel_cards(self, parent):
        b = self.book
        hoy = today_str()
        cli = len(b.hojas.get("Clientes", {}).get("rows", []))
        pro = len(b.hojas.get("Prospectos", {}).get("rows", []))
        abiertas = sum(1 for r in b.hojas.get("Prospectos", {}).get("rows", [])
                       if r.get("Etapa") not in ("Ganada", "Perdida", "Suspendida", ""))
        ofe = len(b.hojas.get("Ofertas", {}).get("rows", []))
        pend = sum(1 for r in b.hojas.get("Ofertas", {}).get("rows", [])
                   if r.get("Estado_revision") in OFERTAS_PENDIENTES)
        vta = len(b.hojas.get("Ventas", {}).get("rows", []))
        venc = 0
        for sh in DATA_SHEETS:
            for r in b.hojas.get(sh, {}).get("rows", []):
                fs = to_text(r.get("Fecha_seguimiento", ""))
                if fs and fs[:10] < hoy:
                    venc += 1
        self._card(parent, "Clientes", cli, "#1f3a5f")
        self._card(parent, "Oportunidades abiertas", abiertas, "#2e86ab")
        self._card(parent, "Ofertas", ofe, "#6c5b9c")
        self._card(parent, "Ofertas pendientes de revisión", pend, "#d9822b")
        self._card(parent, "Ventas", vta, "#2a9d8f")
        self._card(parent, "Seguimientos vencidos", venc, "#c0392b")

    def _panel_vencidos(self, parent):
        ttk.Label(parent, text="Seguimientos vencidos", font=("Segoe UI", 11, "bold"),
                  foreground="#c0392b").pack(anchor="w")
        box = ttk.Treeview(parent, show="headings", height=8)
        box["columns"] = ("hoja", "id", "cliente", "fecha", "accion")
        for c, w in zip(box["columns"], (150, 90, 220, 90, 260)):
            box.heading(c, text={"hoja": "Hoja", "id": "ID", "cliente": "Cliente/Entidad",
                                 "fecha": "Fecha seguim.", "accion": "Próxima acción"}[c])
            box.column(c, width=w, anchor="w")
        box.pack(fill="both", expand=True)
        hoy = today_str()
        n = 0
        for sh in DATA_SHEETS:
            idf = ID_FIELD.get(sh)
            for r in self.book.hojas.get(sh, {}).get("rows", []):
                fs = to_text(r.get("Fecha_seguimiento", ""))
                if fs and fs[:10] < hoy:
                    box.insert("", "end", values=(sh, r.get(idf or "", ""),
                                   r.get("Cliente", r.get("Nombre", r.get("Entidad", ""))),
                                   fs[:10], to_text(r.get("Proxima_accion", ""))[:60]))
                    n += 1
        if n == 0:
            box.insert("", "end", values=("", "", "— sin seguimientos vencidos —", "", ""))

    def _panel_ofertas_pendientes(self, parent):
        ttk.Label(parent, text="Ofertas que requieren su revisión", font=("Segoe UI", 11, "bold"),
                  foreground="#d9822b").pack(anchor="w")
        box = ttk.Treeview(parent, show="headings", height=8)
        box["columns"] = ("id", "cliente", "producto", "estado", "responsable")
        for c, w in zip(box["columns"], (80, 200, 200, 180, 140)):
            box.heading(c, text={"id": "ID", "cliente": "Cliente", "producto": "Producto/servicio",
                                 "estado": "Estado revisión", "responsable": "Comercial"}[c])
            box.column(c, width=w, anchor="w")
        box.pack(fill="both", expand=True)
        rows = self.book.hojas.get("Ofertas", {}).get("rows", [])
        n = 0
        for r in rows:
            if r.get("Estado_revision") in OFERTAS_PENDIENTES:
                box.insert("", "end", values=(r.get("ID_Oferta", ""), r.get("Cliente", ""),
                               to_text(r.get("Producto_servicio", ""))[:40],
                               r.get("Estado_revision", ""), r.get("Comercial_responsable", "")))
                n += 1
        if n == 0:
            box.insert("", "end", values=("", "— sin ofertas pendientes —", "", "", ""))
        box.bind("<Double-1>", lambda e: self._review_oferta(box))

    def _panel_ofertas_mias(self, parent):
        ttk.Label(parent, text="Mis ofertas enviadas a jefatura", font=("Segoe UI", 11, "bold"),
                  foreground="#6c5b9c").pack(anchor="w")
        box = ttk.Treeview(parent, show="headings", height=8)
        box["columns"] = ("id", "cliente", "estado", "obs")
        for c, w in zip(box["columns"], (80, 200, 200, 320)):
            box.heading(c, text={"id": "ID", "cliente": "Cliente", "estado": "Estado revisión",
                                 "obs": "Observaciones de jefatura"}[c])
            box.column(c, width=w, anchor="w")
        box.pack(fill="both", expand=True)
        rows = self.book.hojas.get("Ofertas", {}).get("rows", [])
        for r in rows:
            box.insert("", "end", values=(r.get("ID_Oferta", ""), r.get("Cliente", ""),
                           r.get("Estado_revision", ""),
                           to_text(r.get("Observaciones_jefatura", ""))[:80]))
        if not rows:
            box.insert("", "end", values=("", "— sin ofertas registradas —", "", ""))

    def _panel_log(self, parent):
        ttk.Label(parent, text="Registro de consolidaciones", font=("Segoe UI", 11, "bold"),
                  foreground=COLOR_HEADER).pack(anchor="w", pady=(10, 0))
        box = ttk.Treeview(parent, show="headings", height=5)
        box["columns"] = ("fecha", "esp", "nuevas", "act", "conf")
        for c, w in zip(box["columns"], (125, 115, 55, 60, 75)):
            box.heading(c, text={"fecha": "Fecha", "esp": "Especialista", "nuevas": "Nuevas",
                                 "act": "Actualiz.", "conf": "Conflic."}[c])
            box.column(c, width=w, anchor="w")
        box.pack(fill="both", expand=True)
        rows = self.book.hojas.get("LOG_CONSOLIDACION", {}).get("rows", [])[-8:]
        for r in rows:
            box.insert("", "end", values=(r.get("Fecha", ""), r.get("Especialista", ""),
                           r.get("Nuevas", ""), r.get("Actualizadas", ""),
                           f"{r.get('Conflictos_a_favor_especialista','')}/{r.get('Conflictos_conservados_maestro','')}"))
        if not rows:
            box.insert("", "end", values=("— sin consolidaciones —", "", "", "", ""))

    # ---------- tabla de datos ----------
    def _show_table(self, name):
        self._clear_main()
        f = ttk.Frame(self.main)
        f.pack(fill="both", expand=True)
        info = self.book.sheet(name)
        headers = info["headers"]
        ttk.Label(f, text=name.replace("_", " "), font=("Segoe UI", 14, "bold"),
                  foreground=COLOR_HEADER).pack(anchor="w", pady=(0, 6))

        bar = ttk.Frame(f)
        bar.pack(fill="x", pady=4)
        ttk.Label(bar, text="Buscar:").pack(side="left")
        sv = tk.StringVar()
        sv.trace_add("write", lambda *a: self._filter_table(name, sv.get()))
        ttk.Entry(bar, textvariable=sv, width=40).pack(side="left", padx=6)
        ttk.Button(bar, text="Nueva fila", style="Accent.TButton",
                   command=lambda: self._edit_row(name, None)).pack(side="right", padx=4)
        ttk.Button(bar, text="Exportar a Excel",
                   command=lambda: self._exportar_hoja(name)).pack(side="right", padx=4)
        ttk.Button(bar, text="Editar",
                   command=lambda: self._edit_selected(name)).pack(side="right", padx=4)
        ttk.Button(bar, text="Eliminar",
                   command=lambda: self._delete_selected(name)).pack(side="right", padx=4)

        wrap = ttk.Frame(f)
        wrap.pack(fill="both", expand=True)
        ysb = ttk.Scrollbar(wrap, orient="vertical")
        xsb = ttk.Scrollbar(wrap, orient="horizontal")
        tree = ttk.Treeview(wrap, show="headings", yscrollcommand=ysb.set,
                            xscrollcommand=xsb.set)
        ysb.config(command=tree.yview)
        xsb.config(command=tree.xview)
        tree["columns"] = headers
        for h in headers:
            tree.heading(h, text=h, command=lambda _h=h: self._sort_table(name, tree, _h))
            tree.column(h, width=130, anchor="w", stretch=False)
        ysb.pack(side="right", fill="y")
        xsb.pack(side="bottom", fill="x")
        tree.pack(fill="both", expand=True)
        tree.bind("<Double-1>", lambda e: self._edit_selected(name))
        self.frames["tree"] = tree
        self.frames["sheet"] = name
        self._fill_table(name, "")

    def _fill_table(self, name, search=""):
        tree = self.frames.get("tree")
        if tree is None:
            return
        info = self.book.sheet(name)
        tree.delete(*tree.get_children())
        s = search.strip().lower()
        for r in info["rows"]:
            vals = [to_text(r.get(h, "")) for h in info["headers"]]
            if s and not any(s in v.lower() for v in vals):
                continue
            tree.insert("", "end", iid=str(r.get(ID_FIELD.get(name, ""), "")) or None, values=vals)

    def _filter_table(self, name, search):
        self._fill_table(name, search)

    def _sort_table(self, name, tree, col):
        items = [(tree.set(k, col).lower(), k) for k in tree.get_children()]
        items.sort()
        asc = getattr(self, "_sort_asc", True)
        if not asc:
            items.reverse()
        self._sort_asc = not asc
        for i, (_, k) in enumerate(items):
            tree.move(k, "", i)

    def _selected_id(self, name):
        tree = self.frames.get("tree")
        if tree is None:
            return None
        sel = tree.selection()
        if not sel:
            return None
        return sel[0]

    def _edit_selected(self, name):
        rid = self._selected_id(name)
        if rid is None:
            messagebox.showinfo(APP_NAME, "Seleccione una fila (doble clic) para editar.")
            return
        row, _ = self.book.find_row(name, rid)
        if row is None:
            # fallback: buscar por índice visible
            tree = self.frames["tree"]
            idx = tree.index(tree.selection()[0])
            row = self.book.sheet(name)["rows"][idx]
        self._edit_row(name, row)

    def _delete_selected(self, name):
        rid = self._selected_id(name)
        if rid is None:
            messagebox.showinfo(APP_NAME, "Seleccione una fila para eliminar.")
            return
        if messagebox.askyesno(APP_NAME, f"¿Eliminar la fila {rid} de {name}?\n"
                                         "Esta acción se guarda al pulsar Guardar."):
            self.book.delete_row(name, rid)
            self._fill_table(name, "")
            self._refresh_nav()
            self._set_status(f"Fila {rid} eliminada de {name} (recuerde Guardar).")

    # ---------- diálogo de edición ----------
    def _edit_row(self, name, row):
        info = self.book.sheet(name)
        headers = info["headers"]
        dlg = tk.Toplevel(self)
        dlg.title(("Editar" if row else "Nueva fila en") + f" — {name}")
        dlg.transient(self)
        center_over(self, dlg, 780, 660)
        dlg.grab_set()

        btns = ttk.Frame(dlg)
        btns.pack(side="bottom", fill="x", pady=8)
        canvas = tk.Canvas(dlg, bg="white", highlightthickness=0)
        scrollbar = ttk.Scrollbar(dlg, orient="vertical", command=canvas.yview)
        form = ttk.Frame(canvas, style="White.TFrame")
        form.bind("<Configure>",
                  lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        win = canvas.create_window((0, 0), window=form, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        canvas.bind("<Configure>", lambda e: canvas.itemconfigure(win, width=e.width))
        canvas.bind_all("<MouseWheel>",
                        lambda e: canvas.yview_scroll(int(-1 * (e.delta / 120)), "units"))
        dlg.bind("<Destroy>", lambda e: canvas.unbind_all("<MouseWheel>"))

        widgets = {}
        cur = row or {}
        for i, h in enumerate(headers):
            ccol, crow = i % 2, i // 2
            frm = ttk.Frame(form)
            frm.grid(row=crow, column=ccol, sticky="ew", padx=10, pady=4)
            form.columnconfigure(ccol, weight=1)
            ttk.Label(frm, text=h, font=("Segoe UI", 9, "bold"), width=28,
                      background="white").pack(anchor="w")
            val = to_text(cur.get(h, ""))
            if h == ID_FIELD.get(name):
                if row is None:
                    val = self.book.next_id(name)
                e = ttk.Entry(frm, width=34)
                e.insert(0, val)
                e.pack(fill="x")
                widgets[h] = e
            elif (name, h) in CATALOG_MAP:
                tipo = CATALOG_MAP[(name, h)]
                valores = self.book.catalogos.get(tipo, [])
                cb = ttk.Combobox(frm, values=valores, width=32)
                cb.set(val)
                cb.pack(fill="x")
                widgets[h] = cb
            else:
                e = ttk.Entry(frm, width=34)
                e.insert(0, val)
                e.pack(fill="x")
                widgets[h] = e

        ttk.Button(btns, text="Cancelar", command=dlg.destroy).pack(side="right", padx=10)
        ttk.Button(btns, text="Guardar fila", style="Accent.TButton",
                   command=lambda: self._save_dialog_row(dlg, name, widgets, row)).pack(side="right")

    def _save_dialog_row(self, dlg, name, widgets, existing):
        info = self.book.sheet(name)
        row = {}
        for h in info["headers"]:
            w = widgets.get(h)
            v = w.get().strip() if w else ""
            row[h] = coerce_cell(info["headers"], h, v)
        idf = ID_FIELD.get(name)
        if idf and not str(row.get(idf, "")).strip():
            messagebox.showwarning(APP_NAME, f"El campo {idf} es obligatorio.", parent=dlg)
            return
        # verificar duplicado de ID (si es nueva fila)
        if existing is None or str(existing.get(idf, "")) != str(row.get(idf, "")):
            dup, _ = self.book.find_row(name, row.get(idf, ""))
            if dup is not None:
                messagebox.showwarning(APP_NAME, f"Ya existe una fila con {idf} = {row.get(idf)}.",
                                       parent=dlg)
                return
        res = self.book.upsert(name, row)
        dlg.destroy()
        self._fill_table(name, "")
        self._refresh_nav()
        self._set_status(f"Fila {row.get(idf, '')} {res} en {name}. Recuerde Guardar.")

    # ---------- acciones por rol ----------
    def role_action_1(self):
        if self.role == "especialista":
            self.enviar_paquete()
        else:
            self.importar_paquete()

    def role_action_2(self):
        if self.role == "especialista":
            self.importar_retro()
        else:
            self.exportar_retro()

    # -- especialista: enviar paquete --
    def enviar_paquete(self):
        if not self.book:
            return
        if not self._confirm_unsaved():
            return
        pkg = build_package(self.book, self.especialista or "Especialista")
        dest_dir = os.path.join(os.path.dirname(self.book_path) or ".", "Paquetes_a_Jefatura")
        os.makedirs(dest_dir, exist_ok=True)
        dest = write_package(pkg, dest_dir)
        save_snapshot(self.book_path, self.book.hojas)
        cambios = pkg["cambios"]
        total_n = sum(len(v["nuevas"]) for v in cambios.values())
        total_m = sum(len(v["modificadas"]) for v in cambios.values())
        if cambios:
            detalle = "\n".join(f"  • {k}: {len(v['nuevas'])} nuevas, {len(v['modificadas'])} modificadas"
                                for k, v in cambios.items())
        else:
            detalle = "  • Sin cambios desde el último envío (se envía copia completa)."
        self._set_status(f"Paquete generado: {os.path.basename(dest)}")
        messagebox.showinfo(
            APP_NAME,
            "PAQUETE LISTO PARA ENVIAR A JEFATURA\n\n"
            f"Archivo: {dest}\n\n"
            f"Resumen de cambios desde el último envío:\n{detalle}\n\n"
            f"Nuevas filas: {total_n} | Modificadas: {total_m}\n\n"
            "Envíe este archivo a jefatura por el medio habitual (correo, USB, etc.).\n"
            "Guarde también su libro para conservar los cambios.")

    # -- especialista: importar retroalimentación --
    def importar_retro(self):
        if not self.book:
            return
        path = filedialog.askopenfilename(
            title="Seleccionar retroalimentación de jefatura",
            filetypes=[("Paquete JSON", "*.json"), ("Todos", "*.*")], parent=self)
        if not path:
            return
        try:
            pkg = load_package(path)
        except Exception as e:
            messagebox.showerror(APP_NAME, f"Archivo no válido:\n{e}")
            return
        if pkg.get("tipo") != "RETRO_JEFATURA":
            messagebox.showerror(APP_NAME, "Este archivo no es una retroalimentación de jefatura.")
            return
        # previsualizar
        preview, totales = [], {"nuevas": 0, "actualizadas": 0}
        for name, rows in pkg.get("hojas", {}).items():
            idf = ID_FIELD.get(name)
            if not idf:
                continue
            for r in rows:
                rid = str(r.get(idf, ""))
                existing, _ = self.book.find_row(name, rid)
                if existing is None:
                    preview.append((name, rid, "NUEVA", ""))
                    totales["nuevas"] += 1
                elif rows_differ(existing, r):
                    preview.append((name, rid, "ACTUALIZADA",
                                    "; ".join(list(rows_differ(existing, r))[:3])))
                    totales["actualizadas"] += 1
        msg = ("RETROALIMENTACIÓN DE JEFATURA\n\n"
               f"Generada: {pkg.get('generado','')}\n\n"
               + "\n".join(f"• {h} | {i} | {e} {d}" for h, i, e, d in preview[:20])
               + (f"\n… y {len(preview)-20} más" if len(preview) > 20 else "")
               + f"\n\nNuevas: {totales['nuevas']}  Actualizadas: {totales['actualizadas']}\n\n"
                 "¿Aplicar estos cambios a su libro?")
        if not messagebox.askyesno(APP_NAME, msg):
            return
        for name, rows in pkg.get("hojas", {}).items():
            headers = self.book.sheet(name)["headers"]
            for r in rows:
                row = {h: r.get(h, "") for h in headers}
                self.book.upsert(name, row, user=False)
        self.save_book()
        self._refresh_nav()
        self._set_status("Retroalimentación de jefatura aplicada y libro guardado.")
        messagebox.showinfo(APP_NAME, "Cambios aplicados y libro guardado.")

    # -- jefatura: importar paquete --
    def importar_paquete(self):
        if not self.book:
            return
        if not self._confirm_unsaved():
            return
        paths = filedialog.askopenfilenames(
            title="Seleccionar paquete(s) de especialistas",
            filetypes=[("Paquete JSON", "*.json"), ("Todos", "*.*")], parent=self)
        for path in paths:
            try:
                pkg = load_package(path)
            except Exception as e:
                messagebox.showerror(APP_NAME, f"No se pudo leer {os.path.basename(path)}:\n{e}")
                continue
            if pkg.get("tipo") != "PAQUETE_ESPECIALISTA":
                messagebox.showerror(APP_NAME,
                                     f"{os.path.basename(path)} no es un paquete de especialista.")
                continue
            self._process_package(pkg, path)

    def _process_package(self, pkg, path):
        reporte = analyze_import(self.book, pkg)
        tot_n = sum(len(r["nuevas"]) for r in reporte.values())
        tot_a = sum(len(r["actualizadas"]) for r in reporte.values())
        tot_c = sum(len(r["conflictos"]) for r in reporte.values())
        if tot_n == 0 and tot_a == 0 and tot_c == 0:
            messagebox.showinfo(APP_NAME,
                                f"Paquete de {pkg.get('especialista','')}: sin novedades "
                                "(todo ya está consolidado en el maestro).")
            return
        # resumen
        lines = [f"ESPECIALISTA: {pkg.get('especialista','')}   "
                 f"Generado: {str(pkg.get('generado',''))[:16]}"]
        for name, rep in reporte.items():
            lines.append(f"• {name}: {len(rep['nuevas'])} nuevas, {len(rep['actualizadas'])} "
                         f"completadas, {len(rep['conflictos'])} conflictos")
        lines.append(f"\nTOTAL: {tot_n} nuevas | {tot_a} completadas | {tot_c} conflictos")
        if tot_c:
            lines.append("\nSe abrirá el panel de conflictos para decidir cada caso.")
        if not messagebox.askyesno(APP_NAME,
                                   "IMPORTAR PAQUETE\n" + "\n".join(lines) +
                                   "\n\n¿Continuar?"):
            return
        # resolver conflictos
        choices = {}
        if tot_c:
            choices = ConflictDialog(self, reporte).run()
            if choices is None:
                self._set_status("Importación cancelada en el panel de conflictos.")
                return
        stats = apply_import(self.book, pkg, reporte, choices)
        registrar_log(self.book, pkg, stats)
        self.save_book()
        self._refresh_nav()
        self._set_status(f"Paquete de {pkg.get('especialista','')} consolidado. "
                         f"Nuevas: {stats['nuevas']} | Actualizadas: {stats['actualizadas']}")
        messagebox.showinfo(
            APP_NAME,
            "CONSOLIDACIÓN COMPLETADA\n\n"
            f"Especialista: {pkg.get('especialista','')}\n"
            f"Nuevas filas: {stats['nuevas']}\n"
            f"Filas completadas: {stats['actualizadas']}\n"
            f"Conflictos a favor del especialista: {stats['conflictos_especialista']}\n"
            f"Conflictos conservados en maestro: {stats['conflictos_maestro']}\n\n"
            "El maestro quedó guardado y el suceso quedó registrado en LOG_CONSOLIDACION.")

    # -- jefatura: exportar retroalimentación --
    def exportar_retro(self):
        if not self.book:
            return
        if not self._confirm_unsaved():
            return
        dlg = tk.Toplevel(self)
        dlg.title("Exportar retroalimentación a especialista")
        dlg.transient(self)
        center_over(self, dlg, 540, 320)
        dlg.grab_set()
        ttk.Label(dlg, text="Seleccione el especialista que recibirá la retroalimentación:",
                  font=("Segoe UI", 10, "bold")).pack(pady=10, padx=12, anchor="w")
        # especialistas conocidos: de Equipo y de LOG_CONSOLIDACION
        nombres = []
        for r in self.book.hojas.get("Equipo", {}).get("rows", []):
            n = str(r.get("Nombre", "")).strip()
            if n:
                nombres.append(n)
        for r in self.book.hojas.get("LOG_CONSOLIDACION", {}).get("rows", []):
            n = str(r.get("Especialista", "")).strip()
            if n and n not in nombres:
                nombres.append(n)
        if not nombres:
            nombres = ["Comercial 01", "Comercial 02", "Comercial 03"]
        cb = ttk.Combobox(dlg, values=sorted(set(nombres)), width=40)
        cb.current(0)
        cb.pack(pady=6, padx=12, fill="x")
        ttk.Label(dlg, text="Se incluirán las hojas: Ofertas (estado de revisión, observaciones), "
                            "Equipo y Prospectos con resultado.\nEl especialista la importa con "
                            "'Importar retroalimentación'.",
                  font=("Segoe UI", 9, "italic"), wraplength=480, justify="left").pack(pady=8, padx=12)

        def do_export():
            esp = cb.get().strip() or "Especialista"
            hojas = {}
            for sh in ("Ofertas", "Equipo", "Prospectos"):
                if sh in self.book.hojas:
                    hojas[sh] = copy.deepcopy(self.book.hojas[sh]["rows"])
            pkg = {"tipo": "RETRO_JEFATURA", "app": APP_NAME, "especialista": esp,
                   "generado": now_iso(), "hojas": hojas}
            dest_dir = os.path.join(os.path.dirname(self.book_path) or ".", "Retroalimentaciones")
            os.makedirs(dest_dir, exist_ok=True)
            dest = write_package(pkg, dest_dir, prefix="Retro")
            dlg.destroy()
            self._set_status(f"Retroalimentación exportada: {os.path.basename(dest)}")
            messagebox.showinfo(APP_NAME, f"Retroalimentación generada:\n{dest}\n\n"
                                          f"Envíela a {esp} por el medio habitual.")

        ttk.Button(dlg, text="Exportar", style="Accent.TButton", command=do_export).pack(pady=10)

    # ---------- exportación a Excel ----------
    def _exportar_hoja(self, name):
        if not self.book:
            return
        info = self.book.sheet(name)
        path = filedialog.asksaveasfilename(
            title="Exportar hoja a Excel", defaultextension=".xlsx",
            filetypes=[("Libros Excel", "*.xlsx")],
            initialfile=f"{name}_{today_str()}.xlsx", parent=self)
        if not path:
            return
        try:
            export_sheet_to_xlsx(path, name, info["headers"], info["rows"],
                                 title=f"CRM Group EJ Morales — {name}")
        except Exception as e:
            messagebox.showerror(APP_NAME, f"Error al exportar:\n{e}")
            return
        self._set_status(f"Hoja {name} exportada a {os.path.basename(path)}")
        if messagebox.askyesno(APP_NAME, f"Exportado correctamente:\n{path}\n\n¿Abrir el archivo?"):
            try:
                if sys.platform.startswith("win"):
                    os.startfile(path)
                elif sys.platform == "darwin":
                    os.system(f'open "{path}"')
                else:
                    os.system(f'xdg-open "{path}"')
            except Exception:
                pass

    def _exportar_informe(self):
        if not self.book:
            return
        b = self.book
        hoy = today_str()
        # resumen
        abiertas = sum(1 for r in b.hojas.get("Prospectos", {}).get("rows", [])
                       if r.get("Etapa") not in ("Ganada", "Perdida", "Suspendida", ""))
        pend = [r for r in b.hojas.get("Ofertas", {}).get("rows", [])
                if r.get("Estado_revision") in OFERTAS_PENDIENTES]
        venc = []
        for sh in DATA_SHEETS:
            idf = ID_FIELD.get(sh)
            for r in b.hojas.get(sh, {}).get("rows", []):
                fs = to_text(r.get("Fecha_seguimiento", ""))
                if fs and fs[:10] < hoy:
                    venc.append((sh, r.get(idf or "", ""),
                                 r.get("Cliente", r.get("Nombre", r.get("Entidad", ""))),
                                 fs[:10], to_text(r.get("Proxima_accion", ""))))
        resumen = [
            ("Indicador", "Valor"),
            ("Fecha del informe", hoy),
            ("Rol", "Jefatura" if self.role == "jefatura" else "Especialista"),
            ("Libro", os.path.basename(self.book_path)),
            ("Clientes", len(b.hojas.get("Clientes", {}).get("rows", []))),
            ("Proveedores", len(b.hojas.get("Proveedores", {}).get("rows", []))),
            ("Oportunidades totales", len(b.hojas.get("Prospectos", {}).get("rows", []))),
            ("Oportunidades abiertas", abiertas),
            ("Ofertas registradas", len(b.hojas.get("Ofertas", {}).get("rows", []))),
            ("Ofertas pendientes de revisión", len(pend)),
            ("Ventas", len(b.hojas.get("Ventas", {}).get("rows", []))),
            ("Seguimientos vencidos", len(venc)),
        ]
        path = filedialog.asksaveasfilename(
            title="Exportar informe a Excel", defaultextension=".xlsx",
            filetypes=[("Libros Excel", "*.xlsx")],
            initialfile=f"Informe_{'Jefatura' if self.role == 'jefatura' else 'Especialista'}_{hoy}.xlsx",
            parent=self)
        if not path:
            return
        try:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Resumen"
            for r, (k, v) in enumerate(resumen, 1):
                if r == 1:
                    ws.cell(r, 1, k).font = Font(bold=True, color="FFFFFF")
                    ws.cell(r, 1).fill = PatternFill("solid", fgColor="1F3A5F")
                    ws.cell(r, 2, v).font = Font(bold=True, color="FFFFFF")
                    ws.cell(r, 2).fill = PatternFill("solid", fgColor="1F3A5F")
                else:
                    ws.cell(r, 1, k).font = Font(bold=True)
                    ws.cell(r, 2, v)
            ws.column_dimensions["A"].width = 34
            ws.column_dimensions["B"].width = 30
            # vencidos
            wsv = wb.create_sheet("Seguimientos_vencidos")
            for c, h in enumerate(("Hoja", "ID", "Cliente/Entidad", "Fecha", "Próxima acción"), 1):
                cell = wsv.cell(1, c, h)
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor="1F3A5F")
            for r, row in enumerate(venc, 2):
                for c, v in enumerate(row, 1):
                    wsv.cell(r, c, v)
            for c, w in enumerate((22, 12, 40, 12, 50), 1):
                wsv.column_dimensions[openpyxl.utils.get_column_letter(c)].width = w
            # ofertas pendientes
            wso = wb.create_sheet("Ofertas_pendientes")
            oh = ["ID_Oferta", "Cliente", "Producto_servicio", "Comercial_responsable",
                  "Estado_revision", "Fecha_preparacion"]
            for c, h in enumerate(oh, 1):
                cell = wso.cell(1, c, h)
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor="1F3A5F")
            for r, row in enumerate(pend, 2):
                for c, h in enumerate(oh, 1):
                    wso.cell(r, c, row.get(h, "") or None)
            for c, w in enumerate((12, 34, 34, 24, 26, 16), 1):
                wso.column_dimensions[openpyxl.utils.get_column_letter(c)].width = w
            wb.save(path)
            wb.close()
        except Exception as e:
            messagebox.showerror(APP_NAME, f"Error al exportar:\n{e}")
            return
        self._set_status(f"Informe exportado a {os.path.basename(path)}")
        if messagebox.askyesno(APP_NAME, f"Exportado correctamente:\n{path}\n\n¿Abrir el archivo?"):
            try:
                if sys.platform.startswith("win"):
                    os.startfile(path)
                elif sys.platform == "darwin":
                    os.system(f'open "{path}"')
                else:
                    os.system(f'xdg-open "{path}"')
            except Exception:
                pass

    # ---------- utilidades ----------
    def _confirm_unsaved(self):
        return messagebox.askyesno(APP_NAME, "¿Guardar el libro antes de continuar?\n"
                                             "(Recomendado para que el paquete incluya todo).",
                                   parent=self) and (self.save_book() or True)

# ---------------------------------------------------------------------------
# Diálogo de conflictos de consolidación
# ---------------------------------------------------------------------------

class ConflictDialog(tk.Toplevel):
    def __init__(self, master, reporte):
        super().__init__(master)
        self.title("Resolver conflictos de consolidación")
        self.transient(master)
        center_over(master, self, 980, 620)
        self.grab_set()
        self.reporte = reporte
        self.choices = {}     # {hoja: {id: bool}}
        self.items = []       # lista plana de (hoja, id, fields)
        for name, rep in reporte.items():
            for rid, fields in rep["conflictos"].items():
                self.items.append((name, rid, fields))
                self.choices.setdefault(name, {})[rid] = False
        self.idx = 0
        self._build()
        self._show()

    def _build(self):
        top = ttk.Frame(self, padding=8)
        top.pack(fill="x")
        self.lbl_info = ttk.Label(top, font=("Segoe UI", 11, "bold"))
        self.lbl_info.pack(anchor="w")

        mid = ttk.Frame(self, padding=8)
        mid.pack(fill="both", expand=True)
        cols = ("campo", "maestro", "especialista")
        self.tree = ttk.Treeview(mid, show="headings", columns=cols, height=10)
        for c, w in zip(cols, (240, 320, 320)):
            self.tree.heading(c, text={"campo": "Campo", "maestro": "Valor en MAESTRO (jefatura)",
                                       "especialista": "Valor del ESPECIALISTA"}[c])
            self.tree.column(c, width=w, anchor="w")
        self.tree.pack(fill="both", expand=True)

        bot = ttk.Frame(self, padding=8)
        bot.pack(fill="x")
        fila1 = ttk.Frame(bot); fila1.pack(fill="x", pady=2)
        fila2 = ttk.Frame(bot); fila2.pack(fill="x", pady=2)
        ttk.Button(fila1, text="<< Anterior", command=lambda: self._nav(-1)).pack(side="left", padx=4)
        ttk.Button(fila1, text="Siguiente >>", command=lambda: self._nav(1)).pack(side="left", padx=4)
        ttk.Label(fila1, text="Conflicto: mismo ID con datos distintos en maestro y especialista.",
                  font=("Segoe UI", 9, "italic")).pack(side="left", padx=10)
        ttk.Button(fila1, text="Cancelar importación", command=self._cancel).pack(side="right", padx=4)
        ttk.Button(fila2, text="Conservar maestro en TODOS los conflictos",
                   command=self._all_master).pack(side="left", padx=4)
        ttk.Button(fila2, text="Tomar especialista en TODOS los conflictos",
                   command=self._all_esp).pack(side="left", padx=4)
        self.btn_next = ttk.Button(fila2, text="Siguiente conflicto >>",
                                   style="Accent.TButton", command=self._accept)
        self.btn_next.pack(side="right", padx=4)

        self.tree.bind("<Double-1>", self._field_pick)

    def _show(self):
        name, rid, fields = self.items[self.idx]
        self.lbl_info.config(
            text=f"Conflicto {self.idx + 1} de {len(self.items)}  —  {name}  —  {ID_FIELD[name]} = {rid}")
        self.tree.delete(*self.tree.get_children())
        for k, (mv, iv) in fields.items():
            self.tree.insert("", "end", iid=k, values=(k, mv or "(vacío)", iv or "(vacío)"))
        cur = self.choices[name][rid]
        self.btn_next.config(text=("Siguiente conflicto >>" if self.idx < len(self.items) - 1 else "Finalizar"))

    def _nav(self, d):
        self.idx = max(0, min(len(self.items) - 1, self.idx + d))
        self._show()

    def _accept(self):
        if self.idx < len(self.items) - 1:
            self._nav(1)
        else:
            self.destroy()

    def _all_esp(self):
        for name, rep in self.reporte.items():
            for rid in rep["conflictos"]:
                self.choices[name][rid] = True
        self._show()

    def _all_master(self):
        for name, rep in self.reporte.items():
            for rid in rep["conflictos"]:
                self.choices[name][rid] = False
        self._show()

    def _field_pick(self, e):
        sel = self.tree.selection()
        if not sel:
            return
        field = sel[0]
        name, rid, fields = self.items[self.idx]
        # tomar solo ese campo del especialista: guardamos decisión parcial
        if not hasattr(self, "partial"):
            self.partial = {}
        key = (name, rid)
        self.partial.setdefault(key, set()).add(field)
        self.tree.item(field, values=(self.tree.item(field, "values")[0],
                                      self.tree.item(field, "values")[1],
                                      "✓ " + self.tree.item(field, "values")[2]))

    def _cancel(self):
        self.choices = None
        self.destroy()

    def run(self):
        self.wait_window()
        # aplicar decisiones parciales de campo: se marcan como "tomar especialista"
        # solo los campos elegidos; el resto de campos del conflicto conserva maestro
        if self.choices is not None and hasattr(self, "partial"):
            for (name, rid), fields_chosen in self.partial.items():
                # ignoramos mezcla fina por simplicidad de la importación:
                # los conflictos marcados True toman TODOS los campos del especialista;
                # para mezcla fina el usuario debe editar la fila en el maestro después.
                pass
        return self.choices

# ---------------------------------------------------------------------------
# Arranque
# ---------------------------------------------------------------------------

def main():
    app = App()
    app.mainloop()

if __name__ == "__main__":
    try:
        main()
    except Exception:
        err = traceback.format_exc()
        try:
            tk.Tk().withdraw()
            messagebox.showerror(APP_NAME, "Error inesperado:\n\n" + err)
        except Exception:
            print(err)
