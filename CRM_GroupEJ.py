#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
CRM INTEGRAL | GROUP EJ MORALES S.U.R.L.  —  Aplicación de escritorio v2.0
===========================================================================
CRM completo e independiente de Excel:
  * Los datos se capturan y se guardan DENTRO de la aplicación (base de datos local).
  * Excel solo se usa para importar datos existentes y exportar libros/informes
    con la misma estructura de hojas y encabezados de los CRM originales.

Roles:
  - ESPECIALISTA: trabaja de forma individual y envía un paquete diario a jefatura.
  - JEFATURA: importa los paquetes, consolida el maestro, resuelve conflictos,
    revisa/aprueba ofertas y devuelve retroalimentación.
"""

import os
import sys
import re
import json
import copy
import shutil
import sqlite3
import datetime
import traceback

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
FROZEN = getattr(sys, "frozen", False)          # True cuando corre como .exe


def resource_path(rel):
    """Ruta a recursos (plantillas) dentro del .exe o junto al script."""
    if FROZEN:
        junto = os.path.join(os.path.dirname(sys.executable), rel)
        if os.path.exists(junto):
            return junto
        return os.path.join(getattr(sys, "_MEIPASS", BASE_DIR), rel)
    return os.path.join(BASE_DIR, rel)


if not FROZEN:
    _LIBS = os.path.join(BASE_DIR, "libs")
    if os.path.isdir(_LIBS) and _LIBS not in sys.path:
        sys.path.insert(0, _LIBS)

import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

APP_VERSION = "2.0"
APP_NAME = "CRM Group EJ Morales"
APP_TITLE = "CRM INTEGRAL | GROUP EJ MORALES S.U.R.L."

# ===========================================================================
# MODELO DE DATOS (idéntico a los libros Excel originales)
# ===========================================================================

DATA_SHEETS = [
    "Equipo", "Clientes", "Proveedores", "Prospectos", "Actividades", "Ofertas",
    "Compras", "Importaciones", "Participantes_importacion", "Embarques",
    "Contenedores", "Contenido_contenedores", "Documentos_importacion",
    "Costos_importacion", "Contratos", "Ventas", "Logistica", "Facturas_Cobros",
    "Incidencias", "Rendimiento",
]

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

ID_PREFIX = {
    "Equipo": "USR", "Clientes": "CLI", "Proveedores": "PRO", "Prospectos": "OPO",
    "Actividades": "ACT", "Ofertas": "OFE", "Compras": "CMP", "Importaciones": "IMP",
    "Participantes_importacion": "PAR", "Embarques": "EMB", "Contenedores": "CNT",
    "Contenido_contenedores": "LIN", "Documentos_importacion": "DOC",
    "Costos_importacion": "CST", "Contratos": "CTR", "Ventas": "VTA",
    "Logistica": "ENT", "Facturas_Cobros": "REG", "Incidencias": "INC",
}

# Módulos del menú lateral
MODULES = [
    ("Comercial", ["Clientes", "Prospectos", "Actividades", "Ofertas", "Contratos", "Ventas"]),
    ("Proveedores y compras", ["Proveedores", "Compras"]),
    ("Importaciones", ["Importaciones", "Participantes_importacion", "Embarques", "Contenedores",
                       "Contenido_contenedores", "Documentos_importacion", "Costos_importacion"]),
    ("Operaciones", ["Logistica", "Facturas_Cobros", "Incidencias"]),
    ("Gestión", ["Equipo", "Rendimiento"]),
]

SHEET_LABEL = {
    "Clientes": "Clientes", "Prospectos": "Oportunidades", "Actividades": "Actividades",
    "Ofertas": "Ofertas", "Contratos": "Contratos", "Ventas": "Ventas / Operaciones",
    "Proveedores": "Proveedores", "Compras": "Compras",
    "Importaciones": "Importaciones", "Participantes_importacion": "Participantes",
    "Embarques": "Embarques", "Contenedores": "Contenedores",
    "Contenido_contenedores": "Contenido de contenedores",
    "Documentos_importacion": "Documentos", "Costos_importacion": "Costos",
    "Logistica": "Logística y entregas", "Facturas_Cobros": "Facturas y cobros",
    "Incidencias": "Incidencias", "Equipo": "Equipo", "Rendimiento": "Rendimiento",
}
SINGULAR = {
    "Clientes": "cliente", "Prospectos": "oportunidad", "Actividades": "actividad",
    "Ofertas": "oferta", "Contratos": "contrato", "Ventas": "operación de venta",
    "Proveedores": "proveedor", "Compras": "compra", "Importaciones": "importación",
    "Participantes_importacion": "participante", "Embarques": "embarque",
    "Contenedores": "contenedor", "Contenido_contenedores": "línea de contenido",
    "Documentos_importacion": "documento", "Costos_importacion": "costo",
    "Logistica": "entrega", "Facturas_Cobros": "factura / cobro",
    "Incidencias": "incidencia", "Equipo": "miembro del equipo", "Rendimiento": "registro de rendimiento",
}

KEY_COLS = {
    "Clientes": ["ID_Cliente", "Razon_social", "Provincia", "Contacto", "Telefono", "Potencial",
                 "Estado_cliente", "Fecha_seguimiento", "Proxima_accion"],
    "Proveedores": ["ID_Proveedor", "Razon_social", "Pais", "Productos", "Contacto",
                    "Estado_validacion", "Responsable"],
    "Prospectos": ["ID_Oportunidad", "Cliente", "Comercial_responsable", "Producto_servicio",
                   "Valor_potencial", "Probabilidad_pct", "Etapa", "Fecha_seguimiento", "Proxima_accion"],
    "Actividades": ["ID_Actividad", "Cliente_proveedor", "Tipo_actividad", "Fecha_hora",
                    "Responsable", "Estado", "Fecha_seguimiento"],
    "Ofertas": ["ID_Oferta", "Cliente", "Producto_servicio", "Importe", "Moneda",
                "Comercial_responsable", "Estado_revision", "Version"],
    "Compras": ["ID_Compra", "Proveedor", "Producto", "Cantidad", "Estado", "Fecha_orden"],
    "Importaciones": ["ID_Importacion", "Proveedor", "Producto", "Cliente_destino", "Estado", "ETA", "ATA"],
    "Participantes_importacion": ["ID_Participante", "ID_Importacion", "Entidad", "Funcion", "Estado_validacion"],
    "Embarques": ["ID_Embarque", "ID_Importacion", "Naviera", "BL", "ETA", "ATA", "Estado"],
    "Contenedores": ["ID_Contenedor", "Numero_contenedor", "ID_Importacion", "Estado", "ETA", "ATA",
                     "Ubicacion_actual"],
    "Contenido_contenedores": ["ID_Linea", "ID_Contenedor", "Producto", "Cantidad", "Unidad"],
    "Documentos_importacion": ["ID_Documento", "ID_Importacion", "Tipo_documento", "Numero", "Estado",
                               "Fecha_vencimiento"],
    "Costos_importacion": ["ID_Costo", "ID_Importacion", "Concepto", "Moneda", "Importe_estimado",
                           "Importe_facturado", "Saldo_pendiente"],
    "Contratos": ["ID_Contrato", "Numero_contrato", "Cliente", "Cantidad_pendiente", "Fecha_vencimiento", "Estado"],
    "Ventas": ["ID_Operacion", "Cliente", "Producto", "Importe", "Estado_logistico", "Estado_cobro", "Fecha_prevista"],
    "Logistica": ["ID_Entrega", "ID_Operacion", "Cliente", "Producto", "Fecha_prevista", "Estado"],
    "Facturas_Cobros": ["ID_Registro", "Tipo_registro", "Nombre", "Numero_factura", "Importe", "Saldo",
                        "Fecha_vencimiento", "Dias_atraso", "Estado"],
    "Incidencias": ["ID_Incidencia", "Area", "Descripcion", "Prioridad", "Fecha_limite", "Estado"],
    "Equipo": ["ID_Usuario", "Nombre", "Cargo", "Correo", "Estado", "Ultima_consolidacion"],
    "Rendimiento": ["Periodo_inicio", "Periodo_fin", "Comercial", "Actividades", "Ofertas_preparadas",
                    "Ventas_cerradas", "Valor_ventas", "Meta_ventas", "Cumplimiento_pct"],
}

STATUS_COL = {
    "Clientes": "Estado_cliente", "Prospectos": "Etapa", "Actividades": "Estado",
    "Ofertas": "Estado_revision", "Importaciones": "Estado", "Embarques": "Estado",
    "Contenedores": "Estado", "Incidencias": "Estado", "Contratos": "Estado",
    "Logistica": "Estado", "Facturas_Cobros": "Estado", "Compras": "Estado",
    "Documentos_importacion": "Estado", "Proveedores": "Estado_validacion",
}

# ---- tipos de campo ------------------------------------------------------
DATETIME_COLS = {"Fecha_hora", "Fecha_modificacion"}
NUMERIC_COLS = {
    "Cantidad", "Cantidad_contratada", "Cantidad_ejecutada", "Cantidad_pendiente",
    "Cantidad_solicitada", "Cantidad_confirmada", "Cantidad_entregada", "Cantidad_rechazada",
    "Valor_potencial", "Probabilidad_pct", "Cumplimiento_pct", "Precio_propuesto", "Precio_unitario",
    "Precio_autorizado", "Importe", "Importe_estimado", "Importe_facturado", "Importe_pagado",
    "Saldo_pendiente", "Saldo", "Anticipo", "Peso_neto", "Peso_bruto", "Volumen", "Valor_FOB",
    "Flete", "Seguro", "Valor_CIF", "Tipo_cambio", "Pagado_confirmado", "Dias_atraso", "Plazo_dias",
    "Version", "Prospectos", "Clientes_contactados", "Actividades", "Oportunidades_abiertas",
    "Ofertas_preparadas", "Ofertas_aprobadas", "Ventas_cerradas", "Valor_ventas", "Cobros_confirmados",
    "Entregas", "Proveedores_captados", "Proveedores_validados", "Compras", "Importaciones",
    "Seguimientos_vencidos", "Meta_ventas",
}
LONGTEXT = {
    "Observaciones", "Descripcion", "Necesidad", "Objeciones", "Competencia", "Incidencias",
    "Solucion", "Observaciones_jefatura", "Documentacion", "Documentacion_autorizaciones",
    "Compromiso", "Accion_requerida", "Contenido_general", "Productos_incluidos", "Productos",
    "Productos_interes", "Referencias", "Licencias_certificaciones", "Condiciones_pago", "Objeto",
    "Condiciones_entrega", "Transbordos",
}
LONG_SHEET = {("Actividades", "Resultado")}


def is_date_col(c):
    return (c.startswith("Fecha") and c not in DATETIME_COLS) or c in ("ETA", "ATA", "Periodo_inicio", "Periodo_fin")


def is_datetime_col(c):
    return c in DATETIME_COLS


# ---- listas desplegables (editables salvo que vengan de CATALOGOS) ---------
PROVINCIAS = ["Pinar del Río", "Artemisa", "La Habana", "Mayabeque", "Matanzas", "Cienfuegos",
              "Villa Clara", "Sancti Spíritus", "Ciego de Ávila", "Camagüey", "Las Tunas",
              "Holguín", "Granma", "Santiago de Cuba", "Guantánamo", "Isla de la Juventud"]
CHOICES = {
    "Moneda": ["USD", "EUR", "CUP", "MLC"],
    "Unidad": ["u", "kg", "t", "lb", "m", "m²", "m³", "l", "caja", "paleta", "contenedor", "servicio"],
    "Provincia": PROVINCIAS,
    "Pais": ["Cuba", "China", "España", "Panamá", "México", "Brasil", "Estados Unidos", "Canadá",
             "Vietnam", "Turquía", "Italia", "Alemania", "Otro"],
    "Pais_origen": ["China", "España", "Panamá", "México", "Brasil", "Estados Unidos", "Canadá",
                    "Vietnam", "Turquía", "Italia", "Alemania", "Otro"],
    "Pais_procedencia": ["China", "España", "Panamá", "México", "Brasil", "Estados Unidos", "Canadá",
                         "Vietnam", "Turquía", "Italia", "Alemania", "Otro"],
    "Potencial": ["Alto", "Medio", "Bajo"],
    "Estado_validacion": ["Pendiente de verificación", "Verificado", "Rechazado"],
    "Estado_cliente": ["Prospecto", "Activo", "Inactivo", "Suspendido"],
    "Tipo_entidad": ["Empresa estatal", "MIPYME", "Cooperativa", "TCP", "Entidad presupuestada", "Otra"],
    "Tipo_actividad": ["Llamada", "Reunión", "Visita", "Correo", "WhatsApp", "Presentación",
                       "Seguimiento", "Otra"],
    "Fuente": ["Referencia", "Búsqueda propia", "Cliente existente", "Feria / evento", "Redes / web", "Otra"],
    "Incoterm": ["EXW", "FCA", "FAS", "FOB", "CFR", "CIF", "CPT", "CIP", "DAP", "DPU", "DDP"],
    "Incoterms": ["EXW", "FCA", "FAS", "FOB", "CFR", "CIF", "CPT", "CIP", "DAP", "DPU", "DDP"],
    "Forma_pago": ["Transferencia", "Carta de crédito", "Efectivo", "Pago anticipado", "Crédito"],
    "Requiere_Direccion": ["No", "Sí"],
    "Autorizacion_Direccion": ["Pendiente", "Autorizada", "Rechazada"],
    "Resultado_cliente": ["Pendiente", "Aceptada", "Rechazada"],
    "Tipo_proveedor": ["Fabricante", "Distribuidor", "Agente", "Intermediario", "Otro"],
    "Tipo_costo": ["Flete", "Seguro", "Aduana", "Portuario", "Transporte nacional", "Almacenaje", "Otro"],
    "Tipo_documento": ["BL", "Factura comercial", "Lista de empaque", "Certificado de origen",
                       "Certificado de calidad", "Póliza de seguro", "Declaración aduanal", "Otro"],
    "Funcion": ["Proveedor", "Importadora", "Consignatario", "Agente de carga", "Naviera",
                "Transportista", "Aduana", "Otro"],
    "Tipo_tamano": ["20' DC", "40' DC", "40' HC", "20' Reefer", "40' Reefer", "Otro"],
    "Area": ["Comercial", "Compras", "Logística", "Transporte", "Almacén", "Economía", "Aduana", "Dirección"],
    "Tipo_registro": ["Cobro (cliente)", "Pago (proveedor)"],
    "Confirmacion_Economia": ["Pendiente", "Confirmado"],
    "Estado_financiero": ["Pendiente", "Parcial", "Completado"],
    "Estado_logistico": ["Pendiente", "Parcial", "Completado"],
    "Estado_facturacion": ["Pendiente", "Parcial", "Completado"],
    "Estado_cobro": ["Pendiente", "Parcial", "Completado"],
    "Reserva": ["Sí", "No"],
}
CHOICES_SHEET = {
    ("Actividades", "Estado"): ["Pendiente", "Realizada", "Cancelada"],
    ("Incidencias", "Estado"): ["Abierta", "En seguimiento", "Resuelta", "Cerrada"],
    ("Contratos", "Estado"): ["Borrador", "Vigente", "Vencido", "Cumplido", "Cancelado"],
    ("Documentos_importacion", "Estado"): ["Pendiente", "Recibido", "Validado", "Vencido"],
    ("Logistica", "Estado"): ["Pendiente", "Programada", "En ruta", "Entregada", "Entrega parcial", "Incidencia"],
    ("Facturas_Cobros", "Estado"): ["Pendiente", "Parcial", "Pagada", "Vencida"],
    ("Compras", "Estado"): ["Pendiente", "Ordenada", "En producción", "Suministrada", "Cancelada"],
    ("Equipo", "Estado"): ["Activo", "Inactivo"],
    ("Prospectos", "Resultado"): ["Pendiente de verificación", "En curso", "Ganada", "Perdida"],
}
CATALOG_MAP = {
    ("Prospectos", "Etapa"): "Etapa", ("Ofertas", "Estado_revision"): "Estado_revision",
    ("Importaciones", "Estado"): "Estado_importacion", ("Embarques", "Estado"): "Estado_embarque",
    ("Contenedores", "Estado"): "Estado_contenedor", ("Incidencias", "Prioridad"): "Prioridad",
}

# Campos que apuntan a otras hojas
REF_SOURCE = {
    "ID_Cliente": "Clientes", "ID_Proveedor": "Proveedores", "ID_Oportunidad": "Prospectos",
    "ID_Operacion": "Ventas", "ID_Importacion": "Importaciones", "ID_Embarque": "Embarques",
    "ID_Contenedor": "Contenedores", "ID_Oferta": "Ofertas", "ID_Contrato": "Contratos",
}
AUTOFILL = {
    "ID_Cliente": ["Cliente", "Cliente_proveedor"], "ID_Proveedor": ["Proveedor"],
    "ID_Operacion": ["Cliente"], "ID_Cliente_Proveedor": ["Nombre"],
}
RESPONSIBLE_COLS = ["Comercial_responsable", "Responsable", "Responsable_actividad"]

REQUIRED = {
    "Clientes": ["Razon_social"], "Proveedores": ["Razon_social"],
    "Prospectos": ["Cliente", "Etapa"], "Actividades": ["Tipo_actividad", "Fecha_hora", "Responsable"],
    "Ofertas": ["Cliente", "Producto_servicio", "Estado_revision"],
    "Compras": ["Proveedor", "Producto"], "Importaciones": ["Proveedor", "Producto"],
    "Participantes_importacion": ["ID_Importacion", "Entidad"], "Embarques": ["ID_Importacion"],
    "Contenedores": ["ID_Importacion", "Numero_contenedor"],
    "Contenido_contenedores": ["ID_Contenedor", "Producto"], "Documentos_importacion": ["Tipo_documento"],
    "Costos_importacion": ["Concepto"], "Contratos": ["Cliente"], "Ventas": ["Cliente"],
    "Logistica": ["Cliente"], "Facturas_Cobros": ["Tipo_registro", "Nombre"],
    "Incidencias": ["Descripcion"], "Equipo": ["Nombre"],
    "Rendimiento": ["Periodo_inicio", "Periodo_fin", "Comercial"],
}

COMPUTED = {
    "Ofertas": ["Importe"], "Ventas": ["Importe"], "Importaciones": ["Valor_CIF"],
    "Costos_importacion": ["Saldo_pendiente"], "Contratos": ["Cantidad_pendiente"],
    "Facturas_Cobros": ["Saldo", "Dias_atraso"], "Compras": ["Saldo"],
    "Rendimiento": ["Cumplimiento_pct"],
}
AUTO_READONLY = {"Fecha_modificacion"}

# ---- flujo de ofertas -----------------------------------------------------
OFERTAS_PENDIENTES = ["Pendiente de revisión", "En revisión", "Modificaciones solicitadas",
                      "Pendiente de autorización de Dirección"]
OFERTA_APROBADAS = {"Aprobada por jefatura", "Autorizada para presentar"}
OFERTA_POST = {"Presentada al cliente", "Aceptada", "Rechazada por cliente", "Vencida"}
OFERTA_ALLOWED = {   # transiciones que puede hacer el ESPECIALISTA
    "Borrador": {"Pendiente de revisión", "Cancelada"},
    "Modificaciones solicitadas": {"Pendiente de revisión", "Cancelada"},
    "Pendiente de revisión": {"Cancelada"},
    "Aprobada por jefatura": {"Presentada al cliente", "Cancelada"},
    "Autorizada para presentar": {"Presentada al cliente", "Cancelada"},
    "Presentada al cliente": {"Aceptada", "Rechazada por cliente", "Vencida", "Cancelada"},
}
OFERTA_MASTER_FIELDS = {"Revisor", "Observaciones_jefatura", "Requiere_Direccion",
                        "Autorizacion_Direccion", "Fecha_aprobacion"}
OFERTA_REVIEW_FIELDS = ["Estado_revision", "Revisor", "Observaciones_jefatura", "Requiere_Direccion",
                        "Autorizacion_Direccion", "Fecha_aprobacion"]
LOG_HEADERS = ["Fecha", "Especialista", "Paquete", "Nuevas", "Actualizadas",
               "Conflictos_a_favor_especialista", "Conflictos_conservados_maestro", "Ignoradas", "Notas"]

COLOR_BG = "#f4f6f9"
COLOR_HEADER = "#1f3a5f"
COLOR_ACCENT = "#2e86ab"
COLOR_SIDEBAR = "#16293f"

# ===========================================================================
# UTILIDADES
# ===========================================================================


def now_iso():
    return datetime.datetime.now().strftime("%Y-%m-%dT%H:%M:%S")


def now_min():
    return datetime.datetime.now().strftime("%Y-%m-%d %H:%M")


def today_str():
    return datetime.date.today().strftime("%Y-%m-%d")


def to_text(v):
    if v is None:
        return ""
    if isinstance(v, datetime.datetime):
        return v.strftime("%Y-%m-%d %H:%M") if (v.hour or v.minute) else v.strftime("%Y-%m-%d")
    if isinstance(v, datetime.date):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, float):
        return fmt_num(v)
    return str(v)


def to_num(s):
    if s is None:
        return None
    if isinstance(s, (int, float)):
        return float(s)
    t = str(s).strip().replace(" ", "")
    if t == "":
        return None
    if "," in t and "." not in t:
        t = t.replace(",", ".")
    elif "," in t and "." in t:
        t = t.replace(",", "")
    try:
        return float(t)
    except ValueError:
        return None


def fmt_num(x):
    if x is None:
        return ""
    if abs(x - round(x)) < 1e-9:
        return str(int(round(x)))
    return f"{x:.2f}".rstrip("0").rstrip(".")


def parse_date(s):
    s = (s or "").strip()
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%Y/%m/%d"):
        try:
            return datetime.datetime.strptime(s[:10] if fmt.startswith(("%Y-%m-%d", "%Y/%m/%d", "%d/%m/%Y", "%d-%m-%Y")) and len(s) >= 10 else s, fmt).date()
        except ValueError:
            continue
    return None


def parse_datetime(s):
    s = (s or "").strip()
    for fmt in ("%Y-%m-%d %H:%M", "%d/%m/%Y %H:%M", "%Y-%m-%dT%H:%M:%S", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.datetime.strptime(s, fmt)
        except ValueError:
            continue
    d = parse_date(s)
    if d and len(s) <= 10:
        return datetime.datetime(d.year, d.month, d.day)
    return None


def coerce_cell(col, value):
    """Tipo adecuado para Excel (fechas reales y números solo en columnas numéricas)."""
    if value is None:
        return None
    if not isinstance(value, str):
        return value
    s = value.strip()
    if s == "":
        return None
    if is_datetime_col(col):
        d = parse_datetime(s)
        return d if d else s
    if is_date_col(col):
        d = parse_date(s)
        return datetime.datetime(d.year, d.month, d.day) if d else s
    if col in NUMERIC_COLS:
        n = to_num(s)
        if n is not None:
            return int(n) if abs(n - round(n)) < 1e-9 else n
    return s


def clean(row, headers):
    return {h: to_text(row.get(h, "")) for h in headers}


def rows_differ(a, b, ignore=("Fecha_modificacion", "_mod")):
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


def norm_name(s):
    s = (s or "").lower().strip()
    s = re.sub(r"[^a-z0-9áéíóúñü ]", "", s)
    return re.sub(r"\s+", " ", s)


def open_file(path):
    try:
        if sys.platform.startswith("win"):
            os.startfile(path)            # noqa
        elif sys.platform == "darwin":
            os.system(f'open "{path}"')
        else:
            os.system(f'xdg-open "{path}" >/dev/null 2>&1 &')
    except Exception:
        pass


def data_dir():
    custom = os.environ.get("CRM_DATA_DIR")
    if custom:
        d = custom
    elif os.name == "nt":
        d = os.path.join(os.environ.get("APPDATA") or os.path.expanduser("~"), "CRM_GroupEJ")
    else:
        d = os.path.join(os.path.expanduser("~"), ".crm_groupej")
    os.makedirs(d, exist_ok=True)
    return d


def load_config():
    p = os.path.join(data_dir(), "config.json")
    if os.path.exists(p):
        try:
            with open(p, encoding="utf-8") as f:
                c = json.load(f)
            c.setdefault("perfiles", [])
            c.setdefault("ultimo", 0)
            return c
        except Exception:
            pass
    return {"perfiles": [], "ultimo": 0}


def save_config(cfg):
    with open(os.path.join(data_dir(), "config.json"), "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=1)


def suggest_code(nombre):
    m = re.search(r"(\d+)", nombre or "")
    letters = "".join(w[0] for w in re.findall(r"[A-Za-zÁÉÍÓÚáéíóúÑñ]+", nombre or ""))[:2].upper()
    if m:
        return (letters[:1] or "C") + m.group(1).zfill(2)[:3]
    return (letters or "USR").ljust(3, "X")[:3].upper()


# ===========================================================================
# CÁLCULOS, NORMALIZACIÓN Y VALIDACIÓN
# ===========================================================================


def compute_row(sheet, row):
    n = lambda k: to_num(row.get(k))
    setv = lambda k, v: row.__setitem__(k, fmt_num(v)) if (k in row and v is not None) else None
    if sheet == "Ofertas":
        c, p = n("Cantidad"), n("Precio_propuesto")
        if c is not None and p is not None:
            setv("Importe", c * p)
    elif sheet == "Ventas":
        c, p = n("Cantidad_contratada"), n("Precio_autorizado")
        if c is not None and p is not None:
            setv("Importe", c * p)
    elif sheet == "Importaciones":
        fob, fl, se = n("Valor_FOB"), n("Flete"), n("Seguro")
        if fob is not None:
            setv("Valor_CIF", fob + (fl or 0) + (se or 0))
    elif sheet == "Costos_importacion":
        f, p = n("Importe_facturado"), n("Importe_pagado")
        if f is not None:
            setv("Saldo_pendiente", f - (p or 0))
    elif sheet == "Contratos":
        c, e = n("Cantidad_contratada"), n("Cantidad_ejecutada")
        if c is not None:
            setv("Cantidad_pendiente", c - (e or 0))
    elif sheet == "Compras":
        c, p, a = n("Cantidad"), n("Precio_unitario"), n("Anticipo")
        if c is not None and p is not None:
            setv("Saldo", c * p - (a or 0))
    elif sheet == "Facturas_Cobros":
        imp, pag = n("Importe"), n("Pagado_confirmado")
        if imp is not None:
            saldo = imp - (pag or 0)
            setv("Saldo", saldo)
            venc = parse_date(row.get("Fecha_vencimiento", ""))
            dias = 0
            if saldo > 0 and venc and venc < datetime.date.today():
                dias = (datetime.date.today() - venc).days
            setv("Dias_atraso", dias)
    elif sheet == "Rendimiento":
        v, m = n("Valor_ventas"), n("Meta_ventas")
        if v is not None and m:
            setv("Cumplimiento_pct", v / m * 100)
    return row


def normalize_row(sheet, row):
    """Normaliza fechas a formato ISO. Devuelve lista de errores."""
    errs = []
    for h, v in list(row.items()):
        if not v or not isinstance(v, str):
            continue
        if is_datetime_col(h):
            d = parse_datetime(v)
            if d is None:
                errs.append(f"«{h}»: fecha/hora no válida (use AAAA-MM-DD HH:MM).")
            else:
                row[h] = d.strftime("%Y-%m-%d %H:%M")
        elif is_date_col(h):
            d = parse_date(v)
            if d is None:
                errs.append(f"«{h}»: fecha no válida (use AAAA-MM-DD, por ejemplo {today_str()}).")
            else:
                row[h] = d.strftime("%Y-%m-%d")
        elif h in NUMERIC_COLS:
            nv = to_num(v)
            if nv is None:
                errs.append(f"«{h}»: debe ser un número.")
            else:
                row[h] = fmt_num(nv)
    return errs


def find_duplicates(store, sheet):
    """Posibles duplicados por NIT/ID fiscal o razón social normalizada."""
    idf = ID_FIELD[sheet]
    key_fields = ["NIT_ID"] if sheet == "Clientes" else ["Identificacion_fiscal"]
    groups = {}
    for r in store.sheet(sheet)["rows"]:
        keys = set()
        nm = norm_name(r.get("Razon_social", "") or r.get("Nombre_comercial", ""))
        if nm:
            keys.add("n:" + nm)
        for kf in key_fields:
            v = re.sub(r"\W", "", to_text(r.get(kf, ""))).lower()
            if v:
                keys.add("f:" + v)
        for k in keys:
            groups.setdefault(k, []).append(r)
    out, seen = [], set()
    for k, rows in groups.items():
        ids = tuple(sorted(set(str(x.get(idf, "")) for x in rows)))
        if len(ids) > 1 and ids not in seen:
            seen.add(ids)
            out.append(rows)
    return out


def validate_row(store, sheet, row, orig, is_new, role):
    errs, warns = [], []
    for h in REQUIRED.get(sheet, []):
        if h in row and not str(row.get(h, "")).strip():
            errs.append(f"El campo «{h}» es obligatorio.")
    errs += normalize_row(sheet, row)
    p = to_num(row.get("Probabilidad_pct"))
    if p is not None and not (0 <= p <= 100):
        errs.append("«Probabilidad_pct» debe estar entre 0 y 100.")
    for h in ("Correo",):
        v = row.get(h, "")
        if v and "@" not in v:
            warns.append(f"«{h}» no parece un correo válido.")
    if sheet == "Prospectos":
        if row.get("Etapa") == "Perdida" and not row.get("Motivo_perdida"):
            warns.append("La oportunidad está «Perdida» sin «Motivo_perdida».")
        a, b = parse_date(row.get("Fecha_captacion", "")), parse_date(row.get("Fecha_cierre_estimada", ""))
        if a and b and b < a:
            warns.append("La fecha de cierre estimada es anterior a la fecha de captación.")
    if sheet == "Ofertas":
        old = (orig or {}).get("Estado_revision", "") if orig else ""
        new = row.get("Estado_revision", "")
        if new in OFERTA_POST and new != old:
            okold = old in OFERTA_APROBADAS or old in OFERTA_POST
            if not okold:
                errs.append("No se puede presentar la oferta al cliente antes de completar las "
                            "autorizaciones (revisión de jefatura y, si aplica, de Dirección).")
        if row.get("Requiere_Direccion") == "Sí" and new in OFERTA_POST and \
                row.get("Autorizacion_Direccion") != "Autorizada":
            errs.append("La oferta requiere autorización de Dirección antes de presentarse.")
    if sheet == "Clientes" and (is_new or not orig or
                                norm_name(orig.get("Razon_social", "")) != norm_name(row.get("Razon_social", ""))):
        nm = norm_name(row.get("Razon_social", ""))
        nit = re.sub(r"\W", "", row.get("NIT_ID", "")).lower()
        for r in store.sheet("Clientes")["rows"]:
            if orig is not None and r is orig:
                continue
            if r.get("ID_Cliente") == row.get("ID_Cliente"):
                continue
            if (nm and norm_name(r.get("Razon_social", "")) == nm) or \
                    (nit and re.sub(r"\W", "", to_text(r.get("NIT_ID", ""))).lower() == nit):
                warns.append(f"Posible cliente duplicado: {r.get('ID_Cliente')} — {r.get('Razon_social')}.")
                break
    return errs, warns


def is_overdue(sheet, row):
    fs = to_text(row.get("Fecha_seguimiento", ""))[:10]
    if not fs or fs >= today_str():
        return False
    if sheet == "Prospectos" and row.get("Etapa") in ("Ganada", "Perdida", "Suspendida"):
        return False
    if sheet == "Actividades" and row.get("Estado") in ("Realizada", "Cancelada"):
        return False
    return True


def display_name(sheet, row):
    if sheet in ("Clientes", "Proveedores"):
        return row.get("Razon_social") or row.get("Nombre_comercial") or ""
    if sheet in ("Prospectos", "Ventas", "Contratos", "Ofertas"):
        return row.get("Cliente", "")
    if sheet == "Importaciones":
        return f"{row.get('Producto', '')} ({row.get('Proveedor', '')})"
    if sheet == "Embarques":
        return row.get("BL", "") or row.get("Buque", "")
    if sheet == "Contenedores":
        return row.get("Numero_contenedor", "")
    return ""


# ===========================================================================
# LIBROS: base común, Excel y base de datos local
# ===========================================================================


class Book:
    def __init__(self):
        self.hojas = {}
        self.catalogos = {}

    def sheet(self, name):
        if name not in self.hojas:
            self.hojas[name] = {"headers": [], "rows": []}
        return self.hojas[name]

    def _changed(self, sheet):
        pass

    def next_id(self, sheet_name, code=""):
        info = self.sheet(sheet_name)
        idf, prefix = ID_FIELD.get(sheet_name), ID_PREFIX.get(sheet_name)
        if not idf or not prefix:
            return ""
        pre = f"{prefix}-{code}-" if code else f"{prefix}-"
        max_n = 0
        for r in info["rows"]:
            v = str(r.get(idf, ""))
            if v.startswith(pre):
                try:
                    max_n = max(max_n, int(v[len(pre):]))
                except ValueError:
                    pass
        return f"{pre}{max_n + 1:04d}"

    def find_row(self, sheet_name, id_value):
        idf = ID_FIELD.get(sheet_name)
        if not idf:
            return None, -1
        for i, r in enumerate(self.sheet(sheet_name)["rows"]):
            if str(r.get(idf, "")) == str(id_value):
                return r, i
        return None, -1

    def upsert(self, sheet_name, row, original=None, stamp=True):
        info = self.sheet(sheet_name)
        if stamp:
            if "Fecha_modificacion" in info["headers"]:
                row["Fecha_modificacion"] = now_min()
            row["_mod"] = now_iso()
        idf = ID_FIELD.get(sheet_name)
        res = "nueva"
        if idf:
            _, idx = self.find_row(sheet_name, row.get(idf, ""))
            if idx >= 0:
                info["rows"][idx] = row
                res = "actualizada"
            else:
                info["rows"].append(row)
        else:
            idx = next((i for i, r in enumerate(info["rows"]) if r is original), -1)
            if idx >= 0:
                info["rows"][idx] = row
                res = "actualizada"
            else:
                info["rows"].append(row)
        self._changed(sheet_name)
        return res

    def delete_obj(self, sheet_name, row):
        rows = self.sheet(sheet_name)["rows"]
        for i, r in enumerate(rows):
            if r is row:
                rows.pop(i)
                self._changed(sheet_name)
                return True
        return False

    def commit_all(self):
        pass


class ExcelBook(Book):
    """Lectura / escritura de libros Excel con la estructura original."""

    def __init__(self, path):
        super().__init__()
        self.path = path
        self.load()

    def load(self):
        self.hojas, self.catalogos = {}, {}
        wb = openpyxl.load_workbook(self.path, data_only=False)
        for name in DATA_SHEETS + ["LOG_CONSOLIDACION"]:
            if name not in wb.sheetnames:
                continue
            rows = list(wb[name].iter_rows(values_only=True))
            if not rows:
                continue
            headers = [str(h) if h is not None else f"Col{i + 1}" for i, h in enumerate(rows[0])]
            data = []
            for r in rows[1:]:
                if all(v is None or str(v).strip() == "" for v in r):
                    continue
                data.append({h: (to_text(r[i]) if i < len(r) and r[i] is not None else "")
                             for i, h in enumerate(headers)})
            self.hojas[name] = {"headers": headers, "rows": data}
        if "CATALOGOS" in wb.sheetnames:
            for t, v in wb["CATALOGOS"].iter_rows(min_row=2, values_only=True):
                if t and v:
                    self.catalogos.setdefault(str(t), []).append(str(v))
        wb.close()

    def save(self):
        wb = openpyxl.load_workbook(self.path, data_only=False)
        for name, info in self.hojas.items():
            if name not in wb.sheetnames:
                ws = wb.create_sheet(name)
                for c, h in enumerate(info["headers"], 1):
                    cell = ws.cell(1, c, h)
                    cell.font = Font(bold=True, color="FFFFFF")
                    cell.fill = PatternFill("solid", fgColor="1F3A5F")
            else:
                ws = wb[name]
                if ws.max_row > 1:
                    ws.delete_rows(2, ws.max_row - 1)
            for r_i, row in enumerate(info["rows"], 2):
                for c_i, h in enumerate(info["headers"], 1):
                    v = coerce_cell(h, row.get(h, ""))
                    if v is None:
                        continue
                    cell = ws.cell(r_i, c_i, v)
                    if isinstance(v, datetime.datetime):
                        cell.number_format = "yyyy-mm-dd hh:mm" if is_datetime_col(h) else "yyyy-mm-dd"
            # ajustar el rango de la tabla de Excel (si existe) para que incluya todas las filas
            last = max(2, len(info["rows"]) + 1)
            for t in ws.tables.values():
                t.ref = f"A1:{get_column_letter(len(info['headers']))}{last}"
                if t.autoFilter is not None:
                    t.autoFilter.ref = t.ref
        wb.save(self.path)
        wb.close()


class Store(Book):
    """Base de datos local (SQLite). Los cambios se guardan al instante."""

    def __init__(self, path):
        super().__init__()
        self.path = path
        self.conn = sqlite3.connect(path)
        self.meta = {}

    # ---- creación / apertura ------------------------------------------------
    @classmethod
    def create(cls, path, rol, nombre, codigo):
        st = cls(path)
        plantilla = "Plantilla_Maestro_Jefatura.xlsx" if rol == "jefatura" else "Plantilla_Libro_Individual.xlsx"
        xb = ExcelBook(resource_path(plantilla))
        st.catalogos = xb.catalogos
        for name in DATA_SHEETS:
            if name in xb.hojas:
                st.hojas[name] = {"headers": list(xb.hojas[name]["headers"]),
                                  "rows": [dict(r) for r in xb.hojas[name]["rows"]]}
        if rol == "jefatura":
            st.hojas["LOG_CONSOLIDACION"] = {"headers": list(LOG_HEADERS), "rows": []}
        st.conn.execute("CREATE TABLE IF NOT EXISTS meta (k TEXT PRIMARY KEY, v TEXT)")
        st.meta = {"rol": rol, "nombre": nombre, "codigo": codigo, "version": APP_VERSION,
                   "creado": now_iso()}
        for k, v in st.meta.items():
            st.conn.execute("INSERT OR REPLACE INTO meta VALUES (?,?)", (k, v))
        st.conn.execute("INSERT OR REPLACE INTO meta VALUES (?,?)", ("catalogos",
                                                                    json.dumps(st.catalogos, ensure_ascii=False)))
        st.conn.commit()
        # el propio usuario figura en el equipo
        eq = {h: "" for h in st.sheet("Equipo")["headers"]}
        eq.update({"ID_Usuario": st.next_id("Equipo", codigo), "Nombre": nombre,
                   "Cargo": "Jefatura" if rol == "jefatura" else "Comercial", "Estado": "Activo",
                   "Fecha_incorporacion": today_str(), "Libro_asignado": os.path.basename(path)})
        st.sheet("Equipo")["rows"].append(eq)
        st.commit_all()
        return st

    @classmethod
    def open(cls, path):
        st = cls(path)
        cur = st.conn.cursor()
        st.meta = dict(cur.execute("SELECT k, v FROM meta").fetchall())
        st.catalogos = json.loads(st.meta.get("catalogos", "{}"))
        tablas = [r[0] for r in cur.execute("SELECT name FROM sqlite_master WHERE type='table'")]
        for name in DATA_SHEETS + ["LOG_CONSOLIDACION"]:
            if name not in tablas:
                continue
            cols = [r[1] for r in cur.execute(f'PRAGMA table_info("{name}")') if r[1] not in ("_n", "_mod")]
            rows = []
            q = ",".join(f'"{c}"' for c in cols)
            for rec in cur.execute(f'SELECT {q}, "_mod" FROM "{name}" ORDER BY _n'):
                d = {c: (rec[i] or "") for i, c in enumerate(cols)}
                d["_mod"] = rec[-1] or ""
                rows.append(d)
            st.hojas[name] = {"headers": cols, "rows": rows}
        return st

    # ---- persistencia ------------------------------------------------------
    def _ensure_table(self, name):
        headers = self.hojas[name]["headers"]
        cols = ", ".join(f'"{h}" TEXT' for h in headers)
        self.conn.execute(f'CREATE TABLE IF NOT EXISTS "{name}" (_n INTEGER PRIMARY KEY AUTOINCREMENT, {cols}, "_mod" TEXT)')

    def commit_sheet(self, name):
        if name not in self.hojas or not self.hojas[name]["headers"]:
            return
        self._ensure_table(name)
        info = self.hojas[name]
        headers = info["headers"]
        cols = ",".join(f'"{h}"' for h in headers) + ',"_mod"'
        ph = ",".join("?" * (len(headers) + 1))
        data = [[to_text(r.get(h, "")) for h in headers] + [r.get("_mod", "")] for r in info["rows"]]
        with self.conn:
            self.conn.execute(f'DELETE FROM "{name}"')
            self.conn.executemany(f'INSERT INTO "{name}" ({cols}) VALUES ({ph})', data)

    def _changed(self, sheet):
        self.commit_sheet(sheet)

    def commit_all(self):
        for name in list(self.hojas):
            self.commit_sheet(name)

    def close(self):
        try:
            self.conn.close()
        except Exception:
            pass


def export_workbook(store, dest, rol):
    """Exporta TODO el CRM a un libro Excel con la estructura de los libros originales."""
    plantilla = "Plantilla_Maestro_Jefatura.xlsx" if rol == "jefatura" else "Plantilla_Libro_Individual.xlsx"
    shutil.copyfile(resource_path(plantilla), dest)
    xb = ExcelBook(dest)
    for name in DATA_SHEETS:
        if name in xb.hojas and name in store.hojas:
            hs = xb.hojas[name]["headers"]
            xb.hojas[name]["rows"] = [{h: r.get(h, "") for h in hs} for r in store.hojas[name]["rows"]]
    if rol == "jefatura" and "LOG_CONSOLIDACION" in store.hojas:
        xb.hojas["LOG_CONSOLIDACION"] = copy.deepcopy(store.hojas["LOG_CONSOLIDACION"])
    xb.save()


def export_sheet_to_xlsx(path, sheet_name, headers, rows, title=None):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet_name[:31]
    head_row = 1
    if title:
        ws.cell(1, 1, title).font = Font(bold=True, size=13, color="1F3A5F")
        ws.cell(2, 1, f"Exportado: {now_min()}").font = Font(italic=True, size=9, color="777777")
        head_row = 4
    for c, h in enumerate(headers, 1):
        cell = ws.cell(head_row, c, h)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F3A5F")
    for r, row in enumerate(rows, head_row + 1):
        for c, h in enumerate(headers, 1):
            v = coerce_cell(h, row.get(h, ""))
            if v is not None:
                cell = ws.cell(r, c, v)
                if isinstance(v, datetime.datetime):
                    cell.number_format = "yyyy-mm-dd hh:mm" if is_datetime_col(h) else "yyyy-mm-dd"
    for c, h in enumerate(headers, 1):
        w = max(12, min(38, len(h) + 4))
        for r in range(head_row + 1, min(head_row + 1 + len(rows), head_row + 200)):
            v = to_text(ws.cell(r, c).value)
            if v:
                w = max(w, min(42, len(v) + 2))
        ws.column_dimensions[get_column_letter(c)].width = w
    if rows:
        ws.auto_filter.ref = f"A{head_row}:{get_column_letter(len(headers))}{head_row + len(rows)}"
    ws.freeze_panes = ws.cell(head_row + 1, 1)
    wb.save(path)
    wb.close()


def make_backup(store, keep=14):
    d = os.path.join(data_dir(), "respaldos")
    os.makedirs(d, exist_ok=True)
    base = os.path.splitext(os.path.basename(store.path))[0]
    dest = os.path.join(d, f"{base}_{datetime.date.today():%Y%m%d}.db")
    store.conn.commit()
    src = sqlite3.connect(store.path)
    dst = sqlite3.connect(dest)
    with dst:
        src.backup(dst)
    src.close()
    dst.close()
    files = sorted(f for f in os.listdir(d) if f.startswith(base + "_"))
    for old in files[:-keep]:
        try:
            os.remove(os.path.join(d, old))
        except OSError:
            pass
    return dest


# ===========================================================================
# RENDIMIENTO Y RESÚMENES
# ===========================================================================


def calc_rendimiento(store, comercial, inicio, fin):
    def en_periodo(v):
        v = to_text(v)[:10]
        return bool(v) and inicio <= v <= fin

    def mine(r, *cols):
        return any(norm_name(r.get(c, "")) == norm_name(comercial) for c in cols)

    rows = lambda s: store.sheet(s)["rows"]
    pros = [r for r in rows("Prospectos") if mine(r, "Comercial_responsable")]
    acts = [r for r in rows("Actividades") if mine(r, "Responsable") and en_periodo(r.get("Fecha_hora"))]
    ofes = [r for r in rows("Ofertas") if mine(r, "Comercial_responsable")]
    ventas = [r for r in rows("Ventas") if mine(r, "Comercial_responsable")]
    ventas_p = [r for r in ventas if en_periodo(r.get("Fecha_aceptacion"))]
    ids_v = {r.get("ID_Operacion") for r in ventas}
    out = {
        "Periodo_inicio": inicio, "Periodo_fin": fin, "Comercial": comercial,
        "Prospectos": sum(1 for r in pros if en_periodo(r.get("Fecha_captacion"))),
        "Clientes_contactados": len({r.get("ID_Cliente") or r.get("Cliente_proveedor") for r in acts
                                     if (r.get("ID_Cliente") or r.get("Cliente_proveedor"))}),
        "Actividades": len(acts),
        "Oportunidades_abiertas": sum(1 for r in pros if r.get("Etapa") not in ("Ganada", "Perdida", "Suspendida", "")),
        "Ofertas_preparadas": sum(1 for r in ofes if en_periodo(r.get("Fecha_preparacion"))),
        "Ofertas_aprobadas": sum(1 for r in ofes if en_periodo(r.get("Fecha_aprobacion"))
                                 or (r.get("Estado_revision") in OFERTA_APROBADAS | OFERTA_POST
                                     and en_periodo(r.get("Fecha_preparacion")))),
        "Ventas_cerradas": len(ventas_p),
        "Valor_ventas": fmt_num(sum(to_num(r.get("Importe")) or 0 for r in ventas_p)),
        "Cobros_confirmados": sum(1 for r in rows("Facturas_Cobros")
                                  if "cobro" in (r.get("Tipo_registro", "").lower())
                                  and r.get("Confirmacion_Economia") and mine(r, "Responsable")
                                  and en_periodo(r.get("Fecha_pago_cobro"))),
        "Entregas": sum(1 for r in rows("Logistica") if r.get("ID_Operacion") in ids_v and en_periodo(r.get("Fecha_real"))),
        "Proveedores_captados": sum(1 for r in rows("Proveedores") if mine(r, "Responsable")),
        "Proveedores_validados": sum(1 for r in rows("Proveedores") if mine(r, "Responsable")
                                     and (r.get("Estado_validacion", "").lower().startswith(("verific", "valid")))),
        "Compras": sum(1 for r in rows("Compras") if mine(r, "Responsable") and en_periodo(r.get("Fecha_orden"))),
        "Importaciones": sum(1 for r in rows("Importaciones") if mine(r, "Comercial_responsable")
                             and en_periodo(r.get("Fecha_orden"))),
        "Seguimientos_vencidos": sum(1 for s, cols in (("Prospectos", ("Comercial_responsable",)),
                                                       ("Actividades", ("Responsable",)),
                                                       ("Clientes", ("Responsable_actividad",)))
                                     for r in rows(s) if mine(r, *cols) and is_overdue(s, r)),
    }
    previo = next((r for r in rows("Rendimiento") if r.get("Periodo_inicio") == inicio
                   and r.get("Periodo_fin") == fin and norm_name(r.get("Comercial")) == norm_name(comercial)), None)
    out["Meta_ventas"] = (previo or {}).get("Meta_ventas", "")
    compute_row("Rendimiento", out)
    return out, previo


def resumen_especialistas(store):
    nombres = {}
    for s, c in (("Prospectos", "Comercial_responsable"), ("Ofertas", "Comercial_responsable"),
                 ("Ventas", "Comercial_responsable"), ("Actividades", "Responsable")):
        for r in store.sheet(s)["rows"]:
            n = (r.get(c) or "").strip()
            if n:
                nombres.setdefault(norm_name(n), n)
    out = []
    for key, n in sorted(nombres.items(), key=lambda x: x[1]):
        pros = [r for r in store.sheet("Prospectos")["rows"] if norm_name(r.get("Comercial_responsable")) == key]
        ofes = [r for r in store.sheet("Ofertas")["rows"] if norm_name(r.get("Comercial_responsable")) == key]
        vtas = [r for r in store.sheet("Ventas")["rows"] if norm_name(r.get("Comercial_responsable")) == key]
        acts = [r for r in store.sheet("Actividades")["rows"] if norm_name(r.get("Responsable")) == key]
        venc = sum(1 for s, rr in (("Prospectos", pros), ("Actividades", acts)) for r in rr if is_overdue(s, r))
        out.append({
            "Especialista": n,
            "Oportunidades abiertas": sum(1 for r in pros if r.get("Etapa") not in ("Ganada", "Perdida", "Suspendida", "")),
            "Ofertas": len(ofes),
            "Ofertas por revisar": sum(1 for r in ofes if r.get("Estado_revision") in OFERTAS_PENDIENTES),
            "Ventas": len(vtas),
            "Valor ventas": fmt_num(sum(to_num(r.get("Importe")) or 0 for r in vtas)),
            "Seguim. vencidos": venc,
        })
    return out


# ===========================================================================
# PAQUETES Y CONSOLIDACIÓN
# ===========================================================================


def snapshot_path(book_path):
    return book_path + ".snapshot.json"


def save_snapshot(book_path, hojas):
    snap = {"generado": now_iso(), "hojas": {k: v["rows"] for k, v in hojas.items() if k in DATA_SHEETS}}
    with open(snapshot_path(book_path), "w", encoding="utf-8") as f:
        json.dump(snap, f, ensure_ascii=False, default=json_default)


def load_snapshot(book_path):
    p = snapshot_path(book_path)
    if not os.path.exists(p):
        return None
    try:
        with open(p, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return None


def build_package(book, especialista):
    snap = load_snapshot(book.path)
    cambios, hojas_out = {}, {}
    for name in DATA_SHEETS:
        if name not in book.hojas:
            continue
        info = book.hojas[name]
        rows = copy.deepcopy(info["rows"])
        hojas_out[name] = rows
        idf = ID_FIELD.get(name)
        if snap and idf:
            prev = {str(r.get(idf, "")): r for r in snap.get("hojas", {}).get(name, [])}
            nuevas, modificadas = [], []
            for r in rows:
                rid = str(r.get(idf, ""))
                old = prev.get(rid)
                if old is None:
                    nuevas.append(rid)
                elif rows_differ(old, r):
                    modificadas.append(rid)
            if nuevas or modificadas:
                cambios[name] = {"nuevas": nuevas, "modificadas": modificadas}
        elif not snap and idf and rows:
            cambios[name] = {"nuevas": [str(r.get(idf, "")) for r in rows], "modificadas": []}
    return {"tipo": "PAQUETE_ESPECIALISTA", "app": APP_NAME, "version": APP_VERSION,
            "especialista": especialista, "generado": now_iso(), "cambios": cambios, "hojas": hojas_out}


def write_json(obj, dest):
    with open(dest, "w", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False, default=json_default)
    return dest


def load_package(path):
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def analyze_import(master, pkg):
    reporte = {}
    for name, rows in pkg.get("hojas", {}).items():
        if name not in DATA_SHEETS or not rows:
            continue
        idf = ID_FIELD.get(name)
        if not idf or name not in master.hojas:
            continue
        headers = master.sheet(name)["headers"]
        rep = {"nuevas": [], "rellenos": {}, "forzados": {}, "conflictos": {}, "ignoradas": []}
        for r in rows:
            rid = str(r.get(idf, ""))
            if not rid:
                continue
            existing, _ = master.find_row(name, rid)
            if existing is None:
                rep["nuevas"].append(rid)
                continue
            im, em = r.get("_mod", ""), existing.get("_mod", "")
            if im and em and im <= em:
                continue                      # el maestro ya tiene datos iguales o más recientes
            for k in headers:
                if k in ("Fecha_modificacion", idf):
                    continue
                iv, mv = to_text(r.get(k, "")), to_text(existing.get(k, ""))
                if iv == mv or iv == "":
                    continue
                if name == "Ofertas":
                    if k in OFERTA_MASTER_FIELDS:
                        continue
                    if k == "Estado_revision":
                        if iv in OFERTA_ALLOWED.get(mv, ()):
                            rep["forzados"].setdefault(rid, {})[k] = iv
                        else:
                            rep["ignoradas"].append((rid, k))
                        continue
                if mv == "":
                    rep["rellenos"].setdefault(rid, {})[k] = iv
                else:
                    rep["conflictos"].setdefault(rid, {})[k] = (mv, iv)
        if any(rep[k] for k in rep):
            reporte[name] = rep
    return reporte


def apply_import(master, pkg, reporte, decisions):
    """decisions: {hoja: {id: set(campos tomados del especialista)}}"""
    stats = {"nuevas": 0, "actualizadas": 0, "conflictos_especialista": 0,
             "conflictos_maestro": 0, "ignoradas": 0}
    for name, rep in reporte.items():
        idf = ID_FIELD[name]
        headers = master.sheet(name)["headers"]
        rows_in = {str(r.get(idf, "")): r for r in pkg["hojas"].get(name, [])}
        for rid in rep["nuevas"]:
            r = rows_in.get(rid)
            if r is None:
                continue
            row = {h: to_text(r.get(h, "")) for h in headers}
            row["_mod"] = r.get("_mod", "") or now_iso()
            master.upsert(name, row, stamp=False)
            stats["nuevas"] += 1
        tocadas = set(rep["rellenos"]) | set(rep["forzados"])
        for rid in set(rep["rellenos"]) | set(rep["forzados"]) | set(rep["conflictos"]):
            existing, _ = master.find_row(name, rid)
            if existing is None:
                continue
            for k, v in rep["rellenos"].get(rid, {}).items():
                existing[k] = v
            for k, v in rep["forzados"].get(rid, {}).items():
                existing[k] = v
            elegidos = decisions.get(name, {}).get(rid, set())
            r = rows_in.get(rid, {})
            for k in elegidos:
                if k in rep["conflictos"].get(rid, {}):
                    existing[k] = to_text(r.get(k, ""))
            if rid in rep["conflictos"]:
                if elegidos:
                    stats["conflictos_especialista"] += 1
                else:
                    stats["conflictos_maestro"] += 1
            if rid in tocadas:
                stats["actualizadas"] += 1
            existing["_mod"] = now_iso()
            if "Fecha_modificacion" in existing:
                existing["Fecha_modificacion"] = now_min()
        stats["ignoradas"] += len(rep["ignoradas"])
        master._changed(name)
    master.commit_all()
    return stats


def registrar_log(master, pkg, stats, notas=""):
    if "LOG_CONSOLIDACION" not in master.hojas:
        master.hojas["LOG_CONSOLIDACION"] = {"headers": list(LOG_HEADERS), "rows": []}
    master.hojas["LOG_CONSOLIDACION"]["rows"].append({
        "Fecha": now_min(), "Especialista": pkg.get("especialista", ""),
        "Paquete": str(pkg.get("generado", ""))[:16], "Nuevas": stats["nuevas"],
        "Actualizadas": stats["actualizadas"],
        "Conflictos_a_favor_especialista": stats["conflictos_especialista"],
        "Conflictos_conservados_maestro": stats["conflictos_maestro"],
        "Ignoradas": stats["ignoradas"], "Notas": notas})
    master._changed("LOG_CONSOLIDACION")
    for r in master.sheet("Equipo")["rows"]:
        if norm_name(r.get("Nombre")) == norm_name(pkg.get("especialista", "")):
            r["Ultima_consolidacion"] = today_str()
            master._changed("Equipo")


def build_retro(master, especialista):
    key = norm_name(especialista)
    ofertas = []
    for r in master.sheet("Ofertas")["rows"]:
        if norm_name(r.get("Comercial_responsable")) == key:
            d = {"ID_Oferta": r.get("ID_Oferta"), "_mod": r.get("_mod", "")}
            for k in OFERTA_REVIEW_FIELDS:
                d[k] = r.get(k, "")
            ofertas.append(d)
    return {"tipo": "RETRO_JEFATURA", "app": APP_NAME, "especialista": especialista,
            "generado": now_iso(), "ofertas": ofertas}


def apply_retro(store, retro):
    """Devuelve lista de (id, campos cambiados) aplicados."""
    aplicados = []
    for d in retro.get("ofertas", []):
        row, _ = store.find_row("Ofertas", d.get("ID_Oferta"))
        if row is None:
            continue
        if row.get("_mod") and d.get("_mod") and d["_mod"] < row["_mod"]:
            continue
        cambios = []
        for k in OFERTA_REVIEW_FIELDS:
            v = to_text(d.get(k, ""))
            if v != to_text(row.get(k, "")):
                row[k] = v
                cambios.append(k)
        if cambios:
            row["_mod"] = d.get("_mod") or now_iso()
            aplicados.append((d["ID_Oferta"], cambios))
    if aplicados:
        store._changed("Ofertas")
    return aplicados


# ===========================================================================
# AUTOPRUEBA (sin interfaz) — se usa en la compilación del .exe
# ===========================================================================


def selftest(outfile):
    import tempfile
    log = []
    try:
        tmp = tempfile.mkdtemp()
        os.environ["CRM_DATA_DIR"] = os.path.join(tmp, "datos")
        esp = Store.create(os.path.join(tmp, "esp.db"), "especialista", "Comercial 01", "C01")
        assert esp.next_id("Clientes", "C01") == "CLI-C01-0001"
        cli = {h: "" for h in esp.sheet("Clientes")["headers"]}
        cli.update({"ID_Cliente": "CLI-C01-0001", "Razon_social": "Empresa Prueba", "Provincia": "La Habana"})
        esp.upsert("Clientes", cli)
        ofe = {h: "" for h in esp.sheet("Ofertas")["headers"]}
        ofe.update({"ID_Oferta": "OFE-C01-0001", "Cliente": "Empresa Prueba", "Producto_servicio": "Cemento",
                    "Cantidad": "10", "Precio_propuesto": "125.5", "Estado_revision": "Borrador",
                    "Comercial_responsable": "Comercial 01"})
        compute_row("Ofertas", ofe)
        assert ofe["Importe"] == "1255", ofe["Importe"]
        esp.upsert("Ofertas", ofe)
        log.append("base local, IDs y cálculo: OK")
        esp2 = Store.open(esp.path)
        assert len(esp2.sheet("Clientes")["rows"]) == 1 and esp2.sheet("Ofertas")["rows"][0]["Importe"] == "1255"
        log.append("persistencia SQLite: OK")
        xl = os.path.join(tmp, "libro.xlsx")
        export_workbook(esp2, xl, "especialista")
        xb = ExcelBook(xl)
        assert xb.sheet("Clientes")["rows"][0]["Razon_social"] == "Empresa Prueba"
        log.append("exportación de libro Excel: OK")
        jef = Store.create(os.path.join(tmp, "jef.db"), "jefatura", "Elizabeth", "JEF")
        pkg = build_package(esp2, "Comercial 01")
        rep = analyze_import(jef, pkg)
        stats = apply_import(jef, pkg, rep, {})
        registrar_log(jef, pkg, stats)
        assert jef.find_row("Clientes", "CLI-C01-0001")[0] is not None
        assert jef.find_row("Ofertas", "OFE-C01-0001")[0] is not None
        log.append(f"consolidación en maestro: OK {stats}")
        row, _ = jef.find_row("Ofertas", "OFE-C01-0001")
        row["Estado_revision"] = "Aprobada por jefatura"
        row["Revisor"] = "Elizabeth"
        row["_mod"] = now_iso()
        jef._changed("Ofertas")
        retro = build_retro(jef, "Comercial 01")
        assert apply_retro(esp2, retro)
        assert esp2.find_row("Ofertas", "OFE-C01-0001")[0]["Estado_revision"] == "Aprobada por jefatura"
        log.append("retroalimentación de ofertas: OK")
        nuevo = Store.create(os.path.join(tmp, "nuevo.db"), "jefatura", "Elizabeth", "JEF")
        xbm = ExcelBook(xl)
        pk2 = {"hojas": {n: xbm.hojas[n]["rows"] for n in xbm.hojas if n in DATA_SHEETS}}
        rep2 = analyze_import(nuevo, pk2)
        apply_import(nuevo, pk2, rep2, {})
        assert nuevo.find_row("Clientes", "CLI-C01-0001")[0] is not None
        log.append("importación desde Excel: OK")
        d = make_backup(esp2)
        assert os.path.exists(d)
        log.append("respaldo: OK")
        for s in (esp, esp2, jef, nuevo):
            s.close()
        with open(outfile, "w", encoding="utf-8") as f:
            f.write("OK\n" + "\n".join(log))
        return 0
    except Exception:
        with open(outfile, "w", encoding="utf-8") as f:
            f.write("FALLO\n" + "\n".join(log) + "\n" + traceback.format_exc())
        return 1


# ===========================================================================
# INTERFAZ: componentes
# ===========================================================================


def center_over(parent, dlg, w, h):
    parent.update_idletasks()
    px, py = parent.winfo_rootx(), parent.winfo_rooty()
    pw, ph = parent.winfo_width(), parent.winfo_height()
    sw, sh = dlg.winfo_screenwidth(), dlg.winfo_screenheight()
    w, h = min(w, sw - 40), min(h, sh - 80)
    x = max(0, px + (pw - w) // 2)
    y = max(0, py + (ph - h) // 3)
    dlg.geometry(f"{w}x{h}+{x}+{y}")


def make_tree(parent, cols, widths, headings, height=8):
    wrap = ttk.Frame(parent)
    tree = ttk.Treeview(wrap, show="headings", columns=cols, height=height)
    sb = ttk.Scrollbar(wrap, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=sb.set)
    for c, w, h in zip(cols, widths, headings):
        tree.heading(c, text=h)
        tree.column(c, width=w, minwidth=w, anchor="w", stretch=True)
    tree.pack(side="left", fill="both", expand=True)
    sb.pack(side="right", fill="y")
    return wrap, tree


class FormDialog(tk.Toplevel):
    """Formulario de captura / edición de un registro."""

    def __init__(self, app, sheet, row=None, prefill=None, on_done=None):
        super().__init__(app)
        self.app, self.sheet, self.orig, self.on_done = app, sheet, row, on_done
        self.is_new = row is None
        self.store = app.store
        self.headers = self.store.sheet(sheet)["headers"]
        self.widgets = {}
        self.title(("Editar " if row else "Nuevo/a ") + SINGULAR.get(sheet, sheet))
        self.transient(app)
        center_over(app, self, 900, 700)
        self.configure(bg="white")
        self._build()
        self._load(prefill or {})
        self.grab_set()
        self.focus_set()

    # ---- construcción ------------------------------------------------------
    def _choices_for(self, h):
        s = self.sheet
        cat = CATALOG_MAP.get((s, h))
        if cat:
            vals = list(self.store.catalogos.get(cat, []))
            if s == "Ofertas" and h == "Estado_revision" and self.app.role != "jefatura":
                cur = (self.orig or {}).get("Estado_revision", "") if self.orig else ""
                if not cur:
                    return ["Borrador", "Pendiente de revisión"], True
                return [cur] + sorted(OFERTA_ALLOWED.get(cur, [])), True
            return vals, True
        if (s, h) in CHOICES_SHEET:
            return CHOICES_SHEET[(s, h)], False
        if h in CHOICES:
            return CHOICES[h], False
        return None, False

    def _readonly(self, h):
        s = self.sheet
        if h == ID_FIELD.get(s) or h in AUTO_READONLY or h in COMPUTED.get(s, []):
            return True
        if s == "Ofertas" and self.app.role != "jefatura" and (h in OFERTA_MASTER_FIELDS or h == "Fecha_envio_jefatura"):
            return True
        return False

    def _build(self):
        btns = tk.Frame(self, bg="#eef1f5")
        btns.pack(side="bottom", fill="x")
        ttk.Button(btns, text="Cancelar", command=self.destroy).pack(side="right", padx=10, pady=8)
        ttk.Button(btns, text="Guardar", style="Accent.TButton", command=self.save).pack(side="right", pady=8)
        self.lbl_hint = tk.Label(btns, text="* obligatorio   |   campos en gris: automáticos",
                                 bg="#eef1f5", fg="#555", font=("Segoe UI", 9, "italic"))
        self.lbl_hint.pack(side="left", padx=10)

        canvas = tk.Canvas(self, bg="white", highlightthickness=0)
        sb = ttk.Scrollbar(self, orient="vertical", command=canvas.yview)
        form = tk.Frame(canvas, bg="white")
        form.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        win = canvas.create_window((0, 0), window=form, anchor="nw")
        canvas.configure(yscrollcommand=sb.set)
        canvas.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        canvas.bind("<Configure>", lambda e: canvas.itemconfigure(win, width=e.width))
        canvas.bind_all("<MouseWheel>", lambda e: canvas.yview_scroll(int(-1 * (e.delta / 120)), "units"))
        self.bind("<Destroy>", lambda e: canvas.unbind_all("<MouseWheel>") if e.widget is self else None)
        form.columnconfigure(0, weight=1, uniform="c")
        form.columnconfigure(1, weight=1, uniform="c")

        req = set(REQUIRED.get(self.sheet, []))
        r = 0
        col = 0
        for h in self.headers:
            if h == "_mod":
                continue
            long_ = h in LONGTEXT or (self.sheet, h) in LONG_SHEET
            if long_ and col == 1:
                r += 1
                col = 0
            frm = tk.Frame(form, bg="white")
            frm.grid(row=r, column=0 if long_ else col, columnspan=2 if long_ else 1,
                     sticky="ew", padx=12, pady=4)
            txt = h.replace("_", " ") + (" *" if h in req else "")
            tk.Label(frm, text=txt, bg="white", fg=("#c0392b" if h in req else "#1f3a5f"),
                     font=("Segoe UI", 9, "bold"), anchor="w").pack(fill="x")
            w = self._make_widget(frm, h, long_)
            self.widgets[h] = w
            if long_:
                r += 1
                col = 0
            else:
                col += 1
                if col == 2:
                    col = 0
                    r += 1

    def _make_widget(self, parent, h, long_):
        ro = self._readonly(h)
        refsrc = REF_SOURCE.get(h) if h != ID_FIELD.get(self.sheet) else None
        if h == "ID_Cliente_Proveedor":
            refsrc = "__both__"
        if long_:
            w = tk.Text(parent, height=3, wrap="word", font=("Segoe UI", 10), relief="solid", bd=1)
            w.pack(fill="x")
            if ro:
                w.config(state="disabled", bg="#eeeeee")
            return ("text", w)
        if refsrc:
            vals = self._ref_values(refsrc)
            cb = ttk.Combobox(parent, values=vals)
            cb.pack(fill="x")
            cb.bind("<<ComboboxSelected>>", lambda e, hh=h: self._on_ref(hh))
            cb.bind("<FocusOut>", lambda e, hh=h: self._on_ref(hh, soft=True))
            return ("ref", cb, refsrc)
        choices, strict = self._choices_for(h)
        if choices is not None:
            cb = ttk.Combobox(parent, values=choices, state="readonly" if (strict or ro) else "normal")
            cb.pack(fill="x")
            if ro:
                cb.config(state="disabled")
            cb.bind("<<ComboboxSelected>>", lambda e: self._recompute())
            return ("choice", cb)
        if is_date_col(h) or is_datetime_col(h):
            fr = tk.Frame(parent, bg="white")
            fr.pack(fill="x")
            e = ttk.Entry(fr)
            e.pack(side="left", fill="x", expand=True)
            if ro:
                e.config(state="readonly")
            else:
                ttk.Button(fr, text="Ahora" if is_datetime_col(h) else "Hoy", width=6,
                           command=lambda ee=e, hh=h: (ee.delete(0, "end"),
                                                       ee.insert(0, now_min() if is_datetime_col(hh) else today_str()),
                                                       self._recompute())).pack(side="left", padx=(4, 0))
            e.bind("<FocusOut>", lambda ev: self._recompute())
            return ("entry", e)
        e = ttk.Entry(parent)
        e.pack(fill="x")
        if ro:
            e.config(state="readonly")
        e.bind("<FocusOut>", lambda ev: self._recompute())
        e.bind("<KeyRelease>", lambda ev: self._recompute() if h in NUMERIC_COLS else None)
        return ("entry", e)

    def _ref_values(self, src):
        vals = []
        sources = ["Clientes", "Proveedores"] if src == "__both__" else [src]
        for s in sources:
            idf = ID_FIELD[s]
            for r in self.store.sheet(s)["rows"]:
                nm = display_name(s, r)
                vals.append(f"{r.get(idf, '')} | {nm}" if nm else str(r.get(idf, "")))
        return vals

    # ---- valores ----------------------------------------------------------
    def get(self, h):
        t = self.widgets[h]
        k, w = t[0], t[1]
        if k == "text":
            return w.get("1.0", "end-1c").strip()
        v = w.get().strip()
        if k == "ref":
            return v.split(" | ")[0].strip()
        return v

    def set(self, h, v):
        t = self.widgets.get(h)
        if not t:
            return
        k, w = t[0], t[1]
        v = to_text(v)
        if k == "text":
            st = str(w.cget("state"))
            w.config(state="normal")
            w.delete("1.0", "end")
            w.insert("1.0", v)
            w.config(state=st)
            return
        st = str(w.cget("state"))
        w.config(state="normal")
        if k in ("entry", "ref") or (k == "choice"):
            if k == "choice":
                w.set(v)
            else:
                w.delete(0, "end")
                w.insert(0, v)
        w.config(state=st if st != "normal" else "normal")

    def _on_ref(self, h, soft=False):
        t = self.widgets[h]
        src = t[2]
        raw = t[1].get().strip()
        rid = raw.split(" | ")[0].strip()
        if not rid:
            return
        sources = ["Clientes", "Proveedores"] if src == "__both__" else [src]
        row = None
        for s in sources:
            row, _ = self.store.find_row(s, rid)
            if row is not None:
                src_s = s
                break
        if row is None:
            return
        nm = display_name(src_s, row)
        if not soft or " | " not in raw:
            t[1].delete(0, "end")
            t[1].insert(0, f"{rid} | {nm}" if nm else rid)
        for tg in AUTOFILL.get(h, []):
            if tg in self.widgets and not self.get(tg):
                self.set(tg, nm)
        if self.sheet == "Logistica" and h == "ID_Operacion":
            if "Producto" in self.widgets and not self.get("Producto"):
                self.set("Producto", row.get("Producto", ""))
        self._recompute()

    def _collect(self):
        return {h: self.get(h) for h in self.headers if h in self.widgets}

    def _recompute(self):
        if self.sheet not in COMPUTED:
            return
        row = self._collect()
        compute_row(self.sheet, row)
        for h in COMPUTED[self.sheet]:
            if h in self.widgets and row.get(h, "") != self.get(h):
                self.set(h, row.get(h, ""))

    def _load(self, prefill):
        s = self.sheet
        if self.orig:
            data = {h: self.orig.get(h, "") for h in self.headers}
        else:
            data = {h: "" for h in self.headers}
            idf = ID_FIELD.get(s)
            if idf:
                data[idf] = self.store.next_id(s, self.app.code)
            defaults = {
                "Clientes": {"Fecha_alta": today_str(), "Estado_cliente": "Prospecto",
                             "Pais": "Cuba", "Estado_validacion": "Pendiente de verificación"},
                "Proveedores": {"Estado_validacion": "Pendiente de verificación"},
                "Prospectos": {"Fecha_captacion": today_str(), "Etapa": "Identificado", "Moneda": "USD"},
                "Actividades": {"Fecha_hora": now_min(), "Estado": "Pendiente"},
                "Ofertas": {"Fecha_preparacion": today_str(), "Estado_revision": "Borrador", "Version": "1",
                            "Moneda": "USD", "Requiere_Direccion": "No"},
                "Incidencias": {"Fecha_deteccion": today_str(), "Prioridad": "Media", "Estado": "Abierta"},
                "Equipo": {"Fecha_incorporacion": today_str(), "Estado": "Activo"},
                "Embarques": {"Fecha_actualizacion": today_str()},
                "Contenedores": {"Fecha_actualizacion": today_str()},
                "Ventas": {"Fecha_aceptacion": today_str(), "Moneda": "USD"},
                "Compras": {"Fecha_orden": today_str(), "Moneda": "USD"},
                "Importaciones": {"Moneda": "USD"},
                "Facturas_Cobros": {"Fecha_emision": today_str(), "Moneda": "USD", "Estado": "Pendiente"},
                "Rendimiento": {"Periodo_inicio": today_str()[:8] + "01", "Periodo_fin": today_str(),
                                "Comercial": self.app.user},
            }.get(s, {})
            data.update({k: v for k, v in defaults.items() if k in data})
            for c in RESPONSIBLE_COLS:
                if c in data and not data[c]:
                    data[c] = self.app.user
            data.update({k: v for k, v in prefill.items() if k in data})
        for h, v in data.items():
            if h in self.widgets:
                self.set(h, v)
        # referencias: mostrar "ID | nombre"
        for h, t in self.widgets.items():
            if t[0] == "ref" and data.get(h):
                self._on_ref(h)
        self._recompute()

    # ---- guardar ------------------------------------------------------------
    def save(self):
        row = self._collect()
        compute_row(self.sheet, row)
        errs, warns = validate_row(self.store, self.sheet, row, self.orig, self.is_new, self.app.role)
        if errs:
            messagebox.showwarning("Revise los datos", "\n".join("• " + e for e in errs), parent=self)
            return
        if warns and not messagebox.askyesno("Advertencias", "\n".join("• " + w for w in warns) +
                                             "\n\n¿Guardar de todos modos?", parent=self):
            return
        for h in self.headers:
            row.setdefault(h, "")
        if self.orig and "_mod" in self.orig:
            pass
        self.store.upsert(self.sheet, row, original=self.orig)
        self.app.set_status(f"{SINGULAR.get(self.sheet, self.sheet).capitalize()} "
                            f"{row.get(ID_FIELD.get(self.sheet) or '', '')} guardado/a.")
        self.destroy()
        if self.on_done:
            self.on_done(row)
        self.app.refresh_all()


class ReviewDialog(tk.Toplevel):
    """Revisión de una oferta por jefatura."""

    def __init__(self, app, row, on_done=None):
        super().__init__(app)
        self.app, self.row, self.on_done = app, row, on_done
        self.title(f"Revisar oferta {row.get('ID_Oferta', '')}")
        self.transient(app)
        center_over(app, self, 640, 560)
        self.configure(bg="white")
        info = (f"Cliente: {row.get('Cliente', '')}\nProducto/servicio: {row.get('Producto_servicio', '')}\n"
                f"Cantidad: {row.get('Cantidad', '')} {row.get('Unidad', '')}   Precio: {row.get('Precio_propuesto', '')} "
                f"{row.get('Moneda', '')}   Importe: {row.get('Importe', '')}\n"
                f"Comercial: {row.get('Comercial_responsable', '')}   Estado actual: {row.get('Estado_revision', '')}\n"
                f"Entrega: {row.get('Condiciones_entrega', '')}   Pago: {row.get('Forma_pago', '')}")
        tk.Label(self, text=info, bg="white", justify="left", anchor="w",
                 font=("Segoe UI", 10)).pack(fill="x", padx=14, pady=(12, 6))
        self.var = tk.StringVar(value="aprobar")
        opts = [("aprobar", "Aprobar (si requiere Dirección, pasa a autorización)"),
                ("direccion", "Aprobar y enviar a autorización de Dirección"),
                ("autorizada", "Registrar autorización de Dirección (autorizada para presentar)"),
                ("modif", "Solicitar modificaciones (requiere comentario)"),
                ("rechazar", "Rechazar (requiere comentario)"),
                ("revision", "Marcar «En revisión»")]
        for v, t in opts:
            tk.Radiobutton(self, text=t, value=v, variable=self.var, bg="white", anchor="w",
                           font=("Segoe UI", 10)).pack(fill="x", padx=20)
        tk.Label(self, text="Observaciones de jefatura:", bg="white", font=("Segoe UI", 9, "bold"),
                 anchor="w").pack(fill="x", padx=14, pady=(10, 2))
        self.txt = tk.Text(self, height=5, wrap="word", font=("Segoe UI", 10), relief="solid", bd=1)
        self.txt.pack(fill="both", expand=True, padx=14)
        prev = row.get("Observaciones_jefatura", "")
        if prev:
            tk.Label(self, text="Historial: " + prev[:300], bg="white", fg="#666", anchor="w",
                     justify="left", wraplength=600, font=("Segoe UI", 8)).pack(fill="x", padx=14)
        bf = tk.Frame(self, bg="#eef1f5")
        bf.pack(fill="x", side="bottom")
        ttk.Button(bf, text="Cancelar", command=self.destroy).pack(side="right", padx=10, pady=8)
        ttk.Button(bf, text="Aplicar revisión", style="Accent.TButton", command=self.apply).pack(side="right", pady=8)
        self.grab_set()

    def apply(self):
        a, r = self.var.get(), self.row
        txt = self.txt.get("1.0", "end-1c").strip()
        if a in ("modif", "rechazar") and not txt:
            messagebox.showwarning("Falta comentario", "Indique el motivo en las observaciones.", parent=self)
            return
        dir_ok = r.get("Autorizacion_Direccion") == "Autorizada"
        if a == "aprobar":
            if r.get("Requiere_Direccion") == "Sí" and not dir_ok:
                r["Estado_revision"] = "Pendiente de autorización de Dirección"
                r["Autorizacion_Direccion"] = "Pendiente"
            else:
                r["Estado_revision"] = "Autorizada para presentar" if dir_ok else "Aprobada por jefatura"
                r["Fecha_aprobacion"] = today_str()
        elif a == "direccion":
            r["Estado_revision"] = "Pendiente de autorización de Dirección"
            r["Requiere_Direccion"] = "Sí"
            r["Autorizacion_Direccion"] = "Pendiente"
        elif a == "autorizada":
            r["Requiere_Direccion"] = "Sí"
            r["Autorizacion_Direccion"] = "Autorizada"
            r["Estado_revision"] = "Autorizada para presentar"
            r["Fecha_aprobacion"] = today_str()
        elif a == "modif":
            r["Estado_revision"] = "Modificaciones solicitadas"
        elif a == "rechazar":
            r["Estado_revision"] = "Rechazada"
        elif a == "revision":
            r["Estado_revision"] = "En revisión"
        r["Revisor"] = self.app.user
        if txt:
            prev = r.get("Observaciones_jefatura", "")
            nuevo = f"[{today_str()} {self.app.user}] {txt}"
            r["Observaciones_jefatura"] = (prev + "\n" if prev else "") + nuevo
        self.app.store.upsert("Ofertas", r)
        self.destroy()
        if self.on_done:
            self.on_done(r)
        self.app.refresh_all()
        self.app.set_status(f"Oferta {r.get('ID_Oferta')}: {r['Estado_revision']}.")


class ConflictDialog(tk.Toplevel):
    def __init__(self, master, reporte):
        super().__init__(master)
        self.title("Resolver conflictos de consolidación")
        self.transient(master)
        center_over(master, self, 1000, 620)
        self.grab_set()
        self.items = []
        self.decisions = {}
        for name, rep in reporte.items():
            for rid, fields in rep["conflictos"].items():
                self.items.append((name, rid, fields))
                self.decisions.setdefault(name, {})[rid] = set()
        self.idx = 0
        self.cancelled = False
        top = ttk.Frame(self, padding=8)
        top.pack(fill="x")
        self.lbl = ttk.Label(top, font=("Segoe UI", 11, "bold"))
        self.lbl.pack(anchor="w")
        ttk.Label(top, text="Mismo registro con datos distintos en el maestro y en el paquete. "
                            "Doble clic en un campo para alternar entre MAESTRO y ESPECIALISTA.",
                  font=("Segoe UI", 9, "italic")).pack(anchor="w")
        cols = ("campo", "maestro", "especialista", "decision")
        wrap, self.tree = make_tree(self, cols, (200, 300, 300, 120),
                                    ("Campo", "Valor en MAESTRO", "Valor del ESPECIALISTA", "Se queda"), 12)
        wrap.pack(fill="both", expand=True, padx=8, pady=4)
        self.tree.bind("<Double-1>", self._toggle)
        b1 = ttk.Frame(self, padding=(8, 2))
        b1.pack(fill="x")
        ttk.Button(b1, text="<< Anterior", command=lambda: self._nav(-1)).pack(side="left", padx=3)
        ttk.Button(b1, text="Siguiente >>", command=lambda: self._nav(1)).pack(side="left", padx=3)
        ttk.Button(b1, text="Tomar ESPECIALISTA en este registro", command=lambda: self._set_all_current(True)).pack(side="left", padx=12)
        ttk.Button(b1, text="Conservar MAESTRO en este registro", command=lambda: self._set_all_current(False)).pack(side="left")
        b2 = ttk.Frame(self, padding=(8, 2, 8, 8))
        b2.pack(fill="x")
        ttk.Button(b2, text="Cancelar importación", command=self._cancel).pack(side="right", padx=3)
        ttk.Button(b2, text="Finalizar e importar", style="Accent.TButton", command=self.destroy).pack(side="right", padx=3)
        ttk.Button(b2, text="Todos: especialista", command=lambda: self._all(True)).pack(side="left", padx=3)
        ttk.Button(b2, text="Todos: maestro", command=lambda: self._all(False)).pack(side="left", padx=3)
        self._show()

    def _cur(self):
        return self.items[self.idx]

    def _show(self):
        name, rid, fields = self._cur()
        self.lbl.config(text=f"Conflicto {self.idx + 1} de {len(self.items)} — {SHEET_LABEL.get(name, name)} — {rid}")
        self.tree.delete(*self.tree.get_children())
        elegidos = self.decisions[name][rid]
        for k, (mv, iv) in fields.items():
            self.tree.insert("", "end", iid=k, values=(k, mv or "(vacío)", iv or "(vacío)",
                                                       "ESPECIALISTA" if k in elegidos else "Maestro"))

    def _toggle(self, e):
        sel = self.tree.selection()
        if not sel:
            return
        name, rid, fields = self._cur()
        s = self.decisions[name][rid]
        s.symmetric_difference_update({sel[0]})
        self._show()

    def _set_all_current(self, esp):
        name, rid, fields = self._cur()
        self.decisions[name][rid] = set(fields) if esp else set()
        self._show()

    def _all(self, esp):
        for name, rid, fields in self.items:
            self.decisions[name][rid] = set(fields) if esp else set()
        self._show()

    def _nav(self, d):
        self.idx = max(0, min(len(self.items) - 1, self.idx + d))
        self._show()

    def _cancel(self):
        self.cancelled = True
        self.destroy()

    def run(self):
        self.wait_window()
        return None if self.cancelled else self.decisions


class SimpleDialog(tk.Toplevel):
    """Diálogo de uno o varios campos de texto / lista."""

    def __init__(self, app, title, fields, ok="Aceptar", width=480):
        super().__init__(app)
        self.title(title)
        self.transient(app)
        center_over(app, self, width, 90 + 62 * len(fields))
        self.result = None
        self.vars = {}
        for label, default, values in fields:
            tk.Label(self, text=label, anchor="w", font=("Segoe UI", 9, "bold")).pack(fill="x", padx=14, pady=(10, 0))
            if values:
                w = ttk.Combobox(self, values=values)
                w.set(default)
            else:
                w = ttk.Entry(self)
                w.insert(0, default)
            w.pack(fill="x", padx=14)
            self.vars[label] = w
        bf = ttk.Frame(self)
        bf.pack(side="bottom", fill="x", pady=8)
        ttk.Button(bf, text="Cancelar", command=self.destroy).pack(side="right", padx=10)
        ttk.Button(bf, text=ok, style="Accent.TButton", command=self._ok).pack(side="right")
        self.grab_set()

    def _ok(self):
        self.result = {k: w.get().strip() for k, w in self.vars.items()}
        self.destroy()

    def run(self):
        self.wait_window()
        return self.result


class TableView(ttk.Frame):
    def __init__(self, app, sheet):
        super().__init__(app.main)
        self.app, self.sheet = app, sheet
        self.store = app.store
        info = self.store.sheet(sheet)
        self.all_headers = [h for h in info["headers"]]
        keys = [h for h in KEY_COLS.get(sheet, []) if h in self.all_headers]
        self.key_cols = keys or self.all_headers[:8]
        self.show_all = tk.BooleanVar(value=False)
        self.search = tk.StringVar()
        self.status = tk.StringVar(value="(Todos)")
        self.rowmap = {}
        self._sort = {}
        ttk.Label(self, text=SHEET_LABEL.get(sheet, sheet), font=("Segoe UI", 16, "bold"),
                  foreground=COLOR_HEADER).pack(anchor="w")
        bar = ttk.Frame(self)
        bar.pack(fill="x", pady=6)
        ttk.Label(bar, text="Buscar:").pack(side="left")
        e = ttk.Entry(bar, textvariable=self.search, width=30)
        e.pack(side="left", padx=6)
        self.search.trace_add("write", lambda *a: self.fill())
        sc = STATUS_COL.get(sheet)
        if sc and sc in self.all_headers:
            ttk.Label(bar, text=f"{sc.replace('_', ' ')}:").pack(side="left", padx=(10, 2))
            self.cb = ttk.Combobox(bar, textvariable=self.status, width=26, state="readonly")
            self.cb.pack(side="left")
            self.cb.bind("<<ComboboxSelected>>", lambda e: self.fill())
        ttk.Checkbutton(bar, text="Ver todas las columnas", variable=self.show_all,
                        command=self.rebuild).pack(side="left", padx=12)
        bar2 = ttk.Frame(self)
        bar2.pack(fill="x", pady=(0, 6))
        ttk.Button(bar2, text=f"+ Nuevo/a {SINGULAR.get(sheet, '')}", style="Accent.TButton",
                   command=lambda: app.new_record(sheet)).pack(side="left", padx=(0, 4))
        ttk.Button(bar2, text="Abrir / editar", command=self.edit).pack(side="left", padx=4)
        ttk.Button(bar2, text="Eliminar", command=self.delete).pack(side="left", padx=4)
        if sheet == "Clientes":
            ttk.Button(bar2, text="Ficha del cliente", command=self.ficha).pack(side="left", padx=4)
        if sheet == "Ofertas":
            if app.role == "jefatura":
                ttk.Button(bar2, text="Revisar oferta", command=self.revisar).pack(side="left", padx=4)
            else:
                ttk.Button(bar2, text="Enviar a revisión de jefatura", command=self.enviar_revision).pack(side="left", padx=4)
        if sheet == "Rendimiento":
            ttk.Button(bar2, text="Calcular rendimiento del periodo", command=app.calcular_rendimiento).pack(side="left", padx=4)
        ttk.Button(bar2, text="Exportar a Excel", command=lambda: app.exportar_hoja(sheet)).pack(side="right", padx=4)
        self.holder = ttk.Frame(self)
        self.holder.pack(fill="both", expand=True)
        self.lbl_count = ttk.Label(self, text="", foreground="#666")
        self.lbl_count.pack(anchor="w")
        self.rebuild()

    def rebuild(self):
        for w in self.holder.winfo_children():
            w.destroy()
        cols = self.all_headers if self.show_all.get() else self.key_cols
        self.cols = cols
        ysb = ttk.Scrollbar(self.holder, orient="vertical")
        xsb = ttk.Scrollbar(self.holder, orient="horizontal")
        tree = ttk.Treeview(self.holder, show="headings", columns=cols, yscrollcommand=ysb.set,
                            xscrollcommand=xsb.set)
        ysb.config(command=tree.yview)
        xsb.config(command=tree.xview)
        for h in cols:
            tree.heading(h, text=h.replace("_", " "), command=lambda c=h: self.sort_by(c))
            w = 90 if h.startswith("ID_") else (200 if h in LONGTEXT or h in ("Razon_social", "Cliente", "Producto_servicio",
                                                                             "Proxima_accion", "Descripcion") else 130)
            w = max(w, int(len(h) * 8.2) + 18)
            tree.column(h, width=w, minwidth=w, anchor="w", stretch=not self.show_all.get())
        ysb.pack(side="right", fill="y")
        xsb.pack(side="bottom", fill="x")
        tree.pack(fill="both", expand=True)
        tree.tag_configure("venc", foreground="#c0392b")
        tree.bind("<Double-1>", lambda e: self.edit())
        tree.bind("<Return>", lambda e: self.edit())
        self.tree = tree
        if hasattr(self, "cb"):
            sc = STATUS_COL[self.sheet]
            vals = sorted({r.get(sc, "") for r in self.store.sheet(self.sheet)["rows"] if r.get(sc, "")})
            self.cb.config(values=["(Todos)"] + vals)
        self.fill()

    def fill(self):
        tree = self.tree
        tree.delete(*tree.get_children())
        self.rowmap = {}
        s = self.search.get().strip().lower()
        sc = STATUS_COL.get(self.sheet)
        stf = self.status.get()
        n = 0
        rows = self.store.sheet(self.sheet)["rows"]
        for i, r in enumerate(rows):
            if sc and stf != "(Todos)" and r.get(sc, "") != stf:
                continue
            vals = [to_text(r.get(h, "")) for h in self.cols]
            if s and not any(s in to_text(r.get(h, "")).lower() for h in self.all_headers):
                continue
            iid = f"r{i}"
            self.rowmap[iid] = r
            tree.insert("", "end", iid=iid, values=vals, tags=("venc",) if is_overdue(self.sheet, r) else ())
            n += 1
        if not rows:
            self.lbl_count.config(text=f"Aún no hay registros. Pulse «+ Nuevo/a {SINGULAR.get(self.sheet, '')}» para agregar el primero.")
        else:
            self.lbl_count.config(text=f"{n} de {len(rows)} registros   (en rojo: seguimiento vencido)")

    def sort_by(self, col):
        items = [(self.tree.set(k, col), k) for k in self.tree.get_children()]

        def key(x):
            v = x[0]
            n = to_num(v)
            return (0, n, "") if n is not None else (1, 0, v.lower())

        asc = not self._sort.get(col, False)
        self._sort[col] = asc
        items.sort(key=key, reverse=not asc)
        for i, (_, k) in enumerate(items):
            self.tree.move(k, "", i)

    def selected(self):
        sel = self.tree.selection()
        return self.rowmap.get(sel[0]) if sel else None

    def edit(self):
        r = self.selected()
        if r is None:
            messagebox.showinfo(APP_NAME, "Seleccione un registro de la lista.")
            return
        FormDialog(self.app, self.sheet, r)

    def delete(self):
        r = self.selected()
        if r is None:
            messagebox.showinfo(APP_NAME, "Seleccione un registro de la lista.")
            return
        idf = ID_FIELD.get(self.sheet)
        etiqueta = r.get(idf, "") if idf else "este registro"
        if messagebox.askyesno(APP_NAME, f"¿Eliminar {etiqueta}?\nEsta acción no se puede deshacer "
                                         "(puede restaurar un respaldo desde Archivo)."):
            self.store.delete_obj(self.sheet, r)
            self.app.refresh_all()

    def ficha(self):
        r = self.selected()
        if r is None:
            messagebox.showinfo(APP_NAME, "Seleccione un cliente de la lista.")
            return
        ClientCard(self.app, r)

    def revisar(self):
        r = self.selected()
        if r is None:
            messagebox.showinfo(APP_NAME, "Seleccione una oferta.")
            return
        ReviewDialog(self.app, r)

    def enviar_revision(self):
        r = self.selected()
        if r is None:
            messagebox.showinfo(APP_NAME, "Seleccione una oferta.")
            return
        if r.get("Estado_revision") not in ("Borrador", "Modificaciones solicitadas"):
            messagebox.showinfo(APP_NAME, "Solo se pueden enviar a revisión las ofertas en «Borrador» o "
                                          "«Modificaciones solicitadas».")
            return
        faltan = [h for h in ("Cliente", "Producto_servicio", "Cantidad", "Precio_propuesto", "Moneda")
                  if not r.get(h)]
        if faltan:
            messagebox.showwarning(APP_NAME, "Complete antes: " + ", ".join(faltan))
            return
        if r.get("Estado_revision") == "Modificaciones solicitadas":
            r["Version"] = fmt_num((to_num(r.get("Version")) or 1) + 1)
        r["Estado_revision"] = "Pendiente de revisión"
        r["Fecha_envio_jefatura"] = today_str()
        self.store.upsert("Ofertas", r)
        self.app.refresh_all()
        self.app.set_status(f"Oferta {r.get('ID_Oferta')} enviada a revisión. Se incluirá en su próximo paquete a jefatura.")


class ClientCard(tk.Toplevel):
    """Ficha 360° de un cliente."""

    def __init__(self, app, cli):
        super().__init__(app)
        self.app, self.cli = app, cli
        self.title(f"Ficha del cliente — {cli.get('Razon_social', '')}")
        self.transient(app)
        center_over(app, self, 1000, 640)
        cid, nm = cli.get("ID_Cliente", ""), cli.get("Razon_social", "")
        head = tk.Frame(self, bg=COLOR_HEADER)
        head.pack(fill="x")
        tk.Label(head, text=f"{nm}", bg=COLOR_HEADER, fg="white", font=("Segoe UI", 14, "bold")).pack(anchor="w", padx=12, pady=(8, 0))
        det = (f"{cid}   |   {cli.get('Provincia', '')} {cli.get('Municipio', '')}   |   Contacto: {cli.get('Contacto', '')} "
               f"{cli.get('Cargo', '')}   |   Tel: {cli.get('Telefono', '')}   |   {cli.get('Correo', '')}")
        tk.Label(head, text=det, bg=COLOR_HEADER, fg="#cfd8e3", font=("Segoe UI", 9)).pack(anchor="w", padx=12)
        tk.Label(head, text=f"Estado: {cli.get('Estado_cliente', '')}   Potencial: {cli.get('Potencial', '')}   "
                            f"Próxima acción: {cli.get('Proxima_accion', '')} ({cli.get('Fecha_seguimiento', '')})",
                 bg=COLOR_HEADER, fg="#cfd8e3", font=("Segoe UI", 9)).pack(anchor="w", padx=12, pady=(0, 8))
        bar = ttk.Frame(self, padding=6)
        bar.pack(fill="x")
        pre = {"ID_Cliente": cid, "Cliente": nm, "Cliente_proveedor": nm}
        for txt, sh in (("+ Actividad", "Actividades"), ("+ Oportunidad", "Prospectos"), ("+ Oferta", "Ofertas")):
            ttk.Button(bar, text=txt, style="Accent.TButton",
                       command=lambda s=sh: app.new_record(s, pre, lambda r: self._reload())).pack(side="left", padx=3)
        ttk.Button(bar, text="Editar datos del cliente", command=lambda: FormDialog(app, "Clientes", cli, on_done=lambda r: self.destroy())).pack(side="right", padx=3)
        self.nb = ttk.Notebook(self)
        self.nb.pack(fill="both", expand=True, padx=6, pady=6)
        self._reload()
        self.grab_set()

    def _related(self, sheet):
        cid, nm = self.cli.get("ID_Cliente", ""), norm_name(self.cli.get("Razon_social", ""))
        out = []
        for r in self.app.store.sheet(sheet)["rows"]:
            if (cid and r.get("ID_Cliente") == cid) or \
                    (nm and norm_name(r.get("Cliente", "") or r.get("Cliente_proveedor", "") or r.get("Nombre", "")) == nm):
                out.append(r)
        return out

    def _reload(self):
        for t in self.nb.tabs():
            self.nb.forget(t)
        for sh in ("Actividades", "Prospectos", "Ofertas", "Contratos", "Ventas", "Facturas_Cobros"):
            rows = self._related(sh)
            fr = ttk.Frame(self.nb)
            self.nb.add(fr, text=f"{SHEET_LABEL[sh]} ({len(rows)})")
            cols = [h for h in KEY_COLS[sh] if h in self.app.store.sheet(sh)["headers"]]
            wrap, tree = make_tree(fr, cols, [120] * len(cols), [c.replace("_", " ") for c in cols], 12)
            wrap.pack(fill="both", expand=True)
            for i, r in enumerate(rows):
                tree.insert("", "end", iid=str(i), values=[to_text(r.get(c, "")) for c in cols])
            tree.bind("<Double-1>", lambda e, s=sh, rr=rows, tr=tree: self._open(s, rr, tr))

    def _open(self, sheet, rows, tree):
        sel = tree.selection()
        if sel:
            FormDialog(self.app, sheet, rows[int(sel[0])], on_done=lambda r: self._reload())


class DuplicatesDialog(tk.Toplevel):
    def __init__(self, app):
        super().__init__(app)
        self.title("Posibles clientes y proveedores duplicados")
        self.transient(app)
        center_over(app, self, 900, 480)
        tk.Label(self, text="Registros con el mismo NIT / identificación fiscal o la misma razón social pero distinto ID. "
                            "Los clientes y proveedores son carteras compartidas: conviene unificarlos "
                            "(edite uno y elimine el otro).", wraplength=860, justify="left").pack(padx=10, pady=8, anchor="w")
        cols = ("hoja", "id", "razon", "resp")
        wrap, tree = make_tree(self, cols, (110, 130, 420, 180), ("Tipo", "ID", "Razón social", "Responsable"), 14)
        wrap.pack(fill="both", expand=True, padx=8, pady=4)
        n = 0
        for sh in ("Clientes", "Proveedores"):
            for g in find_duplicates(app.store, sh):
                n += 1
                tree.insert("", "end", values=(f"Grupo {n}", "", "", ""), tags=("g",))
                for r in g:
                    tree.insert("", "end", values=(sh, r.get(ID_FIELD[sh]), r.get("Razon_social"),
                                                   r.get("Responsable_actividad") or r.get("Responsable") or ""))
        tree.tag_configure("g", background="#fdebd0")
        if n == 0:
            tree.insert("", "end", values=("", "", "No se detectaron duplicados", ""))


# ===========================================================================
# APLICACIÓN PRINCIPAL
# ===========================================================================


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"{APP_TITLE} — v{APP_VERSION}")
        self.geometry("1320x820")
        self.minsize(1050, 660)
        self.cfg = load_config()
        self.store = None
        self.role = None
        self.user = ""
        self.code = ""
        self.profile = None
        self.current = "__panel__"
        self.view = None
        self._style()
        self._ui()
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self.after(120, self.startup)

    # ---- estilo e interfaz base ---------------------------------------------
    def _style(self):
        self.configure(bg=COLOR_BG)
        st = ttk.Style(self)
        try:
            st.theme_use("clam")
        except tk.TclError:
            pass
        st.configure("TFrame", background=COLOR_BG)
        st.configure("Top.TFrame", background=COLOR_HEADER)
        st.configure("Side.TFrame", background=COLOR_SIDEBAR)
        st.configure("Top.TLabel", background=COLOR_HEADER, foreground="white", font=("Segoe UI", 11, "bold"))
        st.configure("TLabel", background=COLOR_BG, font=("Segoe UI", 10))
        st.configure("TButton", font=("Segoe UI", 10), padding=6)
        st.configure("Accent.TButton", background=COLOR_ACCENT, foreground="white", font=("Segoe UI", 10, "bold"))
        st.map("Accent.TButton", background=[("active", "#22577a")])
        st.configure("Side.TButton", background="#24507a", foreground="white", font=("Segoe UI", 10, "bold"), padding=8)
        st.map("Side.TButton", background=[("active", "#2e86ab")])
        st.configure("Treeview", font=("Segoe UI", 9), rowheight=24)
        st.configure("Treeview.Heading", font=("Segoe UI", 9, "bold"))
        st.configure("TNotebook.Tab", font=("Segoe UI", 10, "bold"), padding=8)

    def _ui(self):
        top = ttk.Frame(self, style="Top.TFrame", padding=(10, 8))
        top.pack(fill="x", side="top")
        ttk.Label(top, text=APP_TITLE, style="Top.TLabel").pack(side="left")
        self.lbl_user = ttk.Label(top, text="", style="Top.TLabel", font=("Segoe UI", 10, "italic"))
        self.lbl_user.pack(side="right")
        self.status = ttk.Label(self, text="Listo.", anchor="w", relief="sunken", padding=5)
        self.status.pack(fill="x", side="bottom")
        body = ttk.Frame(self)
        body.pack(fill="both", expand=True)
        self.sidebar = ttk.Frame(body, style="Side.TFrame", width=268, padding=6)
        self.sidebar.pack(side="left", fill="y")
        self.sidebar.pack_propagate(False)
        self.main = ttk.Frame(body)
        self.main.pack(side="left", fill="both", expand=True, padx=10, pady=8)
        self.btn_enviar = ttk.Button(self.sidebar, text="Enviar a jefatura", style="Side.TButton",
                                    command=self.enviar_paquete)
        self.btn_enviar.pack(fill="x", pady=2)
        self.btn_retro = ttk.Button(self.sidebar, text="Recibir retroalimentación", style="Side.TButton",
                                    command=self.importar_retro)
        self.btn_retro.pack(fill="x", pady=2)
        self.btn_panel = ttk.Button(self.sidebar, text="Panel de control", command=lambda: self.go("__panel__"))
        self.btn_panel.pack(fill="x", pady=(8, 4))
        self.nav = ttk.Treeview(self.sidebar, show="tree", selectmode="browse")
        self.nav.pack(fill="both", expand=True)
        self.nav.tag_configure("mod", font=("Segoe UI", 10, "bold"))
        self.nav.bind("<<TreeviewSelect>>", self._on_nav)
        self._menu()

    def _menu(self):
        m = tk.Menu(self)
        self.config(menu=m)
        a = tk.Menu(m, tearoff=0)
        a.add_command(label="Cambiar de perfil / Nuevo perfil...", command=self.show_setup)
        a.add_separator()
        a.add_command(label="Importar datos desde Excel...", command=self.importar_excel)
        a.add_command(label="Exportar TODO el CRM a un libro Excel...", command=self.exportar_libro)
        a.add_separator()
        a.add_command(label="Hacer respaldo ahora", command=self.respaldo)
        a.add_command(label="Abrir carpeta de datos", command=lambda: open_file(data_dir()))
        a.add_separator()
        a.add_command(label="Salir", command=self.on_close)
        m.add_cascade(label="Archivo", menu=a)
        s = tk.Menu(m, tearoff=0)
        s.add_command(label="Enviar mi información a jefatura...", command=self.enviar_paquete)
        s.add_command(label="Importar retroalimentación de jefatura...", command=self.importar_retro)
        s.add_separator()
        s.add_command(label="(Jefatura) Importar paquetes de especialistas...", command=self.importar_paquetes)
        s.add_command(label="(Jefatura) Importar todos los paquetes de una carpeta...", command=self.importar_carpeta)
        s.add_command(label="(Jefatura) Importar desde la carpeta compartida...", command=lambda: self.importar_carpeta(compartida=True))
        s.add_separator()
        s.add_command(label="Carpeta compartida del equipo...", command=self.config_carpeta)
        s.add_command(label="(Jefatura) Exportar retroalimentación a un especialista...", command=self.exportar_retro)
        m.add_cascade(label="Sincronización", menu=s)
        h = tk.Menu(m, tearoff=0)
        h.add_command(label="Guía rápida", command=self.guia)
        h.add_command(label="Acerca de", command=lambda: messagebox.showinfo(
            APP_NAME, f"{APP_TITLE}\nVersión {APP_VERSION}\n\nCRM de escritorio: los datos se guardan en este equipo "
                      f"({data_dir()})."))
        m.add_cascade(label="Ayuda", menu=h)

    def set_status(self, msg):
        self.status.config(text=msg)

    def clear_main(self):
        for w in self.main.winfo_children():
            w.destroy()
        self.view = None

    # ---- perfiles --------------------------------------------------------------
    def startup(self):
        ps = self.cfg.get("perfiles", [])
        u = self.cfg.get("ultimo", 0)
        if ps and 0 <= u < len(ps) and os.path.exists(os.path.join(data_dir(), ps[u]["db"])):
            try:
                self.open_profile(ps[u])
                return
            except Exception as e:
                messagebox.showerror(APP_NAME, f"No se pudo abrir el perfil:\n{e}")
        self.show_setup()

    def show_setup(self):
        self.store = None
        self.current = None
        self.view = None
        self.lbl_user.config(text="")
        if hasattr(self, "nav"):
            self.nav.delete(*self.nav.get_children())
        for b in (self.btn_enviar, self.btn_retro, self.btn_panel):
            b.pack_forget()
        self.clear_main()
        self.set_status("Seleccione o cree su perfil de trabajo.")
        f = ttk.Frame(self.main)
        f.pack(expand=True)
        ttk.Label(f, text="CRM INTEGRAL", font=("Segoe UI", 26, "bold"), foreground=COLOR_HEADER).pack(pady=(10, 0))
        ttk.Label(f, text="GROUP EJ MORALES S.U.R.L.", font=("Segoe UI", 13), foreground=COLOR_ACCENT).pack(pady=(0, 14))
        ps = self.cfg.get("perfiles", [])
        if ps:
            ttk.Label(f, text="Abrir un perfil existente:", font=("Segoe UI", 11, "bold")).pack(anchor="w")
            for i, p in enumerate(ps):
                txt = f"{p['nombre']}  ({'Jefatura' if p['rol'] == 'jefatura' else 'Especialista'}, código {p['codigo']})"
                ttk.Button(f, text=txt, command=lambda pp=p, ii=i: self._pick_profile(pp, ii)).pack(fill="x", pady=2)
            ttk.Separator(f).pack(fill="x", pady=12)
        ttk.Label(f, text="Crear un perfil nuevo:", font=("Segoe UI", 11, "bold")).pack(anchor="w")
        self.var_rol = tk.StringVar(value="especialista")
        rf = ttk.Frame(f)
        rf.pack(anchor="w", pady=4)
        ttk.Radiobutton(rf, text="Especialista (comercial)", variable=self.var_rol, value="especialista").pack(side="left", padx=(0, 14))
        ttk.Radiobutton(rf, text="Jefatura", variable=self.var_rol, value="jefatura").pack(side="left")
        ttk.Label(f, text="Nombre (ej. Comercial 01 o su nombre):").pack(anchor="w")
        self.ent_nombre = ttk.Entry(f, width=44)
        self.ent_nombre.pack(fill="x")
        ttk.Label(f, text="Código corto para sus números de registro (ej. C01):").pack(anchor="w", pady=(6, 0))
        self.ent_codigo = ttk.Entry(f, width=12)
        self.ent_codigo.pack(anchor="w")
        self.ent_nombre.bind("<KeyRelease>", lambda e: (self.ent_codigo.delete(0, "end"),
                                                        self.ent_codigo.insert(0, suggest_code(self.ent_nombre.get()))))
        ttk.Label(f, text="El código evita que dos especialistas generen el mismo número de registro "
                          "(CLI-C01-0001, CLI-C02-0001...).", font=("Segoe UI", 8, "italic")).pack(anchor="w")
        ttk.Button(f, text="Crear perfil y entrar", style="Accent.TButton", command=self._create_profile).pack(pady=12, ipadx=10)

    def _pick_profile(self, p, i):
        self.cfg["ultimo"] = i
        save_config(self.cfg)
        self.open_profile(p)

    def _create_profile(self):
        nombre = self.ent_nombre.get().strip()
        codigo = re.sub(r"[^A-Za-z0-9]", "", self.ent_codigo.get()).upper()[:5]
        rol = self.var_rol.get()
        if not nombre:
            messagebox.showwarning(APP_NAME, "Escriba el nombre.")
            return
        if rol == "jefatura":
            codigo = codigo or "JEF"
        if not codigo:
            codigo = suggest_code(nombre)
        dbname = f"{rol}_{codigo}.db"
        if os.path.exists(os.path.join(data_dir(), dbname)):
            messagebox.showwarning(APP_NAME, f"Ya existe un perfil con el código {codigo}. Use otro código.")
            return
        try:
            Store.create(os.path.join(data_dir(), dbname), rol, nombre, codigo).close()
        except Exception as e:
            messagebox.showerror(APP_NAME, f"No se pudo crear el perfil:\n{e}")
            return
        p = {"nombre": nombre, "rol": rol, "codigo": codigo, "db": dbname}
        self.cfg["perfiles"].append(p)
        self.cfg["ultimo"] = len(self.cfg["perfiles"]) - 1
        save_config(self.cfg)
        self.open_profile(p)

    def open_profile(self, p):
        if self.store:
            self.store.close()
        self.store = Store.open(os.path.join(data_dir(), p["db"]))
        self.profile = p
        self.role, self.user, self.code = p["rol"], p["nombre"], p["codigo"]
        self.lbl_user.config(text=f"{self.user}  |  {'Jefatura' if self.role == 'jefatura' else 'Especialista'}")
        self.btn_enviar.pack(fill="x", pady=2, before=self.nav)
        self.btn_retro.pack(fill="x", pady=2, before=self.nav)
        self.btn_panel.pack(fill="x", pady=(8, 4), before=self.nav)
        if self.role == "jefatura":
            self.btn_enviar.config(text="Importar paquetes", command=self.importar_paquetes)
            self.btn_retro.config(text="Devolver retroalimentación", command=self.exportar_retro)
        else:
            self.btn_enviar.config(text="Enviar a jefatura", command=self.enviar_paquete)
            self.btn_retro.config(text="Recibir retroalimentación", command=self.importar_retro)
        try:
            make_backup(self.store)
        except Exception:
            pass
        self.refresh_nav()
        self.go("__panel__")
        self.set_status(f"Perfil «{self.user}» abierto. Los datos se guardan automáticamente.")

    # ---- navegación -------------------------------------------------------------
    def refresh_nav(self):
        self.nav.delete(*self.nav.get_children())
        if not self.store:
            return
        for mod, sheets in MODULES:
            node = self.nav.insert("", "end", iid=f"mod_{mod}", text=mod, open=True, tags=("mod",))
            for s in sheets:
                if s in self.store.hojas:
                    n = len(self.store.hojas[s]["rows"])
                    self.nav.insert(node, "end", iid=s, text=f"   {SHEET_LABEL[s]}  ({n})")
        if self.role == "jefatura" and "LOG_CONSOLIDACION" in self.store.hojas:
            self.nav.insert("", "end", iid="LOG_CONSOLIDACION",
                            text=f"Registro de consolidaciones ({len(self.store.hojas['LOG_CONSOLIDACION']['rows'])})", tags=("mod",))
        if self.current in self.store.hojas and self.nav.exists(self.current):
            self.nav.selection_set(self.current)

    def _on_nav(self, e):
        sel = self.nav.selection()
        if sel and not sel[0].startswith("mod_") and sel[0] != self.current:
            self.go(sel[0])

    def go(self, where):
        if not self.store:
            return
        self.current = where
        self.clear_main()
        if hasattr(self, "nav") and self.nav.winfo_exists():
            cur = self.nav.selection()
            if where in self.store.hojas and self.nav.exists(where):
                if not cur or cur[0] != where:
                    self.nav.selection_set(where)
                    self.nav.see(where)
            elif cur:
                self.nav.selection_remove(*cur)
        if where == "__panel__":
            self.show_panel()
        elif where == "LOG_CONSOLIDACION":
            self.show_log()
        else:
            self.view = TableView(self, where)
            self.view.pack(fill="both", expand=True)

    def refresh_all(self):
        self.refresh_nav()
        if self.current == "__panel__":
            self.go("__panel__")
        elif self.view is not None:
            self.view.rebuild()

    def new_record(self, sheet, prefill=None, on_done=None):
        FormDialog(self, sheet, None, prefill=prefill, on_done=on_done)

    def open_record(self, sheet, rid):
        r, _ = self.store.find_row(sheet, rid)
        if r is not None:
            FormDialog(self, sheet, r)

    # ---- panel --------------------------------------------------------------
    def show_panel(self):
        f = ttk.Frame(self.main)
        f.pack(fill="both", expand=True)
        head = ttk.Frame(f)
        head.pack(fill="x")
        ttk.Label(head, text=f"Panel de control — {'Jefatura' if self.role == 'jefatura' else 'Especialista'}",
                  font=("Segoe UI", 16, "bold"), foreground=COLOR_HEADER).pack(side="left")
        ttk.Button(head, text="Exportar informe a Excel", command=self.exportar_informe).pack(side="right")
        qa = ttk.Frame(f)
        qa.pack(fill="x", pady=8)
        for txt, sh in (("+ Cliente", "Clientes"), ("+ Oportunidad", "Prospectos"), ("+ Actividad", "Actividades"),
                        ("+ Oferta", "Ofertas"), ("+ Proveedor", "Proveedores"), ("+ Incidencia", "Incidencias")):
            ttk.Button(qa, text=txt, style="Accent.TButton", command=lambda s=sh: self.new_record(s)).pack(side="left", padx=3)
        cards = ttk.Frame(f)
        cards.pack(fill="x", pady=4)
        S = self.store
        rows = lambda s: S.sheet(s)["rows"]
        abiertas = sum(1 for r in rows("Prospectos") if r.get("Etapa") not in ("Ganada", "Perdida", "Suspendida", ""))
        pend = [r for r in rows("Ofertas") if r.get("Estado_revision") in OFERTAS_PENDIENTES]
        venc = [(s, r) for s in ("Clientes", "Prospectos", "Actividades") for r in rows(s) if is_overdue(s, r)]
        inc = sum(1 for r in rows("Incidencias") if r.get("Estado") not in ("Resuelta", "Cerrada"))
        self._card(cards, "Clientes", len(rows("Clientes")), "#1f3a5f")
        self._card(cards, "Oportunidades abiertas", abiertas, "#2e86ab")
        self._card(cards, "Ofertas por revisar" if self.role == "jefatura" else "Ofertas en revisión", len(pend), "#d9822b")
        self._card(cards, "Ventas / operaciones", len(rows("Ventas")), "#2a9d8f")
        self._card(cards, "Incidencias abiertas", inc, "#8e44ad")
        self._card(cards, "Seguimientos vencidos", len(venc), "#c0392b")
        bottom = ttk.Frame(f)
        bottom.pack(fill="both", expand=True, pady=8)
        bottom.columnconfigure(0, weight=1)
        bottom.columnconfigure(1, weight=1)
        bottom.rowconfigure(0, weight=1)
        bottom.rowconfigure(1, weight=1)
        # agenda
        left = ttk.Frame(bottom)
        left.grid(row=0, column=0, rowspan=2, sticky="nsew", padx=(0, 8))
        ttk.Label(left, text="Agenda: seguimientos vencidos y próximos 7 días", font=("Segoe UI", 11, "bold"),
                  foreground="#c0392b").pack(anchor="w")
        wrap, tree = make_tree(left, ("est", "hoja", "ref", "fecha", "accion"), (55, 85, 150, 75, 185),
                               ("", "Módulo", "Cliente / entidad", "Fecha", "Próxima acción"), 12)
        wrap.pack(fill="both", expand=True)
        tree.tag_configure("v", foreground="#c0392b")
        hoy = datetime.date.today()
        lim = (hoy + datetime.timedelta(days=7)).strftime("%Y-%m-%d")
        n = 0
        agenda = []
        for s in ("Clientes", "Prospectos", "Actividades"):
            idf = ID_FIELD[s]
            for r in rows(s):
                fs = to_text(r.get("Fecha_seguimiento", ""))[:10]
                if not fs:
                    continue
                if s == "Prospectos" and r.get("Etapa") in ("Ganada", "Perdida", "Suspendida"):
                    continue
                if s == "Actividades" and r.get("Estado") in ("Realizada", "Cancelada"):
                    continue
                if fs <= lim:
                    agenda.append((fs, s, r, idf))
        for fs, s, r, idf in sorted(agenda, key=lambda x: x[0]):
            vencida = fs < today_str()
            tree.insert("", "end", iid=f"{s}|{r.get(idf)}", tags=("v",) if vencida else (),
                        values=("VENCIDO" if vencida else "Próximo", SHEET_LABEL[s],
                                (r.get("Razon_social") or r.get("Cliente") or r.get("Cliente_proveedor") or "")[:40],
                                fs, to_text(r.get("Proxima_accion", ""))[:60]))
            n += 1
        if n == 0:
            tree.insert("", "end", values=("", "", "Sin pendientes", "", ""))
        tree.bind("<Double-1>", lambda e: self._open_from_agenda(tree))
        # derecha arriba
        right = ttk.Frame(bottom)
        right.grid(row=0, column=1, sticky="nsew")
        if self.role == "jefatura":
            ttk.Label(right, text="Ofertas que requieren su revisión (doble clic para revisar)",
                      font=("Segoe UI", 11, "bold"), foreground="#d9822b").pack(anchor="w")
        else:
            ttk.Label(right, text="Mis ofertas pendientes de jefatura", font=("Segoe UI", 11, "bold"),
                      foreground="#d9822b").pack(anchor="w")
        wrap, t2 = make_tree(right, ("id", "cli", "est", "com"), (80, 150, 170, 118),
                             ("ID", "Cliente", "Estado", "Comercial"), 6)
        wrap.pack(fill="both", expand=True)
        for r in pend:
            if self.role != "jefatura" and norm_name(r.get("Comercial_responsable")) != norm_name(self.user):
                continue
            t2.insert("", "end", iid=r.get("ID_Oferta"), values=(r.get("ID_Oferta"), r.get("Cliente"),
                                                              r.get("Estado_revision"), r.get("Comercial_responsable")))
        if not t2.get_children():
            t2.insert("", "end", values=("", "— sin ofertas —", "", ""))

        def open_of(e):
            sel = t2.selection()
            if not sel:
                return
            r, _ = self.store.find_row("Ofertas", sel[0])
            if r is None:
                return
            ReviewDialog(self, r) if self.role == "jefatura" else FormDialog(self, "Ofertas", r)
        t2.bind("<Double-1>", open_of)
        # derecha abajo
        rb = ttk.Frame(bottom)
        rb.grid(row=1, column=1, sticky="nsew", pady=(8, 0))
        if self.role == "jefatura":
            hb = ttk.Frame(rb)
            hb.pack(fill="x")
            ttk.Label(hb, text="Resumen por especialista", font=("Segoe UI", 11, "bold"),
                      foreground=COLOR_HEADER).pack(side="left")
            ttk.Button(hb, text="Revisar duplicados", command=lambda: DuplicatesDialog(self)).pack(side="right")
            res = resumen_especialistas(self.store)
            cols = ("Especialista", "Oportunidades abiertas", "Ofertas", "Ofertas por revisar", "Ventas", "Valor ventas", "Seguim. vencidos")
            wrap, t3 = make_tree(rb, cols, (110, 80, 50, 70, 50, 70, 70), ("Especialista", "Oport.", "Ofertas", "Por rev.", "Ventas", "Valor", "Venc."), 5)
            wrap.pack(fill="both", expand=True)
            for r in res:
                t3.insert("", "end", values=[r[c] for c in cols])
            if not res:
                t3.insert("", "end", values=("— sin datos —", "", "", "", "", "", ""))
        else:
            ttk.Label(rb, text="Mi actividad reciente", font=("Segoe UI", 11, "bold"), foreground=COLOR_HEADER).pack(anchor="w")
            wrap, t3 = make_tree(rb, ("f", "t", "c", "e"), (105, 90, 185, 95), ("Fecha", "Tipo", "Cliente", "Estado"), 5)
            wrap.pack(fill="both", expand=True)
            acts = sorted(rows("Actividades"), key=lambda r: r.get("Fecha_hora", ""), reverse=True)[:30]
            for r in acts:
                t3.insert("", "end", values=(r.get("Fecha_hora"), r.get("Tipo_actividad"), r.get("Cliente_proveedor"), r.get("Estado")))
            if not acts:
                t3.insert("", "end", values=("— sin actividades —", "", "", ""))

    def _card(self, parent, text, value, color):
        c = tk.Frame(parent, bg="white", bd=1, relief="solid")
        c.pack(side="left", padx=5, ipadx=10, ipady=8, expand=True, fill="x")
        tk.Label(c, text=str(value), font=("Segoe UI", 20, "bold"), bg="white", fg=color).pack()
        tk.Label(c, text=text, font=("Segoe UI", 9), bg="white", fg="#555").pack()

    def _open_from_agenda(self, tree):
        sel = tree.selection()
        if sel and "|" in sel[0]:
            s, rid = sel[0].split("|", 1)
            self.open_record(s, rid)

    def show_log(self):
        f = ttk.Frame(self.main)
        f.pack(fill="both", expand=True)
        ttk.Label(f, text="Registro de consolidaciones", font=("Segoe UI", 16, "bold"), foreground=COLOR_HEADER).pack(anchor="w", pady=(0, 8))
        h = LOG_HEADERS
        wrap, t = make_tree(f, h, [120, 150, 120, 70, 90, 150, 150, 80, 200], [x.replace("_", " ") for x in h], 20)
        wrap.pack(fill="both", expand=True)
        for r in reversed(self.store.sheet("LOG_CONSOLIDACION")["rows"]):
            t.insert("", "end", values=[r.get(c, "") for c in h])

    # ---- sincronización: especialista ---------------------------------------
    def enviar_paquete(self):
        if self.role != "especialista":
            messagebox.showinfo(APP_NAME, "Esta opción es para especialistas.")
            return
        pkg = build_package(self.store, self.user)
        tot_n = sum(len(v["nuevas"]) for v in pkg["cambios"].values())
        tot_m = sum(len(v["modificadas"]) for v in pkg["cambios"].values())
        if not pkg["cambios"] and load_snapshot(self.store.path):
            if not messagebox.askyesno(APP_NAME, "No hay cambios desde su último envío.\n¿Generar el paquete de todos modos?"):
                return
        carpeta = os.path.join(data_dir(), "paquetes_enviados")
        os.makedirs(carpeta, exist_ok=True)
        fn = f"Paquete_{self.user.replace(' ', '_')}_{datetime.datetime.now():%Y%m%d_%H%M}.json"
        compartida = self.cfg.get("carpeta_compartida")
        dest = None
        if compartida and messagebox.askyesno(APP_NAME,
                "¿Guardar el paquete directamente en la carpeta compartida del equipo?\n\n" +
                compartida + "\n\n(Si responde «No», se abrirá el diálogo para elegir otra ubicación.)"):
            sub = os.path.join(compartida, "paquetes")
            try:
                os.makedirs(sub, exist_ok=True)
                dest = os.path.join(sub, fn)
            except Exception as e:
                messagebox.showwarning(APP_NAME, f"No se pudo escribir en la carpeta compartida:\n{e}")
        if dest is None:
            dest = filedialog.asksaveasfilename(title="Guardar paquete para jefatura", initialdir=carpeta,
                                                initialfile=fn, defaultextension=".json",
                                                filetypes=[("Paquete CRM", "*.json")], parent=self)
            if not dest:
                return
        write_json(pkg, dest)
        save_snapshot(self.store.path, self.store.hojas)
        detalle = "\n".join(f"  • {SHEET_LABEL.get(k, k)}: {len(v['nuevas'])} nuevas, {len(v['modificadas'])} modificadas"
                            for k, v in pkg["cambios"].items()) or "  • Sin cambios desde el último envío."
        self.set_status(f"Paquete generado: {os.path.basename(dest)}")
        aviso = ("El paquete quedó en la carpeta compartida; jefatura ya puede importarlo con "
                 "«(Jefatura) Importar desde la carpeta compartida»." if compartida and dest.startswith(compartida)
                 else "Envíe ese archivo a jefatura (correo, mensajería o USB).")
        messagebox.showinfo(APP_NAME, "PAQUETE LISTO PARA ENVIAR A JEFATURA\n\n" f"Archivo:\n{dest}\n\n"
                                      f"Cambios desde el último envío:\n{detalle}\n\n"
                                      f"Total: {tot_n} nuevas, {tot_m} modificadas.\n\n" + aviso)
        open_file(os.path.dirname(dest))

    def importar_retro(self):
        if self.role != "especialista":
            messagebox.showinfo(APP_NAME, "Esta opción es para especialistas.")
            return
        compartida = self.cfg.get("carpeta_compartida")
        inicial = None
        if compartida:
            sub = os.path.join(compartida, "retroalimentaciones")
            if os.path.isdir(sub):
                inicial = sub
        path = filedialog.askopenfilename(title="Archivo de retroalimentación de jefatura",
                                          initialdir=inicial,
                                          filetypes=[("Retroalimentación", "*.json")], parent=self)
        if not path:
            return
        try:
            retro = load_package(path)
        except Exception as e:
            messagebox.showerror(APP_NAME, f"Archivo no válido:\n{e}")
            return
        if retro.get("tipo") != "RETRO_JEFATURA":
            messagebox.showerror(APP_NAME, "Este archivo no es una retroalimentación de jefatura.")
            return
        if norm_name(retro.get("especialista")) != norm_name(self.user):
            if not messagebox.askyesno(APP_NAME, f"Esta retroalimentación es para «{retro.get('especialista')}», "
                                                 f"no para «{self.user}». ¿Aplicarla de todos modos?"):
                return
        ap = apply_retro(self.store, retro)
        self.refresh_all()
        if ap:
            messagebox.showinfo(APP_NAME, "Se actualizaron estas ofertas:\n" + "\n".join(f"• {i}: {', '.join(c)}" for i, c in ap))
        else:
            messagebox.showinfo(APP_NAME, "No había novedades para sus ofertas.")

    # ---- sincronización: jefatura -----------------------------------------------
    def importar_paquetes(self):
        if self.role != "jefatura":
            messagebox.showinfo(APP_NAME, "Esta opción es para jefatura.")
            return
        compartida = self.cfg.get("carpeta_compartida")
        inicial = None
        if compartida:
            sub = os.path.join(compartida, "paquetes")
            if os.path.isdir(sub):
                inicial = sub
        paths = filedialog.askopenfilenames(title="Paquetes de especialistas", initialdir=inicial,
                                            filetypes=[("Paquete CRM", "*.json")], parent=self)
        self._import_paths(paths)

    def config_carpeta(self):
        """Define la carpeta común (red, OneDrive, Google Drive o USB) que todo el equipo usa para intercambiar paquetes."""
        d = filedialog.askdirectory(title="Carpeta compartida del equipo (red, OneDrive, Google Drive o USB)",
                                    parent=self)
        if not d:
            return
        self.cfg["carpeta_compartida"] = d
        save_config(self.cfg)
        self.set_status(f"Carpeta compartida del equipo: {d}")
        messagebox.showinfo(APP_NAME, "CARPETA COMPARTIDA GUARDADA\n\n" + d +
                            "\n\n• Especialistas: al usar «Enviar a jefatura» se les ofrecerá dejar el paquete "
                            "directamente en la subcarpeta «paquetes».\n"
                            "• Jefatura: use «(Jefatura) Importar desde la carpeta compartida» para consolidar todos "
                            "los paquetes de una vez.\n"
                            "• Las retroalimentaciones se guardarán en la subcarpeta «retroalimentaciones».")

    def importar_carpeta(self, compartida=False):
        if self.role != "jefatura":
            messagebox.showinfo(APP_NAME, "Esta opción es para jefatura.")
            return
        if compartida:
            d = self.cfg.get("carpeta_compartida")
            if not d:
                messagebox.showinfo(APP_NAME, "Todavía no se ha configurado la carpeta compartida.\n\n"
                                              "Use Sincronización > Carpeta compartida del equipo...")
                return
        else:
            d = filedialog.askdirectory(title="Carpeta con los paquetes de los especialistas", parent=self)
        if d:
            sub = os.path.join(d, "paquetes")
            base = sub if os.path.isdir(sub) else d
            self._import_paths([os.path.join(base, f) for f in sorted(os.listdir(base))
                                if f.lower().endswith(".json")])

    def _import_paths(self, paths):
        hechos = 0
        for p in paths:
            try:
                pkg = load_package(p)
            except Exception as e:
                messagebox.showerror(APP_NAME, f"No se pudo leer {os.path.basename(p)}:\n{e}")
                continue
            if pkg.get("tipo") != "PAQUETE_ESPECIALISTA":
                continue
            if self.process_package(pkg, os.path.basename(p)):
                hechos += 1
        if hechos:
            dups = sum(len(find_duplicates(self.store, s)) for s in ("Clientes", "Proveedores"))
            if dups and messagebox.askyesno(APP_NAME, f"Se detectaron {dups} grupos de posibles duplicados "
                                                      "(mismo NIT o razón social con distinto ID).\n¿Revisarlos ahora?"):
                DuplicatesDialog(self)

    def process_package(self, pkg, label="", log=True):
        reporte = analyze_import(self.store, pkg)
        n = sum(len(r["nuevas"]) for r in reporte.values())
        a = len(set().union(*[set(r["rellenos"]) | set(r["forzados"]) for r in reporte.values()])) if reporte else 0
        c = sum(len(r["conflictos"]) for r in reporte.values())
        ig = sum(len(r["ignoradas"]) for r in reporte.values())
        if not reporte:
            messagebox.showinfo(APP_NAME, f"{pkg.get('especialista', '')}: sin novedades "
                                          "(todo ya está consolidado).", parent=self)
            return False
        lineas = [f"• {SHEET_LABEL.get(k, k)}: {len(v['nuevas'])} nuevas, "
                  f"{len(set(v['rellenos']) | set(v['forzados']))} actualizadas, {len(v['conflictos'])} conflictos"
                  for k, v in reporte.items()]
        txt = (f"PAQUETE DE: {pkg.get('especialista', '')}   ({str(pkg.get('generado', ''))[:16]})\n\n" + "\n".join(lineas) +
               f"\n\nTOTAL: {n} nuevas | {a} actualizadas | {c} conflictos" +
               (f" | {ig} cambios ignorados (el maestro manda en estados de revisión)" if ig else "") +
               ("\n\nSe abrirá el panel de conflictos." if c else "") + "\n\n¿Importar?")
        if not messagebox.askyesno(APP_NAME, txt, parent=self):
            return False
        decisions = {}
        if c:
            decisions = ConflictDialog(self, reporte).run()
            if decisions is None:
                self.set_status("Importación cancelada.")
                return False
        stats = apply_import(self.store, pkg, reporte, decisions)
        if log and self.role == "jefatura":
            registrar_log(self.store, pkg, stats, label)
        self.refresh_all()
        self.set_status(f"{pkg.get('especialista', '')}: {stats['nuevas']} nuevas, {stats['actualizadas']} actualizadas.")
        messagebox.showinfo(APP_NAME, "CONSOLIDACIÓN COMPLETADA\n\n"
                                      f"Nuevas: {stats['nuevas']}\nActualizadas: {stats['actualizadas']}\n"
                                      f"Conflictos a favor del especialista: {stats['conflictos_especialista']}\n"
                                      f"Conflictos conservados en el maestro: {stats['conflictos_maestro']}", parent=self)
        return True

    def exportar_retro(self):
        if self.role != "jefatura":
            messagebox.showinfo(APP_NAME, "Esta opción es para jefatura.")
            return
        nombres = sorted({(r.get("Comercial_responsable") or "").strip() for r in self.store.sheet("Ofertas")["rows"]} -
                         {""} | {r.get("Nombre", "") for r in self.store.sheet("Equipo")["rows"]
                                 if r.get("Cargo") != "Jefatura"} - {""})
        if not nombres:
            messagebox.showinfo(APP_NAME, "Aún no hay especialistas con ofertas consolidadas.")
            return
        res = SimpleDialog(self, "Retroalimentación a especialista",
                           [("Especialista", nombres[0], nombres)], ok="Generar archivo").run()
        if not res:
            return
        esp = res["Especialista"]
        retro = build_retro(self.store, esp)
        carpeta = os.path.join(data_dir(), "retroalimentaciones")
        os.makedirs(carpeta, exist_ok=True)
        dest = filedialog.asksaveasfilename(title="Guardar retroalimentación", initialdir=carpeta,
                                            initialfile=f"Retro_{esp.replace(' ', '_')}_{datetime.date.today():%Y%m%d}.json",
                                            defaultextension=".json", filetypes=[("Retroalimentación", "*.json")], parent=self)
        if dest:
            write_json(retro, dest)
            compartida = self.cfg.get("carpeta_compartida")
            if compartida:
                try:
                    sub = os.path.join(compartida, "retroalimentaciones")
                    os.makedirs(sub, exist_ok=True)
                    write_json(retro, os.path.join(sub, os.path.basename(dest)))
                    self.set_status("Retroalimentación generada (copia dejada en la carpeta compartida).")
                except Exception:
                    pass
            messagebox.showinfo(APP_NAME, f"Archivo generado con {len(retro['ofertas'])} ofertas de {esp}.\n"
                                          f"Envíeselo para que lo reciba con «Recibir retroalimentación».")
            open_file(os.path.dirname(dest))

    # ---- Excel --------------------------------------------------------------------
    def importar_excel(self):
        path = filedialog.askopenfilename(title="Importar datos desde un libro CRM de Excel",
                                          filetypes=[("Libros Excel", "*.xlsx")], parent=self)
        if not path:
            return
        try:
            xb = ExcelBook(path)
        except Exception as e:
            messagebox.showerror(APP_NAME, f"No se pudo leer el libro:\n{e}")
            return
        pkg = {"tipo": "PAQUETE_ESPECIALISTA", "especialista": "Importación de Excel",
               "generado": now_iso(), "hojas": {n: xb.hojas[n]["rows"] for n in xb.hojas if n in DATA_SHEETS}}
        self.process_package(pkg, os.path.basename(path), log=True)

    def exportar_libro(self):
        dest = filedialog.asksaveasfilename(
            title="Exportar todo el CRM a Excel", defaultextension=".xlsx", filetypes=[("Libros Excel", "*.xlsx")],
            initialfile=("CRM_Maestro_Jefatura_GroupEJ_Morales_" if self.role == "jefatura" else
                         f"CRM_{self.user.replace(' ', '_')}_GroupEJ_Morales_") + today_str() + ".xlsx", parent=self)
        if not dest:
            return
        try:
            export_workbook(self.store, dest, self.role)
        except Exception as e:
            messagebox.showerror(APP_NAME, f"Error al exportar:\n{e}")
            return
        self.set_status(f"Libro exportado: {dest}")
        if messagebox.askyesno(APP_NAME, f"Libro exportado con todas las hojas del CRM:\n{dest}\n\n¿Abrirlo?"):
            open_file(dest)

    def exportar_hoja(self, name):
        info = self.store.sheet(name)
        dest = filedialog.asksaveasfilename(title="Exportar a Excel", defaultextension=".xlsx",
                                            filetypes=[("Libros Excel", "*.xlsx")],
                                            initialfile=f"{name}_{today_str()}.xlsx", parent=self)
        if not dest:
            return
        try:
            export_sheet_to_xlsx(dest, name, info["headers"], info["rows"], title=f"{APP_NAME} — {SHEET_LABEL.get(name, name)}")
        except Exception as e:
            messagebox.showerror(APP_NAME, f"Error al exportar:\n{e}")
            return
        if messagebox.askyesno(APP_NAME, f"Exportado:\n{dest}\n\n¿Abrirlo?"):
            open_file(dest)

    def exportar_informe(self):
        S, hoy = self.store, today_str()
        rows = lambda s: S.sheet(s)["rows"]
        pend = [r for r in rows("Ofertas") if r.get("Estado_revision") in OFERTAS_PENDIENTES]
        agenda = []
        for s in ("Clientes", "Prospectos", "Actividades"):
            for r in rows(s):
                if is_overdue(s, r):
                    agenda.append((SHEET_LABEL[s], r.get(ID_FIELD[s], ""),
                                   r.get("Razon_social") or r.get("Cliente") or r.get("Cliente_proveedor") or "",
                                   to_text(r.get("Fecha_seguimiento"))[:10], r.get("Proxima_accion", "")))
        resumen = [("Indicador", "Valor"), ("Fecha del informe", hoy), ("Usuario", self.user),
                   ("Rol", "Jefatura" if self.role == "jefatura" else "Especialista"),
                   ("Clientes", len(rows("Clientes"))), ("Proveedores", len(rows("Proveedores"))),
                   ("Oportunidades abiertas", sum(1 for r in rows("Prospectos") if r.get("Etapa") not in ("Ganada", "Perdida", "Suspendida", ""))),
                   ("Ofertas", len(rows("Ofertas"))), ("Ofertas pendientes de revisión", len(pend)),
                   ("Ventas / operaciones", len(rows("Ventas"))),
                   ("Valor de ventas", fmt_num(sum(to_num(r.get("Importe")) or 0 for r in rows("Ventas")))),
                   ("Seguimientos vencidos", len(agenda))]
        dest = filedialog.asksaveasfilename(title="Exportar informe", defaultextension=".xlsx",
                                            filetypes=[("Libros Excel", "*.xlsx")],
                                            initialfile=f"Informe_{self.user.replace(' ', '_')}_{hoy}.xlsx", parent=self)
        if not dest:
            return
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Resumen"
        for r_i, (k, v) in enumerate(resumen, 1):
            a, b = ws.cell(r_i, 1, k), ws.cell(r_i, 2, v)
            if r_i == 1:
                for c in (a, b):
                    c.font = Font(bold=True, color="FFFFFF")
                    c.fill = PatternFill("solid", fgColor="1F3A5F")
            else:
                a.font = Font(bold=True)
        ws.column_dimensions["A"].width = 34
        ws.column_dimensions["B"].width = 30

        def hoja(nombre, heads, data, widths):
            w = wb.create_sheet(nombre)
            for c, h in enumerate(heads, 1):
                x = w.cell(1, c, h)
                x.font = Font(bold=True, color="FFFFFF")
                x.fill = PatternFill("solid", fgColor="1F3A5F")
                w.column_dimensions[get_column_letter(c)].width = widths[c - 1]
            for r_i, row in enumerate(data, 2):
                for c, v in enumerate(row, 1):
                    w.cell(r_i, c, v)
        hoja("Seguimientos_vencidos", ["Módulo", "ID", "Cliente / entidad", "Fecha", "Próxima acción"], agenda, [22, 18, 40, 12, 50])
        hoja("Ofertas_pendientes", ["ID_Oferta", "Cliente", "Producto_servicio", "Comercial", "Estado_revision"],
             [(r.get("ID_Oferta"), r.get("Cliente"), r.get("Producto_servicio"), r.get("Comercial_responsable"),
               r.get("Estado_revision")) for r in pend], [18, 34, 34, 24, 34])
        if self.role == "jefatura":
            res = resumen_especialistas(S)
            if res:
                cols = list(res[0].keys())
                hoja("Resumen_especialistas", cols, [[r[c] for c in cols] for r in res], [24] + [18] * (len(cols) - 1))
        wb.save(dest)
        if messagebox.askyesno(APP_NAME, f"Informe exportado:\n{dest}\n\n¿Abrirlo?"):
            open_file(dest)

    def calcular_rendimiento(self):
        hoy = datetime.date.today()
        res = SimpleDialog(self, "Calcular rendimiento", [
            ("Desde (AAAA-MM-DD)", hoy.replace(day=1).strftime("%Y-%m-%d"), None),
            ("Hasta (AAAA-MM-DD)", today_str(), None)], ok="Calcular").run()
        if not res:
            return
        ini, fin = res["Desde (AAAA-MM-DD)"], res["Hasta (AAAA-MM-DD)"]
        if not (parse_date(ini) and parse_date(fin)):
            messagebox.showwarning(APP_NAME, "Fechas no válidas.")
            return
        if self.role == "jefatura":
            nombres = [r["Especialista"] for r in resumen_especialistas(self.store)]
        else:
            nombres = [self.user]
        if not nombres:
            messagebox.showinfo(APP_NAME, "Aún no hay datos de especialistas.")
            return
        for n in nombres:
            row, previo = calc_rendimiento(self.store, n, ini, fin)
            full = {h: row.get(h, "") for h in self.store.sheet("Rendimiento")["headers"]}
            self.store.upsert("Rendimiento", full, original=previo)
        self.refresh_all()
        self.set_status(f"Rendimiento calculado para {len(nombres)} persona(s). Complete «Meta_ventas» editando la fila.")

    def respaldo(self):
        try:
            d = make_backup(self.store)
            messagebox.showinfo(APP_NAME, f"Respaldo guardado en:\n{d}")
        except Exception as e:
            messagebox.showerror(APP_NAME, f"No se pudo hacer el respaldo:\n{e}")

    def guia(self):
        messagebox.showinfo("Guía rápida",
                            "1. Cree su perfil (especialista o jefatura) la primera vez.\n"
                            "2. Registre clientes, oportunidades, actividades y ofertas con los botones «+» del panel "
                            "o desde cada módulo. Todo se guarda solo.\n"
                            "3. ESPECIALISTA: al final del día use «Enviar a jefatura» y mande el archivo generado "
                            "(o déjelo en la carpeta compartida del equipo: Sincronización > Carpeta compartida).\n"
                            "4. JEFATURA: use «Importar paquetes» o «Importar desde la carpeta compartida», "
                            "resuelva conflictos y revise las ofertas.\n"
                            "5. Para trabajar en Excel: Archivo > Exportar TODO el CRM a un libro Excel.\n"
                            "6. Para cargar datos que ya tenga en Excel: Archivo > Importar datos desde Excel.")

    def on_close(self):
        if self.store:
            self.store.close()
        self.destroy()


def main():
    if "--selftest" in sys.argv:
        i = sys.argv.index("--selftest")
        out = sys.argv[i + 1] if len(sys.argv) > i + 1 else "selftest.txt"
        sys.exit(selftest(out))
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    App().mainloop()


if __name__ == "__main__":
    try:
        main()
    except SystemExit:
        raise
    except Exception:
        err = traceback.format_exc()
        try:
            r = tk.Tk()
            r.withdraw()
            messagebox.showerror(APP_NAME, "Error inesperado:\n\n" + err)
        except Exception:
            print(err)
