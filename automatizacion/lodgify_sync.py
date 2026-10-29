"""Sincroniza las reservas de Lodgify con el Excel de Canavida y envía el reporte.

Uso:
    python automatizacion/lodgify_sync.py                  # API de Lodgify + correo
    python automatizacion/lodgify_sync.py --csv OTC.csv    # desde un reporte exportado
    python automatizacion/lodgify_sync.py --sin-correo     # solo actualiza el Excel
    python automatizacion/lodgify_sync.py --hoy 2026-10-29 # simula otra fecha

Qué hace en cada corrida:
  1. Descarga las reservas de Lodgify (o lee el CSV "OTC Report").
  2. Mantiene la hoja "Reservas Lodgify" con todas las reservas (también futuras).
  3. Anota en "Ingresos Diarios", en orden de fecha, cada reserva confirmada cuya
     llegada ya ocurrió (o todas, si REGISTRAR_FUTURAS=true), sin duplicar las
     filas que ya existen (se marcan con "Lodgify #<id>" en la columna J).
  4. Corrige las fórmulas rotas conocidas del libro.
  5. Envía a admicanavida@gmail.com el reporte de entradas del día con el Excel.
"""
from __future__ import annotations

import argparse
import csv
import html
import json
import logging
import re
import smtplib
import sys
import time
import unicodedata
import urllib.error
import urllib.parse
import urllib.request
from collections import OrderedDict
from copy import copy
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from email.message import EmailMessage
from pathlib import Path
from zoneinfo import ZoneInfo

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

sys.path.insert(0, str(Path(__file__).resolve().parent))
import config as cfg  # noqa: E402

log = logging.getLogger("lodgify_sync")


# --------------------------------------------------------------------------
# Modelo
# --------------------------------------------------------------------------
@dataclass
class Reserva:
    id: str
    canal: str
    estado: str
    villa: str
    villa_lodgify: str
    huesped: str
    email: str
    llegada: date
    salida: date
    huespedes: int
    monto: float
    moneda: str = "USD"
    avisos: list = field(default_factory=list)

    @property
    def noches(self) -> int:
        return max((self.salida - self.llegada).days, 0)

    @property
    def confirmada(self) -> bool:
        return self.estado.lower() in cfg.ESTADOS_CONFIRMADOS

    @property
    def cancelada(self) -> bool:
        return self.estado.lower() in cfg.ESTADOS_CANCELADOS


def mapear_villa(nombre: str) -> str | None:
    texto = quitar_acentos(nombre).upper()
    for clave, villa in cfg.VILLAS:
        if re.search(rf"\b{clave}\b", texto):
            return villa
    return None


def mapear_canal(*textos: str) -> str:
    for texto in textos:
        t = (texto or "").lower()
        for clave, canal in cfg.CANALES:
            if re.search(rf"\b{re.escape(clave)}", t):
                return canal
    return "Directo / WhatsApp"


def quitar_acentos(texto: str) -> str:
    texto = unicodedata.normalize("NFKD", texto or "")
    return "".join(c for c in texto if unicodedata.category(c)[0] in "LNZP")


def tokens_nombre(nombre: str) -> list[str]:
    return [t for t in re.split(r"[^a-z0-9]+", quitar_acentos(nombre).lower()) if len(t) > 1]


def parse_fecha(valor) -> date | None:
    if valor is None or valor == "":
        return None
    if isinstance(valor, datetime):
        return valor.date()
    if isinstance(valor, date):
        return valor
    texto = str(valor).strip()[:10]
    for formato in ("%Y-%m-%d", "%m/%d/%Y", "%d/%m/%Y"):
        try:
            return datetime.strptime(texto, formato).date()
        except ValueError:
            pass
    return None


def a_numero(valor) -> float:
    """Convierte montos del libro (números o fórmulas simples como =10000/459)."""
    if isinstance(valor, (int, float)):
        return float(valor)
    if isinstance(valor, str) and re.fullmatch(r"=?[\d.+\-*/() ]+", valor.strip()):
        try:
            return float(eval(valor.strip().lstrip("="), {"__builtins__": {}}))  # solo aritmética
        except (SyntaxError, ZeroDivisionError, TypeError):
            return 0.0
    return 0.0


# --------------------------------------------------------------------------
# Fuentes de datos
# --------------------------------------------------------------------------
def leer_csv(ruta: Path) -> list[Reserva]:
    """Lee el reporte "OTC Report" exportado de Lodgify (una fila por concepto)."""
    grupos: "OrderedDict[str, list[dict]]" = OrderedDict()
    with open(ruta, encoding="utf-8-sig", newline="") as f:
        for fila in csv.DictReader(f):
            grupos.setdefault(fila["Id"], []).append(fila)

    reservas = []
    for rid, filas in grupos.items():
        f0 = filas[0]
        canal = mapear_canal(f0["Source"], f0.get("SourceText", ""))
        monto = 0.0
        for linea in filas:
            if linea["LineItem"] == "Tax" and canal in cfg.CANALES_SIN_IMPUESTO:
                continue
            monto += float(linea["GrossAmount"] or 0)
        reservas.append(construir_reserva(
            rid=rid, canal=canal, estado=f0["Status"], villa_lodgify=f0["HouseName"],
            huesped=f0["GuestName"], email=f0["GuestEmail"],
            llegada=parse_fecha(f0["DateArrival"]), salida=parse_fecha(f0["DateDeparture"]),
            huespedes=int(f0.get("NumberOfGuests") or 0), monto=monto,
            moneda=f0.get("Currency") or "USD",
        ))
    return reservas


class LodgifyAPI:
    def __init__(self, api_key: str, base: str = cfg.LODGIFY_API_URL):
        if not api_key:
            raise RuntimeError(
                "Falta LODGIFY_API_KEY. Créela en Lodgify → Configuración → API pública "
                "y guárdela como Secret del repositorio.")
        self.api_key = api_key
        self.base = base.rstrip("/")

    def get(self, ruta: str, **params):
        url = f"{self.base}{ruta}?{urllib.parse.urlencode(params)}"
        peticion = urllib.request.Request(url, headers={
            "X-ApiKey": self.api_key, "Accept": "application/json",
            "User-Agent": "canavida-sync/1.0"})
        for intento in range(5):
            try:
                with urllib.request.urlopen(peticion, timeout=60) as r:
                    return json.loads(r.read().decode("utf-8") or "null")
            except urllib.error.HTTPError as e:
                if e.code in (401, 403):
                    raise RuntimeError(
                        f"Lodgify rechazó la clave de API (HTTP {e.code}). Revise que "
                        "LODGIFY_API_KEY sea la clave de 'API pública' vigente.") from e
                if e.code not in (429, 500, 502, 503, 504) or intento == 4:
                    raise
            except urllib.error.URLError:
                if intento == 4:
                    raise
            espera = 2 ** (intento + 1)
            log.warning("Lodgify no respondió, reintento en %ss", espera)
            time.sleep(espera)

    def paginar(self, ruta: str, **params):
        pagina = 1
        while True:
            datos = self.get(ruta, page=pagina, size=50, includeCount="true", **params)
            items = datos.get("items", []) if isinstance(datos, dict) else (datos or [])
            yield from items
            total = datos.get("count") if isinstance(datos, dict) else None
            if not items or len(items) < 50 or (total is not None and pagina * 50 >= total):
                return
            pagina += 1


def buscar_impuesto(obj) -> float:
    """Suma los importes de impuesto que aparezcan en el detalle de la cotización."""
    total = 0.0
    if isinstance(obj, dict):
        for clave, valor in obj.items():
            k = clave.lower()
            if isinstance(valor, (int, float)) and k in ("tax", "taxes", "vat", "total_tax", "tax_amount"):
                total += float(valor)
            elif isinstance(valor, (dict, list)):
                total += buscar_impuesto(valor)
    elif isinstance(obj, list):
        for v in obj:
            total += buscar_impuesto(v)
    return total


def leer_api(api: LodgifyAPI, desde: date) -> list[Reserva]:
    propiedades = {}
    for p in api.paginar("/v2/properties"):
        propiedades[p.get("id")] = p.get("name") or ""

    reservas = []
    for b in api.paginar("/v2/reservations/bookings", stayFilter="All",
                         includeExternal="true", includeQuoteDetails="true",
                         updatedSince=(desde - timedelta(days=365)).isoformat() + "T00:00"):
        llegada = parse_fecha(b.get("arrival"))
        if not llegada or llegada < desde:
            continue
        guest = b.get("guest") or {}
        canal = mapear_canal(str(b.get("source") or ""), str(b.get("source_text") or ""))
        monto = float(b.get("total_amount") or 0)
        avisos = []
        if canal in cfg.CANALES_SIN_IMPUESTO:
            impuesto = buscar_impuesto(b.get("quote"))
            if impuesto:
                monto -= impuesto
            else:
                avisos.append("Monto de Expedia incluye impuesto (Lodgify no lo desglosó)")
        estado = str(b.get("status") or "")
        if b.get("is_deleted"):
            estado = "Deleted"
        huespedes = sum(int(r.get("people") or 0) for r in b.get("rooms") or [])
        r = construir_reserva(
            rid=str(b.get("id")), canal=canal, estado=estado,
            villa_lodgify=propiedades.get(b.get("property_id"), b.get("property_name") or ""),
            huesped=guest.get("name") or "", email=guest.get("email") or "",
            llegada=llegada, salida=parse_fecha(b.get("departure")) or llegada,
            huespedes=huespedes, monto=monto, moneda=b.get("currency_code") or "USD")
        r.avisos.extend(avisos)
        reservas.append(r)
    return reservas


def construir_reserva(*, rid, canal, estado, villa_lodgify, huesped, email, llegada,
                      salida, huespedes, monto, moneda) -> Reserva:
    villa = mapear_villa(villa_lodgify)
    r = Reserva(id=str(rid), canal=canal, estado=estado or "", villa=villa or villa_lodgify,
                villa_lodgify=villa_lodgify, huesped=(huesped or "").strip(),
                email=(email or "").strip(), llegada=llegada, salida=salida,
                huespedes=huespedes, monto=round(monto, 2), moneda=moneda)
    if not villa:
        r.avisos.append(f"Villa sin equivalencia en 'Listas': {villa_lodgify}")
    if moneda and moneda.upper() != "USD":
        r.avisos.append(f"Moneda {moneda}: el libro trabaja en USD, revisar el monto")
    return r


# --------------------------------------------------------------------------
# Excel
# --------------------------------------------------------------------------
def reparar_formulas(wb) -> list[str]:
    """Corrige errores del libro. Es idempotente: si ya están bien no cambia nada."""
    cambios = []

    # 'Utilidad y Posibles Gastos' F13:F28: apuntaban a filas de otro orden y
    # cuatro quedaron en #REF!. Se busca el gasto real por nombre de categoría.
    ws = wb["Utilidad y Posibles Gastos"]
    for fila in range(13, 29):
        nueva = (f"=SUMIF('Gastos e Ingresos'!$A$13:$A$24,$A{fila},"
                 f"'Gastos e Ingresos'!$E$13:$E$24)")
        if ws.cell(fila, 6).value != nueva:
            ws.cell(fila, 6).value = nueva
            cambios.append(f"Utilidad y Posibles Gastos!F{fila}")

    # 'Arqueo y Resumen Caja' fila 55: COUNTIFS sobre una columna eliminada (#REF!).
    # Ahora cuenta los movimientos de caja con "No" en la columna Notas.
    ws = wb["Arqueo y Resumen Caja"]
    for col, mes in ((2, 5), (3, 6), (4, 7)):
        celda = ws.cell(55, col)
        if isinstance(celda.value, str) and "#REF!" in celda.value:
            celda.value = (f"=COUNTIFS('Caja Efectivo'!$M$6:$M$1005,\"No\","
                           f"'Caja Efectivo'!$A$6:$A$1005,\">=\"&Listas!$L${mes},"
                           f"'Caja Efectivo'!$A$6:$A$1005,\"<=\"&Listas!$M${mes})")
            cambios.append(f"Arqueo y Resumen Caja!{celda.coordinate}")

    # 'Ingresos Diarios'!H4 solo sumaba hasta la fila 200 y las fórmulas del
    # resumen hasta la 1026: se amplían a todo el registro (fila 2026).
    ws = wb[cfg.HOJA_INGRESOS]
    if ws["H4"].value == "=SUM(H6:H200)":
        ws["H4"].value = f"=SUM(H6:H{cfg.ULTIMA_FILA_INGRESOS})"
        cambios.append("Ingresos Diarios!H4")
    for hoja in wb.worksheets:
        for fila in hoja.iter_rows():
            for celda in fila:
                v = celda.value
                if isinstance(v, str) and v.startswith("=") and "'Ingresos Diarios'!" in v and "$1026" in v:
                    celda.value = v.replace("$1026", f"${cfg.ULTIMA_FILA_INGRESOS}")
                    cambios.append(f"{hoja.title}!{celda.coordinate}")
    # 'Gastos e Ingresos' B5:B9 (octubre) tienen montos pegados como valor, por lo
    # que no suman nada nuevo. Se conserva el monto y se le suman las filas que
    # la automatización agrega (nota que empieza con "Lodgify #").
    resumen = wb["Gastos e Ingresos"]
    for fila in range(5, 10):
        celda = resumen.cell(fila, 2)
        if isinstance(celda.value, (int, float)):
            rango = lambda col: f"'Ingresos Diarios'!${col}${cfg.PRIMERA_FILA_SUMADA}:${col}${cfg.ULTIMA_FILA_INGRESOS}"
            celda.value = (f"={celda.value:g}+SUMIFS({rango('H')},{rango('C')},$A{fila},"
                           f"{rango('J')},\"{cfg.MARCA_LODGIFY}*\","
                           f"{rango('A')},\">=\"&Listas!$L$5,{rango('A')},\"<=\"&Listas!$M$5)")
            cambios.append(f"Gastos e Ingresos!{celda.coordinate}")
    for coord in ("B4", "F4"):
        v = ws[coord].value
        if isinstance(v, str) and "1026" in v:
            ws[coord].value = v.replace("1026", str(cfg.ULTIMA_FILA_INGRESOS))
            cambios.append(f"Ingresos Diarios!{coord}")
    return cambios


def fin_del_periodo(wb) -> date:
    fechas = [parse_fecha(wb["Listas"].cell(f, 13).value) for f in (5, 6, 7)]
    return max((f for f in fechas if f), default=date(2026, 12, 31))


def ultima_fila_con_datos(ws) -> int:
    ultima = cfg.PRIMERA_FILA_DATOS - 1
    for fila in range(cfg.PRIMERA_FILA_DATOS, cfg.ULTIMA_FILA_INGRESOS + 1):
        if any(ws.cell(fila, c).value not in (None, "") for c in range(1, cfg.COL_NOTA + 1)):
            ultima = fila
    return ultima


def ids_en_nota(valor) -> set[str]:
    return set(re.findall(re.escape(cfg.MARCA_LODGIFY) + r"(\d+)", str(valor or "")))


def coincide_fila_manual(fila: list, r: Reserva) -> bool:
    """¿Una fila anotada a mano corresponde a la reserva? (misma llegada + nombre)."""
    if parse_fecha(fila[cfg.COL_FECHA - 1][0]) != r.llegada:
        return False
    if fila[cfg.COL_CATEGORIA - 1][0] != cfg.CATEGORIA_HOSPEDAJE:
        return False
    if ids_en_nota(fila[cfg.COL_NOTA - 1][0]):
        return False
    a, b = tokens_nombre(str(fila[cfg.COL_HUESPED - 1][0] or "")), tokens_nombre(r.huesped)
    if not a or not b:
        return False
    if len(set(a) & set(b)) >= 2 or a == b:
        return True
    return a[0] == b[0] and fila[cfg.COL_VILLA - 1][0] == r.villa


def registrar_en_ingresos(wb, reservas: list[Reserva], hoy: date) -> dict:
    """Anota las reservas en 'Ingresos Diarios' en orden de fecha, sin duplicar."""
    ws = wb[cfg.HOJA_INGRESOS]
    fin = fin_del_periodo(wb)
    ultima = ultima_fila_con_datos(ws)
    columnas = range(1, cfg.COL_NOTA + 1)
    filas = [[(ws.cell(f, c).value, copy(ws.cell(f, c)._style)) for c in columnas]
             for f in range(cfg.PRIMERA_FILA_DATOS, ultima + 1)]
    estilo_modelo = [copy(ws.cell(ultima, c)._style) for c in columnas] if filas else None

    resultado = {"anotadas": [], "vinculadas": [], "actualizadas": [], "canceladas": [],
                 "pendientes": [], "fila_de": {}}
    por_id = {}
    for i, fila in enumerate(filas):
        for rid in ids_en_nota(fila[cfg.COL_NOTA - 1][0]):
            por_id[rid] = i

    primera_modificada = None
    for r in sorted(reservas, key=lambda x: (x.llegada, x.villa, x.huesped)):
        if r.id in por_id:
            i = por_id[r.id]
            fila = filas[i]
            nota = str(fila[cfg.COL_NOTA - 1][0] or "")
            automatica = nota.startswith(cfg.MARCA_LODGIFY)
            if r.cancelada and "CANCELADA" not in nota:
                fila[cfg.COL_NOTA - 1] = (f"{nota} · ⚠ CANCELADA en Lodgify – revisar",
                                          fila[cfg.COL_NOTA - 1][1])
                resultado["canceladas"].append(r)
                primera_modificada = min(i, primera_modificada if primera_modificada is not None else i)
            elif automatica and r.confirmada and (
                    abs(a_numero(fila[cfg.COL_MONTO - 1][0]) - r.monto) > 0.009
                    or fila[cfg.COL_NOCHES - 1][0] != r.noches):
                fila[cfg.COL_MONTO - 1] = (r.monto, fila[cfg.COL_MONTO - 1][1])
                fila[cfg.COL_NOCHES - 1] = (r.noches, fila[cfg.COL_NOCHES - 1][1])
                resultado["actualizadas"].append(r)
                primera_modificada = min(i, primera_modificada if primera_modificada is not None else i)
            continue

        if not r.confirmada or r.llegada < cfg.SYNC_DESDE:
            continue
        if r.llegada > fin:
            r.avisos.append("Llegada fuera del periodo del libro: no se anota en Ingresos")
            continue
        if r.llegada > hoy and not cfg.REGISTRAR_FUTURAS:
            resultado["pendientes"].append(r)
            continue

        manual = next((i for i, fila in enumerate(filas) if coincide_fila_manual(fila, r)), None)
        if manual is not None:
            nota = str(filas[manual][cfg.COL_NOTA - 1][0] or "")
            filas[manual][cfg.COL_NOTA - 1] = (
                f"{nota} · Vinculada {cfg.MARCA_LODGIFY}{r.id}".lstrip(" ·"), filas[manual][cfg.COL_NOTA - 1][1])
            por_id[r.id] = manual
            resultado["vinculadas"].append(r)
            primera_modificada = min(manual, primera_modificada if primera_modificada is not None else manual)
            continue

        # Posición: después de la última fila con fecha <= llegada (orden del archivo).
        pos = len(filas)
        for i in range(len(filas) - 1, -1, -1):
            f_fila = parse_fecha(filas[i][0][0])
            if f_fila is not None and f_fila <= r.llegada:
                pos = i + 1
                break
        pos = max(pos, cfg.PRIMERA_FILA_SUMADA - cfg.PRIMERA_FILA_DATOS)
        valores = {
            cfg.COL_FECHA: datetime.combine(r.llegada, datetime.min.time()),
            cfg.COL_VILLA: r.villa, cfg.COL_CATEGORIA: cfg.CATEGORIA_HOSPEDAJE,
            cfg.COL_HUESPED: r.huesped, cfg.COL_CANAL: r.canal, cfg.COL_NOCHES: r.noches,
            cfg.COL_EMAIL: r.email or None, cfg.COL_MONTO: r.monto,
            cfg.COL_METODO: cfg.METODO_COBRO_DEFECTO, cfg.COL_NOTA: f"{cfg.MARCA_LODGIFY}{r.id}",
        }
        nueva = [(valores[c], copy(estilo_modelo[c - 1]) if estilo_modelo else None) for c in columnas]
        filas.insert(pos, nueva)
        por_id = {k: (v + 1 if v >= pos else v) for k, v in por_id.items()}
        por_id[r.id] = pos
        resultado["anotadas"].append(r)
        primera_modificada = min(pos, primera_modificada if primera_modificada is not None else pos)

    if cfg.PRIMERA_FILA_DATOS + len(filas) - 1 > cfg.ULTIMA_FILA_INGRESOS:
        raise RuntimeError("La hoja 'Ingresos Diarios' se quedó sin filas disponibles")

    if primera_modificada is not None:
        for i in range(primera_modificada, len(filas)):
            excel_fila = cfg.PRIMERA_FILA_DATOS + i
            for c, (valor, estilo) in zip(columnas, filas[i]):
                celda = ws.cell(excel_fila, c)
                celda.value = valor
                if estilo is not None:
                    celda._style = estilo
        if resultado["anotadas"] and estilo_modelo:
            formato_fecha = ws.cell(cfg.PRIMERA_FILA_SUMADA, cfg.COL_FECHA).number_format
            for r in resultado["anotadas"]:
                ws.cell(cfg.PRIMERA_FILA_DATOS + por_id[r.id], cfg.COL_FECHA).number_format = formato_fecha

    resultado["fila_de"] = {rid: cfg.PRIMERA_FILA_DATOS + i for rid, i in por_id.items()}
    return resultado


COLUMNAS_RESERVAS = [
    ("Lodgify ID", 12), ("Canal", 16), ("Estado", 11), ("Villa / unidad", 20),
    ("Huésped", 28), ("Email", 34), ("Llegada", 12), ("Salida", 12), ("Noches", 8),
    ("Huéspedes", 10), ("Monto (USD)", 13), ("Detectada", 12),
    ("En Ingresos Diarios", 22), ("Avisos", 50),
]


def actualizar_hoja_reservas(wb, reservas: list[Reserva], filas_ingresos: dict,
                             hoy: date) -> dict:
    """Hoja 'Reservas Lodgify': lista completa; detecta nuevas, canceladas y cambios."""
    previas = {}
    if cfg.HOJA_RESERVAS in wb.sheetnames:
        ws = wb[cfg.HOJA_RESERVAS]
        for fila in ws.iter_rows(min_row=4, values_only=True):
            if fila and fila[0]:
                previas[str(fila[0])] = {
                    "estado": fila[2], "llegada": parse_fecha(fila[6]),
                    "salida": parse_fecha(fila[7]), "monto": fila[10],
                    "detectada": parse_fecha(fila[11]) or hoy}
        wb.remove(ws)
    ws = wb.create_sheet(cfg.HOJA_RESERVAS, index=wb.sheetnames.index(cfg.HOJA_INGRESOS) + 1)

    cambios = {"nuevas": [], "canceladas": [], "modificadas": []}
    actuales = {r.id: r for r in reservas}
    # Las reservas que ya no vienen en la consulta se conservan con su estado anterior.
    for rid in previas.keys() - actuales.keys():
        log.info("Reserva %s no vino en esta consulta; se conserva", rid)

    filas = []
    for r in reservas:
        previa = previas.get(r.id)
        if previa is None:
            if r.confirmada:
                cambios["nuevas"].append(r)
        else:
            if r.cancelada and str(previa["estado"] or "").lower() not in cfg.ESTADOS_CANCELADOS:
                cambios["canceladas"].append(r)
            elif r.confirmada and (previa["llegada"] != r.llegada or previa["salida"] != r.salida
                                   or abs(a_numero(previa["monto"]) - r.monto) > 0.009):
                cambios["modificadas"].append(r)
        if r.id in filas_ingresos:
            en_ingresos = f"Sí – fila {filas_ingresos[r.id]}"
        elif r.cancelada:
            en_ingresos = "No (cancelada)"
        elif r.llegada > hoy:
            en_ingresos = f"Al llegar ({r.llegada:%d/%m})"
        else:
            en_ingresos = "No"
        filas.append([
            int(r.id) if r.id.isdigit() else r.id, r.canal, r.estado, r.villa, r.huesped,
            r.email, r.llegada, r.salida, r.noches, r.huespedes, r.monto,
            previa["detectada"] if previa else hoy, en_ingresos, "; ".join(r.avisos)])

    for rid in sorted(previas.keys() - actuales.keys()):
        p = previas[rid]
        filas.append([int(rid) if rid.isdigit() else rid, None, p["estado"], None, None, None,
                      p["llegada"], p["salida"], None, None, p["monto"], p["detectada"],
                      None, "No vino en la última consulta a Lodgify"])
    filas.sort(key=lambda f: (f[6] or date.max, str(f[3] or "")))

    azul = PatternFill("solid", fgColor="1F4E78")
    ws["A1"] = "Reservas Lodgify (actualizado automáticamente)"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = (f"Última sincronización: {datetime.now(ZoneInfo(cfg.ZONA_HORARIA)):%d/%m/%Y %H:%M} "
                "(hora de Costa Rica). No edite esta hoja: se regenera en cada corrida.")
    ws["A2"].font = Font(italic=True, color="595959")
    for c, (titulo, ancho) in enumerate(COLUMNAS_RESERVAS, 1):
        celda = ws.cell(3, c, titulo)
        celda.font = Font(bold=True, color="FFFFFF")
        celda.fill = azul
        celda.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(c)].width = ancho
    rojo = Font(color="C00000")
    for i, valores in enumerate(filas, 4):
        for c, v in enumerate(valores, 1):
            celda = ws.cell(i, c, v)
            if c in (7, 8, 12):
                celda.number_format = "dd/mm/yyyy"
            elif c == 11:
                celda.number_format = r"\$#,##0.00"
        if str(valores[2] or "").lower() in cfg.ESTADOS_CANCELADOS:
            for c in range(1, len(COLUMNAS_RESERVAS) + 1):
                ws.cell(i, c).font = rojo
    ws.freeze_panes = "A4"
    ws.auto_filter.ref = f"A3:{get_column_letter(len(COLUMNAS_RESERVAS))}{max(len(filas) + 3, 4)}"
    return cambios


# --------------------------------------------------------------------------
# Reporte
# --------------------------------------------------------------------------
def tabla_html(titulo: str, reservas: list[Reserva], vacio: str, hoy: date) -> str:
    h = f'<h3 style="color:#1F4E78;margin:22px 0 6px">{html.escape(titulo)} ({len(reservas)})</h3>'
    if not reservas:
        return h + f'<p style="color:#777;margin:0">{html.escape(vacio)}</p>'
    th = 'style="background:#1F4E78;color:#fff;padding:6px 8px;text-align:left;font-size:13px"'
    td = 'style="padding:6px 8px;border-bottom:1px solid #e3e3e3;font-size:13px"'
    filas = []
    for r in sorted(reservas, key=lambda x: (x.llegada, x.villa)):
        celdas = [r.villa, r.huesped, r.canal, f"{r.llegada:%d/%m}", f"{r.salida:%d/%m}",
                  str(r.noches), str(r.huespedes or ""), f"${r.monto:,.2f}", r.email]
        filas.append("<tr>" + "".join(f"<td {td}>{html.escape(c)}</td>" for c in celdas) + "</tr>")
    cab = ["Villa", "Huésped", "Canal", "Llega", "Sale", "Noches", "Pers.", "Monto", "Email"]
    return (h + '<table style="border-collapse:collapse;width:100%">'
            + "<tr>" + "".join(f"<th {th}>{c}</th>" for c in cab) + "</tr>"
            + "".join(filas) + "</table>")


DIAS = ["lunes", "martes", "miércoles", "jueves", "viernes", "sábado", "domingo"]


def construir_reporte(reservas, ingresos, cambios, reparaciones, hoy, avisos_extra):
    confirmadas = [r for r in reservas if r.confirmada]
    entradas = [r for r in confirmadas if r.llegada == hoy]
    salidas = [r for r in confirmadas if r.salida == hoy]
    ocupadas = {r.villa for r in confirmadas if r.llegada <= hoy < r.salida}
    total_villas = len(cfg.VILLAS)
    manana = [r for r in confirmadas if r.llegada == hoy + timedelta(days=1)]
    mes = [r for r in confirmadas if r.llegada.year == hoy.year and r.llegada.month == hoy.month]

    asunto = (f"Canavida · Entradas del {hoy:%d/%m/%Y}: {len(entradas)} check-in"
              f"{'s' if len(entradas) != 1 else ''} · ${sum(r.monto for r in entradas):,.2f}")
    tarjeta = ('<td style="background:#f2f6fa;border-radius:8px;padding:12px 16px;text-align:center">'
               '<div style="font-size:22px;font-weight:bold;color:#1F4E78">{}</div>'
               '<div style="font-size:12px;color:#555">{}</div></td>')
    resumen = ('<table style="border-spacing:8px;width:100%"><tr>'
               + tarjeta.format(len(entradas), "Entradas hoy")
               + tarjeta.format(len(salidas), "Salidas hoy")
               + tarjeta.format(f"{len(ocupadas)}/{total_villas}", "Villas ocupadas esta noche")
               + tarjeta.format(f"${sum(r.monto for r in mes):,.0f}",
                                f"Reservas con llegada en {hoy:%m/%Y}")
               + "</tr></table>")

    avisos = list(avisos_extra)
    for r in reservas:
        for a in r.avisos:
            avisos.append(f"{r.huesped} ({r.id}): {a}")
    for r in ingresos["canceladas"]:
        avisos.append(f"{r.huesped} ({r.villa}, {r.llegada:%d/%m}) está CANCELADA pero ya figura en "
                      f"Ingresos Diarios (fila {ingresos['fila_de'].get(r.id)}). Revise o elimine esa fila.")

    excel = []
    if ingresos["anotadas"]:
        excel.append(f"{len(ingresos['anotadas'])} reserva(s) anotada(s) en Ingresos Diarios.")
    if ingresos["vinculadas"]:
        excel.append(f"{len(ingresos['vinculadas'])} fila(s) que ya estaban a mano se vincularon "
                     "a su reserva de Lodgify (no se duplicaron).")
    if ingresos["actualizadas"]:
        excel.append(f"{len(ingresos['actualizadas'])} fila(s) actualizadas por cambios en Lodgify.")
    if reparaciones:
        excel.append(f"Se corrigieron {len(reparaciones)} fórmulas con error (#REF!) o rango corto.")
    if not excel:
        excel.append("Sin cambios en Ingresos Diarios.")

    cuerpo = f"""<div style="font-family:Arial,Helvetica,sans-serif;max-width:900px;color:#222">
<h2 style="color:#1F4E78;margin-bottom:2px">Canavida Villas &amp; Resort</h2>
<div style="color:#666;margin-bottom:10px">Reporte diario de reservas · {DIAS[hoy.weekday()]} {hoy:%d/%m/%Y}</div>
{resumen}
{tabla_html("Entradas de hoy (check-in)", entradas, "No hay entradas hoy.", hoy)}
{tabla_html("Salidas de hoy (check-out)", salidas, "No hay salidas hoy.", hoy)}
{tabla_html("Entradas de mañana", manana, "No hay entradas mañana.", hoy)}
{tabla_html("Nuevas reservas desde el último reporte", cambios["nuevas"], "Ninguna.", hoy)}
{tabla_html("Reservas canceladas", cambios["canceladas"], "Ninguna.", hoy)}
{tabla_html("Reservas modificadas (fechas o monto)", cambios["modificadas"], "Ninguna.", hoy)}
<h3 style="color:#1F4E78;margin:22px 0 6px">Excel</h3>
<ul style="margin:0">{"".join(f"<li>{html.escape(x)}</li>" for x in excel)}</ul>
{"<h3 style='color:#C00000;margin:22px 0 6px'>Revisar</h3><ul>" + "".join(f"<li>{html.escape(a)}</li>" for a in avisos) + "</ul>" if avisos else ""}
<p style="color:#999;font-size:11px;margin-top:24px">Generado automáticamente desde Lodgify.
Se adjunta el Excel actualizado.</p></div>"""
    hay_cambios = bool(cambios["nuevas"] or cambios["canceladas"] or cambios["modificadas"]
                       or ingresos["anotadas"] or ingresos["actualizadas"] or avisos)
    return asunto, cuerpo, hay_cambios


def enviar_correo(asunto: str, cuerpo_html: str, adjunto: Path | None = None):
    if not (cfg.GMAIL_USER and cfg.GMAIL_APP_PASSWORD):
        raise RuntimeError("Faltan GMAIL_USER / GMAIL_APP_PASSWORD para enviar el correo")
    msg = EmailMessage()
    msg["Subject"] = asunto
    msg["From"] = f"Canavida Reportes <{cfg.GMAIL_USER}>"
    msg["To"] = cfg.REPORTE_PARA
    msg.set_content(re.sub(r"<[^>]+>", " ", cuerpo_html))
    msg.add_alternative(cuerpo_html, subtype="html")
    if adjunto:
        msg.add_attachment(adjunto.read_bytes(), maintype="application",
                           subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           filename=adjunto.name)
    for intento in range(3):
        try:
            with smtplib.SMTP_SSL("smtp.gmail.com", 465, timeout=60) as s:
                s.login(cfg.GMAIL_USER, cfg.GMAIL_APP_PASSWORD)
                s.send_message(msg)
            log.info("Correo enviado a %s", cfg.REPORTE_PARA)
            return
        except (smtplib.SMTPServerDisconnected, smtplib.SMTPConnectError, OSError):
            if intento == 2:
                raise
            time.sleep(5 * (intento + 1))


# --------------------------------------------------------------------------
# Programa principal
# --------------------------------------------------------------------------
def ejecutar(args) -> int:
    ahora = datetime.now(ZoneInfo(cfg.ZONA_HORARIA))
    hoy = date.fromisoformat(args.hoy) if args.hoy else ahora.date()
    modo = args.modo or ("manana" if ahora.hour < 12 else "tarde")
    excel = Path(args.excel) if args.excel else cfg.EXCEL_PATH
    log.info("Sincronización %s · hoy=%s · libro=%s", modo, hoy, excel)
    for nombre, valor in (("LODGIFY_API_KEY", cfg.LODGIFY_API_KEY), ("GMAIL_USER", cfg.GMAIL_USER),
                          ("GMAIL_APP_PASSWORD", cfg.GMAIL_APP_PASSWORD)):
        log.info("Secret %s: %s", nombre, "configurado" if valor else "FALTA")

    if args.csv:
        reservas = [r for r in leer_csv(Path(args.csv)) if r.llegada and r.llegada >= cfg.SYNC_DESDE]
    else:
        reservas = leer_api(LodgifyAPI(cfg.LODGIFY_API_KEY), cfg.SYNC_DESDE)
    log.info("%d reservas leídas", len(reservas))

    wb = openpyxl.load_workbook(excel)
    reparaciones = reparar_formulas(wb)
    ingresos = registrar_en_ingresos(wb, reservas, hoy)
    cambios = actualizar_hoja_reservas(wb, reservas, ingresos["fila_de"], hoy)
    salida = Path(args.salida) if args.salida else excel
    tmp = salida.with_suffix(".tmp.xlsx")
    wb.save(tmp)
    tmp.replace(salida)  # escritura atómica: el libro nunca queda a medias
    log.info("Excel guardado: %d anotadas, %d vinculadas, %d actualizadas, %d fórmulas corregidas",
             len(ingresos["anotadas"]), len(ingresos["vinculadas"]),
             len(ingresos["actualizadas"]), len(reparaciones))

    asunto, cuerpo, hay_cambios = construir_reporte(reservas, ingresos, cambios, reparaciones, hoy, [])
    if modo == "tarde":
        asunto = asunto.replace("Entradas del", "Actualización de la tarde ·")
    enviar = modo == "manana" or hay_cambios
    if args.reporte_html:
        Path(args.reporte_html).write_text(cuerpo, encoding="utf-8")
    if args.reporte_dir:
        guardar_estado(args.reporte_dir, ok=True, modo=modo, hoy=hoy, asunto=asunto,
                       enviar=enviar, cuerpo=cuerpo)
    if args.sin_correo:
        log.info("Correo omitido (--sin-correo). Asunto: %s", asunto)
    elif enviar:
        enviar_correo(asunto, cuerpo, salida)
    else:
        log.info("Corrida de la tarde sin cambios: no se envía correo")
    return 0


def guardar_estado(carpeta, *, ok, modo, hoy, asunto, enviar, cuerpo="", error=""):
    """Deja el reporte listo para que otro proceso lo envíe (p. ej. por el conector de Gmail)."""
    carpeta = Path(carpeta)
    carpeta.mkdir(parents=True, exist_ok=True)
    (carpeta / "reporte.html").write_text(cuerpo, encoding="utf-8")
    ahora = datetime.now(ZoneInfo(cfg.ZONA_HORARIA))
    (carpeta / "estado.json").write_text(json.dumps({
        "ok": ok, "generado": ahora.isoformat(timespec="seconds"), "hoy": hoy.isoformat(),
        "modo": modo, "asunto": asunto, "enviar": enviar, "error": error,
        "para": cfg.REPORTE_PARA}, ensure_ascii=False, indent=2), encoding="utf-8")


def main(argv=None) -> int:
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--csv", help="Leer reservas de un 'OTC Report' exportado de Lodgify")
    p.add_argument("--excel", help="Ruta del libro (por defecto config.EXCEL_PATH)")
    p.add_argument("--salida", help="Guardar en otra ruta en vez de sobrescribir el libro")
    p.add_argument("--hoy", help="Fecha a usar como 'hoy' (AAAA-MM-DD)")
    p.add_argument("--modo", choices=["manana", "tarde"], help="Por defecto según la hora")
    p.add_argument("--sin-correo", action="store_true", help="No enviar el correo")
    p.add_argument("--reporte-html", help="Guardar también el reporte en este archivo HTML")
    p.add_argument("--reporte-dir", help="Guardar reporte.html y estado.json en esta carpeta")
    args = p.parse_args(argv)
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
    try:
        return ejecutar(args)
    except Exception as e:  # noqa: BLE001 - se informa por correo y se relanza
        log.exception("La sincronización falló")
        if args.reporte_dir:
            hoy = datetime.now(ZoneInfo(cfg.ZONA_HORARIA)).date()
            guardar_estado(args.reporte_dir, ok=False, modo=args.modo or "", hoy=hoy,
                           asunto="⚠ Canavida · Falló la sincronización con Lodgify", enviar=True,
                           cuerpo=f"<p>La actualización automática del Excel falló:</p>"
                                  f"<pre>{html.escape(str(e))}</pre><p>El Excel no se modificó.</p>",
                           error=str(e))
        if not args.sin_correo and cfg.GMAIL_USER and cfg.GMAIL_APP_PASSWORD:
            try:
                enviar_correo("⚠ Canavida · Falló la sincronización con Lodgify",
                              f"<p>La actualización automática del Excel falló:</p>"
                              f"<pre>{html.escape(repr(e))}</pre><p>El Excel no se modificó.</p>")
            except Exception:  # noqa: BLE001
                log.exception("Tampoco se pudo enviar el aviso de error")
        return 1


if __name__ == "__main__":
    sys.exit(main())
