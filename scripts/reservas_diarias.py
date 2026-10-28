"""Reservas nuevas de Lodgify -> hoja Ingresos Diarios del libro. Uso: LODGIFY_API_KEY=... python3 reservas_diarias.py [YYYY-MM-DD]
Revisa las creadas desde 2 días antes de la fecha dada (por defecto hoy, hora Costa Rica); no duplica filas."""
import os, sys, json, urllib.request, urllib.parse
from datetime import date, datetime, timedelta, timezone
import openpyxl

KEY = os.environ["LODGIFY_API_KEY"]
BASE = "https://api.lodgify.com/v2"
LIBRO = os.path.join(os.path.dirname(__file__), "..", "Canavida Gastos Ingresos Oficial (1).xlsx")
UNIDAD = {417302: "Caracara (2 hab)", 417829: "Lapa (3 hab)", 418298: "Loro (4 hab)", 418311: "Love Nest",
          418312: "Quetzal (4 hab)", 418313: "Torogoz (4 hab)", 537300: "Toucan (3 hab)", 537321: "Oropendola (2 hab)",
          537325: "Cusinga (1 hab)", 555489: "Gecko (Condo 3)", 556474: "Lizardo (Condo 5)",
          556475: "Salamandra (Condo 4)", 556478: "Rana (Condo 6)", 598738: "Iguana (Condo 2)",
          733216: "Pavon", 740510: "Tortuga"}
CANAL = {"Directa": "Directo / WhatsApp"}
PLAT = {"AirbnbIntegration": "Airbnb", "BookingCom": "Booking.com", "Manual": "Directo / WhatsApp", "OH": "Web (Lodgify)", "Vrbo": "Vrbo"}

def get(path, **q):
    req = urllib.request.Request(f"{BASE}{path}?{urllib.parse.urlencode(q)}",
                                 headers={"X-ApiKey": KEY, "accept": "application/json"})
    return json.load(urllib.request.urlopen(req, timeout=60))

def actualizar_libro(libro):
    """Agrega las reservas a 'Ingresos Diarios' del libro (sin duplicar fecha + huésped)."""
    from copy import copy
    wb = openpyxl.load_workbook(LIBRO); ws = wb["Ingresos Diarios"]
    ultima = max(r for r in range(6, 2027) if ws.cell(r, 1).value or ws.cell(r, 8).value)
    from collections import Counter
    vistos = Counter((ws.cell(r, 1).value.date() if hasattr(ws.cell(r, 1).value, "date") else None,
                      (ws.cell(r, 4).value or "").strip().lower()) for r in range(6, ultima + 1))
    nuevas = []
    for x in sorted(libro):
        k = (x[0], x[2].strip().lower())
        if vistos[k] > 0:
            vistos[k] -= 1
        else:
            nuevas.append(x)
    for i, (f, unidad, nombre, canal, noches, monto) in enumerate(nuevas, ultima + 1):
        for c in range(1, 11):
            src, dst = ws.cell(ultima, c), ws.cell(i, c)
            dst._style = copy(src._style)
        ws.cell(i, 1, datetime.combine(f, datetime.min.time())); ws.cell(i, 2, unidad)
        ws.cell(i, 3, "Hospedaje"); ws.cell(i, 4, nombre); ws.cell(i, 5, CANAL.get(canal, canal))
        ws.cell(i, 6, noches); ws.cell(i, 8, monto); ws.cell(i, 9, "Tarjeta")
    wb.save(LIBRO)
    return len(nuevas)

def main():
    hoy = date.fromisoformat(sys.argv[1]) if len(sys.argv) > 1 else (datetime.now(timezone.utc) - timedelta(hours=6)).date()
    props = {p["id"]: p["name"] for p in get("/properties", size=100)["items"]}
    rows, libro, page = [], [], 1
    while True:
        d = get("/reservations/bookings", size=50, page=page, includeCount="false",
                updatedSince=(hoy - timedelta(days=3)).isoformat())
        items = d["items"]
        for b in items:
            if b["created_at"][:10] < (hoy - timedelta(days=2)).isoformat() or b["status"] in ("Declined", "Open") or b.get("is_deleted"):
                continue
            nombre = props.get(b["property_id"], str(b["property_id"]))
            tipo = "Condo" if "condo" in nombre.lower() else "Villa"
            n = (date.fromisoformat(b["departure"]) - date.fromisoformat(b["arrival"])).days
            total = b["total_amount"] or round(sum(v for v in b["subtotals"].values() if v), 2)
            src = b["source"]
            libro.append((date.fromisoformat(b["arrival"]), UNIDAD.get(b["property_id"], nombre), b["guest"]["name"],
                          PLAT.get(src, src), n, total))
            rows.append([b["arrival"], f"{tipo} - {nombre.split('|')[0].strip()}", b["guest"]["name"],
                         PLAT.get(src, src), n, total, b["currency_code"]])
        if len(items) < 50:
            break
        page += 1
    agregadas = actualizar_libro(libro)
    print(agregadas, "filas nuevas en el libro")

main()
