"""Configuración de la sincronización Lodgify → Excel de Canavida.

Todo lo que se pueda querer cambiar sin tocar la lógica está aquí.
Las credenciales NUNCA van en este archivo: se leen de variables de entorno
(en GitHub Actions, de los "Secrets" del repositorio).
"""
import os
from datetime import date
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent

# Libro de Excel que se actualiza (relativo a la raíz del repositorio).
EXCEL_PATH = RAIZ / os.environ.get("EXCEL_PATH", "Canavida_Gastos_Ingresos.xlsx")

ZONA_HORARIA = "America/Costa_Rica"

# Destinatario del reporte diario.
REPORTE_PARA = os.environ.get("REPORTE_PARA", "admicanavida@gmail.com")

# Credenciales (variables de entorno / GitHub Secrets).
LODGIFY_API_KEY = os.environ.get("LODGIFY_API_KEY", "")
LODGIFY_API_URL = os.environ.get("LODGIFY_API_URL", "https://api.lodgify.com")
GMAIL_USER = os.environ.get("GMAIL_USER", "")
GMAIL_APP_PASSWORD = os.environ.get("GMAIL_APP_PASSWORD", "")

# Primer día de llegada que se sincroniza con "Ingresos Diarios". Antes de esta
# fecha los ingresos se anotaron a mano y no se tocan (evita duplicados).
SYNC_DESDE = date.fromisoformat(os.environ.get("SYNC_DESDE", "2026-10-25"))

# False (recomendado): una reserva entra en "Ingresos Diarios" el día de su
# llegada, así los meses futuros no muestran ingresos que aún no ocurren. Todas
# las reservas (también las futuras) quedan en la hoja "Reservas Lodgify".
# True: se anotan en "Ingresos Diarios" apenas se detectan.
REGISTRAR_FUTURAS = os.environ.get("REGISTRAR_FUTURAS", "false").lower() == "true"

# Hojas y estructura del libro.
HOJA_INGRESOS = "Ingresos Diarios"
HOJA_RESERVAS = "Reservas Lodgify"
FILA_ENCABEZADO_INGRESOS = 5
PRIMERA_FILA_DATOS = 6
# Las fórmulas del resumen suman desde esta fila; nunca se inserta antes.
PRIMERA_FILA_SUMADA = 27
ULTIMA_FILA_INGRESOS = 2026
MARCA_LODGIFY = "Lodgify #"

# Columnas de "Ingresos Diarios" (1 = A).
COL_FECHA, COL_VILLA, COL_CATEGORIA, COL_HUESPED, COL_CANAL = 1, 2, 3, 4, 5
COL_NOCHES, COL_EMAIL, COL_MONTO, COL_METODO, COL_NOTA = 6, 7, 8, 9, 10

CATEGORIA_HOSPEDAJE = "Hospedaje"
METODO_COBRO_DEFECTO = "Tarjeta"

# Nombre de la villa en Lodgify (palabra clave) → nombre en la hoja "Listas".
VILLAS = [
    ("LOVE NEST", "Love Nest"),
    ("LORO", "Loro (4 hab)"),
    ("CARACARA", "Caracara (2 hab)"),
    ("OROPENDOLA", "Oropendola (2 hab)"),
    ("TOUCAN", "Toucan (3 hab)"),
    ("TOROGOZ", "Torogoz (4 hab)"),
    ("QUETZAL", "Quetzal (4 hab)"),
    ("LAPA", "Lapa (3 hab)"),
    ("CUSINGA", "Cusinga (1 hab)"),
    ("GECKO", "Gecko (Condo 3)"),
    ("SALAMANDRA", "Salamandra (Condo 4)"),
    ("RANA", "Rana (Condo 6)"),
    ("IGUANA", "Iguana (Condo 2)"),
    ("LIZARDO", "Lizardo (Condo 5)"),
    ("COLIBRI", "Colibri"),
    ("TORTUGA", "Tortuga"),
    ("PAVON", "Pavon"),
]

# Canal en Lodgify (texto en minúsculas que contiene) → canal en "Listas".
CANALES = [
    ("airbnb", "Airbnb"),
    ("booking", "Booking.com"),
    ("expedia", "Expedia"),
    ("website", "Web (Lodgify)"),
    ("oh", "Web (Lodgify)"),
    ("canavidaresort", "Web (Lodgify)"),
    ("manual", "Directo / WhatsApp"),
]

# Expedia cobra y entrega el impuesto por su cuenta: en el libro se anota el
# monto sin impuesto (así se hizo a mano, p. ej. Maria Vitória Santos 25/10).
CANALES_SIN_IMPUESTO = {"Expedia"}

ESTADOS_CONFIRMADOS = {"booked"}
ESTADOS_CANCELADOS = {"declined", "cancelled", "canceled", "deleted"}
