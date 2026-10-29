# Automatización Lodgify → Excel → Correo

Dos veces al día (**8:00 a.m. y 8:00 p.m., hora de Costa Rica**) GitHub Actions:

1. Descarga las reservas de Lodgify.
2. Actualiza `Canavida_Gastos_Ingresos.xlsx`:
   - Hoja **Reservas Lodgify** (nueva): todas las reservas desde el 25/10/2026, también las futuras,
     con estado, villa, huésped, llegada/salida, monto, y si ya está en Ingresos Diarios.
   - Hoja **Ingresos Diarios**: anota cada reserva confirmada **el día de su llegada**, en orden de fecha,
     con el mismo formato que las filas escritas a mano (Fecha = llegada, Categoría = Hospedaje,
     Método = Tarjeta, Monto = total de Lodgify; en Expedia sin impuesto). La columna J lleva
     `Lodgify #<id>` para no duplicar nunca. Si una reserva ya estaba anotada a mano, se vincula
     (`Vinculada Lodgify #<id>`) en vez de duplicarla.
   - Cancelaciones: no se borran filas; se marcan `⚠ CANCELADA en Lodgify – revisar` y salen en el correo.
3. Guarda el Excel en este repositorio (historial completo de versiones) y lo envía adjunto.
4. Correo a **admicanavida@gmail.com**:
   - 8:00 a.m.: siempre. Entradas y salidas del día, entradas de mañana, villas ocupadas,
     nuevas reservas, cancelaciones, cambios y avisos.
   - 8:00 p.m.: solo si hubo novedades.
   - Si algo falla, llega un correo "⚠ Falló la sincronización" y el Excel no se toca.

## Configuración (una sola vez)

En GitHub → *Settings → Secrets and variables → Actions → New repository secret*:

| Secret | Valor |
|---|---|
| `LODGIFY_API_KEY` | Lodgify → Configuración → **API pública** → copiar la clave. (La API de Lodgify no acepta usuario y contraseña, solo esta clave.) |
| `GMAIL_USER` | Cuenta Gmail que envía, p. ej. `admicanavida@gmail.com` |
| `GMAIL_APP_PASSWORD` | Contraseña de aplicación de Google (myaccount.google.com → Seguridad → Verificación en 2 pasos → Contraseñas de aplicaciones). No es la contraseña normal. |

Prueba manual: pestaña **Actions → Sincronizar reservas Lodgify → Run workflow**.

## Uso local

```bash
pip install -r automatizacion/requirements.txt
python automatizacion/lodgify_sync.py --csv "OTCReport.csv" --sin-correo   # desde un reporte exportado
python -m pytest automatizacion/tests                                       # pruebas
```

Ajustes en `automatizacion/config.py`: villas, canales, fecha de inicio (`SYNC_DESDE`) y
`REGISTRAR_FUTURAS` (true = anotar en Ingresos apenas se reserva, en vez de al llegar).

## Errores del libro corregidos automáticamente

- `Utilidad y Posibles Gastos` F13:F28: 4 celdas en `#REF!` y el resto tomaba el gasto de otra
  categoría (orden distinto). Ahora busca por nombre de categoría.
- `Arqueo y Resumen Caja` B55:D55: `#REF!` (columna eliminada). Ahora cuenta movimientos con "No" en Notas.
- `Ingresos Diarios` H4 sumaba solo hasta la fila 200; resúmenes hasta la 1026 → ampliados a la 2026.
- `Gastos e Ingresos` B5:B9 (octubre) eran montos pegados como valor: se conservan y se les suman
  las reservas que agrega la automatización.
