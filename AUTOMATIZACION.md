# Automatización Lodgify → Excel → Correo

Dos veces al día, **8:00 a.m. y 8:00 p.m. (hora de Costa Rica)**:

1. **GitHub Actions** (`.github/workflows/lodgify-sync.yml`) descarga las reservas de Lodgify y actualiza
   `Canavida_Gastos_Ingresos.xlsx`:
   - Hoja **Reservas Lodgify**: todas las reservas desde el 25/10/2026, también las futuras, con su estado.
   - Hoja **Ingresos Diarios**: cada reserva confirmada se anota **el día de su llegada**, en orden de fecha,
     con el formato de las filas escritas a mano (Fecha = llegada, Categoría = Hospedaje, Método = Tarjeta,
     Monto = total de Lodgify). La columna J lleva `Lodgify #<id>` para no duplicar nunca; las filas que ya
     estaban a mano se marcan `Vinculada Lodgify #<id>`.
   - Cancelaciones: no se borran filas; se marcan `⚠ CANCELADA en Lodgify – revisar` y salen en el reporte.
   - Guarda el Excel, `reportes/reporte.html` y `reportes/estado.json` en el repositorio.
2. **Claude** (tarea programada a las 8:25 a.m. y 8:25 p.m., con el conector de Gmail) lee
   `reportes/estado.json` y envía el reporte desde **admicanavida@gmail.com** a esa misma dirección, con el
   enlace al Excel. En la tarde solo envía si hubo novedades. Si la sincronización falló o no corrió,
   envía un aviso.

## Configuración

Único secreto necesario, en GitHub → *Settings → Secrets and variables → Actions → Repository secrets*:

| Secret | Valor |
|---|---|
| `LODGIFY_API_KEY` | Lodgify → Configuración → **API pública** (la API no acepta usuario y contraseña). |

Para que GitHub envíe el correo con el Excel adjunto, agregue también:

| Secret | Valor |
|---|---|
| `GMAIL_USER` | `admicanavida@gmail.com` |
| `GMAIL_APP_PASSWORD` | Contraseña de aplicación de Google (myaccount.google.com/apppasswords), 16 letras sin espacios. |

Si estos dos no existen, el workflow solo deja el reporte en `reportes/` y no envía correo.

Ejecución manual: pestaña **Actions → Sincronizar reservas Lodgify → Run workflow**.

## Uso local

```bash
pip install -r automatizacion/requirements.txt
python automatizacion/lodgify_sync.py --csv "OTCReport.csv" --sin-correo   # desde un reporte exportado
python -m pytest automatizacion/tests                                       # pruebas
```

Ajustes en `automatizacion/config.py`: villas, canales, fecha de inicio (`SYNC_DESDE`) y
`REGISTRAR_FUTURAS` (true = anotar en Ingresos apenas se reserva, en vez de al llegar).

## Errores del libro corregidos

- `Utilidad y Posibles Gastos` F13:F28: 4 celdas en `#REF!` y el resto tomaba el gasto de otra
  categoría. Ahora busca por nombre de categoría.
- `Arqueo y Resumen Caja` B55:D55: `#REF!` (columna eliminada). Ahora cuenta movimientos con "No" en Notas.
- `Ingresos Diarios` H4 sumaba solo hasta la fila 200; resúmenes hasta la 1026 → ampliados a la 2026.
- `Gastos e Ingresos` B5:B9 (octubre) eran montos pegados como valor: se conservan y se les suman
  las reservas que agrega la automatización.

## Pendiente de revisar

- Expedia: Lodgify no siempre desglosa el impuesto; el reporte avisa cuando el monto puede incluirlo.
- "Muebleria" no aparece en la tabla de presupuesto de `Utilidad y Posibles Gastos` (total $321 menor).
