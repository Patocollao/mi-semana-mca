# Publicar Mi Semana MCA v3.1

1. Sube estos archivos al repositorio GitHub.
2. En Streamlit Community Cloud publica `app.py`.
3. La app funciona inmediatamente en modo local de prueba.

## Para sincronizar con Google Sheets
1. Crea una hoja Google Sheets vacía.
2. Copia el ID del documento (la parte entre `/d/` y `/edit` de su URL).
3. Crea una cuenta de servicio en Google Cloud y habilita Google Sheets API y Google Drive API.
4. Comparte la hoja con el correo `client_email` de esa cuenta de servicio como Editor.
5. En Streamlit Cloud > App > Settings > Secrets pega el contenido de `.streamlit/secrets.toml.example`, reemplazando los valores.
6. Reinicia la app.

La app crea automáticamente dos pestañas técnicas en el Google Sheet:
- `_app_state`
- `_pdf_snapshots`

No las borres. Son las que permiten que celular y PC compartan la misma información.

## Excel
En la pestaña “Resumen” puedes descargar un Excel de la semana seleccionada.
