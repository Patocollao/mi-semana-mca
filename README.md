# Mi Semana MCA v3.1 — Streamlit

Aplicación personal para planificar visitas MCA desde PDFs MyWork.

## Cambios v3.1
- Semana organizada por pestañas: Lunes, Martes, Miércoles, Jueves y Viernes.
- Las visitas sin día aparecen en una pestaña adicional “Sin asignar”.
- Google Sheets puede usarse como almacenamiento central entre celular y PC.
- Descarga Excel por semana.
- Conserva comparación de PDFs, alertas de visitas nuevas, checks, km y gastos.

## Almacenamiento
Prioridad automática:
1. Google Sheets, si está configurado.
2. Supabase, si está configurado.
3. Archivo local, solo para pruebas.

No se conecta a MCA. El PDF se carga manualmente.
