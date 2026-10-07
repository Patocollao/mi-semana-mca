# Mi Semana MCA v3 — Streamlit

Aplicación personal para planificar visitas MCA desde PDFs MyWork. No se conecta a MCA: el PDF se carga manualmente.

## Qué hace
- Lee PDFs MyWork/MCA.
- Conserva visitas duplicadas al mismo local.
- Compara cada PDF nuevo con el último guardado.
- Alerta visitas nuevas, eliminadas y cambios de horas/estado.
- Conserva checks, día planificado y notas cuando actualizas el PDF.
- Navega por semanas.
- Registra km y gastos por día.
- Registra carga del Tesla una vez por semana.
- Exporta Excel.
- Con Supabase, celular y PC comparten la misma información.

## 1. Probar en PC
```bash
pip install -r requirements.txt
streamlit run app.py
```

## 2. Crear Supabase
1. Crea un proyecto en Supabase.
2. SQL Editor > pega y ejecuta `supabase.sql`.
3. Project Settings > API: copia `Project URL` y **service_role key**.
4. La service_role key es secreta: NO la pongas en GitHub ni en el navegador.

## 3. Publicar en GitHub + Streamlit Community Cloud
1. Crea un repositorio privado o público en GitHub y sube estos archivos.
2. En Streamlit Community Cloud elige `app.py` como archivo principal.
3. En **Advanced settings > Secrets** agrega:
```toml
SUPABASE_URL = "https://TU-PROYECTO.supabase.co"
SUPABASE_SERVICE_ROLE_KEY = "TU_SERVICE_ROLE_KEY"
```
4. Deploy.
5. Abre la URL `https://...streamlit.app` desde celular y PC.

## Seguridad
La app no usa autenticación porque está pensada para un único usuario. La base se accede con la `service_role` únicamente desde el servidor de Streamlit. Las tablas tienen RLS activado y no tienen políticas para acceso público.

Si haces público el repositorio, nunca subas `.streamlit/secrets.toml` real. Solo usa `secrets.toml.example`.
