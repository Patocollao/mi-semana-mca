# Publicación rápida

## A. GitHub
1. Crea un repositorio llamado `mi-semana-mca`.
2. Sube el contenido de esta carpeta, no el ZIP dentro del repositorio.
3. No subas un archivo real `.streamlit/secrets.toml`.

## B. Supabase
1. Crea un proyecto.
2. Ejecuta `supabase.sql` en **SQL Editor**.
3. Copia `Project URL` y `service_role key` desde Project Settings > API.

## C. Streamlit Community Cloud
1. New app > selecciona el repositorio `mi-semana-mca`.
2. Main file path: `app.py`.
3. Advanced settings > Secrets:

```toml
SUPABASE_URL = "https://TU-PROYECTO.supabase.co"
SUPABASE_SERVICE_ROLE_KEY = "TU_SERVICE_ROLE_KEY"
```

4. Deploy.
5. Abre la URL en celular y PC. Ambos verán los mismos datos.

## Flujo de uso
1. Descarga manualmente el PDF desde MCA.
2. Cárgalo en **Actualizar desde MyWork**.
3. Revisa la alerta de cambios.
4. Presiona **Aplicar este PDF y guardar cambios**.
5. La app conserva planificación, checks, km y gastos anteriores.
