# Extractor de datos - Formato Mango

Aplicación en Python para usuario final:

- Sube el archivo de control (`.xlsx`).
- Procesa una hoja específica o todas las hojas con datos.
- Extrae y previsualiza los datos en filas/columnas.
- Descarga un Excel con estructura detectada desde el formato subido.

## Requisitos

```bash
pip install -r requirements.txt
```

## Ejecutar local

```bash
streamlit run app.py
```

## Notas de uso

- La extracción es 100% desde `format_profiles.json` + detección de perfil.
- La opción "Excluir muestras sin datos reales" evita exportar `N°` vacíos.
- La marca de `CONVENCIONAL/ORGANICO` se define desde la interfaz.
- Se convierten valores vacíos lógicos (`None`, `-`, `NA`) a celdas vacías en salida.
- Los encabezados de salida se infieren desde el formato subido.
- La configuración en `.streamlit/config.toml` reduce logs técnicos y oculta trazas en la interfaz.

## Configurar otro formato

Los formatos soportados se configuran en `format_profiles.json` (sin tocar código).

1. Duplica el perfil existente dentro de `profiles`.
2. Cambia `id`, `name`, `match_tokens`, `metadata_cells`, `header_cells` y `row_rules`.
3. (Opcional) cambia `default_profile`.
4. Reinicia la app.

La app detecta el perfil automáticamente y permite elegirlo manualmente.

## Publicar en Streamlit Community Cloud

1. Sube este proyecto a un repositorio GitHub.
2. Verifica que en el repo estén: `app.py`, `extractor.py`, `format_profiles.json`, `requirements.txt`.
3. En Streamlit Cloud: **New app** -> conecta repo -> branch -> `app.py`.
4. Deploy.

Archivo útil incluido:
- `runtime.txt` para fijar Python `3.12`.

## Publicar en servidor propio con Docker

### Build

```bash
docker build -t extractor-mango .
```

### Run

```bash
docker run -p 8501:8501 extractor-mango
```

Abre: `http://localhost:8501`

Archivos de despliegue incluidos:
- `Dockerfile`
- `.dockerignore`
- `Procfile` (plataformas tipo Render/Heroku-like)

## Seguridad recomendada antes de publicar

1. No subas archivos reales con datos sensibles al repositorio.
2. Revisa `.gitignore` para excluir archivos de prueba locales.
3. Si será pública, agrega autenticación o publica solo en red interna.
