from __future__ import annotations

import pandas as pd
import streamlit as st

from extractor import (
    detect_profile_id,
    extract_sheets,
    get_profile_field_order,
    infer_sheet_headers,
    list_data_sheet_names,
    list_format_profiles,
    list_sheet_names,
    normalize_records_for_output,
    render_output_excel,
    rows_to_display_records,
)


def main() -> None:
    st.set_page_config(page_title="Extractor Formato Mango", layout="wide")

    st.title("Extractor de datos - Control Producto Terminado Mango")
    st.write(
        "Sube el formato de control, procesa una hoja o todas las hojas con datos, "
        "previsualiza el resultado y descarga el Excel con estructura detectada desde el formato."
    )

    source_file = st.file_uploader(
        "1) Sube el archivo de control (.xlsx)",
        type=["xlsx"],
        accept_multiple_files=False,
    )

    if source_file is None:
        st.info("Esperando archivo de control para iniciar.")
        return

    try:
        with st.spinner("Analizando archivo cargado..."):
            source_bytes = source_file.getvalue()
            sheet_names = list_sheet_names(source_bytes)
            profiles = list_format_profiles()
            auto_profile_id = detect_profile_id(source_bytes)
    except Exception:  # pragma: no cover
        st.error("No se pudo leer el archivo. Verifica que sea un .xlsx válido.")
        return

    profile_ids = [profile["id"] for profile in profiles]
    profile_labels = {profile["id"]: f'{profile["name"]} ({profile["id"]})' for profile in profiles}
    if auto_profile_id in profile_ids:
        default_profile_index = profile_ids.index(auto_profile_id)
    else:
        default_profile_index = 0

    selected_profile_id = st.selectbox(
        "2) Plantilla de formato",
        options=profile_ids,
        index=default_profile_index,
        format_func=lambda pid: profile_labels[pid],
        help="Se detecta automáticamente, pero puedes cambiarla manualmente.",
    )

    try:
        with st.spinner("Validando hojas del formato seleccionado..."):
            data_sheet_names = list_data_sheet_names(source_bytes, profile_id=selected_profile_id)
    except Exception:
        st.error("No se pudo validar el formato con la plantilla elegida.")
        return

    mode = st.radio(
        "3) Modo de extracción",
        options=["Todas las hojas con datos", "Solo una hoja"],
        index=0,
        horizontal=True,
    )

    if mode == "Todas las hojas con datos":
        default_sheets = data_sheet_names or sheet_names
        selected_sheets = st.multiselect(
            "4) Hojas a procesar",
            options=sheet_names,
            default=default_sheets,
        )
    else:
        selected_sheet = st.selectbox("4) Selecciona la hoja", sheet_names)
        selected_sheets = [selected_sheet]

    tipo_producto = st.radio(
        "5) Marca para columnas CONVENCIONAL / ORGANICO",
        options=["No especificar", "Orgánico", "Convencional"],
        index=0,
        horizontal=True,
    )

    remove_empty = st.checkbox(
        "Excluir muestras sin datos reales (recomendado)",
        value=True,
        help="Mantiene todos los valores de las muestras que sí tienen datos.",
    )

    if st.button("Extraer y previsualizar", type="primary"):
        if not selected_sheets:
            st.warning("Selecciona al menos una hoja para procesar.")
            return

        progress = st.progress(0, text="Iniciando proceso...")
        try:
            with st.spinner("Procesando y extrayendo datos..."):
                progress.progress(15, text="Leyendo estructura de la hoja...")
                header_sheet = selected_sheets[0]
                output_headers = infer_sheet_headers(source_bytes, header_sheet, profile_id=selected_profile_id)
                field_order = get_profile_field_order(selected_profile_id)

                progress.progress(45, text="Extrayendo datos de muestras...")
                records = extract_sheets(
                    file_bytes=source_bytes,
                    sheet_names=selected_sheets,
                    tipo_producto=tipo_producto,
                    drop_empty_samples=remove_empty,
                    profile_id=selected_profile_id,
                )
                output_records = normalize_records_for_output(records)

                progress.progress(70, text="Construyendo previsualización...")
                display_records = rows_to_display_records(output_records, output_headers, field_order)
                df = pd.DataFrame(display_records, columns=output_headers)
                preview_df = df.where(pd.notna(df), "").astype(str)

                progress.progress(85, text="Generando archivo Excel de salida...")
                output_bytes = render_output_excel(
                    output_records,
                    output_headers=output_headers,
                    field_order=field_order,
                )

                progress.progress(100, text="Proceso completado.")
        except Exception:
            progress.empty()
            st.error("No se pudieron extraer los datos. Revisa la hoja seleccionada y el formato del archivo.")
            return

        if not records:
            progress.empty()
            st.warning("No se encontraron registros con datos en las hojas seleccionadas.")
            return

        progress.empty()
        st.success(f"Registros extraídos: {len(records)}")

        st.subheader("Previsualización")
        st.dataframe(preview_df, width="stretch", hide_index=True)

        if mode == "Solo una hoja":
            base_name = selected_sheets[0].replace(" ", "_")
            default_name = f"DATA_PROCESO_{base_name}.xlsx"
        else:
            default_name = "DATA_PROCESO_TODAS_LAS_HOJAS.xlsx"

        st.download_button(
            label="Descargar Excel generado",
            data=output_bytes,
            file_name=default_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


if __name__ == "__main__":
    try:
        main()
    except Exception:
        st.error("Ocurrió un error inesperado. Reinicia la aplicación y vuelve a intentar.")
