import io
import re
from typing import Dict, List, Set, Tuple

import pandas as pd
import pdfplumber
import streamlit as st

st.set_page_config(
    page_title="Control PDM Tambo | Dashboard & Reportes",
    page_icon="🛒",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
    .metric-card {
        background-color: #f8f9fa;
        border-radius: 10px;
        padding: 15px;
        border-left: 5px solid #6f42c1;
    }
    .stDownloadButton button {
        width: 100%;
        background-color: #28a745;
        color: white;
        font-weight: bold;
        border-radius: 8px;
        height: 3em;
    }
</style>
""", unsafe_allow_html=True)

DEFAULT_CODIGOS_PDM: List[str] = [
    "1006666", "1011895", "1005645", "1001529", "1007147",
    "1010150", "1004597", "1007474", "400725017", "1000972",
    "1005799", "1000669", "1016129", "1011235", "1006683",
    "1000973", "400220055", "1016350", "1011832", "1007039",
    "1016105"
]

with st.sidebar:
    st.header("⚙️ Parámetros Operativos")
    
    st.subheader("Códigos PDM Activos")
    pdm_input = st.text_area(
        "SKUs del PDM (uno por línea o separados por coma):",
        value="\n".join(DEFAULT_CODIGOS_PDM),
        height=150,
        help="Permite actualizar los productos del mes sin editar el código fuente."
    )
    codigos_pdm_activos: Set[str] = {
        c.strip().replace("(", "").replace(")", "") 
        for c in re.split(r'[\n,]+', pdm_input) if c.strip()
    }

def parse_pdf_report(
    uploaded_file, 
    codigos_pdm: Set[str]
) -> Tuple[pd.DataFrame, pd.DataFrame, str, Dict[str, int]]:
    
    transacciones_lista = []
    current_trx = None
    in_table = False
    fecha_reporte = "Sin Fecha"
    conteo_skus_pdm: Dict[str, int] = {sku: 0 for sku in codigos_pdm}

    with pdfplumber.open(uploaded_file) as pdf:
        total_paginas = len(pdf.pages)
        progress_bar = st.progress(0, text="Iniciando lectura de páginas...")
        
        for idx, page in enumerate(pdf.pages):
            progress_bar.progress((idx + 1) / total_paginas, text=f"Procesando página {idx + 1} de {total_paginas}...")
            texto = page.extract_text(layout=True)
            if not texto:
                continue

            if fecha_reporte == "Sin Fecha":
                match_fecha = re.search(r'\b(\d{2}/\d{2}/\d{4})\b', texto)
                if match_fecha:
                    fecha_reporte = match_fecha.group(1)

            lineas = texto.split("\n")
            for linea in lineas:
                linea_lower = linea.lower()

                if "trns" in linea_lower and ("art" in linea_lower or "desc" in linea_lower):
                    in_table = True
                    continue

                if not in_table:
                    continue

                match_trx = re.match(r'^ {0,4}(\d{1,6})(?:\s+|$)', linea)
                
                if match_trx:
                    if current_trx is not None:
                        transacciones_lista.append(current_trx)

                    current_trx = {
                        "id": match_trx.group(1),
                        "texto_lineas": [linea],
                        "vendedor": "Desconocido"
                    }

                    match_vendedor = re.search(r'\b(t\d{8})\b', linea_lower)
                    if match_vendedor:
                        current_trx["vendedor"] = match_vendedor.group(1).upper()

                elif current_trx is not None:
                    current_trx["texto_lineas"].append(linea)
                    if current_trx["vendedor"] == "Desconocido":
                        match_vendedor = re.search(r'\b(t\d{8})\b', linea_lower)
                        if match_vendedor:
                            current_trx["vendedor"] = match_vendedor.group(1).upper()

        if current_trx is not None:
            transacciones_lista.append(current_trx)
            
        progress_bar.empty()

    if not transacciones_lista:
        return pd.DataFrame(), pd.DataFrame(), fecha_reporte, {}

    datos_boletas = []
    for data in transacciones_lista:
        texto_completo = " ".join(data["texto_lineas"])
        codigos_encontrados = set(re.findall(r'\b\d{7,9}\b', texto_completo))
        
        pdms_en_boleta = codigos_encontrados.intersection(codigos_pdm)
        tiene_pdm = 1 if len(pdms_en_boleta) > 0 else 0
        
        for sku in pdms_en_boleta:
            conteo_skus_pdm[sku] = conteo_skus_pdm.get(sku, 0) + 1

        cod_vendedor = data["vendedor"]

        datos_boletas.append({
            "Trns_ID": data["id"],
            "Vendedor": cod_vendedor,
            "Contiene_PDM": tiene_pdm,
            "SKUs_Detectados": ", ".join(pdms_en_boleta) if pdms_en_boleta else "Ninguno"
        })

    df_boletas = pd.DataFrame(datos_boletas)

    df_vendedores = df_boletas.groupby("Vendedor").agg(
        Total_Transacciones=("Contiene_PDM", "count"),
        PDM=("Contiene_PDM", "sum")
    ).reset_index()

    df_vendedores["Porcentaje"] = (
        df_vendedores["PDM"] / df_vendedores["Total_Transacciones"]
    ).fillna(0)

    df_vendedores = df_vendedores.sort_values(by="Porcentaje", ascending=False)

    return df_boletas, df_vendedores, fecha_reporte, conteo_skus_pdm

def create_excel_report(df_vendedores: pd.DataFrame, fecha_reporte: str) -> bytes:
    output = io.BytesIO()
    
    total_trx = df_vendedores["Total_Transacciones"].sum()
    total_pdm = df_vendedores["PDM"].sum()
    porcentaje_total = (total_pdm / total_trx) if total_trx > 0 else 0.0

    filas_excel = [{
        "Fecha": fecha_reporte,
        "Total Transacciones": total_trx,
        "PDM": total_pdm,
        "Porcentaje": porcentaje_total,
        "Vendedor Responsable": "TOTAL CAJA"
    }]

    for _, row in df_vendedores.iterrows():
        filas_excel.append({
            "Fecha": fecha_reporte,
            "Total Transacciones": int(row["Total_Transacciones"]),
            "PDM": int(row["PDM"]),
            "Porcentaje": float(row["Porcentaje"]),
            "Vendedor Responsable": str(row["Vendedor"])
        })

    df_export = pd.DataFrame(filas_excel)

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df_export.to_excel(writer, sheet_name="Reporte PDM", index=False, header=False, startrow=1)
        workbook = writer.book
        worksheet = writer.sheets["Reporte PDM"]

        f_header = workbook.add_format({
            "bold": True, "align": "center", "valign": "vcenter",
            "bg_color": "#1A2530", "font_color": "#FFFFFF", "border": 1, "font_size": 11
        })
        f_total_row = workbook.add_format({
            "bold": True, "align": "center", "bg_color": "#E9ECEF", "border": 1
        })
        f_total_pct = workbook.add_format({
            "bold": True, "align": "center", "bg_color": "#E9ECEF", "border": 1, "num_format": "0.0%"
        })
        f_data_center = workbook.add_format({"align": "center", "border": 1})
        f_data_num = workbook.add_format({"align": "center", "border": 1, "num_format": "#,##0"})
        f_data_pct = workbook.add_format({"align": "center", "border": 1, "num_format": "0.0%"})

        headers = ["Fecha", "Total Transacciones", "PDM", "Porcentaje", "Vendedor Responsable"]
        for col_idx, col_name in enumerate(headers):
            worksheet.write(0, col_idx, col_name, f_header)

        worksheet.write(1, 0, df_export.iloc[0, 0], f_total_row)
        worksheet.write(1, 1, df_export.iloc[0, 1], f_total_row)
        worksheet.write(1, 2, df_export.iloc[0, 2], f_total_row)
        worksheet.write(1, 3, df_export.iloc[0, 3], f_total_pct)
        worksheet.write(1, 4, df_export.iloc[0, 4], f_total_row)

        for r_idx in range(1, len(df_export)):
            worksheet.write(r_idx + 1, 0, df_export.iloc[r_idx, 0], f_data_center)
            worksheet.write(r_idx + 1, 1, df_export.iloc[r_idx, 1], f_data_num)
            worksheet.write(r_idx + 1, 2, df_export.iloc[r_idx, 2], f_data_num)
            worksheet.write(r_idx + 1, 3, df_export.iloc[r_idx, 3], f_data_pct)
            worksheet.write(r_idx + 1, 4, df_export.iloc[r_idx, 4], f_data_center)

        worksheet.set_column("A:A", 14)
        worksheet.set_column("B:B", 22)
        worksheet.set_column("C:C", 14)
        worksheet.set_column("D:D", 16)
        worksheet.set_column("E:E", 28)

    return output.getvalue()

st.title("📊 Control de Penetración de Ventas PDM")
st.markdown("Carga el cierre/reporte de caja en PDF para auditar la efectividad por cajero y generar el consolidado de PDM.")

archivo_pdf = st.file_uploader("Arrastra o selecciona el archivo PDF de caja:", type=["pdf"])

if archivo_pdf is not None:
    file_id = f"{archivo_pdf.name}_{archivo_pdf.size}"
    
    if "last_file_id" not in st.session_state or st.session_state.last_file_id != file_id:
        with st.spinner("Decodificando estructura del reporte..."):
            df_boletas, df_vendedores, fecha_rep, conteo_skus = parse_pdf_report(
                archivo_pdf, codigos_pdm_activos
            )
            st.session_state.df_boletas = df_boletas
            st.session_state.df_vendedores = df_vendedores
            st.session_state.fecha_rep = fecha_rep
            st.session_state.conteo_skus = conteo_skus
            st.session_state.last_file_id = file_id

    df_boletas = st.session_state.df_boletas
    df_vendedores = st.session_state.df_vendedores
    fecha_rep = st.session_state.fecha_rep
    conteo_skus = st.session_state.conteo_skus

    if df_boletas.empty:
        st.error("⚠️ No se identificaron transacciones válidas en el documento. Revisa que el PDF contenga las columnas 'Trns' y 'Art/Desc'.")
    else:
        total_trx = df_vendedores["Total_Transacciones"].sum()
        total_pdm = df_vendedores["PDM"].sum()
        penetracion_global = (total_pdm / total_trx * 100) if total_trx > 0 else 0
        top_vendedor = df_vendedores.iloc[0]["Vendedor"] if not df_vendedores.empty else "N/A"
        top_pct = (df_vendedores.iloc[0]["Porcentaje"] * 100) if not df_vendedores.empty else 0

        st.markdown(f"### 📅 Fecha de Operación detectada: **{fecha_rep}**")
        
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Total Transacciones", f"{total_trx:,}")
        col2.metric("Ventas con PDM", f"{total_pdm:,}")
        col3.metric("Penetración Global", f"{penetracion_global:.1f}%")
        col4.metric("Mejor Rendimiento", f"{top_vendedor}", f"{top_pct:.1f}% efectividad")

        st.markdown("---")

        tab_resumen, tab_graficos, tab_auditoria = st.tabs([
            "📋 Resumen de Caja", "📈 Gráficos de Rendimiento", "🔍 Auditoría Boleta por Boleta"
        ])

        with tab_resumen:
            st.subheader("Rendimiento por Colaborador (Código)")
            df_display = df_vendedores.copy()
            df_display["Porcentaje"] = df_display["Porcentaje"].apply(lambda x: f"{x * 100:.1f}%")
            
            st.dataframe(
                df_display.rename(columns={
                    "Total_Transacciones": "Total Boletas",
                    "PDM": "Boletas con PDM",
                    "Porcentaje": "% Efectividad"
                }),
                use_container_width=True,
                hide_index=True
            )

            excel_bytes = create_excel_report(df_vendedores, fecha_rep)
            nombre_archivo = f"Reporte_PDM_Tambo_{fecha_rep.replace('/', '-')}.xlsx"
            
            st.download_button(
                label=f"📥 Descargar Reporte Formateado ({nombre_archivo})",
                data=excel_bytes,
                file_name=nombre_archivo,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        with tab_graficos:
            col_g1, col_g2 = st.columns(2)
            
            with col_g1:
                st.subheader("Efectividad (%) por Vendedor")
                chart_data = df_vendedores.set_index("Vendedor")[["Porcentaje"]] * 100
                st.bar_chart(chart_data)

            with col_g2:
                st.subheader("Top SKUs de PDM Detectados")
                skus_vendidos = {k: v for k, v in conteo_skus.items() if v > 0}
                if skus_vendidos:
                    df_skus = pd.DataFrame(
                        list(skus_vendidos.items()), 
                        columns=["Código SKU", "Unidades Vendidas"]
                    ).sort_values(by="Unidades Vendidas", ascending=False)
                    st.bar_chart(df_skus.set_index("Código SKU"))
                else:
                    st.info("No se registraron ventas de los SKUs PDM configurados.")

        with tab_auditoria:
            st.subheader("Inspección de Boletas Extraídas")
            
            filtro_cajero = st.selectbox(
                "Filtrar por vendedor:", 
                options=["Todos"] + list(df_vendedores["Vendedor"].unique())
            )
            
            df_filtrado = df_boletas if filtro_cajero == "Todos" else df_boletas[df_boletas["Vendedor"] == filtro_cajero]
            
            st.dataframe(
                df_filtrado.rename(columns={
                    "Trns_ID": "N° Transacción",
                    "Contiene_PDM": "Lleva PDM (1/0)",
                    "SKUs_Detectados": "Productos PDM Encontrados"
                }),
                use_container_width=True,
                hide_index=True
            )
else:
    st.info("Esperando el archivo PDF de caja para comenzar el análisis.")
