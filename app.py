import streamlit as st
import pandas as pd
import pdfplumber
import re
import io

# Configuración de la página
st.set_page_config(page_title="Procesador PDM Tambo", layout="wide")
st.title("Procesador de Ventas PDM")
st.write("Sube el reporte de caja (PDF) para extraer los datos y descargar el Excel.")

# Diccionario para traducir códigos a nombres
NOMBRES_VENDEDORES = {
    "T72758473": "Britney",
    "T70729978": "Jhony",
    "T70962854": "Lucia"
}

def limpiar_texto(texto):
    texto = str(texto).lower()
    tildes = {'á': 'a', 'é': 'e', 'í': 'i', 'ó': 'o', 'ú': 'u'}
    for con_tilde, sin_tilde in tildes.items():
        texto = texto.replace(con_tilde, sin_tilde)
    texto = re.sub(r'\s+', ' ', texto)
    return texto.strip()

LISTA_PDM = {
    "Papas Lays Clásicas": ["lays", "clasicas"],
    "Doritos Queso Atrevido": ["doritos", "atrevido"],
    "Mix Piqueo Snax": ["piqueo", "snax"],
    "Chocman Doble Manjar": ["chocman", "doble"],
    "Gomitas Ambrosia": ["ambrosia"],
    "BonoBn": ["bonobn"],
    "Mogul Pastillas Frutales": ["mogul", "pastilla"],
    "Mogul Sandía": ["mogul", "sandia"],
    "Jelly Beans Extreme": ["jelly", "beans"],
    "Osito Extreme": ["osito", "extreme"],
    "Morochas XL": ["morochas", "xl"],
    "Sublime Sonrisa": ["sublime", "sonrisa"],
    "Cappuccino 37gr": ["cappuccino", "37"],
    "Blanco 40gr": ["blanco", "40"],
    "Cabanossi Braedt": ["cabanossi"],
    "Chips Ahoy": ["chips", "ahoy"],
    "Oreo Fresa": ["oreo", "fresa"],
    "Chocolate 108gr": ["chocolate", "108"],
    "Regular Rollo 135gr": ["regular", "rollo"]
}

archivo_pdf = st.file_uploader("Sube tu archivo de Reporte (PDF)", type=["pdf"])

if archivo_pdf is not None:
    st.info("Procesando reporte de caja... Esto puede tomar unos segundos.")
    
    transacciones_procesadas = []
    trx_actual_texto = []
    vendedor_actual = "Desconocido"
    fecha_reporte = "Sin Fecha"
    
    with pdfplumber.open(archivo_pdf) as pdf:
        for page in pdf.pages:
            texto = page.extract_text(layout=True) or page.extract_text()
            if texto:
                if fecha_reporte == "Sin Fecha":
                    match_fecha = re.search(r'(\d{2}/\d{2}/\d{4})', texto)
                    if match_fecha:
                        fecha_reporte = match_fecha.group(1)

                for linea in texto.split('\n'):
                    linea_basica = linea.strip().lower()
                    
                    match_nueva_trx = re.match(r'^(\d{1,6})\s+(\d{5,8})\s+', linea_basica)
                    
                    if match_nueva_trx:
                        if trx_actual_texto:
                            transacciones_procesadas.append({
                                'texto': " ".join(trx_actual_texto),
                                'vendedor': vendedor_actual
                            })
                        trx_actual_texto = [linea_basica]
                        
                        match_vendedor = re.search(r'\b(t\d{8})\b', linea_basica)
                        if match_vendedor:
                            vendedor_actual = match_vendedor.group(1).upper()
                        else:
                            vendedor_actual = "Desconocido"
                            
                    elif trx_actual_texto:
                        trx_actual_texto.append(linea_basica)
                        if vendedor_actual == "Desconocido":
                            match_vendedor = re.search(r'\b(t\d{8})\b', linea_basica)
                            if match_vendedor:
                                vendedor_actual = match_vendedor.group(1).upper()

    if trx_actual_texto:
        transacciones_procesadas.append({
            'texto': " ".join(trx_actual_texto),
            'vendedor': vendedor_actual
        })

    if not transacciones_procesadas:
        st.error("No se detectaron transacciones. Verifica el formato del PDF.")
    else:
        datos_finales = []
        
        for trx in transacciones_procesadas:
            texto_trx_limpio = limpiar_texto(trx['texto'])
            pdms_en_boleta = set()
            
            for pdm_nombre, palabras_clave in LISTA_PDM.items():
                if all(palabra in texto_trx_limpio for palabra in palabras_clave):
                    pdms_en_boleta.add(pdm_nombre)
            
            datos_finales.append({
                "Vendedor": trx['vendedor'],
                "Contiene_PDM": 1 if len(pdms_en_boleta) > 0 else 0
            })
            
        df = pd.DataFrame(datos_finales)
        
        total_trx_dia = len(df)
        total_pdm_dia = df['Contiene_PDM'].sum()
        porcentaje_dia = (total_pdm_dia / total_trx_dia) if total_trx_dia > 0 else 0
        
        vendedores_unicos = df['Vendedor'].unique().tolist()
        if "Desconocido" in vendedores_unicos and len(vendedores_unicos) > 1:
            vendedores_unicos.remove("Desconocido")
            
        # Traducir los códigos a nombres reales
        nombres_traducidos = [NOMBRES_VENDEDORES.get(v, v) for v in vendedores_unicos]
        vendedores_str = ", ".join(nombres_traducidos)
        
        filas_excel = [{
            "Fecha": fecha_reporte,
            "Total Transacciones": total_trx_dia,
            "PDM": total_pdm_dia,
            "Porcentaje": porcentaje_dia,
            "Vendedor Responsable": vendedores_str
        }]
        
        df_excel = pd.DataFrame(filas_excel)
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_excel.to_excel(writer, sheet_name="Reporte PDM", index=False, header=False, startrow=1)
            
            workbook = writer.book
            worksheet = writer.sheets["Reporte PDM"]
            
            formato_cabecera = workbook.add_format({'bold': True, 'align': 'center', 'border': 1, 'bg_color': '#D9D9D9'})
            formato_verde = workbook.add_format({'bg_color': '#32CD32', 'align': 'center', 'border': 1})
            formato_verde_porcentaje = workbook.add_format({'bg_color': '#32CD32', 'num_format': '0.0%', 'align': 'center', 'border': 1})
            formato_blanco_centro = workbook.add_format({'align': 'center', 'border': 1})
            
            worksheet.set_column('A:A', 12)
            worksheet.set_column('B:B', 20)
            worksheet.set_column('C:C', 10)
            worksheet.set_column('D:D', 15)
            worksheet.set_column('E:E', 40) # Columna de vendedores más ancha por si hay varios nombres
            
            for col_num, value in enumerate(df_excel.columns.values):
                worksheet.write(0, col_num, value, formato_cabecera)
                
            worksheet.write(1, 0, df_excel.iloc[0, 0], formato_blanco_centro)
            worksheet.write(1, 1, df_excel.iloc[0, 1], formato_verde)
            worksheet.write(1, 2, df_excel.iloc[0, 2], formato_verde)
            worksheet.write(1, 3, df_excel.iloc[0, 3], formato_verde_porcentaje)
            worksheet.write(1, 4, df_excel.iloc[0, 4], formato_blanco_centro)
                
        excel_data = output.getvalue()
        
        st.success(f"Analisis completado. Se detectaron {total_trx_dia} transacciones en total.")
        st.download_button(
            label="Descargar Reporte en Excel",
            data=excel_data,
            file_name="Reporte_Ventas_PDM_Tambo.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )