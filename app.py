import streamlit as st
import pandas as pd
import pdfplumber
import re
import io

# Configuración de la página
st.set_page_config(page_title="Procesador PDM Tambo", layout="wide")
st.title("Procesador de Ventas PDM")
st.write("Sube el reporte de caja (PDF) para extraer los datos y descargar el Excel.")

# Diccionario de vendedores
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

# Lista de palabras clave optimizada
LISTA_PDM = {
    "Papas Lays Clásicas": ["lays", "clasicas"],
    "Doritos Queso Atrevido": ["doritos", "atrevido"],
    "Mix Piqueo Snax": ["piqueo", "snax"],
    "Chocman Doble Manjar": ["chocman", "doble"],
    "Gomitas Ambrosia": ["ambrosia"],
    "Bon o Bon": ["bon", "o", "bon"], 
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
    
    with pdfplumber.open(archivo_pdf) as pdf:
        texto_completo = ""
        for page in pdf.pages:
            texto_completo += (page.extract_text() or "") + "\n"
            
    # Extraer la fecha del reporte
    fecha_reporte = "Sin Fecha"
    match_fecha = re.search(r'(\d{2}/\d{2}/\d{4})', texto_completo)
    if match_fecha:
        fecha_reporte = match_fecha.group(1)

    # NUEVO CEREBRO: Separa todo el texto palabra por palabra ignorando el formato
    tokens = [t for t in re.split(r'\s+|\|', texto_completo) if t]
    
    transacciones_dict = {}
    current_trx = None
    
    for i, token in enumerate(tokens):
        # Detecta el Trns: Es un número de 3 a 6 dígitos seguido inmediatamente por un Código de Artículo (7 a 9 dígitos)
        if re.match(r'^\d{3,6}$', token):
            if i + 1 < len(tokens) and re.match(r'^\d{7,9}$', tokens[i+1]):
                current_trx = token
                transacciones_dict[current_trx] = []
                
        # Agrupa toda la data dentro de su número de Trns
        if current_trx:
            transacciones_dict[current_trx].append(token)

    if not transacciones_dict:
        st.error("No se detectaron transacciones. Verifica el formato del PDF.")
    else:
        datos_finales = []
        
        for trns_id, tokens_trx in transacciones_dict.items():
            texto_trx_limpio = limpiar_texto(" ".join(tokens_trx))
            
            # Buscar al vendedor dentro de la transacción
            cod_vendedor = "Desconocido"
            for t in tokens_trx:
                if re.match(r'^T\d{8}$', t, re.IGNORECASE):
                    cod_vendedor = t.upper()
                    break
                    
            nombre_vendedor = NOMBRES_VENDEDORES.get(cod_vendedor, cod_vendedor)
            
            # Buscar PDMs
            pdms_en_boleta = set()
            for pdm_nombre, palabras_clave in LISTA_PDM.items():
                if all(palabra in texto_trx_limpio for palabra in palabras_clave):
                    pdms_en_boleta.add(pdm_nombre)
                    
            datos_finales.append({
                "Trns_ID": trns_id,
                "Vendedor": nombre_vendedor,
                "Contiene_PDM": 1 if len(pdms_en_boleta) > 0 else 0
            })
            
        df = pd.DataFrame(datos_finales)
        
        # Agrupación exacta por vendedor
        df_vendedores = df.groupby('Vendedor').agg(
            Total_Transacciones=('Contiene_PDM', 'count'),
            PDM=('Contiene_PDM', 'sum')
        ).reset_index()
        
        total_trx_dia = df_vendedores['Total_Transacciones'].sum()
        total_pdm_dia = df_vendedores['PDM'].sum()
        
        filas_excel = []
        
        # 1. Fila de TOTAL CAJA
        filas_excel.append({
            "Fecha": fecha_reporte,
            "Total Transacciones": total_trx_dia,
            "PDM": total_pdm_dia,
            "Porcentaje": (total_pdm_dia / total_trx_dia) if total_trx_dia > 0 else 0,
            "Vendedor Responsable": "TOTAL CAJA"
        })
        
        # 2. Filas divididas por Vendedor
        for index, row in df_vendedores.iterrows():
            filas_excel.append({
                "Fecha": fecha_reporte,
                "Total Transacciones": row['Total_Transacciones'],
                "PDM": row['PDM'],
                "Porcentaje": (row['PDM'] / row['Total_Transacciones']) if row['Total_Transacciones'] > 0 else 0,
                "Vendedor Responsable": row['Vendedor']
            })
            
        df_excel = pd.DataFrame(filas_excel)
        
        # Generar Excel final
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
            worksheet.set_column('E:E', 25)
            
            for col_num, value in enumerate(df_excel.columns.values):
                worksheet.write(0, col_num, value, formato_cabecera)
                
            for row_num in range(len(df_excel)):
                worksheet.write(row_num + 1, 0, df_excel.iloc[row_num, 0], formato_blanco_centro)
                worksheet.write(row_num + 1, 1, df_excel.iloc[row_num, 1], formato_verde)
                worksheet.write(row_num + 1, 2, df_excel.iloc[row_num, 2], formato_verde)
                worksheet.write(row_num + 1, 3, df_excel.iloc[row_num, 3], formato_verde_porcentaje)
                worksheet.write(row_num + 1, 4, df_excel.iloc[row_num, 4], formato_blanco_centro)
                
        excel_data = output.getvalue()
        
        st.success(f"Analisis completado. Se detectaron {total_trx_dia} transacciones exactas.")
        st.download_button(
            label="Descargar Reporte en Excel",
            data=excel_data,
            file_name="Reporte_Ventas_PDM_Tambo.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
