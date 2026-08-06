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

# Nueva lista de PDM basada estrictamente en Códigos de Artículo
CODIGOS_PDM = {
    "1002403", "1000986", "1006886", "1010945", "1010944",
    "1016699", "1014585", "1005799", "1005644", "1000918",
    "1001529", "1007039", "1001613", "1004275", "400150017",
    "1010148", "1016708"
}

archivo_pdf = st.file_uploader("Sube tu archivo de Reporte (PDF)", type=["pdf"])

if archivo_pdf is not None:
    st.info("Procesando reporte de caja... Esto puede tomar unos segundos.")
    
    transacciones_lista = []
    current_trx_data = None
    in_table = False
    fecha_reporte = "Sin Fecha"
    
    with pdfplumber.open(archivo_pdf) as pdf:
        for page in pdf.pages:
            texto = page.extract_text(layout=True)
            if not texto: continue
            
            # Extraer la fecha (solo la primera vez que la encuentra)
            if fecha_reporte == "Sin Fecha":
                match_fecha = re.search(r'(\d{2}/\d{2}/\d{4})', texto)
                if match_fecha:
                    fecha_reporte = match_fecha.group(1)

            # Analizamos línea por línea
            for linea in texto.split('\n'):
                linea_lower = linea.lower()
                
                # Detectamos dónde empieza realmente la tabla para no contar resúmenes
                if "trns" in linea_lower and ("art" in linea_lower or "desc" in linea_lower):
                    in_table = True
                    continue
                    
                if not in_table:
                    continue
                    
                # Busca de 1 a 6 números anclados en el extremo izquierdo de la fila.
                match_trx = re.match(r'^ {0,4}(\d{1,6})(?:\s+|$)', linea)
                
                if match_trx:
                    if current_trx_data is not None:
                        transacciones_lista.append(current_trx_data)
                        
                    current_trx_data = {
                        'id': match_trx.group(1), 
                        'texto_lineas': [linea], 
                        'vendedor': 'Desconocido'
                    }
                    
                    match_vendedor = re.search(r'\b(t\d{8})\b', linea_lower)
                    if match_vendedor:
                        current_trx_data['vendedor'] = match_vendedor.group(1).upper()
                        
                elif current_trx_data is not None:
                    current_trx_data['texto_lineas'].append(linea)
                    if current_trx_data['vendedor'] == 'Desconocido':
                        match_vendedor = re.search(r'\b(t\d{8})\b', linea_lower)
                        if match_vendedor:
                            current_trx_data['vendedor'] = match_vendedor.group(1).upper()

    # Guardar la última transacción cuando se acaba el documento
    if current_trx_data is not None:
        transacciones_lista.append(current_trx_data)

    if not transacciones_lista:
        st.error("No se detectaron transacciones. Verifica el formato del PDF.")
    else:
        datos_finales = []
        
        for data in transacciones_lista:
            # Unimos todas las líneas de la transacción
            texto_trx_completo = " ".join(data['texto_lineas'])
            
            # NUEVO CEREBRO: Extrae todos los números sueltos de 7 a 9 dígitos
            codigos_en_boleta = set(re.findall(r'\b\d{7,9}\b', texto_trx_completo))
            
            # Compara los códigos de la boleta con tu lista de PDM y se queda con las coincidencias
            pdms_en_boleta = codigos_en_boleta.intersection(CODIGOS_PDM)
            
            # Traducir código al nombre real
            cod_vendedor = data['vendedor']
            nombre_vendedor = NOMBRES_VENDEDORES.get(cod_vendedor, cod_vendedor)
            
            datos_finales.append({
                "Trns_ID": data['id'],
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
        
        # Generar Excel final con formato visual
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
        
        st.success(f"Analisis completado. Se detectaron exactamente {total_trx_dia} transacciones.")
        st.download_button(
            label="Descargar Reporte en Excel",
            data=excel_data,
            file_name="Reporte_Ventas_PDM_Tambo.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
