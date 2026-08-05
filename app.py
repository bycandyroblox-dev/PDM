import streamlit as st
import pandas as pd
import pdfplumber
import re
import io

# Configuración de la página
st.set_page_config(page_title="Procesador PDM Tambo", layout="wide")
st.title("Procesador de Ventas PDM")
st.write("Sube el reporte de caja (PDF) para extraer los datos y descargar el Excel.")

# Diccionario para traducir códigos a nombres de los vendedores
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
    
    # Ahora usamos un diccionario para asegurar que cada Trns sea único y no se pierdan datos por saltos de página
    transacciones_dict = {}
    trx_actual_id = None
    fecha_reporte = "Sin Fecha"
    
    with pdfplumber.open(archivo_pdf) as pdf:
        for page in pdf.pages:
            texto = page.extract_text(layout=True) or page.extract_text()
            if texto:
                # Extraer fecha
                if fecha_reporte == "Sin Fecha":
                    match_fecha = re.search(r'(\d{2}/\d{2}/\d{4})', texto)
                    if match_fecha:
                        fecha_reporte = match_fecha.group(1)

                for linea in texto.split('\n'):
                    linea_basica = linea.strip().lower()
                    
                    # Buscar el inicio estricto de una transacción por su número de Trns
                    match_nueva_trx = re.match(r'^(\d{1,6})\s+(\d{5,8})\b', linea_basica)
                    
                    if match_nueva_trx:
                        trx_actual_id = match_nueva_trx.group(1)
                        # Si el Trns es nuevo, creamos su espacio en el diccionario
                        if trx_actual_id not in transacciones_dict:
                            transacciones_dict[trx_actual_id] = {'texto_lineas': [], 'vendedor': 'Desconocido'}
                            
                    # Si ya estamos dentro de un Trns, agrupamos sus datos sin importar en qué página esté
                    if trx_actual_id:
                        transacciones_dict[trx_actual_id]['texto_lineas'].append(linea_basica)
                        # Buscamos al vendedor si aún no lo tenemos registrado en esta transacción
                        if transacciones_dict[trx_actual_id]['vendedor'] == "Desconocido":
                            match_vendedor = re.search(r'\b(t\d{8})\b', linea_basica)
                            if match_vendedor:
                                transacciones_dict[trx_actual_id]['vendedor'] = match_vendedor.group(1).upper()

    if not transacciones_dict:
        st.error("No se detectaron transacciones. Verifica el formato del PDF.")
    else:
        datos_finales = []
        
        # Procesamos cada transacción de forma individual y segura
        for trns_id, data in transacciones_dict.items():
            texto_trx_limpio = limpiar_texto(" ".join(data['texto_lineas']))
            pdms_en_boleta = set()
            
            for pdm_nombre, palabras_clave in LISTA_PDM.items():
                if all(palabra in texto_trx_limpio for palabra in palabras_clave):
                    pdms_en_boleta.add(pdm_nombre)
            
            # Traducir código al nombre real
            cod_vendedor = data['vendedor']
            nombre_vendedor = NOMBRES_VENDEDORES.get(cod_vendedor, cod_vendedor)
            
            datos_finales.append({
                "Trns_ID": trns_id,
                "Vendedor": nombre_vendedor,
                "Contiene_PDM": 1 if len(pdms_en_boleta) > 0 else 0
            })
            
        df = pd.DataFrame(datos_finales)
        
        # Agrupamos los datos separando por vendedor
        df_vendedores = df.groupby('Vendedor').agg(
            Total_Transacciones=('Contiene_PDM', 'count'),
            PDM=('Contiene_PDM', 'sum')
        ).reset_index()
        
        # Calculamos los totales absolutos de la caja
        total_trx_dia = df_vendedores['Total_Transacciones'].sum()
        total_pdm_dia = df_vendedores['PDM'].sum()
        
        filas_excel = []
        
        # 1. Agregamos primero la fila del TOTAL CAJA general
        filas_excel.append({
            "Fecha": fecha_reporte,
            "Total Transacciones": total_trx_dia,
            "PDM": total_pdm_dia,
            "Porcentaje": (total_pdm_dia / total_trx_dia) if total_trx_dia > 0 else 0,
            "Vendedor Responsable": "TOTAL CAJA"
        })
        
        # 2. Agregamos las filas separadas por cada vendedor que operó ese día
        for index, row in df_vendedores.iterrows():
            filas_excel.append({
                "Fecha": fecha_reporte,
                "Total Transacciones": row['Total_Transacciones'],
                "PDM": row['PDM'],
                "Porcentaje": (row['PDM'] / row['Total_Transacciones']) if row['Total_Transacciones'] > 0 else 0,
                "Vendedor Responsable": row['Vendedor']
            })
            
        df_excel = pd.DataFrame(filas_excel)
        
        # Generar archivo Excel con estilos
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_excel.to_excel(writer, sheet_name="Reporte PDM", index=False, header=False, startrow=1)
            
            workbook = writer.book
            worksheet = writer.sheets["Reporte PDM"]
            
            formato_cabecera = workbook.add_format({'bold': True, 'align': 'center', 'border': 1, 'bg_color': '#D9D9D9'})
            formato_verde = workbook.add_format({'bg_color': '#32CD32', 'align': 'center', 'border': 1})
            formato_verde_porcentaje = workbook.add_format({'bg_color': '#32CD32', 'num_format': '0.0%', 'align': 'center', 'border': 1})
            formato_blanco_centro = workbook.add_format({'align': 'center', 'border': 1})
            
            # Anchos de columna
            worksheet.set_column('A:A', 12)
            worksheet.set_column('B:B', 20)
            worksheet.set_column('C:C', 10)
            worksheet.set_column('D:D', 15)
            worksheet.set_column('E:E', 25)
            
            # Pintar Cabeceras
            for col_num, value in enumerate(df_excel.columns.values):
                worksheet.write(0, col_num, value, formato_cabecera)
                
            # Pintar las filas
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
