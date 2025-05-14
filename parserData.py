import requests
import pandas as pd
import re

# URL de la API de la tabla Ninja (AJAX)
url_ajax = "https://distribuidores.carmotion.com.mx/wp-admin/admin-ajax.php?action=wp_ajax_ninja_tables_public_action&table_id=177&target_action=get-all-data&default_sorting=old_first&ninja_table_public_nonce=6f68928eb3&chunk_number=0"

response = requests.get(url_ajax)

if response.status_code == 200:
    data = response.json()
    print("Extrayendo datos.....")
    # Extraer los datos relevantes (SKU, DESCRIPCION, TOTAL, PRECIO)
    datos = []
    for row in data:
        value = row.get('value', {})
        datos.append([
            value.get('a', ''),  # SKU
            value.get('b', ''),  # DESCRIPCION
            value.get('c', ''),  # TOTAL
            value.get('d', '')   # PRECIO
        ])

    # Guardar en CSV
    df = pd.DataFrame(datos, columns=["SKU", "DESCRIPCION", "TOTAL", "PRECIO"])

    
    df["DESCRIPCION"] = df["DESCRIPCION"].apply(lambda x: re.sub(r'\s+', ' ', x).strip())

    df.to_csv("datos_extraidos.csv", index=False)

    print("Extracción completada. Datos guardados en 'datos_extraidos.csv'.")

    ########################3Leer el archivo de Excel###########################
    excel_path = "CARMOTION_GOMMAS.xlsx"
    xls = pd.ExcelFile(excel_path)
    df_excel = pd.read_excel(xls, sheet_name="CAR MOT", header=1)

    df_excel_copy = df_excel.copy()

    for index, row in df.iterrows():
        sku = row["SKU"]
        mask = df_excel_copy["."] == sku

        if mask.any():
           
            df_excel_copy.loc[mask, "DESCRIPCION"] = row["DESCRIPCION"]

            
            df_excel_copy.loc[mask, "Columna1"] = pd.to_numeric(row["TOTAL"], errors='coerce')


            precio = row["PRECIO"].strip()
            try:
                precio_num = int(float(precio.replace("$", "").replace(",", "")))
            except ValueError:
                precio_num = 0

            df_excel_copy.loc[mask, "Columna2"] = precio_num

    # Guardar el nuevo Excel actualizado
    df_excel_copy.to_excel("CARMOTION_GOMMAS_ACTUALIZADO.xlsx", sheet_name="CAR MOT", index=False, startrow=1)
    print("Actualización completada en el archivo 'CARMOTION_GOMMAS_ACT.xlsx'.")

else:
    print(f"Error al obtener los datos. Código de estado: {response.status_code}")
