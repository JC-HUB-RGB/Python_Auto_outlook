"""SETUPPP"""

import win32com.client as win32
import pandas as pd
import openpyxl
import os
from datetime import date
import shutil


"""Info general"""
outlook = win32.Dispatch('outlook.application')
directorio_actual = os.path.dirname(__file__)
hoy = date.today()
fecha_formateada = hoy.strftime("%m-%d-%y")
fecha_ayer = hoy - pd.Timedelta(days=1)
fecha_ayer = fecha_ayer.strftime("%m-%d-%y")
df_columns_to_drop = ['NOA Rec Date', 'NOA Assigment', 'NOA Sent', 'NOA Sent User','Debtor NOA Document','CSR.1', 'Office','Notice', 'Notice Contact Email','NOA Entered Date', 'Last Inv Date', 'Last Inv #', 'First Inv Date', 'Relationship Age', 'First Funded', 'Client Age', 'Funded Balance', 'Non Funded Balance']
Ruta_nube = "C:\\Users\\cokek\\OneDrive - C.R. England\\Documents\\"+fecha_formateada +"_ECS_Factoring_NOARecdDate.xlsx"



RUTA_LOCAL = os.path.join(directorio_actual,"..","data")

"""Ruta Archivo Nuevo"""

ruta_df1 = os.path.join(directorio_actual,"..","data",fecha_formateada+"_ECS_Factoring_NOARecdDate" ".xlsx")
#print(ruta_df1)

"""Ruta Archivo Anterior"""

ruta_df2 = os.path.join(directorio_actual,"..","data",fecha_ayer+"_ECS_Factoring_NOARecdDate" ".xlsx")
#print(ruta_df2)

"""Lectura de archivos Excel"""

try:
    df_hoy = pd.read_excel(ruta_df1)
    df_ayer = pd.read_excel(ruta_df2)
    print("Archivos Excel leídos correctamente.")

except Exception as e:
    print(f"Error al leer los archivos Excel: {e}")
    print(f"Copiando el archivo desde la nube a la ruta local, intentelo de nuevo.")
    shutil.copy2(Ruta_nube, ruta_df1)
    exit(1)
 

""""Adaptacion del df_hoy: eliminar columnas, insertar columna 'CA NOTES' y guardar el archivo para luego realizar el merge con df_ayer"""

try:
        
        df_hoy.drop(columns=df_columns_to_drop, axis='columns', inplace=True, errors='ignore')
        print("Columnas eliminadas")

        df_hoy.insert(9, 'CA NOTES', '')
        print("Columna 'CA NOTES' insertada")
        df_hoy.to_excel(ruta_df1, index=False)

        #print(df_hoy.head())
except ValueError as e:
        Error_Cachado = 'Columna ya existe'
except Exception as e:
        Error_Cachado = 'Error inesperado'
finally:
        """Buscar y reemplazar texto específico en la hoja de Excel"""

        df_hoy.replace(to_replace="Golden Moon Transport Inc //Only Wire or RTP", value="Golden Moon Transport Inc", inplace=True, regex=False)
        df_hoy.replace(to_replace="MBM Global Inc (RTP Only)", value="MBM Global Inc", inplace=True, regex=False)
        df_hoy.replace(to_replace="Gholia Logistics Inc (NO ACH FEE)", value="Gholia Logistics Inc", inplace=True, regex=False)
        df_hoy.replace(to_replace="Debtors@englandlogistics.com", value=" ", inplace=True, regex=False)
        df_hoy.replace(to_replace="paperwork@englandlogistics.com", value=" ", inplace=True, regex=False)
        df_hoy.replace(to_replace="PS#  inquiries@fusiontransport.com ", value="noa@fusiontransport.com", inplace=True, regex=False)
        df_hoy.replace(to_replace="PS:carrierinquiries@challenger.com", value="noa@fusiontransport.com", inplace=True, regex=False)
        df_hoy.replace(to_replace="quickpay@uslfreight.com PS-ap@uslfreight.com TRIUM", value="NOAs:uslogisticsllc@noa.triumphpay.com", inplace=True, regex=False)
        df_hoy.replace(to_replace="PS Triumph: (469) 312-7222", value="Paullog@noa.triumphpay.com", inplace=True, regex=False)
        df_hoy.replace(to_replace="NOA transplacetx@noa.triumphpay.com", value="transplacestuttgart@noa.triumphpay.com", inplace=True, regex=False)
        df_hoy.replace(to_replace="Classic Freight Transportation Inc (NO ACH FEE)", value="Classic Freight Transportation Inc", inplace=True, regex=False)
        df_hoy.replace(to_replace="Dhillon Bros Carrier LLC (RTP Only)", value="Dhillon Bros Carrier LLC", inplace=True, regex=False)
        df_hoy.to_excel(ruta_df1, index=False)

        print("Reemplazo de texto específico completado y archivo guardado.")

        if Error_Cachado == 'Columna ya existe':
            print("La columna 'CA NOTES' ya existe. Continuando con el proceso...")
        
        elif Error_Cachado == 'Error inesperado':
            print(f"Ocurrió un error inesperado: {e}")
        else:
            print("Proceso completado sin errores.")


"""Realizar merge entre df_hoy y df_ayer utilizando 'Last PO #' como clave"""

try:
    matriz_ayer = df_ayer[['Last PO #', 'CA NOTES']]
    df_merged = pd.merge(df_hoy, matriz_ayer, on='Last PO #', how='left')
    df_hoy.to_excel(ruta_df1, index=False)

except Exception as e:
        print(f"Error al realizar el merge: {e}")







            
    


