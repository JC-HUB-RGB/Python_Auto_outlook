"""SETUPPP"""
import win32com.client as win32
import pandas as pd
import openpyxl
import os
from datetime import date
import shutil
import time
import tkinter as tk
import archives_loader as CA
import email_creator as EC

"""Info general"""
onedrive_path_noas = os.path.expanduser('~\\OneDrive - C.R. England\\NOAs')
onedrive_path = os.path.expanduser('~\\OneDrive - C.R. England')
date_name = date.weekday(date.today())
outlook = win32.Dispatch('outlook.application')
directorio_actual = os.path.dirname(__file__)
hoy = date.today()
fecha_hoy = hoy.strftime("%m-%d-%y")
fecha_ayer = hoy - pd.Timedelta(days=1)
fecha_ayer = fecha_ayer.strftime("%m-%d-%y")
terminacion_archivo_procesado = "_ECS_Factoring_NOARecdDate_Procesado.xlsx"
terminacion_archivo_sin_procesar = "_ECS_Factoring_NOARecdDate.xlsx"

adaptacion_deudores = pd.read_excel(os.path.join(onedrive_path_noas, "Debtors Info.xlsx"), sheet_name='Debtors')
adaptacion_carriers = pd.read_excel(os.path.join(onedrive_path_noas, "Debtors Info.xlsx"), sheet_name='Carriers')

df_adaptacion_deudores = pd.DataFrame(adaptacion_deudores)
df_adaptacion_carriers = pd.DataFrame(adaptacion_carriers)




"""Rutas Generales"""

ruta_excel_hoy_sin_procesar = os.path.join(directorio_actual,"..","data",fecha_hoy +terminacion_archivo_sin_procesar)
ruta_excel_hoy_procesado = os.path.join(directorio_actual,"..","data",fecha_hoy+terminacion_archivo_procesado)
ruta_descargar_excel_hoy_procesado = os.path.join(onedrive_path_noas, "FILE NOA PROCESSED", fecha_hoy + terminacion_archivo_procesado)

""""Dataframes y variables globales"""
df_columns_to_drop = ['NOA Rec Date', 'NOA Assigment', 'NOA Sent','Last Inv Date', 'NOA Sent User','Debtor NOA Document','CSR.1', 'Office','Notice', 'Notice Contact Email','NOA Entered Date', 'Last Inv #', 'First Inv Date', 'First Funded', 'Client Age', 'Funded Balance', 'Non Funded Balance']


"""Definicion de funciones"""

def Guardar_archivo_excel(nombre_dataframe, ruta_guardar):
   nombre_dataframe.to_excel(ruta_guardar, index=False)      

columnas_a_modificar_debtors = ['Debtor Email Address', 'Attention Note', 'Warning Note']

mapa_reemplazo_carriers = dict(zip(df_adaptacion_carriers['Unprocessed'], df_adaptacion_carriers['Processed']))

while True:
        try:
                CSR_INPUT = input("Porfavor seleccione el CSR que desea procesar: 1.SPAREDES 2.DMONTANO 3.tcorona 4.AAGUILAR 5.jehoug 6.VGUERRERO 7.dagarcia 8.VMANRIQUEZ\n")

                if CSR_INPUT == '1':
                        CSR_name = 'SPAREDES'
                        break
                elif CSR_INPUT == '2':
                        CSR_name = 'DMONTANO'
                        break
                elif CSR_INPUT == '3':
                        CSR_name = 'tcorona'
                        break
                elif CSR_INPUT == '4':
                        CSR_name = 'AAGUILAR'
                        break 
                elif CSR_INPUT == '5':
                        CSR_name = 'jehoug'
                        break
                elif CSR_INPUT == '6':
                        CSR_name = 'VGUERRERO'
                        break
                elif CSR_INPUT == '7':
                        CSR_name = 'dagarcia'
                        break
                elif CSR_INPUT == '8':
                        CSR_name = 'VMANRIQUEZ'
                        break
                elif CSR_INPUT == '9':
                        CSR_name = 'BSANCHEZ'
                        break

                else: 
                        raise ValueError("Entrada no válida. Por favor, seleccione 1, 2, 3, 4, 5, 6 o 7.")
                
        except ValueError:
                print("Entrada no válida. Por favor, seleccione 1, 2, 3, 4, 5, 6 o 7.")


       
df_hoy, df_ayer, decision = CA.carga_archivos_excel()

if decision == 0: #no se encontro un archivo procesado de hoy
        try:    
                """"Adaptacion del df_hoy: eliminar columnas, insertar columna 'CA NOTES' """
                df_hoy.drop(columns=df_columns_to_drop, axis='columns', inplace=True, errors='ignore')
                print("Columnas eliminadas")
                try:
                        df_hoy.insert(9, 'CA NOTES', '')
                        df_hoy.insert(14, '1st attempt', '')
                        df_hoy.insert(16, '2nd attempt', '')
                        df_hoy.insert(18, '3rd attempt', '')
                                                                                                
                        print("Columnas 'CA NOTES', '1st attempt', '2nd attempt', '3rd attempt' insertadas correctamente")
                except Exception as e :
                        print("La columna 'CA NOTES' ya existe. Continuando con el proceso...")

                print("Creacion del nuevo archivo con las columnas 'CA NOTES', '1st attempt', '2nd attempt', '3rd attempt' sin errores")
                try:
                        for indice, fila in df_adaptacion_carriers.iterrows():
                                carrier_real = fila['Unprocessed']
                                carrier_limpio = fila['Processed']
                                df_hoy = df_hoy.replace({carrier_real: carrier_limpio})
                except Exception as e :
                        print("Error cambiando valores de Carriers")
                        exit(1)


                print("Reemplazo de nombres de Carriers completado")
                df_ayer = df_ayer.drop_duplicates(subset=['Last PO #'], keep='last')
                df_hoy = df_hoy.drop_duplicates(subset=['Last PO #'], keep='last')
                try:
                        for indice, fila in df_adaptacion_deudores.iterrows():
                                #print(f'Remplazo de texto para deudor {deudor}')
                                deudor_real = fila['Debtor Name']
                                valores_reemplazo = fila[columnas_a_modificar_debtors].values
                                df_hoy.loc[df_hoy['Debtor Name'] == deudor_real, columnas_a_modificar_debtors] = valores_reemplazo

                except Exception as e :
                        print("Error cambiando valores de debtors")
                        exit(1)

                #print(df_hoy.head()) 

        except Exception as e:
                print(f"Ocurrió un error inesperado en la adaptacion del archivo de hoy: {e}")
                exit(1)

        """Realizar merge entre df_hoy y df_ayer utilizando 'Last PO #' como clave"""
        columnas_a_copiar = ['CA NOTES', '1st attempt', '2nd attempt', '3rd attempt']
        try:
            for columna in columnas_a_copiar:    
                mapeo_notas = df_ayer.set_index('Last PO #')[columna]
                df_hoy[columna] = df_hoy['Last PO #'].map(mapeo_notas)
        except Exception as e:
                print(f"Error al realizar el merge: {e}")
                exit(1)

        """FILTRO DE REGISTROS PARA ENVIAR CORREO"""

        try:    

                Condicion1 = df_hoy['CSR'] == CSR_name
                Condicion2 = df_hoy['CA NOTES'].isna()
                Condicion_Edad = df_hoy['Relationship Age'].isin([7, 14, 21])
                Condicion_1st = ~df_hoy['1st attempt'].astype(str).str.contains(fecha_hoy, na=False)
                Condicion_2nd = ~df_hoy['2nd attempt'].astype(str).str.contains(fecha_hoy, na=False)
                Condicion_3rd = ~df_hoy['3rd attempt'].astype(str).str.contains(fecha_hoy, na=False)
                df_correos_enviar = df_hoy[Condicion1 & (Condicion2 | Condicion_Edad) & (Condicion_1st & Condicion_2nd & Condicion_3rd)].copy()
                size = df_correos_enviar['Last PO #'].size
                print(f'Existen {size} correos por enviar')
                time.sleep(10)
        except Exception as e:
                print(f"Error al filtrar los registros para el correo: {e}")
                exit(1)

elif decision == 1: #se econtro un archivo procesado de hoy, pero el debtors file tiene cambios.
                print("Reemplazo de texto específico completado")
                try:
                        for indice, fila in df_adaptacion_deudores.iterrows():
                                #print(f'Remplazo de texto para deudor {deudor}')
                                deudor_real = fila['Debtor Name']
                                valores_reemplazo = fila[columnas_a_modificar_debtors].values
                                df_hoy.loc[df_hoy['Debtor Name'] == deudor_real, columnas_a_modificar_debtors] = valores_reemplazo

                except Exception as e :
                        print("Error cambiando valores de debtors")
                        exit(1)


                try:
                        Condicion1 = df_hoy['CSR'] == CSR_name
                        Condicion2 = df_hoy['CA NOTES'].isna()
                        Condicion_Edad = df_hoy['Relationship Age'].isin([7, 14, 21])
                        Condicion_1st = ~df_hoy['1st attempt'].astype(str).str.contains(fecha_hoy, na=False)
                        Condicion_2nd = ~df_hoy['2nd attempt'].astype(str).str.contains(fecha_hoy, na=False)
                        Condicion_3rd = ~df_hoy['3rd attempt'].astype(str).str.contains(fecha_hoy, na=False)
                        df_correos_enviar = df_hoy[Condicion1 & (Condicion2 | Condicion_Edad) & (Condicion_1st & Condicion_2nd & Condicion_3rd)].copy()
                        size = df_correos_enviar['Last PO #'].size
                        print(f'Existen {size} correos por enviar')
                        time.sleep(10)
                except Exception as e:
                        print(f"Error al filtrar los registros para el correo: {e}")
                        exit(1)

        
elif decision == 2: #se encontro un archivo procesado de hoy, y el debtors file no tiene cambios.
        """FILTRO DE REGISTROS PARA ENVIAR CORREO"""
        try:    

                Condicion1 = df_hoy['CSR'] == CSR_name
                Condicion2 = df_hoy['CA NOTES'].isna()
                Condicion_Edad = df_hoy['Relationship Age'].isin([7, 14, 21])
                Condicion_1st = ~df_hoy['1st attempt'].astype(str).str.contains(fecha_hoy, na=False)
                Condicion_2nd = ~df_hoy['2nd attempt'].astype(str).str.contains(fecha_hoy, na=False)
                Condicion_3rd = ~df_hoy['3rd attempt'].astype(str).str.contains(fecha_hoy, na=False)
                df_correos_enviar = df_hoy[Condicion1 & (Condicion2 | Condicion_Edad) & (Condicion_1st & Condicion_2nd & Condicion_3rd)].copy()
                size = df_correos_enviar['Last PO #'].size
                print(f'Existen {size} correos por enviar')
                time.sleep(10)
        except Exception as e:
                print(f"Error al filtrar los registros para el correo: {e}")
                exit(1)

EC.enviar_correos(df_correos_enviar, fecha_hoy, onedrive_path)

for col_actualizar in ['CA NOTES', '1st attempt', '2nd attempt', '3rd attempt']:
    df_final = df_hoy['Last PO #'].map(df_correos_enviar.set_index('Last PO #')[col_actualizar])
    df_hoy[col_actualizar] = df_final.combine_first(df_hoy[col_actualizar])
Guardar_archivo_excel(df_hoy,ruta_excel_hoy_procesado)
shutil.copy2(ruta_excel_hoy_procesado,ruta_descargar_excel_hoy_procesado)