"""SETUPPP"""
import win32com.client as win32
import pandas as pd
import openpyxl
import os
from datetime import date
import shutil
import time
import tkinter as tk
import pywintypes
import Cargador_de_archivos as CA

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

df_columns_to_drop = ['NOA Rec Date', 'NOA Assigment', 'NOA Sent', 'NOA Sent User','Debtor NOA Document','CSR.1', 'Office','Notice', 'Notice Contact Email','NOA Entered Date', 'Last Inv Date', 'Last Inv #', 'First Inv Date', 'Relationship Age', 'First Funded', 'Client Age', 'Funded Balance', 'Non Funded Balance']
Palabras_Buscar = ["noa", "doesnt verify", "@noa.triumphpay", "web", "triumphpay", "website", "use triumph pay", "triumph", "tp", "epay"]
nombres_no_touch = ["avensis energy services, llc", "metro parcel & freight inc", "mvk transport corporation", "cmd energy, llc"]


"""Definicion de funciones"""

def Guardar_archivo_excel(nombre_dataframe, ruta_guardar):
   nombre_dataframe.to_excel(ruta_guardar, index=False)      

columnas_a_modificar_debtors = ['Debtor Email Address', 'Attention Note', 'Warning Note']

mapa_reemplazo_carriers = dict(zip(df_adaptacion_carriers['Unprocessed'], df_adaptacion_carriers['Processed']))

while True:
        try:
                CSR_INPUT = input("Porfavor seleccione el CSR que desea procesar: 1.VGUERRERO 2.MPALMER 3.SPAREDES 4.dagarcia 5.jehoug 6.tcorona 7.AAGUILAR 8.VMANRIQUEZ\n")

                if CSR_INPUT == '1':
                        CSR_name = 'VGUERRERO'
                        break
                elif CSR_INPUT == '2':
                        CSR_name = 'MPALMER'
                        break
                elif CSR_INPUT == '3':
                        CSR_name = 'SPAREDES'
                        break
                elif CSR_INPUT == '4':
                        CSR_name = 'dagarcia'
                        break 
                elif CSR_INPUT == '5':
                        CSR_name = 'jehoug'
                        break
                elif CSR_INPUT == '6':
                        CSR_name = 'tcorona'
                        break
                elif CSR_INPUT == '7':
                        CSR_name = 'AAGUILAR'
                        break
                elif CSR_INPUT == '8':
                        CSR_name = 'VMANRIQUEZ'
                        break

                else: 
                        raise ValueError("Entrada no válida. Por favor, seleccione 1, 2, 3, 4, 5, 6 o 7.")
                
        except ValueError:
                print("Entrada no válida. Por favor, seleccione 1, 2, 3, 4, 5, 6 o 7.")
       

#df_hoy, df_ayer, decision = CA.carga_archivos_excel()
df_hoy, df_ayer, decision = CA.carga_archivos_excel_nuevo()

if decision == 0:
        try:    
                """"Adaptacion del df_hoy: eliminar columnas, insertar columna 'CA NOTES' """
                df_hoy.drop(columns=df_columns_to_drop, axis='columns', inplace=True, errors='ignore')
                print("Columnas eliminadas")
                try:
                        df_hoy.insert(9, 'CA NOTES', '')
                        print("Columna 'CA NOTES' insertada")
                except Exception as e :
                        print("La columna 'CA NOTES' ya existe. Continuando con el proceso...")

                print("Creacion del nuevo archivo con la columna 'CA NOTES' sin errores")
                try:
                        for indice, fila in df_adaptacion_carriers.iterrows():
                                carrier_real = fila['Unprocessed']
                                carrier_limpio = fila['Processed']
                                df_hoy = df_hoy.replace({carrier_real: carrier_limpio})
                except Exception as e :
                        print("Error cambiando valores de Carriers")
                        exit(1)


                print("Reemplazo de nombres de Carriers completado")
                time.sleep(5)


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
                print(f"Ocurrió un error inesperado: {e}")
                exit(1)

        """Realizar merge entre df_hoy y df_ayer utilizando 'Last PO #' como clave"""
        try:
                mapeo_notas = df_ayer.set_index('Last PO #')['CA NOTES']
                df_hoy['CA NOTES'] = df_hoy['Last PO #'].map(mapeo_notas)
        except Exception as e:
                print(f"Error al realizar el merge: {e}")
                exit(1)

        """FILTRO DE REGISTROS PARA ENVIAR CORREO"""

        try:    

                Condicion1 = df_hoy['CSR'] == CSR_name
                Condicion2 = df_hoy['CA NOTES'].isna()
                df_correos_enviar = df_hoy[(Condicion1) & (Condicion2)].copy()
                size = df_correos_enviar['Last PO #'].size
                print(f'Existen {size} correos por enviar')
                time.sleep(10)
        except Exception as e:
                print(f"Error al filtrar los registros para el correo: {e}")
                exit(1)

elif decision == 1:
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
                        df_correos_enviar = df_hoy[(Condicion1) & (Condicion2)].copy()
                        size = df_correos_enviar['Last PO #'].size
                        print(f'Existen {size} correos por enviar')
                        time.sleep(5)
                except Exception as e:
                        print(f"Error al filtrar los registros para el correo: {e}")
                        exit(1)

        
elif decision == 2:
        """FILTRO DE REGISTROS PARA ENVIAR CORREO"""
        try:    

                Condicion1 = df_hoy['CSR'] == CSR_name
                Condicion2 = df_hoy['CA NOTES'].isna()
                df_correos_enviar = df_hoy[(Condicion1) & (Condicion2)].copy()
                size = df_correos_enviar['Last PO #'].size
                print(f'Existen {size} correos por enviar')
                time.sleep(5)
        except Exception as e:
                print(f"Error al filtrar los registros para el correo: {e}")
                exit(1)

"""Armador de correos"""
for indice, fila in df_correos_enviar.iterrows():

        Client_Name = fila['Client Name']
        MC_number = fila['MCNumber']
        PO_Number = fila['Last PO #']
        Email_Destinatario = fila['Debtor Email Address']       
        correo = outlook.CreateItem(0)

        """Se convierte los valores de las celdas en minusuculas"""
        
        nota_attention = str(fila['Attention Note']).lower()
        nota_warning = str(fila['Warning Note']).lower()
        correos_no_touch = str(fila['Client Name']).lower()
        correos_englandlogistics = str(fila['Debtor Email Address']).lower()

        """ciclos para verificar el valor de las celdas no tiene palabras de atencion."""
        tiene_nombre_no_touch = any(palabra in correos_no_touch for palabra in nombres_no_touch)
        tiene_palabra_attention = any(palabra in nota_attention for palabra in Palabras_Buscar)
        tiene_palabra_warning = any(palabra in nota_warning for palabra in Palabras_Buscar)
        
        correo_nulo = pd.isna(fila['Debtor Email Address'])
        

        print(f"Procesando correo para {Client_Name} , con la carga {PO_Number}")

        if tiene_palabra_attention or tiene_palabra_warning:

                df_correos_enviar.at[indice, 'CA NOTES'] = 'Checar Manualmente'
                print(f"Correo no enviado, checar manualmente")
                #time.sleep(5)
                continue  # Si se encuentra alguna de las palabras en las notas, se omite el envío del correo para este registro
                
        elif correo_nulo :
                df_correos_enviar.at[indice, 'CA NOTES'] = 'NO EMAIL'
                #time.sleep(5)
                continue

        elif 'debtors@englandlogistics.com' in correos_englandlogistics or 'paperwork@englandlogistics.com' in correos_englandlogistics or '@englandlogistics.com' in correos_englandlogistics:
                df_correos_enviar.at[indice, 'CA NOTES'] = 'Correo de England Logistics, checar manualmente'
                #time.sleep(5)
                print(f"Correo no enviado, checar manualmente")
                continue
        
        elif tiene_nombre_no_touch:
                df_correos_enviar.at[indice, 'CA NOTES'] = 'Cliente no touch'
                #time.sleep(5)
                print(f"No touch cliente, correo no enviado")
                continue
        try:
                correo.To = Email_Destinatario
                correo.Subject = f"PLEASE REPLY NOA confirmation required {Client_Name} // MC {MC_number} // Load {PO_Number}"
                correo.Body = f"""
Good morning

                                
Attached you'll find our NOA for the carrier. Please confirm when received.

                        
Thank you and have a great day! 🙂

                        
*Please note if you are seeing this message again, it is because we have not received confirmation."""
        
                ruta_attachment = os.path.join(onedrive_path,"Documents","NOA",Client_Name + " - NOA.pdf" )
                correo.Attachments.Add(ruta_attachment)
                correo.importance = 2
                correo.Send()
                #correo.Display()
                if '@noa.triumphpay.com' in correos_englandlogistics: 
                        df_correos_enviar.at[indice, 'CA NOTES'] = f'{fecha_hoy} SENT to TP'
                        time.sleep(2)
                        print(f"Correo enviado a Triumphpay")
                else:
                        df_correos_enviar.at[indice, 'CA NOTES'] = f'{fecha_hoy} SENT'
                        time.sleep(2)
                        #input("Press Enter to continue to the next email...")

              
        except pywintypes.com_error as e:
                print(f"attachment not found for {Client_Name}: {e}")
                df_correos_enviar.at[indice, 'CA NOTES'] = 'File not found'
                continue
        except Exception as e:
                print(f"Error al enviar el correo para {Client_Name}: {e}")
                df_correos_enviar.at[indice, 'CA NOTES'] = 'Error al enviar el correo'
                continue

df_final = df_hoy['Last PO #'].map(df_correos_enviar.set_index('Last PO #')['CA NOTES'])
df_hoy['CA NOTES'] = df_final.combine_first(df_hoy['CA NOTES'])
Guardar_archivo_excel(df_hoy,ruta_excel_hoy_procesado)
shutil.copy2(ruta_excel_hoy_procesado,ruta_descargar_excel_hoy_procesado)