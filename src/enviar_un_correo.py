"""SETUPPP"""

import win32com.client as win32
import pandas as pd
import openpyxl
import os
from datetime import date
import shutil
import time


"""Info general"""
outlook = win32.Dispatch('outlook.application')
correo = outlook.CreateItem(0)
directorio_actual = os.path.dirname(__file__)
hoy = date.today()
fecha_formateada = hoy.strftime("%m-%d-%y")
fecha_ayer = hoy - pd.Timedelta(days=1)
fecha_ayer = fecha_ayer.strftime("%m-%d-%y")
df_columns_to_drop = ['NOA Rec Date', 'NOA Assigment', 'NOA Sent', 'NOA Sent User','Debtor NOA Document','CSR.1', 'Office','Notice', 'Notice Contact Email','NOA Entered Date', 'Last Inv Date', 'Last Inv #', 'First Inv Date', 'Relationship Age', 'First Funded', 'Client Age', 'Funded Balance', 'Non Funded Balance']
Ruta_nube = r"C:\\Users\\cokek\\OneDrive - C.R. England\\Documents\\"+ fecha_formateada +"_ECS_Factoring_NOARecdDate.xlsx"
ruta_attachment = os.path.join(directorio_actual,"..","attachments",fecha_formateada+"_ECS_Factoring_NOARecdDate" ".xlsx")
Palabras_Buscar = ["noa", "DOESNT VERIFY", "@noa.triumphpay", "WEB", "triumphpay", "WEBSITE", "USE TRIUMPH PAY", "triumph", "TP",]


"""Ruta Archivo Nuevo"""

#ruta_df1 = os.path.join(directorio_actual,"..","data",fecha_formateada+"_ECS_Factoring_NOARecdDate" ".xlsx")
ruta_df1 = os.path.join(directorio_actual,"..","data",fecha_ayer+"_ECS_Factoring_NOARecdDate" ".xlsx")
#print(ruta_df1)

"""Ruta Archivo Anterior"""

ruta_df2 = os.path.join(directorio_actual,"..","data",fecha_ayer+"_ECS_Factoring_NOARecdDate" ".xlsx")
#print(ruta_df2)

"""Definicion de funciones"""
def carga_archivos_excel( intentos=0, max_intentos=3):
        try:
                df_hoy = pd.read_excel(ruta_df1)
                df_ayer = pd.read_excel(ruta_df2)
                print("Archivos Excel leídos correctamente.")
                return True, df_hoy, df_ayer, 0  # Si la lectura es exitosa, retornamos True para salir del ciclo
        except Exception as e:
                print(f"Error al leer los archivos Excel: {e}")
                print(f"Copiando el archivo desde la nube a la ruta local, intentelo de nuevo.")
                intentos += 1
                shutil.copy2(Ruta_nube, ruta_df1)
                return False, None, None, intentos  # Si ocurre un error, retornamos False para intentar nuevamente
        except FileNotFoundError as e:
                print(f"Archivo no encontrado: {e}, º")
                intentos += 1
                return False, None, None, intentos

def Guardar_archivo_excel():
   df_hoy.to_excel(ruta_df1, index=False)      





"""Lectura de archivos Excel"""

Estado_carga, df_hoy, df_ayer, intentos = carga_archivos_excel()

"""Reintentos para cargar los archivos Excel, con una pausa de 3 segundos entre cada intento"""

while Estado_carga == False and intentos < 3:
        print("Esperando 3 segundos antes de intentar cargar los archivos nuevamente...")
        time.sleep(3)  # Espera 3 segundos antes de intentar cargar los archivos nuevamente

        Estado_carga, df_hoy, df_ayer, intentos = carga_archivos_excel()

""""Adaptacion del df_hoy: eliminar columnas, insertar columna 'CA NOTES' y guardar el archivo para luego realizar el merge con df_ayer"""

try:
        df_hoy.drop(columns=df_columns_to_drop, axis='columns', inplace=True, errors='ignore')
        print("Columnas eliminadas")

        df_hoy.insert(9, 'CA NOTES', '')
        print("Columna 'CA NOTES' insertada")
        Guardar_archivo_excel()
        print("Creacion del nuevo archivo con la columna 'CA NOTES' sin errores")

        #print(df_hoy.head())
except ValueError as e:
        print("La columna 'CA NOTES' ya existe. Continuando con el proceso...")
        
except Exception as e:
        print(f"Ocurrió un error inesperado: {e}")
        exit(1)
finally:
        """Buscar y reemplazar texto específico en la hoja de Excel"""

        df_hoy.replace(to_replace="Golden Moon Transport Inc //Only Wire or RTP", value="Golden Moon Transport Inc", inplace=True, regex=False)
        df_hoy.replace(to_replace="MBM Global Inc (RTP Only)", value="MBM Global Inc", inplace=True, regex=False)
        df_hoy.replace(to_replace="Gholia Logistics Inc (NO ACH FEE)", value="Gholia Logistics Inc", inplace=True, regex=False)
        df_hoy.replace(to_replace="Debtors@englandlogistics.com", value=" ", inplace=True, regex=False)
        df_hoy.replace(to_replace="paperwork@englandlogistics.com", value=" ", inplace=True, regex=False)
        df_hoy.replace(to_replace="invoices@amtransexpedite.com;Pods@fusiontransport.com;invoices@amtransexpedite.com", value="noa@fusiontransport.com", inplace=True, regex=False)       
        df_hoy.replace(to_replace="invoices@amtransexpedite.com;pod@amtransexpedite.com;pods@fusiontransport.com;payables@fusiontransport.com", value="noa@fusiontransport.com", inplace=True, regex=False)
        df_hoy.replace(to_replace="PS:carrierinquiries@challenger.com", value="noa@fusiontransport.com", inplace=True, regex=False)
        df_hoy.replace(to_replace="quickpay@uslfreight.com PS-ap@uslfreight.com TRIUM", value="NOAs:uslogisticsllc@noa.triumphpay.com", inplace=True, regex=False)
        df_hoy.replace(to_replace="PS Triumph: (469) 312-7222", value="Paullog@noa.triumphpay.com", inplace=True, regex=False)
        df_hoy.replace(to_replace="NOA transplacetx@noa.triumphpay.com", value="transplacestuttgart@noa.triumphpay.com", inplace=True, regex=False)
        df_hoy.replace(to_replace="Classic Freight Transportation Inc (NO ACH FEE)", value="Classic Freight Transportation Inc", inplace=True, regex=False)
        df_hoy.replace(to_replace="jd@droverlogisticsgroup.com; ap@shipdrover.com", value="ap@shipdrover.com", inplace=True, regex=False)
        df_hoy.replace(to_replace="ari.accounting@actn.com", value="ari.ap@actn.com; action@noa.triumphpay.com", inplace=True, regex=False)
        df_hoy.replace(to_replace="carriers@its4logistics.com;paperwork@its4logistics.com", value="inquiries@its4logistics.com; cody.chapman@its4logistics.com", inplace=True, regex=False)
        df_hoy.replace(to_replace="accounting@journeyfreight.com", value="apinquiries@journeyfreight.com", inplace=True, regex=False)
        df_hoy.replace(to_replace="Dhillon Bros Carrier LLC (RTP Only)", value="Dhillon Bros Carrier LLC", inplace=True, regex=False)
        df_hoy.replace(to_replace="ap@transportationone.com", value="paymentstatus@transportationone.com; billing@transportationone.com", inplace=True, regex=False)
        df_hoy.replace(to_replace="usa-accounting@scotlynn.com", value="paystatus@scotlynn.com", inplace=True, regex=False)
        df_hoy.replace(to_replace="; spotbilling@spotinc.com", value="paymentstatus@spotinc.com", inplace=True, regex=False)
        

        Guardar_archivo_excel()

        print("Reemplazo de texto específico completado y archivo guardado.")
        df_ayer_unico = df_ayer.drop_duplicates(subset=['Last PO #'], keep='last')
        

"""Realizar merge entre df_hoy y df_ayer utilizando 'Last PO #' como clave"""

try:
        mapeo_notas = df_ayer_unico.set_index('Last PO #')['CA NOTES']
        df_hoy['CA NOTES'] = df_hoy['Last PO #'].map(mapeo_notas)
        Guardar_archivo_excel()

except Exception as e:
        print(f"Error al realizar el merge: {e}")
        exit(1)


"""FILTRO DE REGISTROS PARA ENVIAR CORREO: CSR = VGUERRERO y CA NOTES = #N/A"""

try:
        Condicion1 = df_hoy['CSR'] == 'VGUERRERO'
        Condicion2 = df_hoy['CA NOTES'].isna()
        df_final = df_hoy[(Condicion1) & (Condicion2)]
        size = df_final['Last PO #'].size
        print(size)
except Exception as e:
        print(f"Error al filtrar los registros para el correo: {e}")
        exit(1)

"""Armador de correos"""

for indice, fila in df_final.iterrows():

        Client_Name = fila['Client Name']
        MC_number = fila['MCNumber']
        PO_Number = fila['Last PO #']
        Email_Destinatario = fila['Debtor Email Address']       


        nota_attention = str(fila['Attention Note']).lower()
        nota_warning = str(fila['Warning Note']).lower()

        tiene_palabra_attention = any(palabra in nota_attention for palabra in Palabras_Buscar)
        tiene_palabra_warning = any(palabra in nota_warning for palabra in Palabras_Buscar)

        if tiene_palabra_attention or tiene_palabra_warning:

                fila['CA NOTES'] = "Revisar notas de atención o advertencia"
                
                Guardar_archivo_excel()

                continue  # Si se encuentra alguna de las palabras en las notas, se omite el envío del correo para este registro

        try:
                correo.To = Email_Destinatario
                correo.Subject = f"PLEASE REPLY NOA confirmation required Carrier {Client_Name} // MC {MC_number} // Load {PO_Number}"
                correo.Body = f"""
Good morning

                                
Attached you'll find our NOA for the carrier. Please confirm when received.

                        
Thank you and have a great day! 🙂

                        
*Please note if you are seeing this message again, it is because we have not received confirmation."""
        
                ruta_attachment = os.path.join(directorio_actual,"..","attachments")
                correo.Attachments.Add(ruta_attachment)
                correo.display()
                fila['CA NOTES'] = f"{fecha_formateada} SENT"
                input("Press Enter to continue to the next email...")                       

        except FileNotFoundError as e:
                print(f"attachment not found for {Client_Name}: {e}")
                fila['CA NOTES'] = "File attachment not found"
                continue  # Si el archivo adjunto no se encuentra, se omite el envío del correo para este registro y se continúa con el siguiente
        except Exception as e:
                print(f"Error al enviar el correo para {Client_Name}: {e}")
                fila['CA NOTES'] = "Error sending email"
                continue
               
       

                
       






            
    


