"""SETUPPP"""
import win32com.client as win32
import pandas as pd
import openpyxl
import os
from datetime import date
import shutil
import time
import tkinter as tk
from tkinter import filedialog
import pywintypes

"""Info general"""

date_name = date.weekday(date.today())
outlook = win32.Dispatch('outlook.application')

directorio_actual = os.path.dirname(__file__)
hoy = date.today()
fecha_hoy = hoy.strftime("%m-%d-%y")
fecha_ayer = hoy - pd.Timedelta(days=1)
fecha_ayer = fecha_ayer.strftime("%m-%d-%y")

"""Fechas en Fin de Semana"""
"""Fechas cuando es Domingo"""
fecha_viernes_en_domingo = hoy - pd.Timedelta(days=2)
fecha_jueves_en_domingo = fecha_viernes_en_domingo - pd.Timedelta(days=1)
fecha_jueves_en_domingo = fecha_jueves_en_domingo.strftime("%m-%d-%y")
fecha_viernes_en_domingo = fecha_viernes_en_domingo.strftime("%m-%d-%y")

"""Fechas cuando es sabado"""
fecha_viernes_en_sabado = hoy - pd.Timedelta(days=1)
fecha_jueves_en_sabado = fecha_viernes_en_sabado - pd.Timedelta(days=1)
fecha_jueves_en_sabado = fecha_jueves_en_sabado.strftime("%m-%d-%y")
fecha_viernes_en_sabado = fecha_viernes_en_sabado.strftime("%m-%d-%y")

"""Fechas cuando es lunes"""
fecha_viernes_en_lunes = hoy - pd.Timedelta(days=3)
fecha_viernes_en_lunes = fecha_viernes_en_lunes.strftime("%m-%d-%y")

"""RUTAS"""

"""Rutas Generales"""
Ruta_nube_hoy = r"C:\\Users\\Jorge\\OneDrive - C.R. England\Documents\\"+ fecha_hoy +"_ECS_Factoring_NOARecdDate.xlsx"
Ruta_nube_ayer = r"C:\\Users\\Jorge\\OneDrive - C.R. England\Documents\\"+ fecha_ayer +"_ECS_Factoring_NOARecdDate.xlsx"

ruta_df_hoy = os.path.join(directorio_actual,"..","data",fecha_hoy+"_ECS_Factoring_NOARecdDate" ".xlsx")
ruta_df_hoy_procesado = os.path.join(directorio_actual,"..","data",fecha_hoy+"_ECS_Factoring_NOARecdDate_Procesado" ".xlsx")

ruta_df_ayer = os.path.join(directorio_actual,"..","data",fecha_ayer+"_ECS_Factoring_NOARecdDate" ".xlsx")
ruta_df_ayer_procesado = os.path.join(directorio_actual,"..","data",fecha_ayer+"_ECS_Factoring_NOARecdDate_Procesado" ".xlsx")

"""Rutas en domingo"""
Ruta_nube_hoy_domingo = r"C:\\Users\\cokek\\OneDrive - C.R. England\\Documents\\"+ fecha_viernes_en_domingo+"_ECS_Factoring_NOARecdDate.xlsx"
Ruta_nube_ayer_domingo = r"C:\\Users\\cokek\\OneDrive - C.R. England\\Documents\\"+ fecha_jueves_en_domingo+"_ECS_Factoring_NOARecdDate.xlsx"
ruta_df_hoy_domingo = os.path.join(directorio_actual,"..","data",fecha_viernes_en_domingo +"_ECS_Factoring_NOARecdDate" ".xlsx")
ruta_df_ayer_domingo = os.path.join(directorio_actual,"..","data",fecha_jueves_en_domingo +"_ECS_Factoring_NOARecdDate" ".xlsx")

"""Rutas en sábado"""
Ruta_nube_hoy_sabado = r"C:\\Users\\cokek\\OneDrive - C.R. England\\Documents\\"+ fecha_viernes_en_sabado+"_ECS_Factoring_NOARecdDate.xlsx"
Ruta_nube_ayer_sabado = r"C:\\Users\\cokek\\OneDrive - C.R. England\\Documents\\"+ fecha_jueves_en_sabado+"_ECS_Factoring_NOARecdDate.xlsx"
ruta_df_hoy_sabado = os.path.join(directorio_actual,"..","data",fecha_viernes_en_sabado +"_ECS_Factoring_NOARecdDate" ".xlsx")
ruta_df_ayer_sabado = os.path.join(directorio_actual,"..","data",fecha_jueves_en_sabado +"_ECS_Factoring_NOARecdDate" ".xlsx")

"""Rutas en Lunes"""
Ruta_nube_ayer_lunes = r"C:\\Users\\Jorge\\OneDrive - C.R. England\Documents\\"+ fecha_viernes_en_lunes+"_ECS_Factoring_NOARecdDate.xlsx"
ruta_df_ayer_lunes = os.path.join(directorio_actual,"..","data",fecha_viernes_en_lunes +"_ECS_Factoring_NOARecdDate" ".xlsx")

"""Miscelaneo"""

ruta_nuevo_archivo = os.path.join(directorio_actual,"..","RESULTADOS",fecha_hoy+ "_Archivo_Procesado.xlsx" )
Ruta_nube_archivo_procesado = r"C:\\Users\\Jorge\\OneDrive - C.R. England\\Documents\\"+ fecha_hoy +"_ECS_Factoring_NOARecdDate_Procesado.xlsx"
df_columns_to_drop = ['NOA Rec Date', 'NOA Assigment', 'NOA Sent', 'NOA Sent User','Debtor NOA Document','CSR.1', 'Office','Notice', 'Notice Contact Email','NOA Entered Date', 'Last Inv Date', 'Last Inv #', 'First Inv Date', 'Relationship Age', 'First Funded', 'Client Age', 'Funded Balance', 'Non Funded Balance']
Palabras_Buscar = ["noa", "doesnt verify", "@noa.triumphpay", "web", "triumphpay", "website", "use triumph pay", "triumph", "tp", "epay"]

"""Definicion de funciones"""
def carga_archivos_excel_semanal( intentos=0, max_intentos=3):
        try:
                df_hoy = pd.read_excel(ruta_df_hoy)
                df_ayer = pd.read_excel(ruta_df_ayer)
                print("Archivos Excel leídos correctamente.")
                return True, df_hoy, df_ayer, 0  # Si la lectura es exitosa, retornamos True para salir del ciclo
        except Exception as e:
                print(f"Error al leer los archivos Excel: {e}")
                print(f"Copiando el archivo desde la nube a la ruta local, intentelo de nuevo.")
                intentos += 1
                shutil.copy2(Ruta_nube_hoy, ruta_df_hoy)
                shutil.copy2(Ruta_nube_ayer, ruta_df_ayer)
                return False, None, None, intentos  # Si ocurre un error, retornamos False para intentar nuevamente

def carga_archivos_excel_domingo(intentos=0, max_intentos=3):
        try:
                df_hoy = pd.read_excel(ruta_df_hoy_domingo)
                df_ayer = pd.read_excel(ruta_df_ayer_domingo)
                print("Archivos Excel leídos correctamente.")
                return True, df_hoy, df_ayer, 0
        except Exception as e:
                print(f"Error al leer los archivos Excel: {e}")
                print(f"Copiando el archivo desde la nube a la ruta local, intentelo de nuevo.")
                intentos += 1
                shutil.copy2(Ruta_nube_hoy_domingo, ruta_df_hoy_domingo)
                shutil.copy2(Ruta_nube_ayer_domingo, ruta_df_ayer_domingo)
                return False, None, None, intentos
def carga_archivos_excel_sabado(intentos=0, max_intentos=3):
        try:
                df_hoy = pd.read_excel(ruta_df_hoy_sabado)
                df_ayer = pd.read_excel(ruta_df_ayer_sabado)
                print("Archivos Excel leídos correctamente.")
                return True, df_hoy, df_ayer, 0
        except Exception as e:
                print(f"Error al leer los archivos Excel: {e}")
                print(f"Copiando el archivo desde la nube a la ruta local, intentelo de nuevo.")
                intentos += 1
                shutil.copy2(Ruta_nube_hoy_sabado, ruta_df_hoy_sabado)
                shutil.copy2(Ruta_nube_ayer_sabado, ruta_df_ayer_sabado)
                return False, None, None, intentos
def carga_archivos_excel_lunes(intentos=0, max_intentos=3):
        try:
                df_hoy = pd.read_excel(ruta_df_hoy)
                df_ayer = pd.read_excel(ruta_df_ayer_lunes)
                print("Archivos Excel leídos correctamente.")
                return True, df_hoy, df_ayer, 0
        except Exception as e:
                print(f"Error al leer los archivos Excel: {e}")
                print(f"Copiando el archivo desde la nube a la ruta local, intentelo de nuevo.")
                intentos += 1
                shutil.copy2(Ruta_nube_hoy, ruta_df_hoy)
                shutil.copy2(Ruta_nube_ayer_lunes, ruta_df_ayer_lunes)
                return False, None, None, intentos

def Guardar_archivo_excel(nombre_dataframe, ruta_guardar):
   nombre_dataframe.to_excel(ruta_guardar, index=False)      

Palabras_reemplazo = {
        'Golden Moon Transport Inc //Only Wire or RTP' : 'Golden Moon Transport Inc',
        'Gholia Logistics Inc (NO ACH FEE)':'Gholia Logistics Inc',
        'Classic Freight Transportation Inc (NO ACH FEE)':'Classic Freight Transportation Inc',
        'Dhillon Bros Carrier LLC (RTP Only)':'Dhillon Bros Carrier LLC',
        'MBM Global Inc (RTP Only)':'MBM Global Inc',
        'Debtors@englandlogistics.com':'',
        'Debtors@Englandlogistics.com':'',
        'debtors@englandlogistics.com':'',
        'paperwork@englandlogistics.com':'',
}

columnas_a_modificar = ['Debtor Email Address', 'Attention Note', 'Warning Note']

Adaptación_deudores = {

        'AM Trans Expedite Inc - IL':['noa@fusiontransport.com', '', ''],
        'AM Trans Expedite Inc - NJ':['noa@fusiontransport.com', '', ''],
        'Challenger Motor Freight Inc':['vendorsetup@challenger.com', '', ''],
        'Challenger Logistics Inc - ON':['vendorsetup@challenger.com', '', ''],
        'US Logistics LLC':['uslogisticsllc@noa.triumphpay.com', '', ''],
        'Transplace Texas LP - KY':['transplacestuttgart@noa.triumphpay.com', '', ''],
        'Paul Logistics Inc':['Paullog@noa.triumphpay.com', '', ''],
        'Drover Logistics Corp':['team@shipdrover.com;ap@shipdrover.com', '', ''],
        'Ari Logistics LLC dba Action Enterprise Logistics':['ari.ap@actn.com; action@noa.triumphpay.com', '', ''],
        'ITS National LLC dba ITS National':['inquiries@its4logistics.com; cody.chapman@its4logistics.com', '', ''],
        'Transportation One - IL':['paymentstatus@transportationone.com; billing@transportationone.com', '', ''],
        'Scotlynn USA Division Inc':['paystatus@scotlynn.com','',''],
        'Spot Freight Inc':['paymentstatus@spotinc.com','',''],
        'Frontline Logistics':['cvella@frontlinelogistics.com','',''],
        'Jear Logistics LLC':['jear@noa.triumphpay.com','','',],
        'Transend Logistics LLC -IL':['dmims@transendlogistics.com','',''],
        'Mohawk Global Logistics':['mstoddard@mohawkglobal.com','', ''],
        'R2 LOGISTICS':['carrierservices@r2logistics.com','EPAY', ''],
        'Venture Connect / Transcorr Ntl Logistics LLC':['carrierinv@venturelogistics.com; paystatus@venturelogistics.com','', ''],
        'Andover Logistics LLC':['mark@andoverlogistics.com;accounting@andoverlogistics.com','EPAY', ''],
        'Synchrogistics LLC - NC':['accounting@synchrogistics.com', '', ''],
        'Circle Logistics Inc - IN':['PAYSTATUS@CIRCLEDELIVERS.COM','', ''],
        'Unlimited Logistics LLC':['noa@unlimitedlogistics.com','', ''],
        'Estrella Dispatch Inc':['Estrelladispatch@noa.triumphpay.com; ap@estrelladispatch.com','',''],
        'R2 LOGISTICS':['cs@r2logistics.com','', ''],
        'Titanium American Logistics':['TITANIUM@noa.triumphpay.com','', ''],
        'Transend Logistics LLC -IL':['dmims@transendlogistics.com; accounting@transendlogistics.com','', ''],
}
"""     
        
        
        '':['','', ''],
        '':['','', ''],
        '':['','', ''],
        '':['','', ''],
        '':['','', ''],
        '':['','', ''],
        '':['','', ''],
        '':['','', ''],
        '':['','', ''],
                
        {'accounting@journeyfreight.com':'apinquiries@journeyfreight.com'},"""
        

"""def seleccionar_archivo_excel():
    # Abre una ventana emergente para que el usuario seleccione un archivo de Excel
    # 1. Crear una ventana oculta de Tkinter (para que no salga una ventana vacía de fondo)
    root = tk.Tk()
    root.withdraw()
    
    # 2. Configurar y abrir el explorador de archivos
    ruta_archivo = filedialog.askopenfilename(
        title="Selecciona el reporte de Excel",
        filetypes=[("Archivos de Excel", "*.xlsx *.xls")] # Solo permite ver archivos de Excel
    )
    
    # 3. Retornar la ruta elegida (si el usuario cancela, devolverá una cadena vacía)
    return ruta_archivo"""


"""Empieza el código"""

"""Lectura de archivos Excel"""

if date_name == 6: #Domingo
        Estado_carga, df_hoy, df_ayer, intentos = carga_archivos_excel_domingo()
        try:
                while Estado_carga == False and intentos < 3:
                        print("Esperando 3 segundos antes de intentar cargar los archivos nuevamente...")
                        time.sleep(3)  # Espera 3 segundos antes de intentar cargar los archivos nuevamente
                        Estado_carga, df_hoy, df_ayer, intentos = carga_archivos_excel_domingo()
        except Exception as e:
                        print(f"NO SE ENCUENTRAN LOS ARCHIVOS A CARGAR, POR FAVOR INTENTE MAS TARDE.")
elif date_name == 5: #Sabado
        Estado_carga, df_hoy, df_ayer, intentos = carga_archivos_excel_sabado()
        try:
                while Estado_carga == False and intentos < 3:
                        print("Esperando 3 segundos antes de intentar cargar los archivos nuevamente...")
                        time.sleep(3)  # Espera 3 segundos antes de intentar cargar los archivos nuevamente
                        Estado_carga, df_hoy, df_ayer, intentos = carga_archivos_excel_sabado()
        except Exception as e:
                        print(f"NO SE ENCUENTRAN LOS ARCHIVOS A CARGAR, POR FAVOR INTENTE MAS TARDE.")
elif date_name == 0: #Lunes
        Estado_carga, df_hoy, df_ayer, intentos = carga_archivos_excel_lunes()
        try:
                while Estado_carga == False and intentos < 3:
                        print("Esperando 3 segundos antes de intentar cargar los archivos nuevamente...")
                        time.sleep(3)  # Espera 3 segundos antes de intentar cargar los archivos nuevamente
                        Estado_carga, df_hoy, df_ayer, intentos = carga_archivos_excel_lunes()
        except Exception as e:
                        print(f"NO SE ENCUENTRAN LOS ARCHIVOS A CARGAR, POR FAVOR INTENTE MAS TARDE.")
else:
        Estado_carga, df_hoy, df_ayer, intentos = carga_archivos_excel_semanal()
        try:
                while Estado_carga == False and intentos < 3:
                        print("Esperando 3 segundos antes de intentar cargar los archivos nuevamente...")
                        time.sleep(3)  # Espera 3 segundos antes de intentar cargar los archivos nuevamente
                        Estado_carga, df_hoy, df_ayer, intentos = carga_archivos_excel_semanal()
        except Exception as e:
                print(f"NO SE ENCUENTRAN LOS ARCHIVOS A CARGAR, POR FAVOR INTENTE MAS TARDE.")


""""Adaptacion del df_hoy: eliminar columnas, insertar columna 'CA NOTES' """

try:
        df_hoy.drop(columns=df_columns_to_drop, axis='columns', inplace=True, errors='ignore')
        print("Columnas eliminadas")
        try:
                df_hoy.insert(9, 'CA NOTES', '')
                print("Columna 'CA NOTES' insertada")
        except Exception as e :
                 print("La columna 'CA NOTES' ya existe. Continuando con el proceso...")

        print("Creacion del nuevo archivo con la columna 'CA NOTES' sin errores")
        df_hoy.replace(Palabras_reemplazo, inplace=True, regex=False)
        

        print("Reemplazo de texto específico completado")
        time.sleep(5)
        df_ayer = df_ayer.drop_duplicates(subset=['Last PO #'], keep='last')
        df_hoy = df_hoy.drop_duplicates(subset=['Last PO #'], keep='last')
        try:
                for deudor, lista_valores in Adaptación_deudores.items():
                        #print(f'Remplazo de texto para deudor {deudor}')
                        df_hoy.loc[df_hoy['Debtor Name'] == deudor, columnas_a_modificar] = lista_valores

        except Exception as e :
                print("Error cambiando valores de debtors")
                exit(1)

        Guardar_archivo_excel(df_hoy,ruta_nuevo_archivo)
        #print(df_hoy.head()) 

except Exception as e:
        print(f"Ocurrió un error inesperado: {e}")
        exit(1)
        





"""Realizar merge entre df_hoy y df_ayer utilizando 'Last PO #' como clave"""
try:
        mapeo_notas = df_ayer.set_index('Last PO #')['CA NOTES']
        df_hoy['CA NOTES'] = df_hoy['Last PO #'].map(mapeo_notas)
        Guardar_archivo_excel(df_hoy,ruta_df_hoy)

except Exception as e:
        print(f"Error al realizar el merge: {e}")
        exit(1)





"""FILTRO DE REGISTROS PARA ENVIAR CORREO"""

try:    
        Condicion1 = df_hoy['CSR'] == 'VGUERRERO'
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
        

        """ciclos para verificar el valor de las celdas no tiene palabras de atencion."""

        tiene_palabra_attention = any(palabra in nota_attention for palabra in Palabras_Buscar)
        tiene_palabra_warning = any(palabra in nota_warning for palabra in Palabras_Buscar)
        correo_nulo = pd.isna(fila['Debtor Email Address'])

        print(f"Procesando correo para {Client_Name} , con la carga {PO_Number}")

        if tiene_palabra_attention or tiene_palabra_warning:

                df_correos_enviar.at[indice, 'CA NOTES'] = 'Checar Manualmente'
                print(f"Correo no enviado, checar manualmente")
                
                continue  # Si se encuentra alguna de las palabras en las notas, se omite el envío del correo para este registro
                
        elif correo_nulo :
                df_correos_enviar.at[indice, 'CA NOTES'] = 'NO EMAIL'
                continue


        try:
                correo.To = Email_Destinatario
                correo.Subject = f"PLEASE REPLY NOA confirmation required {Client_Name} // MC {MC_number} // Load {PO_Number}"
                correo.Body = f"""
Good morning

                                
Attached you'll find our NOA for the carrier. Please confirm when received.

                        
Thank you and have a great day! 🙂

                        
*Please note if you are seeing this message again, it is because we have not received confirmation."""
        
                ruta_attachment = os.path.join(directorio_actual,"..","attachments",Client_Name + " - NOA.pdf" )
                correo.Attachments.Add(ruta_attachment)
                correo.Send()
                #correo.Display()
                df_correos_enviar.at[indice, 'CA NOTES'] = f'{fecha_hoy} SENT'
                #time.sleep(3)
                #input("Press Enter to continue to the next email...")


               

        except pywintypes.com_error as e:
                print(f"attachment not found for {Client_Name}: {e}")
                df_correos_enviar.at[indice, 'CA NOTES'] = f'File not found'
                continue
        except Exception as e:
                print(f"Error al enviar el correo para {Client_Name}: {e}")
                df_correos_enviar.at[indice, 'CA NOTES'] = f'Error al enviar el correo'
                continue

df_final = df_hoy['Last PO #'].map(df_correos_enviar.set_index('Last PO #')['CA NOTES'])
df_hoy['CA NOTES'] = df_final.combine_first(df_hoy['CA NOTES'])
print(df_hoy)
Guardar_archivo_excel(df_hoy,ruta_df_hoy_procesado)
shutil.copy2(ruta_df_hoy_procesado,Ruta_nube_archivo_procesado)
