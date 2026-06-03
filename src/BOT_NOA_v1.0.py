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

onedrive_path = os.path.expanduser('~\\OneDrive - C.R. England\\Documents')

date_name = date.weekday(date.today())
outlook = win32.Dispatch('outlook.application')
directorio_actual = os.path.dirname(__file__)
hoy = date.today()
fecha_hoy = hoy.strftime("%m-%d-%y")
fecha_ayer = hoy - pd.Timedelta(days=1)
fecha_ayer = fecha_ayer.strftime("%m-%d-%y")

terminacion_archivo_procesado = "_ECS_Factoring_NOARecdDate_Procesado.xlsx"
terminacion_archivo_sin_procesar = "_ECS_Factoring_NOARecdDate.xlsx"

"""Rutas Generales"""

ruta_excel_hoy_sin_procesar = os.path.join(directorio_actual,"..","data",fecha_hoy +terminacion_archivo_sin_procesar)
ruta_excel_hoy_procesado = os.path.join(directorio_actual,"..","data",fecha_hoy+terminacion_archivo_procesado)
Ruta_nube_archivo_procesado = onedrive_path + '\\' + fecha_hoy + terminacion_archivo_procesado

""""Dataframes y variables globales"""

df_columns_to_drop = ['NOA Rec Date', 'NOA Assigment', 'NOA Sent', 'NOA Sent User','Debtor NOA Document','CSR.1', 'Office','Notice', 'Notice Contact Email','NOA Entered Date', 'Last Inv Date', 'Last Inv #', 'First Inv Date', 'Relationship Age', 'First Funded', 'Client Age', 'Funded Balance', 'Non Funded Balance']
Palabras_Buscar = ["noa", "doesnt verify", "@noa.triumphpay", "web", "triumphpay", "website", "use triumph pay", "triumph", "tp", "epay"]
nombres_no_touch = ["avensis energy services, llc", "metro parcel & freight inc", "mvk transport corporation", "cmd energy, llc"]


"""Definicion de funciones"""

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
        'TAN Transport Inc - REACTIVATION':'TAN Transport Inc',
        'paperwrok@englandlogistics.com':'',
        'Mehreen Enterprises LTD (US Currency) (WIRE ONLY)':'Mehreen Enterprises LTD',
        'MVI Transport (dba Main Venture Investing LLC) ACH preference':'MVI Transport (dba Main Venture Investing LLC)',
        '24X7 Riders Transport Inc (CAD CURRENCY)':'24X7 Riders Transport Inc',
        'Canadian Transport Network Inc (CAD Currency)':'Canadian Transport Network Inc',
        'HS Carrier LLC - RTP ONLY':'HS Carrier LLC',
}

columnas_a_modificar = ['Debtor Email Address', 'Attention Note', 'Warning Note']

while True:
        try:
                CSR_INPUT = input("Porfavor seleccione el CSR que desea procesar: 1.VGUERRERO 2.MPALMER 3.SPAREDES")

                if CSR_INPUT == '1':
                        CSR_name = 'VGUERRERO'
                        break
                elif CSR_INPUT == '2':
                        CSR_name = 'MPALMER'
                        break
                elif CSR_INPUT == '3':
                        CSR_name = 'SPAREDES'
                        break                        
                else: 
                        raise ValueError("Entrada no válida. Por favor, seleccione 1, 2 o 3.")
                
        except ValueError:
                print("Entrada no válida. Por favor, seleccione 1, 2 o 3.")

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
        'Beemac Logistics':['beemaclogistics@noa.triumphpay.com','', ''],
        'Deploy Solutions (dba of CWH LLC)':['marzenae@deploysolutionsgroup.com; ap@deploysolutionsgroup.com','', ''],
        'Heniff Logistics LLC':['eblackman@heniff.com','', ''],
        'Hirschbach Transportation Services Inc - IA':['NOA@hirschbach.com','', ''],
        'Smartway Transportation / Overland Park, KS':['smartwaytransportationllc@noa.triumphpay.com','', ''],
        'Taimen Transport LLC':['ap@taimentransport.com; ship@taimentransport.com','', ''],
        'Traffic Tech Inc -IL':['ap@traffictech.com; noa@traffictech.com','', ''],
        'West Michigan Transport LLC dba Windmill Transport':['windmilltransport@noa.triumphpay.com','', ''],
        'England Logistics Inc':['','', ''],
        'Specialized Transport  Co':['','', ''],
        'Chief Logistics, LLC':['accounting@chieflogistics.com; claudia@chieflogistics.com','', ''],
        'Navajo Express Inc':['carrierpaystatus@navajo.com','', ''],
        'Ruan Transport Corporation':['carrierbilling@ruan.com; freightaudit@ruan.com; carrierprocurement@ruan.com','', ''],
        'Rubicon Transportation LLC':['payments@rubicontransportation.com; rubicon@noa.triumphpay.com','', ''],
        'Ship Amino LLC dba OGRE-S':['ap@shipamino.com; shipamino@noa.triumphpay.com; invoice@shipamino.com','', ''],
        'RJS Logistics, Inc.':['billing@rjslogistics.com; payment@rjslogistics.com','', ''],
        'Best Logistic Services aka Reich Logistics services Inc':['NOA@shipwithbest.com','', ''],
        'Longship (dba of Quality Logistics LLC)':['longship@noa.triumphpay.com; accounting@longship.us','', ''],
        'Emerge Transportation dba of Emergetech Inc':['Noa@emergemarket.com','', ''],
        'First Call Logistics, LLC.':['	firstcallrn@noa.triumphpay.com;accounting@gofclogistics.com','', ''],
        'FLS TRANSPORTATION SVCS':['flstransportation@noa.triumphpay.com; ap@flstransport.com','', ''],
        'HD Shipping Solutions LLC':['paperwork@hdships.com','', ''],
        'Integrity Express Logistics':['intxlog@noa.triumphpay.com','', ''],
        'Landstar Ranger Inc':['carriermaintenance@landstar.com','', ''],
        'Motus Freight LLC':['MotusFreight@noa.triumphpay.com','', ''],
        'Raven Cargo Inc':['ravencargo@noa.triumphpay.com;accounting@raven-cargo.com','', ''],
        'Sage Freight LLC':['sagefreight@noa.triumphpay.com','', ''],
        'Trinity Logistics Inc - DE':['trinitylogistics@noa.triumphpay.com','', ''],
        'Brock LLC dba Brock Transportation LLC':['accounting@brockweb.com','', ''],
        'RXO / XPO - NC':['rxo@noa.triumphpay.com','', ''],
        'Ryan Transportation Service, Inc.':['ryanap@ryantrans.com','', ''],
        'Stonearch Logistics, LLC.':['stonearchlogistics@noa.triumphpay.com; acct@stonearchlogistics.com','', ''],
        'Select Transport Partners LLC':['ap@goselect.com','', ''],
        'King of Freight LLC':['carrierrelations@kingoffreight.com','', ''],
        'Logistic Dynamics LLC':['CarrierDev@ShipLDi.com','', ''],
        'Hope Transportation':['hopetransportation@noa.triumphpay.com; payables@hopetrans.com','', ''],
        'OTR TRANSPORTATION':['accountspayable@loadotr.com; ap@loadotr.com','', ''],
        'Triple T Transport Inc':['AP@tripleTTransport.com','', ''],
        'Pivot Supply Chain Solutions':['accounting@loadpivot.com','', ''],
        'RJ LOGISTICS ASSETS, LLC':['RJLogistics@noa.triumphpay.com','', ''],
        'Reinsfelder Inc':['billing@reinsfelder.com','', ''],
        'Neon Logistics LLC':['neon@noa.triumphpay.com','', ''],
        'Core Logistics Brokerage LLC':['corelogistics@noa.triumphpay.com','', ''],
        'Edge Logistics LLC':['paperwork@edgelogistics.com; edgelogistics@noa.triumphpay.com; accounting@edgelogistics.com','', ''],
        'Emerge Transportation dba of Emergetech Inc':['Loaddocs@emergemarket.com; noa@emergemarket.com','', ''],
        'Freight All Kinds, LLC':['ap@fakinc.com; assignments@fakinc.com; paperwork@fakinc.com','', ''],
        'Landstar Transportation Logistics, Inc':['carriermaintenance@landstar.com','', ''],
        'United Transportation Services Inc':['UnitedTransportation@noa.triumphpay.com; accounting@shiputs.net','', ''],
        'Wel Companies Inc':['wel.logisticsap@WELCOMPANIES.com','', ''],
        'XTL Logistics USA Inc':['cbtinc@noa.triumphpay.com;invoices.us@xtl.com','', ''],
        'Nolan Transportation Group LLC':['pod@ntgfreight.com;payables@ntgfreight.com','', ''],
        'Jarrett Logistics Systems':['accounting@gojarrettlogistics.com','', ''],
        'TAB LLC dba of TA Brokerage LLC':['TABACCOUNTING@ARTUREXPRESS.COM','', ''],
        "Leonard's Express Inc":['ap@leonardsexpress.com;leitrips@leonardsexpress.com','',''],
        'Southwest Logistics Management-Oklahoma':['ap@sw-logitics;chthompson@sw-logistics.com;cbellamy@sw-logistics.com','', ''],
        'Premier Global Transportation LLC':['accounting@premierglobaltransportation.com','', ''],
        'Tin Goose Logistics LLC':['ap@tingooselogistics.com','', ''],
        'Value Logistics LLC - TX':['getpaid@valuetruckaz.com','', ''],
        'Pathmark Transportation':['apinquiries@pathmarktrans.com','', ''],
        'NFI Logistics LLC (dba of National Freight Truck Lines Inc)':['RELAYINVOICES@nfiindustries.com; nfi@noa.triumphpay.com; carrier.relations@nfiindustries.com','', ''],
        'Native American Worldwide Logistics LLC':['freightbills@nalww.com; Krista.monroe@nalww.com','', ''],
        'MLM Supply Chain LLC':['wschneider@mlmtrans.com; kristy.larson@mlmtrucking.com','', ''],
        'MountainMovers Transportation & Logistics':['acct@mountainmoverstl.com;status@mountainmoverstl.com','', ''],
        'Martins Refrigerated Express Inc':['dedel@martinmilk.com ; BILLINGMRE@martinmilk.com','', ''],
        'High Tide Logistics LLC - IL':[' hightide@noa.triumphpay.com; ap@hightidelogistics.com','', ''],
        'UCW Logistics':['UCW@noa.triumphpay.com','', ''],
        'HTL Freight Heritage Trucking Matchmaker Logistics':['htlfreight@noa.triumphpay.com;dkalejr@bellsouth.net;billing@htlfreight.com','', ''],
        'Schmieding Produce / HC Schmieding Produce':['ap@Schmieding.com','', ''],
        'Standard Transportation Services, Inc.':['ap@standardtransinc.com; mtilton@stdtrans.com; ap@stdtrans.com','', ''],
        'Fuel Transport Inc.':['acctspayable@fueltransport.com; billing@fueltransport.com','',
                                ''],
        'Corporate Traffic (FL)':['carrierinvoices@corporatetraffic.com; corporatetraffic@noa.triumphpay.com','', ''],
        'Ally Logistics LLC':['docs@allylogistics.com; AP@allylogistics.com ;allylogistics@noa.triumphpay.com','', ''],
        'Colossal Transport Solutions, LLC':['carrierinvoicing@colossaltransport.com','', ''],
        'CTS Logistics Solutions Texas':['accounting@ctsls-usa.com','', ''],
        'Quality Freight Logistics,Inc':['qualityfreight@noa.triumphpay.com; info@qflteam.com','', ''],
        'Open Road Transportation':['openroad@noa.triumphpay.com; noa@openroadtrans.com','', ''],
        'BBI Logistics LLC':['invoices@bbilogistics.com;paymentstatus@bbilogistics.com;ap@bbilogistics.com','', ''],
        'Universal Freight Systems Inc':['invoicesubmission@ufsystems.com; universalfreight@noa.triumphpay.com','', ''],
        'Archer Cargo LLC':['accounting@archercargo.net','', ''],
        'Five Star Trucking Ltd':['invoices@fivestartrucking.com;ap@fivestartrucking.com','', ''],
        'Flock Freight Inc':['documents@flockfreight.com','', ''],
        'Johanson Transportation Service':['factoringNOA@johansontrans.com','', ''],
        'Kingsgate Logistics Inc':['accounting@kingsgatelogistics.com','', ''],
        'Mode Transportation LLC - Dallas':['factornotice@modetransportation.com','', ''],
        'Reliable Transportation Solutions':['ap@relyonrts.com','', ''],
        'ROCK CITY LOGISTICS, LLC':['rockcitylogistics@noa.triumphpay.com','', ''],
        'SPC Transport Co.':['billing@spctran.com;Dhartford@spctran.com','', ''],
        'Unicron Logistics Solutions':['unicronlogistics@gmail.com','', ''],
        'Drover Logistics Group LLC':['ap@shipdrover.com','', ''],

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
        

               
        {'accounting@journeyfreight.com':'apinquiries@journeyfreight.com'},"""
        

df_hoy, df_ayer, decision = CA.carga_archivos_excel()

if decision == True:
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

                Guardar_archivo_excel(df_hoy,ruta_excel_hoy_sin_procesar)
                #print(df_hoy.head()) 

        except Exception as e:
                print(f"Ocurrió un error inesperado: {e}")
                exit(1)

        """Realizar merge entre df_hoy y df_ayer utilizando 'Last PO #' como clave"""
        try:
                mapeo_notas = df_ayer.set_index('Last PO #')['CA NOTES']
                df_hoy['CA NOTES'] = df_hoy['Last PO #'].map(mapeo_notas)
                Guardar_archivo_excel(df_hoy,ruta_excel_hoy_sin_procesar)

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
                time.sleep(5)
        except Exception as e:
                print(f"Error al filtrar los registros para el correo: {e}")
                exit(1)

        
else:
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

        elif 'debtors@englandlogistics.com' in correos_englandlogistics or 'paperwork@englandlogistics.com' in correos_englandlogistics:
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
        
                ruta_attachment = os.path.join(directorio_actual,"..","attachments",Client_Name + " - NOA.pdf" )
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
shutil.copy2(ruta_excel_hoy_procesado,Ruta_nube_archivo_procesado)
