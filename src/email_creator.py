import win32com.client as win32
import os
import pywintypes
import time
import pandas as pd
outlook = win32.Dispatch('outlook.application')
Palabras_Buscar = ["noa", "doesnt verify", "@noa.triumphpay", "web", "triumphpay", "website", "use triumph pay", "triumph", "tp", "epay"]
nombres_no_touch = ["avensis energy services, llc", "cmd energy, llc", "inlog inc","bestway transfer & storage inc","elite lv express llc","freight gutu broker llc","hustler logistics llc","iraheta 2 bros inc","ivester's transportation, inc","metro parcel & freight inc","tampa heavy haul llc","topline trucking inc","werden logistics corp", "mvk transport corporation", ]


def enviar_correos(df_correos_enviar, fecha_hoy, onedrive_path):

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
            correos = str(fila['Debtor Email Address']).lower()

            """ciclos para verificar el valor de las celdas no tiene palabras de atencion."""
            tiene_nombre_no_touch = any(palabra in correos_no_touch for palabra in nombres_no_touch)
            tiene_palabra_attention = any(palabra in nota_attention for palabra in Palabras_Buscar)
            tiene_palabra_warning = any(palabra in nota_warning for palabra in Palabras_Buscar)
            debtor_not_processed = not 'changed on debtors file' in nota_attention
            correo_nulo = pd.isna(fila['Debtor Email Address'])
            

            print(f"Procesando correo para {Client_Name} , con la carga {PO_Number}")

            if tiene_palabra_attention or tiene_palabra_warning and debtor_not_processed:

                    df_correos_enviar.at[indice, 'CA NOTES'] = 'Checar Manualmente'
                    print(f"Correo no enviado, checar manualmente")
                    #time.sleep(5)
                    continue  # Si se encuentra alguna de las palabras en las notas, se omite el envío del correo para este registro
                    
            elif correo_nulo :
                    df_correos_enviar.at[indice, 'CA NOTES'] = 'NO EMAIL'
                    #time.sleep(5)
                    continue

            elif 'debtors@englandlogistics.com' in correos or 'paperwork@englandlogistics.com' in correos or '@englandlogistics.com' in correos:
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
                    #correo.Send()
                    #correo.Display()
                    if '@noa.triumphpay.com' in correos: 
                            df_correos_enviar.at[indice, 'CA NOTES'] = f'{fecha_hoy} SENT to TP'
                            #time.sleep(2)
                            print(f"Correo enviado a Triumphpay")
                    if 'website' in nota_warning and not debtor_not_processed:
                            df_correos_enviar.at[indice, 'CA NOTES'] = f'{fecha_hoy} SENT to WEB'
                            #time.sleep(2)
                            print(f"Correo enviado a WEB")
                    else:
                            df_correos_enviar.at[indice, 'CA NOTES'] = f'{fecha_hoy} SENT'
                            #time.sleep(2)
                    for col_intentos in ['1st attempt','2nd attempt','3rd attempt']:
                        valor_raw = df_correos_enviar.at[indice, col_intentos]
                        valor_columna_intento =str(valor_raw).strip() 
                        if pd.isna(valor_raw) or valor_columna_intento in ['', 'nan', 'None']:
                            df_correos_enviar.at[indice, col_intentos] = f'{fecha_hoy} SENT to {Email_Destinatario}'

                            #print("Usando loc:", df_correos_enviar.loc[indice, col_intentos])
                            break
                    #time.sleep(2)
                            #input("Press Enter to continue to the next email...")
                    

                
            except pywintypes.com_error as e:
                    print(f"attachment not found for {Client_Name}: {e}")
                    df_correos_enviar.at[indice, 'CA NOTES'] = 'File not found'
                    continue
            except Exception as e:
                    print(f"Error al enviar el correo para {Client_Name}: {e}")
                    df_correos_enviar.at[indice, 'CA NOTES'] = 'Error al enviar el correo'
                    continue