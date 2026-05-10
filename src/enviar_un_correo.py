import win32com.client as win32
import os
import glob
directorio_actual = os.path.dirname(__file__)
outlook = win32.Dispatch('outlook.application')
excel = win32.Dispatch('Excel.Application')
ruta_data = os.path.join(directorio_actual,"..","data","ECS_Factoring_NOARecdDate 5-01-26.xlsx")
libro = excel.Workbooks.Open(ruta_data)
hoja = libro.Sheets('ECS_Factoring_NOARecdDate')
ultima_fila = hoja.Cells(hoja.Rows.Count, 1).End(-4162).Row

#Todas 

for fila in range(456, ultima_fila + 1):
    
    # 3. Verificamos si la fila está oculta (por un filtro o manualmente)
    if hoja.Rows(fila).EntireRow.Hidden:
        print(f"Saltando la fila {fila} porque está oculta...")
        continue         
    
    dato_nombre = str(hoja.Cells(fila, 2).Value).strip()
    dato_correo = str(hoja.Cells(fila, 8).Value).strip()
    dato_PO = str(hoja.Cells(fila, 10).Value).strip()
    dato_MC = str(hoja.Cells(fila, 3).Value).strip()
    print(f"Enviando correo a: {dato_nombre} ({dato_correo})")
    

#dato_correo = hoja.Cells(456,8).Value
#dato_PO = hoja.Cells(456,10).Value
#dato_MC = hoja.Cells(456,3).Value
#dato_Carrier = hoja.Cells(456,2).Value


    ruta_adjunto = os.path.join(directorio_actual,"..","attachments",f"{dato_nombre}* - NOA.pdf")
    #ruta_absoluta_adjunto = os.path.abspath(ruta_adjunto)
    #print("la ruta es:" + ruta_absoluta_adjunto)

    coincidencias = glob.glob(ruta_adjunto)

    if coincidencias:
    # 3. Extraemos la RUTA COMPLETA del primer archivo encontrado
    # Aquí ya NO hay asterisco, es el nombre real: "C:/Carpeta/Jorge De La Rosa.pdf"
        ruta_archivo_real = coincidencias[0]
    
    # 4. Si solo quieres el nombre (para un print o para el asunto del correo)
        nombre_limpio_sin_ruta = os.path.basename(ruta_archivo_real)
    
        print(f"Archivo encontrado correctamente: {nombre_limpio_sin_ruta}")
    
    # USAR 'ruta_archivo_real' para adjuntar en Outlook
    # mail.Attachments.Add(ruta_archivo_real)
    else:
        print(f"No se encontró nada que empezara con {dato_nombre}")




    correo = outlook.CreateItem(0)
    correo.To = dato_correo
    correo.Subject = 'PLEASE REPLY NOA confirmation required Carrier ' + dato_nombre + ' // MC ' + dato_MC + ' // Load ' + dato_PO
    correo.HTMLBody = 'Good morning,<br><br><br>Attached you will find our NOA for the carrier. Please confirm when received.<br><br><br>Thank you and have a great day! 🙂<br><br><br>*Please note if you are seeing this message again, it is because we have not received confirmation.'
    correo.Attachments.Add(ruta_archivo_real)
    print(f"Adjuntando el archivo: {ruta_archivo_real}")
    correo.display()
    

# Pausa de seguridad
    confirmacion = input("Presiona ENTER para enviar este correo o escribe 's' para saltarlo: ")
    
    if confirmacion.lower() == 's':
        print(f"Fila {fila} saltada por el usuario.")
        continue