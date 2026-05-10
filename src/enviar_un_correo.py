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
hoja.Cells.Replace(What="Golden Moon Transport Inc //Only Wire or RTP",Replacement="Golden Moon Transport Inc", 
                LookAt=2, 
                SearchOrder=1, 
                MatchCase=False)

#Todas 

for fila in range(456, ultima_fila + 1):
    
    
    if hoja.Rows(fila).EntireRow.Hidden:
        print(f"Saltando la fila {fila} porque está oculta...")
        continue         
    
    dato_nombre = str(hoja.Cells(fila, 2).Value).strip()
    dato_correo = "cokeklo.delarosa@gmail.com"
    #dato_correo = str(hoja.Cells(fila, 8).Value).strip()
    dato_PO = str(hoja.Cells(fila, 10).Value).strip()
    dato_MC = str(hoja.Cells(fila, 3).Value).strip()
    print(f"Enviando correo a: {dato_nombre} ({dato_correo})")
    correo = outlook.CreateItem(0)
    correo.To = dato_correo
    correo.Subject = 'PLEASE REPLY NOA confirmation required Carrier ' + dato_nombre + ' // MC ' + dato_MC + ' // Load ' + dato_PO
    correo.HTMLBody = 'Good morning,<br><br><br>Attached you will find our NOA for the carrier.,' \
    ' Please confirm when received.<br><br><br>Thank you and have a great day!,' \
    ' 🙂<br><br><br>*Please note if you are seeing this message again, it is because we have not received confirmation.'

    ruta_adjunto = os.path.join(directorio_actual,"..","attachments",f"{dato_nombre}* - NOA.pdf")
    coincidencias = glob.glob(ruta_adjunto)

    if coincidencias:
        # SI SE ENCUENTRA EL ARCHIVO
        ruta_final = coincidencias[0]
        correo.Attachments.Add(ruta_final)
        print(f"✅ Archivo encontrado para {dato_nombre}. Enviando automáticamente...")
        correo.Send() 
    else:
        # NO SE ENCUENTRA EL ARCHIVO
        print(f"⚠️ No se encontró archivo para {dato_nombre}. Mostrando correo para revisión manual.")
        correo.Display() # Abre la ventana de Outlook para que tú lo manejes
    

# Pausa de seguridad
    confirmacion = input("Presiona ENTER para enviar este correo o escribe 's' para saltarlo: ")
    
    if confirmacion.lower() == 's':
        print(f"Fila {fila} saltada por el usuario.")
        continue