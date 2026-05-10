import win32com.client as win32
import os
directorio_actual = os.path.dirname(__file__)


outlook = win32.Dispatch('outlook.application')
excel = win32.Dispatch('Excel.Application')

ruta_data = os.path.join(directorio_actual,"..","data","ECS_Factoring_NOARecdDate 5-01-26.xlsx")

libro = excel.workbooks.open(ruta_data)

hoja = libro.Sheets('ECS_Factoring_NOARecdDate')

dato_correo = hoja.Cells(456,8).Value
dato_PO = hoja.Cells(456,10).Value
dato_MC = hoja.Cells(456,3).Value
dato_Carrier = hoja.Cells(456,2).Value

nombre_archivo = f"{dato_Carrier} - NOA.pdf"
ruta_adjunto = os.path.join(directorio_actual,"..","attachtments",nombre_archivo)
ruta_absoluta_adjunto = os.path.abspath(ruta_adjunto)
print("la ruta es: " + ruta_absoluta_adjunto)


correo = outlook.CreateItem(0)
correo.To = dato_correo
correo.Subject = 'PLEASE REPLY NOA confirmation required Carrier ' + dato_Carrier + ' // MC ' + dato_MC + ' // Load ' + dato_PO
correo.HTMLBody = 'Good morning,<br><br><br>Attached you will find our NOA for the carrier. Please confirm when received.<br><br><br>Thank you and have a great day! 🙂<br><br><br>*Please note if you are seeing this message again, it is because we have not received confirmation.'


correo.Attachments.Add(ruta_absoluta_adjunto)


correo.display()