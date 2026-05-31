
import os
import pandas as pd
from datetime import date, timedelta
import shutil
import time

onedrive_path = os.path.expanduser('~\\OneDrive - C.R. England\\Documents')
date_name = date.weekday(date.today())
directorio_actual = os.path.dirname(__file__)
hoy = date.today()
terminacion_archivo_procesado = "_ECS_Factoring_NOARecdDate_Procesado.xlsx"
terminacion_archivo_sin_procesar = "_ECS_Factoring_NOARecdDate.xlsx"

"""Rutas de Martes a Viernes"""
#ruta_descargar_excel_ayer = os.path.join(onedrive_path, fecha_ayer + terminacion_archivo_procesado)
#ruta_excel_ayer_sin_procesar = os.path.join(directorio_actual,"..","data",fecha_ayer +terminacion_archivo_sin_procesar)



def carga_archivos_excel():
    date_name = date.weekday(date.today())
    hoy = date.today()
    if date_name == 6: #Domingo
        fecha_hoy = hoy-timedelta(days=2)
        fecha_ayer = hoy-timedelta(days=3)
        print("Hoy es Domingo, se cargaran los archivos de Viernes")
        time.sleep(5)
    elif date_name == 5: #Sabado
        fecha_hoy = hoy - timedelta(days=1)
        fecha_ayer = hoy - timedelta(days=2)
        print("Hoy es Sabado, se cargaran los archivos de Viernes")
        time.sleep(5)
    elif date_name == 0: #Lunes
        fecha_hoy = hoy
        fecha_ayer = hoy - timedelta(days=3)
        print("Hoy es Lunes, se cargaran los archivos de Lunes y Viernes")
        time.sleep(5)
    else:
            fecha_hoy = hoy
            fecha_ayer = hoy - timedelta(days=1)
            print("Hoy es un dia entre Martes y Viernes, se cargaran los archivos de hoy y ayer")
            time.sleep(5)

    fecha_hoy = fecha_hoy.strftime("%m-%d-%y")
    fecha_ayer = fecha_ayer.strftime("%m-%d-%y")

    ruta_descargar_excel_hoy_sin_procesar = os.path.join(onedrive_path, fecha_hoy + terminacion_archivo_sin_procesar)
    ruta_descargar_excel_hoy_procesado = os.path.join(onedrive_path, fecha_hoy + terminacion_archivo_procesado)
    ruta_excel_hoy_sin_procesar = os.path.join(directorio_actual,"..","data",fecha_hoy +terminacion_archivo_sin_procesar)
    ruta_excel_hoy_procesado = os.path.join(directorio_actual,"..","data",fecha_hoy+terminacion_archivo_procesado)
    ruta_excel_ayer_procesado = os.path.join(directorio_actual,"..","data",fecha_ayer + terminacion_archivo_procesado)
    

    if not os.path.exists(ruta_excel_hoy_procesado):#tiene el ultimo archivo procesado de correos?
        if not os.path.exists(ruta_descargar_excel_hoy_procesado):#tiene el ultimo archivo procesado de correos en onedrive?
            if os.path.exists(ruta_excel_hoy_sin_procesar):#Tienes el archivo de NOA de hoy sin procesar?
                if os.path.exists(ruta_excel_ayer_procesado):#Tienes el archivo de ayer procesado?
                    df_hoy = pd.read_excel(ruta_excel_hoy_sin_procesar)
                    df_ayer = pd.read_excel(ruta_excel_ayer_procesado)
                    bandera_carga = True
                    print("Archivo sin procesar de hoy y archivo procesado de ayer encontrados en carpeta data, cargando los archivos...")
                    time.sleep(5)
                    return df_hoy, df_ayer, bandera_carga
                else:#Buscar el ultimo archivo procesado de correos en onedrive
                    print("No se encuentra el archivo procesado de ayer, por favor revise la carpeta data y agregalo manualmente")
                    time.sleep(5)
                    exit(1)
            else:#Tienes el archivo de ayer procesado?
                if os.path.exists(ruta_excel_ayer_procesado):#Tienes el archivo de ayer procesado?
                    shutil.copy2(ruta_descargar_excel_hoy_sin_procesar, ruta_excel_hoy_sin_procesar)
                    df_hoy = pd.read_excel(ruta_excel_hoy_sin_procesar)
                    df_ayer = pd.read_excel(ruta_excel_ayer_procesado)
                    bandera_carga = True
                    print("Archivo sin procesar de hoy encontrado en OneDrive y archivo procesado de ayer encontrado en carpeta data, cargando los archivos...")
                    time.sleep(5)
                    return df_hoy, df_ayer, bandera_carga
                else:#Buscar el ultimo archivo procesado de correos en onedrive
                    shutil.copy2(ruta_descargar_excel_hoy_sin_procesar, ruta_excel_hoy_sin_procesar)           
                    print("No se encuentra el archivo procesado de ayer, por favor revise la carpeta data y agregalo manualmente")
                    exit(1)
        shutil.copy2(ruta_descargar_excel_hoy_procesado, ruta_excel_hoy_procesado)
        df_hoy = pd.read_excel(ruta_excel_hoy_procesado)
        bandera_carga = False
        print("Archivo procesado de hoy encontrado en OneDrive, cargando el archivo...")
        time.sleep(5)
        return df_hoy, None, bandera_carga
    else:
        print("Archivo procesado de hoy encontrado en carpeta data, verificando si es el más actualizado...")
        if os.path.exists(ruta_descargar_excel_hoy_procesado):
            if os.path.getmtime(ruta_descargar_excel_hoy_procesado) > os.path.getmtime(ruta_excel_hoy_procesado):
                shutil.copy2(ruta_descargar_excel_hoy_procesado, ruta_excel_hoy_procesado)
                df_hoy = pd.read_excel(ruta_excel_hoy_procesado)
                bandera_carga = False
                print("Archivo procesado de hoy encontrado en OneDrive es más actualizado que el de carpeta data, cargando el archivo de OneDrive...")
                time.sleep(5)
                return df_hoy, None, bandera_carga
            elif os.path.getmtime(ruta_descargar_excel_hoy_procesado) < os.path.getmtime(ruta_excel_hoy_procesado):
                df_hoy = pd.read_excel(ruta_excel_hoy_procesado)
                shutil.copy2(ruta_excel_hoy_procesado,ruta_descargar_excel_hoy_procesado)
                bandera_carga = False
                print("Archivo procesado de hoy encontrado en carpeta data es más actualizado que el de OneDrive, cargando el archivo de carpeta data...")
                time.sleep(5)
                return df_hoy, None, bandera_carga
        else:
            df_hoy = pd.read_excel(ruta_excel_hoy_procesado)
            bandera_carga = False
            time.sleep(5)
            return df_hoy, None, bandera_carga
