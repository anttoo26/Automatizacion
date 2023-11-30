import pandas as pd
from datetime import datetime
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication

# Conexión a la base de datos
import pyodbc
direccion_servidor = '10.108.32.12'
nombre_bd = 'LPA'
nombre_usuario = 'anto'
password = '&9X.J@Bd8yT9'
try:
    conexion = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=' + 
                              direccion_servidor+';DATABASE='+nombre_bd+';UID='+nombre_usuario+';PWD=' + password)
    print("Conexión Exitosa")
    
except Exception as e:
    print("Ocurrió un error al conectar a SQL Server: ", e)


# Ruta del archivo Excel original
archivo_original = r"LPA HERMOSILLO.xlsx"

# Leer datos desde la hoja "BD"
datos_originalBD = pd.read_excel(archivo_original, sheet_name="BD")

# Seleccionar el rango de celdas desde "I2" hasta "AS10000"
datos_seleccionadosBD = datos_originalBD.iloc[0:10000, 8:53]

# Crear un nuevo DataFrame para pegar los datos
nuevos_datosBD = pd.DataFrame()

# Pegar los datos seleccionados en el nuevo DataFrame
nuevos_datosBD = pd.concat([nuevos_datosBD, datos_seleccionadosBD], axis=1)

columnas_fecha = datos_originalBD.iloc[0:1000, 18:21]  

# Dar formato de fecha corta a las columnas
columnas_fecha = columnas_fecha.apply(lambda x: pd.to_datetime(x, errors='coerce', format='%d/%m/%Y'))

# Reemplazar las columnas originales con las nuevas columnas formateadas
datos_originalBD.iloc[0:1000, 18:21] = columnas_fecha

# Guardar el archivo CSV con la sintaxis deseada
fecha_hoy = datetime.now().strftime('%y.%m.%d')
nombre_archivo_csv = f"LPA {fecha_hoy} HMO.csv"
archivo_nuevo_csv =  nombre_archivo_csv
nuevos_datosBD.to_csv(archivo_nuevo_csv, index=False, header=False)

print(f"Archivo CSV guardado como {archivo_nuevo_csv}")


# Configurar detalles del correo electrónico
#si la persona tiene la verificacion de dos pasos activa no funcionara
smtp_server = 'smtp.office365.com'
smtp_port = 587
smtp_user = "anttoo-13@outlook.es"
smtp_password = ''

# Lista de destinatarios
destinatarios = ["anttoo-13@outlook.es"]

# Configurar el mensaje de correo
mensaje = MIMEMultipart()
mensaje['From'] = 'anttoo-13@outlook.es'
mensaje['To'] = ', '.join(destinatarios)
mensaje['Subject'] = 'R3IN02T106521 C07005'

# Cuerpo del correo
cuerpo_correo = MIMEText('Adjunto encontrarás el archivo CSV solicitado.')
mensaje.attach(cuerpo_correo)

# Adjuntar el archivo CSV al correo
with open(archivo_nuevo_csv, 'rb') as archivo:
    adjunto = MIMEApplication(archivo.read(), _subtype='csv')
    adjunto.add_header('Content-Disposition', f'attachment; filename={nombre_archivo_csv}')
    mensaje.attach(adjunto)

# Conectar al servidor SMTP y enviar el correo
with smtplib.SMTP(smtp_server, smtp_port) as server:
    server.starttls()
    server.login(smtp_user, smtp_password)
    server.sendmail(smtp_user, destinatarios, mensaje.as_string())

print(f"Correo electrónico enviado con el archivo adjunto a {', '.join(destinatarios)}")

#BackOrder

# Leer datos desde la hoja "BACK ORDER"
datos_originalBO = pd.read_excel(archivo_original, sheet_name="BACK ORDER")

# Copiamos toda la tabla de BACK ORDER
datos_seleccionadosBO = datos_originalBO.iloc[0:9, 0:15]

# Crear un nuevo DataFrame para pegar los datos
nuevos_datosBO = pd.DataFrame()

# Pegar los datos seleccionados en el nuevo DataFrame
nuevos_datosBO = pd.concat([nuevos_datosBO, datos_seleccionadosBO], axis=1)

# Guardar el archivo CSV con la sintaxis deseada
fecha_hoy = datetime.now().strftime('%y.%m.%d')
nombre_archivo_csv = f"BACK ORDER {fecha_hoy}.csv"
archivo_nuevo_csv =  nombre_archivo_csv
nuevos_datosBO.to_csv(archivo_nuevo_csv, index=False, header=False)

print(f"Archivo CSV guardado como {archivo_nuevo_csv}")




