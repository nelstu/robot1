import os
import time
import pandas as pd

from openpyxl import load_workbook
import glob
import smtplib
from email.message import EmailMessage
from selenium import webdriver
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.firefox.service import Service
# Ruta del geckodriver y carpeta de descarga
ruta_geckodriver = "/Users/nelsonstuardo/Downloads/geckodriver"
carpeta_descarga = "/Users/nelsonstuardo/Downloads/sii"  # crea esta carpeta antes



# Crear opciones de Firefox
options = Options()
profile = webdriver.FirefoxProfile()

# Configurar perfil para descargas automáticas
profile.set_preference("browser.download.folderList", 2)
profile.set_preference("browser.download.dir", carpeta_descarga)
profile.set_preference("browser.helperApps.neverAsk.saveToDisk", "application/pdf,application/octet-stream,text/csv,application/vnd.ms-excel")
profile.set_preference("pdfjs.disabled", True)

# Asignar el perfil a las opciones
options.profile = profile

# Crear el navegador
service = Service(ruta_geckodriver)
driver = webdriver.Firefox(service=service, options=options)

# Ir a la página
driver.get("https://www.sii.cl/servicios_online/1039-3256.html")

# Esperar que cargue bien
time.sleep(5)

# Hacer clic en el enlace deseado (ejemplo: por texto visible)
enlace = driver.find_element(By.LINK_TEXT, "Ingresar al Registro de Compras y Ventas")  # cambia el texto aquí
enlace.click()

time.sleep(5)
# Buscar el campo de RUT por su id y escribir un valor
campo_rut = driver.find_element(By.ID, "rutcntr")
campo_rut.send_keys("93915000-5")  # Puedes cambiar este valor
time.sleep(2)

campo_clave = driver.find_element(By.ID, "clave")
campo_clave.send_keys("profesor2")  # Puedes cambiar este valor

# Esperar un poco para simular usuario
time.sleep(1)

# Hacer clic en el botón Ingresar
boton_ingresar = driver.find_element(By.ID, "bt_ingresar")
boton_ingresar.click()




time.sleep(5)
# Hacer clic en el botón "Consultar" (usando XPath con texto)
boton_consultar = driver.find_element(By.XPATH, "//button[contains(text(), 'Consultar')]")
boton_consultar.click()

# Esperar para ver el resultado
time.sleep(4)
# Esperar y hacer clic usando el atributo ui-sref
elemento_venta = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.CSS_SELECTOR, 'a[ui-sref="venta"]'))
)
elemento_venta.click()


time.sleep(5)

# Esperar y hacer clic en el botón "Descargar Detalles"
boton_descargar = WebDriverWait(driver, 15).until(
    EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Descargar Detalles')]"))
)
boton_descargar.click()


# --- Buscar el archivo descargado más reciente ---
def archivo_mas_reciente(directorio):
    lista_archivos = [
        f for f in glob.glob(os.path.join(directorio, "*"))
        if os.path.isfile(f)
    ]
    if not lista_archivos:
        return None
    return max(lista_archivos, key=os.path.getctime)


def enviar_archivo_por_email(destinatario, archivo_path):
    import smtplib
    from email.message import EmailMessage
    import os

    remitente = "envios@iopa.cl"
    password = "$NSloteria2015$"

    msg = EmailMessage()
    msg['Subject'] = 'Archivo de Rechazos SII'
    msg['From'] = remitente
    msg['To'] = destinatario
    msg.set_content('Te adjunto el archivo descargado.')

    with open(archivo_path, 'rb') as f:
        contenido = f.read()
        nombre_archivo = os.path.basename(archivo_path)
        msg.add_attachment(contenido, maintype='application', subtype='octet-stream', filename=nombre_archivo)

    with smtplib.SMTP_SSL('mail.iopa.cl', 465) as smtp:
        smtp.login(remitente, password)
        smtp.send_message(msg)
        print("Correo enviado con éxito.")


# Buscar el archivo más reciente
import glob
import os




carpeta_descarga = "/Users/nelsonstuardo/Downloads/sii"
archivo = archivo_mas_reciente(carpeta_descarga)

# Llamar a la función si hay archivo
if archivo:

    # Leer el CSV
    df = pd.read_csv(archivo, sep=";", dtype=str)
    df.columns = df.columns.str.strip()  # limpiar nombres de columna

    # Guardar como Excel (todos los registros)
    archivo_excel = "archivo_completo_con_filtro.xlsx"
    df.to_excel(archivo_excel, index=False)

    # Activar filtros en el archivo Excel
    wb = load_workbook(archivo_excel)
    ws = wb.active
    ws.auto_filter.ref = ws.dimensions  # Aplica el filtro a toda la tabla
    wb.save(archivo_excel)

    time.sleep(5)
    enviar_archivo_por_email("nstuardo@gmail.com", archivo_excel)
else:
    print("No se encontró ningún archivo para enviar.")
