from selenium import webdriver
from selenium.webdriver.firefox.service import Service
import time



from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


# Ruta al geckodriver
ruta_geckodriver = "/Users/nelsonstuardo/Desktop/geckodriver"
service = Service(ruta_geckodriver)

# Iniciar navegador
driver = webdriver.Firefox(service=service)

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

# Esperar para ver el resultado

time.sleep(2)
# Hacer clic en el botón "Consultar" (usando XPath con texto)
boton_consultar = driver.find_element(By.XPATH, "//button[contains(text(), 'Consultar')]")
boton_consultar.click()
time.sleep(2)

# Esperar y hacer clic en "Descargar Detalles"
WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Descargar Detalles')]"))
).click()



# Cerrar el navegador
#driver.quit()

