import os
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Constants: ajusta estas rutas según tu proyecto
CHROME_PROFILE_DIR = os.path.join(os.getcwd(), "chrome_selenium_profile")
CHROMEDRIVER_PATH = os.path.join(os.path.dirname(__file__), "driver", "chromedriver.exe")

def iniciar_navegador(profile_dir=CHROME_PROFILE_DIR, driver_path=CHROMEDRIVER_PATH):
    """
    Inicia una instancia de Chrome con un perfil persistente.
    """
    opts = webdriver.ChromeOptions()
    opts.add_argument(f"--user-data-dir={profile_dir}")
    opts.add_argument("--start-maximized")
    return webdriver.Chrome(service=Service(driver_path), options=opts)

def esperar_whatsapp(driver, timeout=120):
    """
    Espera hasta que la interfaz de WhatsApp Web esté lista para usarse.
    """
    WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located((By.XPATH,
            '//div[@contenteditable="true" and contains(@data-tab,"3")]'))
    )

def buscar_contacto(driver, nombre, timeout=30):
    """
    Busca el chat de un contacto por título EXACTO y hace click en él.
    """
    # 1) Esperar y localizar la caja de búsqueda
    caja = WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located((By.XPATH,
            '//div[@contenteditable="true" and contains(@data-tab,"3")]'))
    )
    caja.click()
    caja.clear()
    caja.send_keys(nombre)
    time.sleep(2)  # tiempo para que carguen los resultados dinámicos

    # 2) Esperar hasta que el span con @title exacto sea clickable y hacer click
    xpath_contacto = f'//span[@title="{nombre}"]'
    contacto_elem = WebDriverWait(driver, timeout).until(
        EC.element_to_be_clickable((By.XPATH, xpath_contacto))
    )
    contacto_elem.click()

def enviar_texto(driver, texto, timeout=15):
    """
    Escribe y envía texto en el chat activo.
    """
    caja = WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located((By.XPATH,
            '//footer//div[@contenteditable="true"]'))
    )
    caja.click()
    caja.send_keys(texto)

def enviar_imagen(driver, ruta_archivo, timeout=10):
    """
    Adjunta y envía una imagen en el chat activo.
    """
    # Clic en el botón de adjuntar
    plus_btn = WebDriverWait(driver, timeout).until(
        EC.element_to_be_clickable((By.XPATH,
            '//button[@aria-label="Adjuntar" or @title="Adjuntar"]'))
    )
    plus_btn.click()

    # Cargar el archivo
    file_input = WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located((By.XPATH,
            '//input[@type="file" and contains(@accept,"image")]'))
    )
    file_input.send_keys(ruta_archivo)

    # Enviar la imagen
    send_btn = WebDriverWait(driver, timeout + 5).until(
        EC.element_to_be_clickable((By.XPATH,
            '//div[@role="button"][@aria-label="Enviar"]'))
    )
    time.sleep(1)
    send_btn.click()

def enviar_captura_whatsapp(driver, contacto, ruta_imagen, mensaje_texto=""):
    """
    Orquesta el envío de un mensaje y una imagen al contacto dado.
    """
    driver.get("https://web.whatsapp.com")
    esperar_whatsapp(driver)
    buscar_contacto(driver, contacto)
    if mensaje_texto:
        enviar_texto(driver, mensaje_texto + "\n")
    enviar_imagen(driver, ruta_imagen)
    # opcional: esperar unos segundos antes de cerrar
    time.sleep(2)