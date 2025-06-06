import pygetwindow as gw
import pyautogui
import time
import os
import pytz
import traceback
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager
from pathlib import Path

# === CONFIGURACIÓN ===
CONTACTO = "+57 321 7877709"
POWERBI_TITULO = "JMRS.DB - VD DIARIO Y MENSUAL"
REGION = (180, 120, 960, 530)
CAPTURA_DIR = Path(__file__).parent / "Capturas"
CAPTURA_DIR.mkdir(exist_ok=True)

# === FUNCIONES ===
def buscar_ventana(titulo):
    print("[INFO] Buscando Power BI...")
    ventanas = gw.getWindowsWithTitle(titulo)
    if not ventanas:
        raise Exception("No se encontró Power BI abierto con el titulo especificado.")
    ventana = ventanas[0]
    ventana.activate()
    ventana.maximize()
    time.sleep(2)
    print("[OK] Ventana de Power BI activa.")
    return ventana

def tomar_captura(ventana, region):
    print("[INFO] Tomando captura...")
    captura = pyautogui.screenshot(region=(ventana.left + region[0], ventana.top + region[1], region[2], region[3]))
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    ruta = CAPTURA_DIR / f"captura_{timestamp}.png"
    captura.save(ruta)
    print("[OK] Captura guardada en:", ruta)
    return str(ruta)

def iniciar_navegador():
    print("[INFO] Iniciando Chrome...")
    options = webdriver.ChromeOptions()
    options.add_argument(f"--user-data-dir={os.getcwd()}/chrome_selenium_profile")
    options.add_argument("--start-maximized")
    service = Service(ChromeDriverManager().install())
    return webdriver.Chrome(service=service, options=options)

def esperar_whatsapp(driver):
    print("[INFO] Esperando que cargue WhatsApp Web...")
    WebDriverWait(driver, 120).until(
        lambda d: d.find_elements(By.XPATH, '//div[@contenteditable="true" and contains(@data-tab,"3")]')
    )
    print("[OK] WhatsApp Web está listo para buscar contactos.")

def buscar_contacto(driver, nombre):
    print("[INFO] Buscando contacto seleccionado...")
    search_box = WebDriverWait(driver, 15).until(
        EC.presence_of_element_located((By.XPATH, '//div[@contenteditable="true" and contains(@data-tab,"3")]'))
    )
    search_box.click()
    search_box.clear()
    search_box.send_keys(nombre)
    time.sleep(2)
    search_box.send_keys(Keys.DOWN)
    search_box.send_keys(Keys.ENTER)
    print("[OK] Chat abierto con coincidencia.")

def escribir_mensaje(driver, mensaje):
    print("[INFO] Escribiendo mensaje...")
    input_box = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, '//div[@contenteditable="true" and @data-tab="10"]'))
    )
    input_box.send_keys(mensaje)
    time.sleep(0.5)

def adjuntar_imagen(driver, ruta):
    print("[INFO] Adjuntando imagen...")
    plus_btn = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//button[@aria-label="Adjuntar" or @title="Adjuntar"]'))
    )
    plus_btn.click()
    file_input = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, '//input[@type="file" and @accept="image/*,video/mp4,video/3gpp,video/quicktime"]'))
    )
    file_input.send_keys(ruta)
    print("[OK] Imagen seleccionada.")

def enviar_mensaje(driver):
    print("[INFO] Enviando...")
    send_button = WebDriverWait(driver, 15).until(
        EC.element_to_be_clickable((By.XPATH, '//div[@role="button"][@aria-label="Enviar"]'))
    )
    time.sleep(1)
    send_button.click()

def obtener_hora_es():
    return datetime.now(pytz.timezone("Europe/Madrid")).strftime("%H:%M")

# === EJECUCIÓN PRINCIPAL ===
try:
    ventana = buscar_ventana(POWERBI_TITULO)
    ruta_captura = tomar_captura(ventana, REGION)

    driver = iniciar_navegador()
    driver.get("https://web.whatsapp.com")
    esperar_whatsapp(driver)

    buscar_contacto(driver, CONTACTO)
    hora = obtener_hora_es()
    escribir_mensaje(driver, hora)
    adjuntar_imagen(driver, ruta_captura)
    enviar_mensaje(driver)

    try:
        print("[OK] Mensaje enviado correctamente.")
    except UnicodeEncodeError:
        print("[OK] Mensaje enviado correctamente (caracteres especiales en contacto).")

except Exception:
    print("[ERROR] Ocurrió un problema al enviar el mensaje:")
    traceback.print_exc()

finally:
    print("[FIN] Proceso finalizado.")
    print("[INFO] Esperando 20 segundos antes de cerrar...")
    time.sleep(20)
    # driver.quit()
