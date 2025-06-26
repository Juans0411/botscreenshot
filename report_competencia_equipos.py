import time
import pyautogui
import pygetwindow as gw
import subprocess
import sys
import os
import pytz
from datetime import datetime
from pathlib import Path
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

sys.stdout.reconfigure(line_buffering=True)

# Contacto recibido por argumento
if len(sys.argv) < 2:
    print("[ERROR] No se recibió el contacto.")
    exit()

CONTACTO = sys.argv[1]

# Configuración
PBIX_TITULO = "JMRS.DB - COMPETENCIA EQUIPOS"
IMG_ACTUALIZAR = "actualizar.png"
IMG_PESTANA = "pestana_captacion.png"
TIEMPO_ESPERA_ACTUALIZACION = 30
CAPTURA_REGION = (171, 184, 957, 249)
MENSAJE = "10 Agentes Cap Europa"
CAPTURA_DIR = Path(__file__).parent / "Capturas"
CAPTURA_DIR.mkdir(exist_ok=True)

# --- Funciones auxiliares ---
def buscar_y_activar_ventana(nombre):
    print("[INFO] Buscando Power BI...")
    ventanas = gw.getWindowsWithTitle(nombre)
    if not ventanas:
        raise Exception("No se encontró la ventana de Power BI.")
    ventana = ventanas[0]
    if ventana.isMinimized:
        ventana.restore()
    ventana.activate()
    time.sleep(3)
    print("[OK] Ventana activada.")
    return ventana

def actualizar_powerbi():
    print("[INFO] Buscando botón 'Actualizar'...")
    btn = pyautogui.locateCenterOnScreen(IMG_ACTUALIZAR, confidence=0.8)
    if not btn:
        pyautogui.screenshot("pantalla_debug.png")
        raise Exception("[ERROR] No se encontró el botón 'Actualizar'. Se guardó pantalla_debug.png")
    pyautogui.click(btn)
    print("[OK] Botón 'Actualizar' clicado.")
    print(f"[INFO] Esperando {TIEMPO_ESPERA_ACTUALIZACION} segundos para completar actualización...")
    time.sleep(TIEMPO_ESPERA_ACTUALIZACION)

def ir_a_pestana():
    print(f"[INFO] Buscando pestaña '{MENSAJE}'...")
    pestana = pyautogui.locateCenterOnScreen(IMG_PESTANA, confidence=0.8)
    if not pestana:
        raise Exception(f"[ERROR] No se encontró la pestaña '{MENSAJE}' en pantalla.")
    pyautogui.click(pestana)
    time.sleep(3)
    print(f"[OK] Pestaña '{MENSAJE}' seleccionada.")

def tomar_captura(ventana):
    print("[INFO] Tomando captura...")
    ahora = datetime.now(pytz.timezone("Europe/Madrid"))
    hora_str = ahora.strftime("%H%M")
    fecha_str = ahora.strftime("%Y%m%d")
    captura = pyautogui.screenshot(region=(ventana.left + CAPTURA_REGION[0], ventana.top + CAPTURA_REGION[1], CAPTURA_REGION[2], CAPTURA_REGION[3]))
    ruta = CAPTURA_DIR / f"captura_{fecha_str}_{hora_str}.png"
    captura.save(ruta)
    print(f"[OK] Captura guardada: {ruta}")
    return str(ruta)

def enviar_por_whatsapp(ruta_captura, mensaje):
    print("[INFO] Iniciando navegador...")
    options = webdriver.ChromeOptions()
    options.add_argument(f"--user-data-dir={os.getcwd()}/chrome_selenium_profile")
    options.add_argument("--start-maximized")
    driver_path = str(Path(__file__).parent / "driver" / "chromedriver.exe")
    driver = webdriver.Chrome(service=Service(driver_path), options=options)
    
    driver.get("https://web.whatsapp.com")
    print("[INFO] Esperando WhatsApp Web...")
    WebDriverWait(driver, 120).until(
        lambda d: d.find_elements(By.XPATH, '//div[@contenteditable="true" and contains(@data-tab,"3")]')
    )

    print(f"[INFO] Buscando contacto: {CONTACTO}...")
    search = WebDriverWait(driver, 15).until(
        EC.presence_of_element_located((By.XPATH, '//div[@contenteditable="true" and contains(@data-tab,"3")]'))
    )
    search.click()
    search.clear()
    search.send_keys(CONTACTO)
    time.sleep(2)

    resultados = driver.find_elements(By.XPATH, f'//span[contains(@title, "{CONTACTO}")]')
    if not resultados:
        print(f"[ERROR] No se encontró el contacto '{CONTACTO}' en WhatsApp.")
        driver.quit()
        return
    resultados[0].click()

    input_box = WebDriverWait(driver, 15).until(
        EC.presence_of_element_located((By.XPATH, '//footer//div[@contenteditable="true"]'))
    )
    input_box.click()
    input_box.send_keys(mensaje)
    time.sleep(0.5)

    print("[INFO] Adjuntando imagen...")
    plus_btn = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//button[@aria-label="Adjuntar" or @title="Adjuntar"]'))
    )
    plus_btn.click()
    file_input = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, '//input[@type="file"]'))
    )
    file_input.send_keys(ruta_captura)

    send_btn = WebDriverWait(driver, 15).until(
        EC.element_to_be_clickable((By.XPATH, '//div[@role="button"][@aria-label="Enviar"]'))
    )
    send_btn.click()
    print("[OK] Mensaje enviado.")
    time.sleep(5)
    driver.quit()

# --- Ejecución principal ---
try:
    ventana = buscar_y_activar_ventana(PBIX_TITULO)
    actualizar_powerbi()
    ir_a_pestana()
    captura = tomar_captura(ventana)
    enviar_por_whatsapp(captura, MENSAJE)
    print("[FIN] Proceso completado.")
except Exception as e:
    print(str(e))
    print("[FIN] Proceso finalizado con errores.")


