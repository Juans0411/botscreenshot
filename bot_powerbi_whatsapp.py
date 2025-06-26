import pygetwindow as gw
import pyautogui
import time
import os
import pytz
import sys
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from pathlib import Path

# Forzar impresión inmediata
sys.stdout.reconfigure(line_buffering=True)

# Obtener contacto desde argumento
if len(sys.argv) < 2:
    print("[ERROR] No se recibió el contacto.")
    exit()

CONTACTO = sys.argv[1]

# CONFIGURACIÓN
POWERBI_TITULO = "JMRS.DB - VD DIARIO Y MENSUAL"
REGION = (168, 120, 985, 543)
CAPTURA_DIR = Path(__file__).parent / "Capturas"
CAPTURA_DIR.mkdir(exist_ok=True)

def buscar_ventana(titulo):
    print("[INFO] Buscando Power BI...")
    ventanas = gw.getWindowsWithTitle(titulo)
    if not ventanas:
        raise Exception("No se encontró Power BI abierto con el título especificado.")
    ventana = ventanas[0]
    ventana.activate()
    ventana.maximize()
    time.sleep(2)
    print("[OK] Ventana de Power BI activa.")
    return ventana

def obtener_hora_corte_es():
    ahora = datetime.now(pytz.timezone("Europe/Madrid"))
    minuto_redondeado = 0 if ahora.minute < 30 else 30
    return f"{ahora.hour:02d}:{minuto_redondeado:02d}"

def tomar_captura(ventana, region):
    print("[INFO] Tomando captura...")
    hora_corte = obtener_hora_corte_es().replace(":", "")
    captura = pyautogui.screenshot(region=(ventana.left + region[0], ventana.top + region[1], region[2], region[3]))
    timestamp = datetime.now().strftime("%Y%m%d")
    ruta = CAPTURA_DIR / f"captura_{timestamp}_{hora_corte}.png"
    captura.save(ruta)
    print("[OK] Captura guardada en:", ruta)
    return str(ruta)

def iniciar_navegador():
    print("[INFO] Iniciando Chrome...")
    options = webdriver.ChromeOptions()
    options.add_argument(f"--user-data-dir={os.getcwd()}/chrome_selenium_profile")
    options.add_argument("--start-maximized")
    driver_path = str(Path(__file__).parent / "driver" / "chromedriver.exe")
    return webdriver.Chrome(service=Service(driver_path), options=options)

def esperar_whatsapp(driver):
    print("[INFO] Esperando que cargue WhatsApp Web...")
    WebDriverWait(driver, 120).until(
        lambda d: d.find_elements(By.XPATH, '//div[@contenteditable="true" and contains(@data-tab,"3")]')
    )
    print("[OK] WhatsApp Web está listo.")

def buscar_contacto(driver, nombre):
    print(f"[INFO] Buscando contacto: {nombre}...")
    try:
        search_box = WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.XPATH, '//div[@contenteditable="true" and contains(@data-tab,"3")]'))
        )
        search_box.click()
        search_box.clear()
        search_box.send_keys(nombre)
        time.sleep(2)

        resultados = driver.find_elements(By.XPATH, f'//span[contains(@title, "{nombre}")]')
        if not resultados:
            raise Exception(f"[ERROR] No se encontró el contacto '{nombre}' en WhatsApp.")

        resultados[0].click()
        print(f"[OK] Chat abierto con: {nombre}")

    except Exception as e:
        raise e  # Se maneja desde el bloque principal

def escribir_mensaje(driver, mensaje):
    print("[INFO] Escribiendo mensaje...")
    try:
        input_box = WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.XPATH, '//footer//div[@contenteditable="true"]'))
        )
        input_box.click()
        input_box.send_keys(mensaje)
        time.sleep(0.5)
    except Exception as e:
        print("[ERROR] No se pudo escribir el mensaje.")
        print(e)

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
    print("[INFO] Enviando mensaje...")
    send_button = WebDriverWait(driver, 15).until(
        EC.element_to_be_clickable((By.XPATH, '//div[@role="button"][@aria-label="Enviar"]'))
    )
    time.sleep(1)
    send_button.click()

# EJECUCIÓN PRINCIPAL
try:
    ventana = buscar_ventana(POWERBI_TITULO)
    hora = obtener_hora_corte_es()
    ruta_captura = tomar_captura(ventana, REGION)

    driver = iniciar_navegador()
    driver.get("https://web.whatsapp.com")
    esperar_whatsapp(driver)

    buscar_contacto(driver, CONTACTO)
    escribir_mensaje(driver, hora)
    adjuntar_imagen(driver, ruta_captura)
    enviar_mensaje(driver)

    print("[OK] Mensaje enviado correctamente.")

except Exception as e:
    print(str(e))
    if "[ERROR] No se encontró el contacto" in str(e):
        print("[SUGERENCIA] Verifica que hayas escrito el nombre o número exactamente como aparece en WhatsApp Web.")

finally:
    print("[FIN] Proceso finalizado.")
    print("[INFO] Esperando 20 segundos antes de cerrar...")
    time.sleep(20)
    # driver.quit()