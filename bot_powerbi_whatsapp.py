import sys
sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)
sys.stderr.reconfigure(encoding="utf-8", line_buffering=True)

import time, os
import pyautogui, pygetwindow as gw, pytz
from datetime import datetime
from pathlib import Path
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Parámetro: contacto
if len(sys.argv) < 2:
    print("[ERROR] Falta el contacto.")
    sys.exit(1)
CONTACTO = sys.argv[1]

# Configuración
POWERBI_TITULO = "JMRS.DB - VD DIARIO Y MENSUAL"
REGION        = (168, 120, 985, 543)
CAPTURA_DIR   = Path(__file__).parent / "Capturas"
CAPTURA_DIR.mkdir(exist_ok=True)
CH_PROFILE    = os.path.join(os.getcwd(), "chrome_selenium_profile")
CH_DRIVER     = str(Path(__file__).parent / "driver" / "chromedriver.exe")

def buscar_ventana(titulo):
    print("[INFO] Buscando Power BI...")
    ventanas = gw.getWindowsWithTitle(titulo)
    if not ventanas:
        raise RuntimeError(f"No se encontró Power BI con título: '{titulo}'.")
    ventana = ventanas[0]
    if ventana.isMinimized:
        ventana.restore()
    ventana.activate()
    ventana.maximize()
    time.sleep(2)
    print("[OK] Ventana de Power BI activa.")
    return ventana

def obtener_hora_corte_es():
    ahora = datetime.now(pytz.timezone("Europe/Madrid"))
    minuto = 0 if ahora.minute < 30 else 30
    return f"{ahora.hour:02d}:{minuto:02d}"

def tomar_captura(ventana, region):
    print("[INFO] Tomando captura...")
    hora_corte = obtener_hora_corte_es().replace(":", "")
    timestamp  = datetime.now().strftime("%Y%m%d")
    nombre     = f"captura_{timestamp}_{hora_corte}.png"
    ruta       = CAPTURA_DIR / nombre
    img        = pyautogui.screenshot(region=(
        ventana.left + region[0],
        ventana.top  + region[1],
        region[2], region[3]
    ))
    img.save(ruta)
    print(f"[OK] Captura guardada en: {ruta}")
    return str(ruta)

def iniciar_navegador():
    print("[INFO] Iniciando Chrome...")
    opts = webdriver.ChromeOptions()
    opts.add_argument(f"--user-data-dir={CH_PROFILE}")
    opts.add_argument("--start-maximized")
    return webdriver.Chrome(service=Service(CH_DRIVER), options=opts)

def esperar_whatsapp(driver):
    print("[INFO] Esperando WhatsApp Web...")
    WebDriverWait(driver, 120).until(
        EC.presence_of_element_located((By.XPATH,
            '//div[@contenteditable="true" and contains(@data-tab,"3")]'))
    )
    print("[OK] WhatsApp Web lista.")

def buscar_contacto(driver, nombre):
    """
    Escribe en la caja de búsqueda y abre el primer resultado,
    prefiriendo un título exacto si existe.
    """
    print(f"[INFO] Buscando contacto: {nombre}...")
    search_box = WebDriverWait(driver, 15).until(
        EC.presence_of_element_located((By.XPATH,
            '//div[@contenteditable="true" and contains(@data-tab,"3")]'))
    )
    search_box.click()
    search_box.clear()
    search_box.send_keys(nombre)
    time.sleep(2)
    # Intentar match exacto
    try:
        exact = driver.find_element(By.XPATH, f'//span[@title="{nombre}"]')
        exact.click()
        print(f"[OK] Chat abierto con: {nombre}")
    except:
        # Fallback: ENTER abre el primer resultado
        search_box.send_keys(Keys.ENTER)
        print(f"[OK] Chat abierto con la primera sugerencia para: {nombre}")
    time.sleep(1)

def escribir_mensaje(driver, mensaje):
    print("[INFO] Escribiendo mensaje...")
    input_box = WebDriverWait(driver, 15).until(
        EC.presence_of_element_located((By.XPATH,
            '//footer//div[@contenteditable="true"]'))
    )
    input_box.click()
    input_box.send_keys(mensaje)
    time.sleep(0.5)
    print("[OK] Texto listo (no enviado aún).")

def adjuntar_imagen(driver, ruta):
    print("[INFO] Adjuntando imagen...")
    plus_btn = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH,
            '//button[@aria-label="Adjuntar" or @title="Adjuntar"]'))
    )
    plus_btn.click()
    file_input = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH,
            '//input[@type="file" and contains(@accept,"image")]'))
    )
    file_input.send_keys(ruta)
    print("[OK] Imagen seleccionada.")

def enviar_mensaje(driver):
    print("[INFO] Enviando mensaje...")
    send_btn = WebDriverWait(driver, 15).until(
        EC.element_to_be_clickable((By.XPATH,
            '//div[@role="button"][@aria-label="Enviar"]'))
    )
    time.sleep(1)
    send_btn.click()
    print("[OK] Mensaje enviado.")

if __name__ == "__main__":
    try:
        ventana = buscar_ventana(POWERBI_TITULO)
        ruta     = tomar_captura(ventana, REGION)

        driver = iniciar_navegador()
        driver.get("https://web.whatsapp.com")
        esperar_whatsapp(driver)

        buscar_contacto(driver, CONTACTO)
        escribir_mensaje(driver, obtener_hora_corte_es())
        adjuntar_imagen(driver, ruta)
        enviar_mensaje(driver)

        print("[OK] Proceso completado correctamente.")
    except Exception as e:
        print(f"[ERROR] {e}")
    finally:
        print("[FIN] Esperando 10 segundos antes de cerrar...")
        time.sleep(10)
        try:
            driver.quit()
        except:
            pass