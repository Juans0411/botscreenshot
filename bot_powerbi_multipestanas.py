import sys
# Forzar impresión inmediata en UTF-8
try:
    sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)
    sys.stderr.reconfigure(encoding="utf-8", line_buffering=True)
except:
    pass

import time
import pyautogui
import pygetwindow as gw
import pytz
from datetime import datetime
from pathlib import Path
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Importar utilidades del bot de WhatsApp
from bot_powerbi_whatsapp import (
    buscar_ventana,
    tomar_captura,
    buscar_contacto,
    iniciar_navegador,
    esperar_whatsapp,
    escribir_mensaje,
    adjuntar_imagen,
    enviar_mensaje
)

# Parámetro: contacto
def main():
    if len(sys.argv) < 2:
        print("[ERROR] Falta el contacto.")
        sys.exit(1)
    CONTACTO = sys.argv[1]

    # Configuración de Power BI
    PB_TITLE  = "JMRS.DB - COMPETENCIA EQUIPOS"
    BTN_IMG   = "actualizar.png"
    TAB_IMG   = "pestana_captacion.png"
    TAB_LABEL = "10 Agentes Cap Europa"
    REGION    = (171, 184, 957, 249)

    def pulsar_actualizar(img_path):
        print("[INFO] Pulsando 'Actualizar'...")
        loc = pyautogui.locateCenterOnScreen(img_path, confidence=0.8)
        if not loc:
            raise RuntimeError("No encontré el botón 'Actualizar'.")
        pyautogui.click(loc)
        print("[OK] Actualización lanzada. Espera 30s")
        time.sleep(30)

    def seleccionar_pestana(img_path):
        print(f"[INFO] Seleccionando pestaña '{TAB_LABEL}'...")
        loc = pyautogui.locateCenterOnScreen(img_path, confidence=0.8)
        if not loc:
            raise RuntimeError(f"No encontré la pestaña '{TAB_LABEL}'.")
        pyautogui.click(loc)
        time.sleep(2)
        print(f"[OK] Pestaña '{TAB_LABEL}' seleccionada.")

    driver = None
    try:
        # 1) Abrir y actualizar Power BI
        win = buscar_ventana(PB_TITLE)
        pulsar_actualizar(BTN_IMG)

        # 2) Seleccionar la pestaña deseada
        win = buscar_ventana(PB_TITLE)
        seleccionar_pestana(TAB_IMG)

        # 3) Tomar captura
        win = buscar_ventana(PB_TITLE)
        ruta = tomar_captura(win, REGION)
        print(f"[INFO] Captura guardada en: {ruta}")

        # 4) Enviar por WhatsApp
        driver = iniciar_navegador()
        driver.get("https://web.whatsapp.com")
        esperar_whatsapp(driver)
        buscar_contacto(driver, CONTACTO)
        escribir_mensaje(driver, TAB_LABEL)
        adjuntar_imagen(driver, ruta)
        enviar_mensaje(driver)

        print("[OK] Proceso multipestañas completado.")

    except Exception as e:
        print(f"[ERROR] {e}")
    finally:
        if driver:
            try:
                driver.quit()
            except:
                pass
        print("[FIN] Proceso terminado.")

if __name__ == '__main__':
    main()