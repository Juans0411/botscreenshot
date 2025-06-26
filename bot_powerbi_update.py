import time
import pyautogui
import pygetwindow as gw
import subprocess
import sys

# Forzar salida inmediata de print()
sys.stdout.reconfigure(line_buffering=True)

# Obtener el contacto desde argumentos
if len(sys.argv) < 2:
    print("[ERROR] No se recibió el contacto.")
    exit()

CONTACTO = sys.argv[1]

# CONFIGURACIÓN
nombre_archivo_pbix = "JMRS.DB - VD DIARIO Y MENSUAL"
imagen_boton = 'actualizar.png'
tiempo_espera_actualizacion = 30  # segundos

print(f"Buscando ventana de Power BI...")
ventanas = gw.getWindowsWithTitle(nombre_archivo_pbix)

if ventanas:
    ventana = ventanas[0]
    if ventana.isMinimized:
        ventana.restore()
    ventana.activate()
    print(f"Ventana activada.")
    time.sleep(3)
else:
    print(f"No se encontró ninguna ventana de Power BI con el nombre: '{nombre_archivo_pbix}'.")
    exit()

print("Buscando el botón 'Actualizar' en pantalla...")
try:
    button_location = pyautogui.locateCenterOnScreen(imagen_boton, confidence=0.8)
    if button_location:
        pyautogui.moveTo(button_location)
        pyautogui.click()
        print("Botón 'Actualizar' encontrado y en ejecución.")
    else:
        raise Exception("Botón no encontrado.")
except Exception as e:
    print("Error:", e)
    pyautogui.screenshot("pantalla_debug.png")
    print("Captura guardada como 'pantalla_debug.png'. Revisa la imagen.")
    exit()

print(f"\nEsperando {tiempo_espera_actualizacion} segundos para completar la actualización...\n")
time.sleep(tiempo_espera_actualizacion)

print("[OK] Proceso de actualización finalizado. Enviando captura por WhatsApp...")

subprocess.run(["python", "bot_powerbi_whatsapp.py", CONTACTO], check=True)