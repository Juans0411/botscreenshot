import time
import pyautogui
import pygetwindow as gw
import subprocess


# CONFIGURACIÓN

nombre_archivo_pbix = "JMRS.DB - VD DIARIO Y MENSUAL"
imagen_boton = 'actualizar.png'
tiempo_espera_actualizacion = 30  # Tiempo estimado para completar el refresh


# PASO 1: Traer la ventana de Power BI al frente

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
    print(f"No se encontro ninguna ventana de Power BI con el nombre: '{nombre_archivo_pbix}'.")
    exit()


# PASO 2: Buscar y hacer clic en el botón 'Actualizar'

print("Buscando el boton 'Actualizar' en pantalla...")
try:
    button_location = pyautogui.locateCenterOnScreen(imagen_boton, confidence=0.8)
    if button_location:
        pyautogui.moveTo(button_location)
        pyautogui.click()
        print("Boton 'Actualizar' encontrado y en ejecucion.")
    else:
        raise Exception("Boton no encontrado.")
except Exception as e:
    print("Error:", e)
    print("Guardando captura de pantalla para depuracion...")
    pyautogui.screenshot("pantalla_debug.png")
    print("Captura guardada como 'pantalla_debug.png'. Revisa la imagen.")
    exit()


# PASO 3: Esperar la actualización

print(f"Esperando {tiempo_espera_actualizacion} segundos para completar la actualizacion...")
time.sleep(tiempo_espera_actualizacion)

print("[OK] Proceso de actualizacion finalizado. Listo para la siguiente accion.")

subprocess.run(["python", "bot_powerbi_whatsapp.py"], check=True)