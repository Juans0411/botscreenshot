import time, pyautogui, pygetwindow as gw
from datetime import datetime
import pytz

def buscar_ventana(titulo, wait_activate=3):
    ventanas = gw.getWindowsWithTitle(titulo)
    if not ventanas:
        raise RuntimeError(f"No hallé Power BI con título “{titulo}”.")
    v = ventanas[0]
    if v.isMinimized: v.restore()
    v.activate()
    time.sleep(wait_activate)
    return v

def pulsar_actualizar(imagen_boton, espera=30, confidence=0.8):
    loc = pyautogui.locateCenterOnScreen(imagen_boton, confidence=confidence)
    if not loc:
        raise RuntimeError("No encontré el botón de “Actualizar”.")
    pyautogui.click(loc)
    time.sleep(espera)

def seleccionar_pestana(imagen_pestana, confidence=0.8, post_wait=2):
    loc = pyautogui.locateCenterOnScreen(imagen_pestana, confidence=confidence)
    if not loc:
        raise RuntimeError("No encontré la pestaña.")
    pyautogui.click(loc)
    time.sleep(post_wait)

def tomar_captura(ventana, region, prefix="captura"):
    # region = (x_offset, y_offset, width, height)
    now = datetime.now(pytz.timezone("Europe/Madrid"))
    ts = now.strftime("%Y%m%d_%H%M")
    fn = f"{prefix}_{ts}.png"
    path = ventana._parent.parent / "Capturas" / fn  # o donde quieras
    img = pyautogui.screenshot(region=(
        ventana.left + region[0],
        ventana.top  + region[1],
        region[2], region[3]
    ))
    img.save(path)
    return str(path)