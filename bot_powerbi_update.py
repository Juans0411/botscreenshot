import time, sys, subprocess
import pyautogui, pygetwindow as gw

# salida en tiempo real
sys.stdout.reconfigure(line_buffering=True)
sys.stderr.reconfigure(line_buffering=True)

if len(sys.argv) < 2:
    print("[ERROR] Falta el contacto.")
    sys.exit(1)
CONTACTO = sys.argv[1]

PB_TITLE = "JMRS.DB - VD DIARIO Y MENSUAL"
BTN_IMG = "actualizar.png"
WAIT_SEC = 30

def activate_window(title):
    print("[INFO] Buscando ventana de Power BI...")
    wins = gw.getWindowsWithTitle(title)
    if not wins:
        print(f"[ERROR] No hallé '{title}'."); sys.exit(1)
    w = wins[0]
    if w.isMinimized: w.restore()
    w.activate(); time.sleep(3)
    print("[OK] Ventana lista.")
    return w

def click_update(btn_img):
    print("[INFO] Buscando botón 'Actualizar'...")
    loc = pyautogui.locateCenterOnScreen(btn_img, confidence=0.8)
    if not loc:
        print("[ERROR] Botón no encontrado."); sys.exit(1)
    pyautogui.click(loc)
    print(f"[OK] Actualizando...espera {WAIT_SEC}s")
    time.sleep(WAIT_SEC)

if __name__ == "__main__":
    win = activate_window(PB_TITLE)
    click_update(BTN_IMG)
    print("[OK] Actualización lista. Llamo al envío por WhatsApp.")
    subprocess.run(["python", "bot_powerbi_whatsapp.py", CONTACTO], check=True)