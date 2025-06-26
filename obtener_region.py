import pyautogui

print("🟩 PASO 1: Mueve el mouse a la ESQUINA SUPERIOR IZQUIERDA de la región deseada.")
input("Presiona ENTER cuando estés listo...")
x1, y1 = pyautogui.position()
print(f"Esquina superior izquierda: ({x1}, {y1})")

print("🟩 PASO 2: Ahora mueve el mouse a la ESQUINA INFERIOR DERECHA.")
input("Presiona ENTER cuando estés listo...")
x2, y2 = pyautogui.position()
print(f"Esquina inferior derecha: ({x2}, {y2})")

width = x2 - x1
height = y2 - y1
print(f"\n✅ Región lista: ({x1}, {y1}, {width}, {height})")