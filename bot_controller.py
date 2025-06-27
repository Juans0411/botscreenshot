# bot_controller.py
import sys, time
from powerbi_utils import buscar_ventana, pulsar_actualizar, seleccionar_pestana, tomar_captura
from whatsapp_utils import iniciar_navegador, enviar_captura_whatsapp

if len(sys.argv) < 3:
    print("uso: python bot_controller.py [diario|competencia|oficinas_diario|oficinas_mensual] <contacto>")
    sys.exit(1)

bot_type, CONTACTO = sys.argv[1], sys.argv[2]

# Configuración para cada bot
CONFIGS = {
    "diario": {
        "title": "JMRS.DB - VD DIARIO Y MENSUAL",
        "update_image": "actualizar.png",
        "tabs": [None],
        "regions": [(168,120,985,543)],
        "messages": [""],
        "prefixes": ["vdDiario"]
    },
    "competencia": {
        "title": "JMRS.DB - COMPETENCIA EQUIPOS",
        "update_image": "actualizar.png",
        "tabs": ["pestana_captacion.png"],
        "regions": [(171,184,957,249)],
        "messages": ["10 Agentes Cap Europa"],
        "prefixes": ["competencia"]
    },
     "competencia": {
        "title": "JMRS.DB - OFICINAS DIARIO",
        "update_image": "actualizar.png",
        "tabs": ["pestana_captacion.png"],
        "regions": [(171,184,957,249)],
        "messages": ["10 Agentes Cap Europa"],
        "prefixes": ["competencia"]
    },
     "competencia": {
        "title": "JMRS.DB - OFICINAS MENSUAL",
        "update_image": "actualizar.png",
        "tabs": ["pestana_captacion.png"],
        "regions": [(171,184,957,249)],
        "messages": ["10 Agentes Cap Europa"],
        "prefixes": ["competencia"]
    },
    }

if bot_type not in CONFIGS:
    print("bot desconocido:", bot_type)
    sys.exit(1)

cfg = CONFIGS[bot_type]

try:
    # 1. Power BI: activar y actualizar
    win = buscar_ventana(cfg["title"])
    pulsar_actualizar(cfg["update_image"])

    # 2. WhatsApp: levantar navegador *una sola vez*
    driver = iniciar_navegador()

    # 3. Para cada “pestaña” (o si tabs[0] is None, solo una captura)
    for tab_img, region, msg, prefix in zip(cfg["tabs"], cfg["regions"], cfg["messages"], cfg["prefixes"]):
        if tab_img:
            seleccionar_pestana(tab_img)
        path = tomar_captura(win, region, prefix)
        enviar_captura_whatsapp(driver, CONTACTO, path, msg)

    print("[OK] Todas las capturas enviadas.")
except Exception as e:
    print("[ERROR]", e)
finally:
    # un breve descanso antes de cerrar
    time.sleep(5)
    try:
        driver.quit()
    except:
        pass