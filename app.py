from flask import Flask, render_template, jsonify, request, send_from_directory
import subprocess
import os
import glob

app = Flask(__name__)
process = None

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/start", methods=["POST"])
def start_script():
    global process
    data = request.get_json()
    contacto = data.get("contacto", "").strip()

    if not contacto:
        return jsonify({"status": "missing_contact"})

    if process is None or process.poll() is not None:
        process = subprocess.Popen(
            ["python", "bot_powerbi_update.py", contacto],
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True
        )
        return jsonify({"status": "started"})
    else:
        return jsonify({"status": "already running"})

@app.route("/stop", methods=["POST"])
def stop_script():
    global process
    if process and process.poll() is None:
        process.terminate()
        return jsonify({"status": "terminated"})
    return jsonify({"status": "not running"})

@app.route("/logs")
def get_logs():
    global process
    if process and process.stdout:
        output = process.stdout.readline()
        return jsonify({"log": output})
    return jsonify({"log": ""})

@app.route("/ultima_captura")
def ultima_captura():
    carpeta = os.path.join(app.root_path, "Capturas")
    lista = sorted(glob.glob(os.path.join(carpeta, "*.png")), reverse=True)
    if lista:
        nombre_archivo = os.path.basename(lista[0])
        return send_from_directory(carpeta, nombre_archivo)
    else:
        return "No hay capturas", 404

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
