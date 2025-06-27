from flask import Flask, render_template, jsonify, request, send_from_directory
import subprocess, os, glob

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
    script   = data.get("script", "bot_powerbi_update.py").strip()

    if not contacto:
        return jsonify({"status": "missing_contact"})

    if process is None or process.poll() is not None:
        process = subprocess.Popen(
            ["python", script, contacto],
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
        line = process.stdout.readline()
        return jsonify({"log": line})
    return jsonify({"log": ""})

@app.route("/ultima_captura")
def ultima_captura():
    carpeta = os.path.join(app.root_path, "Capturas")
    files = glob.glob(os.path.join(carpeta, "*.png"))
    if not files:
        return "No hay capturas", 404
    # Ordenar por fecha de modificación, descendente
    files.sort(key=lambda x: os.path.getmtime(x), reverse=True)
    nombre = os.path.basename(files[0])
    return send_from_directory(carpeta, nombre)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)

