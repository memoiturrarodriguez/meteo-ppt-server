from flask import Flask, request, send_file, jsonify
from pathlib import Path
import subprocess
import json
import sys
import time
import traceback
import threading
import os

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 25 * 1024 * 1024  # 25 MB

BASE_DIR = Path(__file__).parent
JSON_FILE = BASE_DIR / "reporte_meteo.json"
SCRIPT_PPT = BASE_DIR / "generar_ppt_meteo_final_v3.py"

# Evita que dos usuarios/exportaciones pisen el mismo reporte_meteo.json al mismo tiempo.
_GENERATE_LOCK = threading.Lock()


@app.route("/", methods=["GET"])
def inicio():
    return jsonify({
        "status": "ok",
        "mensaje": "Servidor PPT Meteo funcionando correctamente.",
        "endpoint": "/generar-ppt",
    })


@app.route("/health", methods=["GET"])
def health():
    return "ok", 200


@app.route("/generar-ppt", methods=["POST"])
def generar_ppt():
    with _GENERATE_LOCK:
        try:
            data = request.get_json(silent=True)

            if data is None:
                return jsonify({"error": "No se recibió JSON válido."}), 400

            # Guardar JSON recibido desde la app.
            with open(JSON_FILE, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)

            if not SCRIPT_PPT.exists():
                return jsonify({
                    "error": "No se encontró el script generador PPT.",
                    "detalle": str(SCRIPT_PPT),
                }), 500

            inicio_generacion = time.time()

            resultado = subprocess.run(
                [sys.executable, str(SCRIPT_PPT)],
                cwd=str(BASE_DIR),
                capture_output=True,
                text=True,
                timeout=180,
            )

            print("STDOUT:", resultado.stdout)
            print("STDERR:", resultado.stderr)

            if resultado.returncode != 0:
                return jsonify({
                    "error": "Error al generar PPT.",
                    "detalle": resultado.stderr or resultado.stdout,
                }), 500

            # Buscar el PPT más reciente generado después de ejecutar el script.
            ppt_generados = sorted(
                BASE_DIR.glob("*.pptx"),
                key=lambda p: p.stat().st_mtime,
                reverse=True,
            )

            if not ppt_generados:
                return jsonify({"error": "No se encontró ningún PPT generado."}), 500

            ppt_final = ppt_generados[0]

            for ppt in ppt_generados:
                if ppt.stat().st_mtime >= inicio_generacion - 2:
                    ppt_final = ppt
                    break

            print(f"ENVIANDO PPT: {ppt_final}")

            return send_file(
                ppt_final,
                as_attachment=True,
                download_name=ppt_final.name,
                mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            )

        except subprocess.TimeoutExpired:
            return jsonify({
                "error": "Tiempo excedido generando PPT.",
                "detalle": "El servidor demoró más de 180 segundos en generar la presentación.",
            }), 504
        except Exception as e:
            print("ERROR GENERAL:", str(e))
            print(traceback.format_exc())
            return jsonify({
                "error": str(e),
                "detalle": traceback.format_exc(),
            }), 500


if __name__ == "__main__":
    port = int(os.environ.get("PORT", "5000"))
    app.run(host="0.0.0.0", port=port)
