from flask import Flask, request, jsonify
from flask_cors import CORS
import os
from analyzer import analizar_documento
from comparador import comparar_documentos
from pdf_generator import generar_pdf
import base64

app = Flask(__name__)
CORS(app)

@app.route('/health', methods=['GET'])
def health():
    return jsonify({"status": "ok"})

@app.route('/analizar', methods=['POST'])
def analizar():
    if 'archivo' not in request.files:
        return jsonify({"error": "No se recibió archivo"}), 400
    
    archivo = request.files['archivo']
    if not archivo.filename.endswith('.docx'):
        return jsonify({"error": "Solo se aceptan archivos .docx"}), 400

    ruta_temp = f"/tmp/{archivo.filename}"
    archivo.save(ruta_temp)

    try:
        resultado = analizar_documento(ruta_temp)
        return jsonify(resultado)
    except Exception as e:
        return jsonify({"error": str(e)}), 500
    finally:
        if os.path.exists(ruta_temp):
            os.remove(ruta_temp)

@app.route('/comparar', methods=['POST'])
def comparar():
    if 'version1' not in request.files or 'version2' not in request.files:
        return jsonify({"error": "Se necesitan dos archivos"}), 400

    v1 = request.files['version1']
    v2 = request.files['version2']

    ruta1 = f"/tmp/v1_{v1.filename}"
    ruta2 = f"/tmp/v2_{v2.filename}"
    v1.save(ruta1)
    v2.save(ruta2)

    try:
        resultado = comparar_documentos(ruta1, ruta2)
        return jsonify(resultado)
    except Exception as e:
        return jsonify({"error": str(e)}), 500
    finally:
        for ruta in [ruta1, ruta2]:
            if os.path.exists(ruta):
                os.remove(ruta)

@app.route('/generar-pdf', methods=['POST'])
def pdf():
    datos = request.json
    if not datos:
        return jsonify({"error": "No se recibieron datos"}), 400
    
    try:
        pdf_bytes = generar_pdf(datos)
        pdf_b64 = base64.b64encode(pdf_bytes).decode('utf-8')
        return jsonify({"pdf": pdf_b64})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
