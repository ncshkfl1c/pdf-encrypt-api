from flask import Flask, request, jsonify
from pypdf import PdfReader, PdfWriter
import base64, io

app = Flask(__name__)

@app.route('/')
def home():
    return "API is running"

@app.route('/encrypt', methods=['POST'])
def encrypt_pdf():
    try:
        data = request.json
        password = data.get("password", "123456")
        file_base64 = data.get("file")

        file_bytes = base64.b64decode(file_base64)
        reader = PdfReader(io.BytesIO(file_bytes))
        writer = PdfWriter()

        for page in reader.pages:
            writer.add_page(page)

        writer.encrypt(password)

        output = io.BytesIO()
        writer.write(output)

        return jsonify({
            "status": "success",
            "file": base64.b64encode(output.getvalue()).decode()
        })

    except Exception as e:
        return jsonify({"status": "error", "message": str(e)})
