from flask import Flask, request, jsonify
from pypdf import PdfReader, PdfWriter
import io, base64

app = Flask(__name__)

@app.route('/encrypt', methods=['POST'])
def encrypt_pdf():
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

    encrypted_base64 = base64.b64encode(output.getvalue()).decode()

    return jsonify({
        "file": encrypted_base64
    })

if __name__ == '__main__':
    app.run(port=5000)