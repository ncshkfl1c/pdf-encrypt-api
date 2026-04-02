from flask import Flask, request, jsonify
import base64, io, tempfile

from pypdf import PdfReader, PdfWriter
import msoffcrypto

app = Flask(__name__)

@app.route('/')
def home():
    return "API is running"

@app.route('/encrypt', methods=['POST'])
def encrypt_file():
    try:
        data = request.get_json(force=True)

        file_base64 = data.get("file")
        password = data.get("password", "123456")

        if not file_base64:
            return jsonify({"error": "Missing file"}), 400

        # remove prefix nếu có
        if file_base64.startswith("data:"):
            file_base64 = file_base64.split(",")[1]

        file_bytes = base64.b64decode(file_base64)

        # 🧠 Detect file type
        if file_bytes[:4] == b'%PDF':
            # ===== PDF =====
            reader = PdfReader(io.BytesIO(file_bytes))
            writer = PdfWriter()

            for page in reader.pages:
                writer.add_page(page)

            writer.encrypt(password)

            output = io.BytesIO()
            writer.write(output)

            result_bytes = output.getvalue()
            file_type = "pdf"

        else:
            # ===== Excel =====
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_in:
                tmp_in.write(file_bytes)
                input_path = tmp_in.name

            output_path = input_path.replace(".xlsx", "_encrypted.xlsx")

            with open(input_path, "rb") as f_in:
                office_file = msoffcrypto.OfficeFile(f_in)
                office_file.encrypt(password=password)

                with open(output_path, "wb") as f_out:
                    office_file.save(f_out)

            with open(output_path, "rb") as f:
                result_bytes = f.read()

            file_type = "excel"

        return jsonify({
            "status": "success",
            "type": file_type,
            "file": base64.b64encode(result_bytes).decode()
        })

    except Exception as e:
        return jsonify({"error": str(e)}), 400


if __name__ == '__main__':
    app.run()
