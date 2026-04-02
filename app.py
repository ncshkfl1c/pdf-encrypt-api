from flask import Flask, request, jsonify
import base64, io, tempfile

from pypdf import PdfReader, PdfWriter
import msoffcrypto
import pandas as pd

app = Flask(__name__)

@app.route('/')
def home():
    return "API is running"

@app.route('/process', methods=['POST'])
def process_file():
    try:
        data = request.get_json(force=True)

        file_base64 = data.get("file")
        html = data.get("html")
        password = data.get("password", "")
        file_name = data.get("fileName", "").lower()

        result_bytes = None
        file_type = None

        # =========================
        # 1. HTML → EXCEL
        # =========================
        if html:
            tables = pd.read_html(html)

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                for i, table in enumerate(tables):
                    table.to_excel(writer, sheet_name=f"Sheet{i+1}", index=False)

            result_bytes = output.getvalue()
            file_type = "excel"

        # =========================
        # 2. FILE (PDF / EXCEL)
        # =========================
        elif file_base64:
            if file_base64.startswith("data:"):
                file_base64 = file_base64.split(",")[1]

            file_bytes = base64.b64decode(file_base64)

            # detect theo tên file
            if file_name.endswith(".pdf") or file_bytes[:4] == b'%PDF':
                reader = PdfReader(io.BytesIO(file_bytes))
                writer = PdfWriter()

                for page in reader.pages:
                    writer.add_page(page)

                if password:
                    writer.encrypt(password)

                output = io.BytesIO()
                writer.write(output)

                result_bytes = output.getvalue()
                file_type = "pdf"

            elif file_name.endswith(".xlsx") or file_bytes[:2] == b'PK':
                with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                    tmp.write(file_bytes)
                    input_path = tmp.name

                output_path = input_path.replace(".xlsx", "_enc.xlsx")

                if password:
                    with open(input_path, "rb") as f_in:
                        office = msoffcrypto.OfficeFile(f_in)
                        office.encrypt(password=password)

                        with open(output_path, "wb") as f_out:
                            office.save(f_out)

                    with open(output_path, "rb") as f:
                        result_bytes = f.read()
                else:
                    result_bytes = file_bytes

                file_type = "excel"

            else:
                return jsonify({"error": "Unsupported file type"}), 400

        else:
            return jsonify({"error": "No input provided"}), 400

        # =========================
        # 3. RETURN
        # =========================
        return jsonify({
            "status": "success",
            "type": file_type,
            "file": base64.b64encode(result_bytes).decode()
        })

    except Exception as e:
        return jsonify({"error": str(e)}), 400
