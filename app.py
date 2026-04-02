from flask import Flask, request, jsonify
import base64, io

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
        # 1. HTML → Excel
        # =========================
        if html:
            # clean HTML
            html = html.replace('\n', '').replace('\t', '')

            tables = pd.read_html(html)

            if not tables:
                return jsonify({"error": "No table found in HTML"}), 400

            excel_stream = io.BytesIO()

            with pd.ExcelWriter(excel_stream, engine='openpyxl') as writer:
                for i, table in enumerate(tables):
                    table.to_excel(writer, sheet_name=f"Sheet{i+1}", index=False)

            excel_bytes = excel_stream.getvalue()

            # 👉 Encrypt nếu có password
            if password:
                input_stream = io.BytesIO(excel_bytes)
                output_stream = io.BytesIO()

                office = msoffcrypto.OfficeFile(input_stream)
                office.encrypt(password=password)
                office.save(output_stream)

                result_bytes = output_stream.getvalue()
            else:
                result_bytes = excel_bytes

            file_type = "excel"

        # =========================
        # 2. FILE (PDF / EXCEL)
        # =========================
        elif file_base64:

            # remove prefix nếu có
            if file_base64.startswith("data:"):
                file_base64 = file_base64.split(",")[1]

            file_bytes = base64.b64decode(file_base64)

            # ===== PDF =====
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

            # ===== EXCEL =====
            elif file_name.endswith(".xlsx") or file_bytes[:2] == b'PK':
                if password:
                    input_stream = io.BytesIO(file_bytes)
                    output_stream = io.BytesIO()

                    office = msoffcrypto.OfficeFile(input_stream)
                    office.encrypt(password=password)
                    office.save(output_stream)

                    result_bytes = output_stream.getvalue()
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


if __name__ == '__main__':
    app.run()
