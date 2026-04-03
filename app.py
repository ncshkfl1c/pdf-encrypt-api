from flask import Flask, request, jsonify
import base64, io, traceback

import pandas as pd
import msoffcrypto

from pypdf import PdfReader, PdfWriter

from openpyxl import load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter

app = Flask(__name__)

# =========================
# UTIL
# =========================
def detect_file_type(file_bytes):
    if file_bytes[:4] == b'%PDF':
        return "pdf"
    elif file_bytes[:2] == b'PK':
        return "excel"
    return "unknown"

def retry(func, times=2):
    for i in range(times):
        try:
            return func()
        except Exception as e:
            if i == times - 1:
                raise

# =========================
# HTML → Excel (Safe)
# =========================
def process_html_to_excel(html, password):

    # CLEAN HTML
    html = html.replace('\n', '').replace('\t', '')
    html = html.replace('<br>', '\n').replace('<br/>', '\n').replace('<br />', '\n')

    # Parse HTML
    try:
        tables = pd.read_html(io.StringIO(html))
    except Exception as e:
        raise Exception(f"HTML parse failed: {str(e)}")

    if not tables:
        raise Exception("No table found in HTML")

    # Write Excel
    excel_stream = io.BytesIO()
    with pd.ExcelWriter(excel_stream, engine='openpyxl') as writer:
        for i, table in enumerate(tables):
            table.to_excel(writer, sheet_name=f"Sheet{i+1}", index=False)

    excel_stream.seek(0)

    # Format Excel
    try:
        wb = load_workbook(excel_stream)
    except Exception as e:
        raise Exception(f"Excel load failed: {str(e)}")

    for idx, ws in enumerate(wb.worksheets, start=1):
        max_row = ws.max_row
        max_col = ws.max_column

        table_range = f"A1:{get_column_letter(max_col)}{max_row}"

        tab = Table(displayName=f"Table{idx}", ref=table_range)
        style = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
        tab.tableStyleInfo = style
        ws.add_table(tab)

        for col in ws.columns:
            max_length = 0
            col_letter = col[0].column_letter

            for cell in col:
                if cell.value:
                    cell.alignment = Alignment(wrap_text=True)
                    max_length = max(max_length, len(str(cell.value)))

            ws.column_dimensions[col_letter].width = min(max_length + 2, 50)

    # Save formatted
    formatted_stream = io.BytesIO()
    wb.save(formatted_stream)

    result_bytes = formatted_stream.getvalue()

    # Encrypt Excel
    if password:
        try:
            input_stream = io.BytesIO(result_bytes)
            output_stream = io.BytesIO()

            office = msoffcrypto.OfficeFile(input_stream)
            office.encrypt(password, output_stream)

            result_bytes = output_stream.getvalue()
        except Exception as e:
            raise Exception(f"Excel encrypt failed: {str(e)}")

    return result_bytes, "excel"

# =========================
# PDF (Safe)
# =========================
def process_pdf(file_bytes, password):

    if file_bytes[:4] != b'%PDF':
        raise Exception("Invalid PDF file")

    def _process():
        reader = PdfReader(io.BytesIO(file_bytes))

        # xử lý nếu PDF đã bị encrypt
        if reader.is_encrypted:
            try:
                reader.decrypt("")
            except:
                raise Exception("PDF is encrypted and cannot be opened")

        writer = PdfWriter()

        for page in reader.pages:
            writer.add_page(page)

        if password:
            writer.encrypt(password)

        output = io.BytesIO()
        writer.write(output)

        return output.getvalue()

    return retry(_process), "pdf"

# =========================
# EXCEL (Safe)
# =========================
def process_excel(file_bytes, password):

    if not password:
        return file_bytes, "excel"

    try:
        input_stream = io.BytesIO(file_bytes)
        output_stream = io.BytesIO()

        office = msoffcrypto.OfficeFile(input_stream)
        office.encrypt(password, output_stream)

        return output_stream.getvalue(), "excel"
    except Exception as e:
        raise Exception(f"Excel encrypt failed: {str(e)}")

# =========================
# MAIN API
# =========================
@app.route('/process', methods=['POST'])
def process_file():
    try:
        data = request.get_json(force=True)

        file_base64 = data.get("file")
        html = data.get("html")
        password = data.get("password", "")

        # VALIDATE INPUT
        if not html and not file_base64:
            return jsonify({"error": "No input"}), 400

        # =========================
        # HTML FLOW
        # =========================
        if html:
            if len(html) < 50:
                return jsonify({"error": "HTML too short"}), 400

            result_bytes, file_type = retry(lambda: process_html_to_excel(html, password))

        # =========================
        # FILE FLOW
        # =========================
        else:
            if file_base64.startswith("data:"):
                file_base64 = file_base64.split(",")[1]

            file_bytes = base64.b64decode(file_base64)
            file_bytes = bytes(file_bytes)

            if len(file_bytes) < 100:
                return jsonify({"error": "File corrupted"}), 400

            if len(file_bytes) > 10 * 1024 * 1024:
                return jsonify({"error": "File too large"}), 400

            file_type = detect_file_type(file_bytes)

            if file_type == "pdf":
                result_bytes, file_type = process_pdf(file_bytes, password)

            elif file_type == "excel":
                result_bytes, file_type = process_excel(file_bytes, password)

            else:
                return jsonify({"error": "Unsupported file format"}), 400

        # =========================
        # RETURN
        # =========================
        return jsonify({
            "status": "success",
            "type": file_type,
            "fileName": f"output.{file_type}",
            "file": base64.b64encode(result_bytes).decode()
        })

    except Exception as e:
        print(traceback.format_exc())
        return jsonify({"error": str(e)}), 400


if __name__ == '__main__':
    app.run()
