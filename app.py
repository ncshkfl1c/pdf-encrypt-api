from flask import Flask, request, jsonify
import base64, io

import pandas as pd
import msoffcrypto

from pypdf import PdfReader, PdfWriter

from openpyxl import load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter

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

        # =====================================================
        # 1. HTML → EXCEL (TABLE + FORMAT + ENCRYPT)
        # =====================================================
        if html:
            # Clean HTML
            html = html.replace('\n', '').replace('\t', '')
            html = html.replace('<br>', '\n').replace('<br/>', '\n').replace('<br />', '\n')

            # Read HTML
            tables = pd.read_html(io.StringIO(html))
            if not tables:
                return jsonify({"error": "No table found"}), 400

            # Write Excel
            excel_stream = io.BytesIO()
            with pd.ExcelWriter(excel_stream, engine='openpyxl') as writer:
                for i, table in enumerate(tables):
                    table.to_excel(writer, sheet_name=f"Sheet{i+1}", index=False)

            # Load lại để format
            excel_stream.seek(0)
            wb = load_workbook(excel_stream)

            for idx, ws in enumerate(wb.worksheets, start=1):
                max_row = ws.max_row
                max_col = ws.max_column

                # Create table
                table_range = f"A1:{get_column_letter(max_col)}{max_row}"
                tab = Table(displayName=f"Table{idx}", ref=table_range)

                style = TableStyleInfo(
                    name="TableStyleMedium9",
                    showRowStripes=True
                )
                tab.tableStyleInfo = style
                ws.add_table(tab)

                # Format
                for col in ws.columns:
                    max_length = 0
                    col_letter = col[0].column_letter

                    for cell in col:
                        if cell.value:
                            cell.alignment = Alignment(wrap_text=True)
                            length = len(str(cell.value))
                            if length > max_length:
                                max_length = length

                    ws.column_dimensions[col_letter].width = min(max_length + 2, 50)

            # Save formatted Excel
            formatted_stream = io.BytesIO()
            wb.save(formatted_stream)
            result_bytes = formatted_stream.getvalue()

            # Encrypt Excel
            if password:
                input_stream = io.BytesIO(result_bytes)
                output_stream = io.BytesIO()

                office = msoffcrypto.OfficeFile(input_stream)
                office.encrypt(password, output_stream)

                result_bytes = output_stream.getvalue()

            file_type = "excel"

        # =====================================================
        # 2. FILE (PDF / EXCEL)
        # =====================================================
        elif file_base64:

            if file_base64.startswith("data:"):
                file_base64 = file_base64.split(",")[1]

            file_bytes = base64.b64decode(file_base64)

            # ================= PDF =================
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

            # ================= EXCEL =================
            elif file_name.endswith(".xlsx") or file_bytes[:2] == b'PK':
                if password:
                    input_stream = io.BytesIO(file_bytes)
                    output_stream = io.BytesIO()

                    office = msoffcrypto.OfficeFile(input_stream)
                    office.encrypt(password, output_stream)

                    result_bytes = output_stream.getvalue()
                else:
                    result_bytes = file_bytes

                file_type = "excel"

            else:
                return jsonify({"error": "Unsupported file type"}), 400

        else:
            return jsonify({"error": "No input provided"}), 400

        # =====================================================
        # RETURN
        # =====================================================
        return jsonify({
            "status": "success",
            "type": file_type,
            "file": base64.b64encode(result_bytes).decode()
        })

    except Exception as e:
        return jsonify({"error": str(e)}), 400


if __name__ == '__main__':
    app.run()
