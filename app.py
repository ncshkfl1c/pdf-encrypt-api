@app.route('/encrypt', methods=['GET', 'POST'])
def encrypt_pdf():
    if request.method == 'GET':
        return "API OK - use POST"

    try:
        data = request.get_json(force=True)

        file_base64 = data.get("file")
        password = data.get("password", "123456")

        import base64, io
        from pypdf import PdfReader, PdfWriter

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
        return jsonify({"error": str(e)}), 400
