import os
from io import BytesIO
from zipfile import ZipFile
from flask import Flask, jsonify, request, send_file
from werkzeug.utils import secure_filename
import pandas as pd
import excel_afip

file_path = excel_afip.ARCHIVO_CLIENTES
extension = os.path.splitext(file_path)[-1].lower()

if extension == ".xls":
    engine = "xlrd"
else:
    engine = "openpyxl"

app = Flask(__name__)

@app.route('/health', methods=['GET'])
def health():
    return jsonify({'status': 'ok'})

@app.route('/process', methods=['POST'])
def process():
    parallel = request.args.get('parallel', 'true').lower() == 'true'
    excel_afip.procesar_archivos(parallel=parallel)
    return jsonify({'processed': True, 'parallel': parallel})


@app.route('/process-files', methods=['POST'])
def process_files():
    """Accepts multiple uploaded xlsx files and returns them processed as a zip."""
    if 'files' not in request.files:
        return jsonify({'error': 'No files provided'}), 400

    files = request.files.getlist('files')
    if not files:
        return jsonify({'error': 'No files provided'}), 400

    os.makedirs(excel_afip.DIR_ENTRADA, exist_ok=True)
    os.makedirs(excel_afip.DIR_SALIDA, exist_ok=True)

    clientes_excel = pd.read_excel(file_path, sheet_name="Sheet1", engine="openpyxl")
    processed_paths = []

    for f in files:
        filename = secure_filename(f.filename)
        input_path = os.path.join(excel_afip.DIR_ENTRADA, filename)
        f.save(input_path)
        out_path = excel_afip.procesar_archivo(filename, clientes_excel)
        if out_path:
            processed_paths.append(out_path)

    if not processed_paths:
        return jsonify({'error': 'Processing failed'}), 500

    zip_buffer = BytesIO()
    with ZipFile(zip_buffer, 'w') as zf:
        for path in processed_paths:
            zf.write(path, os.path.basename(path))
    zip_buffer.seek(0)

    return send_file(zip_buffer, mimetype='application/zip', as_attachment=True,
                     download_name='processed_files.zip')

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 8000))
    app.run(host='0.0.0.0', port=port)
