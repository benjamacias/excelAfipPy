import os
from flask import Flask, jsonify, request
import excel_afip

app = Flask(__name__)

@app.route('/health', methods=['GET'])
def health():
    return jsonify({'status': 'ok'})

@app.route('/process', methods=['POST'])
def process():
    parallel = request.args.get('parallel', 'true').lower() == 'true'
    excel_afip.procesar_archivos(parallel=parallel)
    return jsonify({'processed': True, 'parallel': parallel})

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 8000))
    app.run(host='0.0.0.0', port=port)
