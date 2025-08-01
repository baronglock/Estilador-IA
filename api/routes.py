from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
from werkzeug.utils import secure_filename
import os
import json  # <-- ADICIONE ESTA LINHA
from backend.main import WordStylerProcessor
from backend.config import Config

app = Flask(__name__)
CORS(app)

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in Config.ALLOWED_EXTENSIONS

@app.route('/api/process', methods=['POST'])
def process_document():
    """Endpoint principal para processar documento"""
    # Verifica se arquivo foi enviado
    if 'file' not in request.files:
        return jsonify({'error': 'Nenhum arquivo enviado'}), 400
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'Nenhum arquivo selecionado'}), 400
    
    if not allowed_file(file.filename):
        return jsonify({'error': 'Tipo de arquivo não permitido. Use .docx'}), 400
    
    # Obtém dados do formulário
    data = request.form
    book_name = data.get('book_name')
    api_key = data.get('api_key')
    styles = json.loads(data.get('styles', '[]'))
    removal_prompts = json.loads(data.get('removal_prompts', '[]'))
    
    if not all([book_name, api_key, styles]):
        return jsonify({'error': 'Dados incompletos'}), 400
    
    # Salva arquivo temporariamente
    filename = secure_filename(file.filename)
    file_path = os.path.join(Config.UPLOAD_DIR, filename)
    file.save(file_path)
    
    try:
        # Processa documento
        processor = WordStylerProcessor()
        result = processor.process_document(
            file_path, book_name, api_key, styles, removal_prompts
        )
        
        # Remove arquivo temporário
        os.remove(file_path)
        
        return jsonify(result)
        
    except Exception as e:
        # Remove arquivo em caso de erro
        if os.path.exists(file_path):
            os.remove(file_path)
        return jsonify({'error': str(e)}), 500

@app.route('/api/download/<path:filename>', methods=['GET'])
def download_file(filename):
    """Endpoint para download de arquivos"""
    file_path = os.path.join(Config.OUTPUT_DIR, filename)
    
    if not os.path.exists(file_path):
        return jsonify({'error': 'Arquivo não encontrado'}), 404
    
    return send_file(file_path, as_attachment=True)

@app.route('/api/health', methods=['GET'])
def health_check():
    """Verifica se a API está funcionando"""
    return jsonify({
        'status': 'ok',
        'message': 'Word AI Styler API está funcionando'
    })

if __name__ == '__main__':
    Config.create_directories()
    app.run(debug=True, port=5000)