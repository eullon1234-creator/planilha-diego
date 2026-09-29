import os
import sys
import uuid
from flask import Flask, render_template, request, jsonify, send_file
from werkzeug.utils import secure_filename
from excel_processor import process_inventory_excel

# Inicialização do Flask
app = Flask(__name__, template_folder='.', static_folder='static')
app.config['SECRET_KEY'] = 'planilha-diego-secret-key-2026'

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, 'uploads')
OUTPUT_FOLDER = os.path.join(BASE_DIR, 'outputs')

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['OUTPUT_FOLDER'] = OUTPUT_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # Até 100MB

ALLOWED_EXTENSIONS = {'xlsx', 'xls'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({'success': False, 'error': 'Nenhum arquivo enviado.'}), 400

    file = request.files['file']
    if file.filename == '':
        return jsonify({'success': False, 'error': 'Nome de arquivo inválido.'}), 400

    if not allowed_file(file.filename):
        return jsonify({
            'success': False,
            'error': 'Formato não suportado. Por favor, envie uma planilha em formato .xlsx ou .xls baixada do sistema.'
        }), 400

    try:
        # Salvar arquivo original temporariamente com nome seguro
        raw_filename = secure_filename(file.filename)
        unique_prefix = uuid.uuid4().hex[:8]
        safe_filename = f"{unique_prefix}_{raw_filename}"
        input_filepath = os.path.join(app.config['UPLOAD_FOLDER'], safe_filename)
        file.save(input_filepath)

        # Processar com o motor inteligente
        output_filepath, stats = process_inventory_excel(input_filepath, app.config['OUTPUT_FOLDER'])
        out_filename = os.path.basename(output_filepath)

        return jsonify({
            'success': True,
            'message': 'Planilha processada com sucesso!',
            'download_url': f'/download/{out_filename}',
            'stats': stats
        })

    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({
            'success': False,
            'error': f'Erro ao processar o inventário: {str(e)}'
        }), 500

@app.route('/download/<filename>')
def download_file(filename):
    safe_name = secure_filename(filename)
    filepath = os.path.join(app.config['OUTPUT_FOLDER'], safe_name)
    if not os.path.exists(filepath):
        return "Arquivo não encontrado ou expirado.", 404
    return send_file(filepath, as_attachment=True, download_name=safe_name)

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    print(f"=======================================================")
    print(f"       PLANILHA DIEGO - APLICATIVO INICIADO            ")
    print(f"  Acesse no navegador: http://localhost:{port}        ")
    print(f"=======================================================")
    app.run(host='0.0.0.0', port=port, debug=False)
