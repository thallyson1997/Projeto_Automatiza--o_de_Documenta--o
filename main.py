from flask import Flask, render_template, request, send_file
from functions.document_generator import gerar_documento
import io

app = Flask(__name__, template_folder='templates')


@app.route('/')
def index():
    """Rota da página inicial"""
    return render_template('index.html')


@app.route('/upload')
def upload():
    """Rota da página de upload de documentação"""
    return render_template('upload.html')


@app.route('/gerar-documento', methods=['POST'])
def gerar_doc():
    """Rota para gerar e fazer download do documento"""
    try:
        # Obtém os campos (obrigatórios)
        unidade = request.form.get('unidade', '').strip()
        data_input = request.form.get('data', '').strip()
        legenda = request.form.get('legenda', '').strip()
        
        # Valida campos obrigatórios
        if not unidade or not data_input or not legenda:
            return {'erro': 'Todos os campos devem ser preenchidos'}, 400
        
        # Formata a data
        data_formatada = convertar_data(data_input)
        
        # Processa as imagens (obrigatórias)
        imagens = []
        if 'imagens' not in request.files or len(request.files.getlist('imagens')) == 0:
            return {'erro': 'Pelo menos uma imagem deve ser enviada'}, 400
        
        arquivos = request.files.getlist('imagens')
        
        # Limita a 4 imagens
        if len(arquivos) > 4:
            return {'erro': 'Máximo de 4 imagens permitidas'}, 400
        
        for arquivo in arquivos:
            if arquivo and arquivo.filename != '':
                imagens.append(arquivo.read())
        
        if not imagens:
            return {'erro': 'Pelo menos uma imagem válida deve ser enviada'}, 400
        
        # Gera o documento
        documento_bytes = gerar_documento(unidade, data_formatada, legenda, imagens)
        
        # Retorna o documento como download
        return send_file(
            io.BytesIO(documento_bytes),
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            as_attachment=True,
            download_name='documento.docx'
        )
    
    except Exception as e:
        return {'erro': str(e)}, 500


def convertar_data(data_iso):
    """Converte data de formato ISO (YYYY-MM-DD) para DD.MM.YYYY"""
    try:
        from datetime import datetime
        data_obj = datetime.strptime(data_iso, '%Y-%m-%d')
        return data_obj.strftime('%d.%m.%Y')
    except Exception as e:
        raise Exception(f"Erro ao formatar data: {str(e)}")


if __name__ == '__main__':
    app.run(debug=True)
