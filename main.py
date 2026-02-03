from flask import Flask, render_template, request, send_file
from functions.document_generator import gerar_documento, gerar_documento_multiplo
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
    """Rota para gerar e fazer download do documento único com múltiplas páginas"""
    try:
        # Coleta todos os formulários do request
        formularios = []
        form_index = 0
        
        # Procura por formulários no formato formulario-0, formulario-1, etc
        while True:
            unidade = request.form.get(f'unidade-{form_index}', '').strip()
            data_input = request.form.get(f'data-{form_index}', '').strip()
            legenda = request.form.get(f'legenda-{form_index}', '').strip()
            
            if not unidade and not data_input and not legenda:
                break
            
            # Valida campos obrigatórios do formulário
            if not unidade or not data_input or not legenda:
                return {'erro': f'Todos os campos do formulário {form_index + 1} devem ser preenchidos'}, 400
            
            # Formata a data
            data_formatada = convertar_data(data_input)
            
            # Processa as imagens deste formulário
            imagens = []
            imagens_key = f'imagens-{form_index}'
            
            if imagens_key not in request.files or len(request.files.getlist(imagens_key)) == 0:
                return {'erro': f'Pelo menos uma imagem deve ser enviada no formulário {form_index + 1}'}, 400
            
            arquivos = request.files.getlist(imagens_key)
            
            # Limita a 4 imagens
            if len(arquivos) > 4:
                return {'erro': f'Máximo de 4 imagens permitidas no formulário {form_index + 1}'}, 400
            
            for arquivo in arquivos:
                if arquivo and arquivo.filename != '':
                    imagens.append(arquivo.read())
            
            if not imagens:
                return {'erro': f'Pelo menos uma imagem válida deve ser enviada no formulário {form_index + 1}'}, 400
            
            formularios.append({
                'unidade': unidade,
                'data': data_formatada,
                'legenda': legenda,
                'imagens': imagens
            })
            
            form_index += 1
        
        if not formularios:
            return {'erro': 'Pelo menos um formulário deve ser preenchido'}, 400
        
        # Gera o documento único com múltiplas páginas
        documento_bytes = gerar_documento_multiplo(formularios)
        
        # Retorna o documento como download
        return send_file(
            io.BytesIO(documento_bytes),
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            as_attachment=True,
            download_name='documentacao.docx'
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
