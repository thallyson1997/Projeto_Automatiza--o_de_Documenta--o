from flask import Flask, render_template, request, send_file
from functions.document_generator import gerar_documento, gerar_documento_multiplo
from functions.document_generator2 import gerar_documento_modelo2_empresa
import io
import re

app = Flask(__name__, template_folder='templates')


@app.route('/')
def index():
    """Rota da página inicial"""
    return render_template('index.html')


@app.route('/upload')
def upload():
    """Rota da página de upload de documentação"""
    return render_template('upload.html')


@app.route('/upload2')
def upload2():
    """Rota da página de upload para fiscalização cozinha"""
    return render_template('upload2.html')


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
                    # Lê os bytes do arquivo e validação
                    arquivo.seek(0)  # Reseta o ponteiro para o início
                    arquivo_bytes = arquivo.read()
                    if len(arquivo_bytes) > 0:
                        imagens.append(arquivo_bytes)
            
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
        
        # Se houver apenas 1 formulário, usa gerar_documento
        if len(formularios) == 1:
            form = formularios[0]
            documento_bytes = gerar_documento(form['unidade'], form['data'], form['legenda'], form['imagens'])
        else:
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


@app.route('/gerar-documento2', methods=['POST'])
def gerar_doc_modelo2():
    """Rota para gerar modelo2 com capa fixa e secoes de formularios repetidas"""
    try:
        from datetime import datetime

        empresa = request.form.get('empresa', '').strip()
        if not empresa:
            empresa = request.form.get('unidade-0', '').strip()

        if not empresa:
            return {'erro': 'O campo Empresa deve ser preenchido'}, 400

        data_inicio_iso = request.form.get('data_inicio', '').strip()
        data_fim_iso = request.form.get('data_fim', '').strip()

        if not data_inicio_iso or not data_fim_iso:
            return {'erro': 'Os campos de Período (Data Início e Data Fim) devem ser preenchidos'}, 400

        try:
            data_inicio_obj = datetime.strptime(data_inicio_iso, '%Y-%m-%d')
            data_fim_obj = datetime.strptime(data_fim_iso, '%Y-%m-%d')
        except Exception:
            return {'erro': 'Formato de data inválido. Use o formato AAAA-MM-DD.'}, 400

        if data_inicio_obj > data_fim_obj:
            return {'erro': 'A Data Início não pode ser maior que a Data Fim.'}, 400

        indices_formulario = []
        for key in request.form.keys():
            match = re.fullmatch(r'data_formulario-(\d+)', key)
            if match:
                indices_formulario.append(int(match.group(1)))

        if not indices_formulario:
            return {'erro': 'Pelo menos um formulário deve existir.'}, 400

        total_formularios = max(indices_formulario) + 1
        datas_formulario = []

        for form_index in range(total_formularios):
            data_formulario_iso = request.form.get(f'data_formulario-{form_index}', '').strip()

            if not data_formulario_iso:
                datas_formulario.append('')
                continue

            try:
                data_formulario_obj = datetime.strptime(data_formulario_iso, '%Y-%m-%d')
            except Exception:
                return {'erro': f'Formato da Data do Formulário {form_index + 1} inválido. Use AAAA-MM-DD.'}, 400

            if data_formulario_obj < data_inicio_obj or data_formulario_obj > data_fim_obj:
                return {'erro': f'A data do Formulário {form_index + 1} deve estar entre Data Início e Data Fim.'}, 400

            datas_formulario.append(data_formulario_obj.strftime('%d/%m/%Y'))

        data_inicio = data_inicio_obj.strftime('%d/%m/%Y')
        data_fim = data_fim_obj.strftime('%d/%m/%Y')

        imagens_formularios = []
        for index in range(len(datas_formulario)):
            imagem_lanche = request.files.get(f'imagem_lanche-{index}')
            imagem_ceia = request.files.get(f'imagem_ceia-{index}')

            imagem_lanche_bytes = None
            imagem_ceia_bytes = None

            if imagem_lanche and imagem_lanche.filename:
                imagem_lanche_bytes = imagem_lanche.read()

            if imagem_ceia and imagem_ceia.filename:
                imagem_ceia_bytes = imagem_ceia.read()

            dados_form = {
                'imagem_lanche': imagem_lanche_bytes,
                'imagem_ceia': imagem_ceia_bytes,
                'legenda_lanche': request.form.get(f'legenda_lanche-{index}', '').strip(),
                'legenda_ceia': request.form.get(f'legenda_ceia-{index}', '').strip()
            }
            for n in range(1, 5):
                f_almoco = request.files.get(f'imagem_almoco_{n}-{index}')
                dados_form[f'imagem_almoco_{n}'] = f_almoco.read() if f_almoco and f_almoco.filename else None
                dados_form[f'proteina_almoco_{n}'] = request.form.get(f'proteina_almoco_{n}-{index}', '').strip()
                dados_form[f'peso_almoco_{n}'] = request.form.get(f'peso_almoco_{n}-{index}', '').strip()
                if n <= 2:
                    dados_form[f'acompanhamento_almoco_{n}'] = request.form.get(f'acompanhamento_almoco_{n}-{index}', '').strip()
                f_jantar = request.files.get(f'imagem_jantar_{n}-{index}')
                dados_form[f'imagem_jantar_{n}'] = f_jantar.read() if f_jantar and f_jantar.filename else None
                dados_form[f'proteina_jantar_{n}'] = request.form.get(f'proteina_jantar_{n}-{index}', '').strip()
                dados_form[f'peso_jantar_{n}'] = request.form.get(f'peso_jantar_{n}-{index}', '').strip()
                if n <= 2:
                    dados_form[f'acompanhamento_jantar_{n}'] = request.form.get(f'acompanhamento_jantar_{n}-{index}', '').strip()
            imagens_formularios.append(dados_form)

        documento_bytes = gerar_documento_modelo2_empresa(
            empresa,
            data_inicio,
            data_fim,
            datas_formulario,
            imagens_formularios
        )

        return send_file(
            io.BytesIO(documento_bytes),
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            as_attachment=True,
            download_name='modelo2.docx'
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
    app.run(host='0.0.0.0', port=10000, debug=True)
