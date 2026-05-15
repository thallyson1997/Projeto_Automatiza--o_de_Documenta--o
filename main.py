from flask import Flask, render_template, request, send_file, jsonify
from functions.document_generator import gerar_documento, gerar_documento_multiplo
from functions.document_generator2 import gerar_documento_modelo2_empresa
from functions.document_generator3 import gerar_documento_modelo3_alipen
from docx2pdf import convert
import io
import re
import tempfile
import os
import platform

app = Flask(__name__, template_folder='templates')

# Rota para pré-visualização em PDF do documento
@app.route('/preview-documento-pdf', methods=['POST'])
def preview_documento_pdf():
    """Gera um PDF temporário do documento preenchido e retorna uma URL para visualização"""
    try:
        com_initialized = False
        if platform.system() == 'Windows':
            import pythoncom
            pythoncom.CoInitialize()
            com_initialized = True

        formularios_preview = []

        if request.is_json:
            data = request.get_json() or {}
            formularios_input = data.get('formularios', [])

            for form in formularios_input:
                unidade = (form.get('unidade') or '').strip() or '[UNIDADE]'
                data_input = (form.get('data') or '').strip()
                data_str = convertar_data(data_input) if data_input else '[DATA]'
                legenda = (form.get('legenda') or '').strip() or '[LEGENDA]'

                formularios_preview.append({
                    'unidade': unidade,
                    'data': data_str,
                    'legenda': legenda,
                    'imagens': []
                })
        else:
            indices = set()

            for key in request.form.keys():
                match = re.match(r'^(unidade|data|legenda)-(\d+)$', key)
                if match:
                    indices.add(int(match.group(2)))

            for key in request.files.keys():
                match = re.match(r'^imagens-(\d+)$', key)
                if match:
                    indices.add(int(match.group(1)))

            for idx in sorted(indices):
                unidade = (request.form.get(f'unidade-{idx}', '') or '').strip() or '[UNIDADE]'
                data_input = (request.form.get(f'data-{idx}', '') or '').strip()
                data_str = convertar_data(data_input) if data_input else '[DATA]'
                legenda = (request.form.get(f'legenda-{idx}', '') or '').strip() or '[LEGENDA]'

                imagens = []
                arquivos = request.files.getlist(f'imagens-{idx}')
                for arquivo in arquivos[:4]:
                    if arquivo and arquivo.filename:
                        arquivo.seek(0)
                        conteudo = arquivo.read()
                        if conteudo:
                            imagens.append(conteudo)

                formularios_preview.append({
                    'unidade': unidade,
                    'data': data_str,
                    'legenda': legenda,
                    'imagens': imagens
                })

        if not formularios_preview:
            formularios_preview = [{
                'unidade': '[UNIDADE]',
                'data': '[DATA]',
                'legenda': '[LEGENDA]',
                'imagens': []
            }]

        # Gera DOCX de 1 ou varias paginas conforme quantidade de formularios.
        if len(formularios_preview) == 1:
            form = formularios_preview[0]
            docx_bytes = gerar_documento(form['unidade'], form['data'], form['legenda'], form['imagens'])
        else:
            docx_bytes = gerar_documento_multiplo(formularios_preview)
        with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp_docx:
            tmp_docx.write(docx_bytes)
            tmp_docx_path = tmp_docx.name

        # Gerar PDF temporário
        tmp_pdf_path = tmp_docx_path.replace('.docx', '.pdf')
        convert(tmp_docx_path, tmp_pdf_path)

        # Guardar caminho para servir depois
        filename = os.path.basename(tmp_pdf_path)
        # Salva em pasta temporária pública
        static_tmp_dir = os.path.join('static', 'tmp')
        os.makedirs(static_tmp_dir, exist_ok=True)
        final_pdf_path = os.path.join(static_tmp_dir, filename)
        os.replace(tmp_pdf_path, final_pdf_path)
        os.unlink(tmp_docx_path)

        # Retorna URL para o PDF
        url = f'/static/tmp/{filename}'
        return jsonify({'pdf_url': url})
    except Exception as e:
        return jsonify({'erro': f'Erro ao gerar PDF: {str(e)}'}), 500
    finally:
        if platform.system() == 'Windows':
            try:
                if com_initialized:
                    pythoncom.CoUninitialize()
            except Exception:
                pass



# Rota para pré-visualização do documento
@app.route('/preview-documento', methods=['POST'])
def preview_documento():
    """Gera uma pré-visualização HTML do documento com os dados enviados"""
    try:
        data = request.get_json()
        formularios = data.get('formularios', [])
        if not formularios:
            # Modelo vazio
            unidade = '[UNIDADE]'
            data_str = '[DATA]'
            legenda = '[LEGENDA]'
            html = render_template('preview_modelo.html', unidade=unidade, data=data_str, legenda=legenda)
            return html
        # Só mostra o primeiro formulário para preview
        form = formularios[0]
        unidade = form.get('unidade') or '[UNIDADE]'
        data_input = (form.get('data') or '').strip()
        data_str = convertar_data(data_input) if data_input else '[DATA]'
        legenda = form.get('legenda') or '[LEGENDA]'
        html = render_template('preview_modelo.html', unidade=unidade, data=data_str, legenda=legenda)
        return html
    except Exception as e:
        return f'<div class="preview-placeholder">Erro ao gerar pré-visualização: {str(e)}</div>'


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


@app.route('/upload3')
def upload3():
    """Rota da página de upload para relatório ALIPEN"""
    return render_template('upload3.html')


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


@app.route('/gerar-documento3', methods=['POST'])
def gerar_doc_modelo3():
    """Rota para gerar modelo3 (ALIPEN) com 4 campos: café, lanche, almoço, jantar"""
    try:
        from datetime import datetime

        indices_formulario = []
        for key in request.form.keys():
            match = re.fullmatch(r'data_formulario-(\d+)', key)
            if match:
                indices_formulario.append(int(match.group(1)))

        if not indices_formulario:
            return {'erro': 'Pelo menos um formulário deve existir.'}, 400

        total_formularios = max(indices_formulario) + 1
        datas_formulario = []
        unidades_formulario = []

        for form_index in range(total_formularios):
            data_formulario_iso = request.form.get(f'data_formulario-{form_index}', '').strip()
            unidade_formulario = request.form.get(f'unidade_formulario-{form_index}', '').strip()

            if not data_formulario_iso:
                datas_formulario.append('')
                unidades_formulario.append(unidade_formulario)
                continue

            try:
                data_formulario_obj = datetime.strptime(data_formulario_iso, '%Y-%m-%d')
            except Exception:
                return {'erro': f'Formato da Data do Formulário {form_index + 1} inválido. Use AAAA-MM-DD.'}, 400

            datas_formulario.append(data_formulario_obj.strftime('%d/%m/%Y'))
            unidades_formulario.append(unidade_formulario)

        imagens_formularios = []
        for index in range(len(datas_formulario)):
            imagem_cafe = request.files.get(f'imagem_cafe-{index}')
            imagem_lanche = request.files.get(f'imagem_lanche-{index}')
            imagem_almoco = request.files.get(f'imagem_almoco-{index}')
            imagem_jantar = request.files.get(f'imagem_jantar-{index}')

            dados_form = {
                'imagem_cafe': imagem_cafe.read() if imagem_cafe and imagem_cafe.filename else None,
                'legenda_cafe': request.form.get(f'legenda_cafe-{index}', '').strip(),
                'imagem_lanche': imagem_lanche.read() if imagem_lanche and imagem_lanche.filename else None,
                'legenda_lanche': request.form.get(f'legenda_lanche-{index}', '').strip(),
                'imagem_almoco': imagem_almoco.read() if imagem_almoco and imagem_almoco.filename else None,
                'proteina_almoco': request.form.get(f'proteina_almoco-{index}', '').strip(),
                'peso_almoco': request.form.get(f'peso_almoco-{index}', '').strip(),
                'acompanhamento_almoco': request.form.get(f'acompanhamento_almoco-{index}', '').strip(),
                'imagem_jantar': imagem_jantar.read() if imagem_jantar and imagem_jantar.filename else None,
                'proteina_jantar': request.form.get(f'proteina_jantar-{index}', '').strip(),
                'peso_jantar': request.form.get(f'peso_jantar-{index}', '').strip(),
                'acompanhamento_jantar': request.form.get(f'acompanhamento_jantar-{index}', '').strip()
            }
            imagens_formularios.append(dados_form)

        documento_bytes = gerar_documento_modelo3_alipen(
            datas_formulario,
            unidades_formulario,
            imagens_formularios
        )

        return send_file(
            io.BytesIO(documento_bytes),
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            as_attachment=True,
            download_name='modelo3.docx'
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
