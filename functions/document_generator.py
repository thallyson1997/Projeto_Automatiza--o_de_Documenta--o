from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import io
import os


def gerar_documento(unidade, data, legenda, imagens=None):
    """
    Carrega o modelo.docx, substitui os placeholders e insere as imagens
    e retorna o documento modificado.
    
    Args:
        unidade (str): Texto para substituir [UNIDADE]
        data (str): Data formatada (DD.MM.YYYY) para substituir [DATA]
        legenda (str): Texto para substituir [LEGENDA]
        imagens (list): Lista com até 4 imagens em bytes
    
    Returns:
        bytes: Documento Word em bytes
    """
    try:
        # Caminho do modelo
        modelo_path = os.path.join(os.path.dirname(__file__), '..', 'documento', 'modelo.docx')
        
        # Verifica se o arquivo modelo existe
        if not os.path.exists(modelo_path):
            raise FileNotFoundError(f"Arquivo modelo não encontrado em {modelo_path}")
        
        # Abre o documento modelo
        doc = Document(modelo_path)
        
        if imagens is None:
            imagens = []
        
        # Substitui placeholders em parágrafos top-level
        for paragrafo in doc.paragraphs:
            texto_completo = paragrafo.text
            
            if '[UNIDADE]' in texto_completo or '[DATA]' in texto_completo or '[LEGENDA]' in texto_completo:
                texto_completo = texto_completo.replace('[UNIDADE]', unidade)
                texto_completo = texto_completo.replace('[DATA]', data)
                texto_completo = texto_completo.replace('[LEGENDA]', legenda)
                
                for run in paragrafo.runs:
                    run.text = ''
                
                if paragrafo.runs:
                    paragrafo.runs[0].text = texto_completo
                else:
                    paragrafo.add_run(texto_completo)
        
        # Procura e substitui em tabelas
        for tabela in doc.tables:
            for linha in tabela.rows:
                for celula in linha.cells:
                    for paragrafo in celula.paragraphs:
                        texto_completo = paragrafo.text
                        
                        # Se encontrar [IMAGENS], insere as imagens nessa célula
                        if '[IMAGENS]' in texto_completo:
                            # Limpa o parágrafo
                            for run in paragrafo.runs:
                                run.text = ''
                            
                            # Insere as imagens
                            if imagens:
                                inserir_imagens_na_celula(celula, imagens)
                        
                        # Substitui outros placeholders
                        elif '[UNIDADE]' in texto_completo or '[DATA]' in texto_completo or '[LEGENDA]' in texto_completo:
                            texto_completo = texto_completo.replace('[UNIDADE]', unidade)
                            texto_completo = texto_completo.replace('[DATA]', data)
                            texto_completo = texto_completo.replace('[LEGENDA]', legenda)
                            
                            for run in paragrafo.runs:
                                run.text = ''
                            
                            if paragrafo.runs:
                                paragrafo.runs[0].text = texto_completo
                            else:
                                paragrafo.add_run(texto_completo)
        
        # Converte o documento para bytes
        doc_bytes = io.BytesIO()
        doc.save(doc_bytes)
        doc_bytes.seek(0)
        
        return doc_bytes.getvalue()
    
    except Exception as e:
        raise Exception(f"Erro ao gerar documento: {str(e)}")


def inserir_imagens_na_celula(celula, imagens):
    """
    Insere as imagens em uma célula de tabela de forma apropriada:
    - 1 imagem: centralizada, com altura 1/3 da célula e largura 2/3
    - 2 imagens: uma em cima da outra
    - 3 imagens: 2 em cima, 1 embaixo
    - 4 imagens: 2 em cima, 2 embaixo
    
    Args:
        celula: Célula da tabela do documento Word
        imagens: Lista com bytes das imagens
    """
    num_imagens = len(imagens)
    
    # Limpa a célula
    for paragrafo in celula.paragraphs:
        for run in paragrafo.runs:
            run.text = ''
    
    # Obtém o primeiro parágrafo da célula
    primeiro_paragrafo = celula.paragraphs[0]
    primeiro_paragrafo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # Remove espaçamento do parágrafo
    remover_espacamento_paragrafo(primeiro_paragrafo)
    
    if num_imagens == 1:
        # 1 imagem: 9.6cm x 7.5cm
        # Centraliza verticalmente a célula
        centralizar_celula_verticalmente(celula)
        
        # Tamanhos fixos em polegadas (convertidos de cm)
        largura_imagem = Inches(9.6 / 2.54)  # 9.6 cm em polegadas
        altura_imagem = Inches(7.5 / 2.54)   # 7.5 cm em polegadas
        
        adicionar_imagem_com_dimensoes(primeiro_paragrafo, imagens[0], largura_imagem, altura_imagem)
    
    elif num_imagens == 2:
        # 2 imagens uma em cima da outra: 9.6cm x 7.5cm cada
        centralizar_celula_verticalmente(celula)
        
        # Tamanhos fixos em polegadas (convertidos de cm)
        largura_imagem = Inches(9.6 / 2.54)  # 9.6 cm em polegadas
        altura_imagem = Inches(7.5 / 2.54)   # 7.5 cm em polegadas
        
        adicionar_imagem_com_dimensoes(primeiro_paragrafo, imagens[0], largura_imagem, altura_imagem)
        
        # Adiciona espaço entre as imagens
        espaco_p = celula.add_paragraph()
        espaco_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        remover_espacamento_paragrafo(espaco_p)
        espaco_p.paragraph_format.space_after = Pt(6)  # Espaço de 6 pontos
        
        novo_p = celula.add_paragraph()
        novo_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        remover_espacamento_paragrafo(novo_p)
        adicionar_imagem_com_dimensoes(novo_p, imagens[1], largura_imagem, altura_imagem)
    
    elif num_imagens == 3:
        # 3 imagens: 2 em cima (lado a lado), 1 embaixo
        # Cada imagem: 6.8cm x 7.9cm
        centralizar_celula_verticalmente(celula)
        
        # Tamanhos fixos em polegadas (convertidos de cm)
        largura_imagem = Inches(6.8 / 2.54)  # 6.8 cm em polegadas
        altura_imagem = Inches(7.9 / 2.54)   # 7.9 cm em polegadas
        
        # Primeira linha com 2 imagens lado a lado
        adicionar_imagem_com_dimensoes(primeiro_paragrafo, imagens[0], largura_imagem, altura_imagem)
        primeiro_paragrafo.add_run('  ')
        adicionar_imagem_com_dimensoes(primeiro_paragrafo, imagens[1], largura_imagem, altura_imagem)
        primeiro_paragrafo.paragraph_format.space_after = Pt(6)
        
        # Segunda linha com 1 imagem centralizada
        novo_p = celula.add_paragraph()
        novo_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        remover_espacamento_paragrafo(novo_p)
        adicionar_imagem_com_dimensoes(novo_p, imagens[2], largura_imagem, altura_imagem)
    
    elif num_imagens >= 4:
        # 4+ imagens: 2 em cima (lado a lado), 2 embaixo (lado a lado)
        # Cada imagem: 6.8cm x 7.9cm
        centralizar_celula_verticalmente(celula)
        
        # Tamanhos fixos em polegadas (convertidos de cm)
        largura_imagem = Inches(6.8 / 2.54)  # 6.8 cm em polegadas
        altura_imagem = Inches(7.9 / 2.54)   # 7.9 cm em polegadas
        
        # Primeira linha com 2 imagens lado a lado
        adicionar_imagem_com_dimensoes(primeiro_paragrafo, imagens[0], largura_imagem, altura_imagem)
        primeiro_paragrafo.add_run('  ')
        adicionar_imagem_com_dimensoes(primeiro_paragrafo, imagens[1], largura_imagem, altura_imagem)
        primeiro_paragrafo.paragraph_format.space_after = Pt(6)
        
        # Segunda linha com 2 imagens lado a lado
        novo_p = celula.add_paragraph()
        novo_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        remover_espacamento_paragrafo(novo_p)
        adicionar_imagem_com_dimensoes(novo_p, imagens[2], largura_imagem, altura_imagem)
        novo_p.add_run('  ')
        adicionar_imagem_com_dimensoes(novo_p, imagens[3], largura_imagem, altura_imagem)


def obter_largura_celula(celula):
    """
    Obtém a largura da célula em polegadas.
    Se não conseguir obter, retorna um valor padrão.
    """
    try:
        tcPr = celula._element.tcPr
        if tcPr is not None:
            tcW = tcPr.find(qn('w:tcW'))
            if tcW is not None:
                # Obtém o valor em twips (1/20 de um ponto)
                w_val = tcW.get(qn('w:w'))
                if w_val:
                    # Converte twips para polegadas (1440 twips = 1 polegada)
                    twips = int(w_val)
                    polegadas = twips / 1440
                    return polegadas
    except:
        pass
    
    # Retorna valor padrão se não conseguir obter
    return 3.0  # 3 polegadas como padrão


def adicionar_imagem_com_dimensoes(paragrafo, imagem_bytes, largura, altura):
    """
    Adiciona uma imagem com largura e altura específicas em polegadas.
    
    Args:
        paragrafo: Parágrafo onde adicionar a imagem
        imagem_bytes: Bytes da imagem
        largura: Largura em polegadas
        altura: Altura em polegadas
    """
    imagem_stream = io.BytesIO(imagem_bytes)
    
    run = paragrafo.add_run()
    picture = run.add_picture(imagem_stream, width=largura, height=altura)


def centralizar_celula_verticalmente(celula):
    """Centraliza o conteúdo da célula verticalmente"""
    tcPr = celula._element.tcPr
    if tcPr is None:
        tcPr = OxmlElement('w:tcPr')
        celula._element.insert(0, tcPr)
    
    # Remove vAlign existente se houver
    vAlign = tcPr.find(qn('w:vAlign'))
    if vAlign is not None:
        tcPr.remove(vAlign)
    
    # Adiciona vAlign com valor center
    vAlign = OxmlElement('w:vAlign')
    vAlign.set(qn('w:val'), 'center')
    tcPr.append(vAlign)


def remover_espacamento_paragrafo(paragrafo):
    """Remove espaçamento antes e depois do parágrafo"""
    pPr = paragrafo._element.get_or_add_pPr()
    pSpacing = pPr.find(qn('w:spacing'))
    
    if pSpacing is None:
        pSpacing = OxmlElement('w:spacing')
        pPr.append(pSpacing)
    
    # Define espaçamento antes e depois como 0
    pSpacing.set(qn('w:before'), '0')
    pSpacing.set(qn('w:after'), '0')
    pSpacing.set(qn('w:line'), '240')  # Single spacing (240 twips)
    pSpacing.set(qn('w:lineRule'), 'auto')


def adicionar_imagem_ao_paragrafo(paragrafo, imagem_bytes, largura_inches):
    """Adiciona uma imagem a um parágrafo"""
    imagem_stream = io.BytesIO(imagem_bytes)
    paragrafo.add_run().add_picture(imagem_stream, width=Inches(largura_inches))
