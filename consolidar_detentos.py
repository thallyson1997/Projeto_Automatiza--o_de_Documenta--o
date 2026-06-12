import pandas as pd
import glob
import sys
from pathlib import Path
from datetime import datetime
import warnings

# Suprimir avisos do openpyxl
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

print("PROGRAMA EXECUTANDO...")

# Caminho da pasta com os arquivos Excel
pasta_entrada = r"c:\Users\thallyson.fontenelle\Documents\GitHub\programas\entrada"
pasta_saida = r"c:\Users\thallyson.fontenelle\Documents\GitHub\programas\saida"

# Encontrar todos os arquivos .xlsx na pasta de entrada
arquivos = sorted(glob.glob(f"{pasta_entrada}\\quantitativoDiario-*.xlsx"))

# Lista para armazenar os dados consolidados
dados_consolidados = []
dia_count = 0

# Processar cada arquivo
for arquivo in arquivos:
    try:
        # Extrair a data do nome do arquivo (formato: quantitativoDiario-DD-MM-YYYY.xlsx)
        nome_arquivo = Path(arquivo).stem
        data_str = nome_arquivo.replace("quantitativoDiario-", "")
        
        # Ler o Excel, pulando a primeira linha (título)
        df = pd.read_excel(arquivo, header=0, skiprows=1)
        
        # Adicionar a coluna de data extraída do nome do arquivo
        df['DATA_ARQUIVO'] = pd.to_datetime(data_str, format="%d-%m-%Y")
        
        # Agrupar por UNIDADE PRISIONAL e contar o número de detentos
        agrupado = df.groupby(['DATA_ARQUIVO', 'UNIDADE PRISIONAL']).size().reset_index(name='QUANTIDADE')
        
        dados_consolidados.append(agrupado)
        dia_count += 1
        print(f"DIA {dia_count} EXECUTADO...")
        
    except Exception as e:
        # Erros são ignorados para não poluir a saída
        pass

# Consolidar todos os dados
if dados_consolidados:
    resultado = pd.concat(dados_consolidados, ignore_index=True)
    
    # Pivotar os dados: DATA nas linhas, UNIDADE PRISIONAL nas colunas
    tabela_resultado = resultado.pivot_table(
        index='DATA_ARQUIVO',
        columns='UNIDADE PRISIONAL',
        values='QUANTIDADE',
        aggfunc='sum'
    )
    
    # Ordenar por data
    tabela_resultado = tabela_resultado.sort_index()
    
    # Renomear o índice para uma label mais clara
    tabela_resultado.index.name = 'DATA'
    
    # Gerar nome do arquivo de saída dinamicamente
    primeira_data = tabela_resultado.index[0]
    nome_mes = primeira_data.strftime('%B').upper()
    ano = primeira_data.year
    
    # Mapear nomes de meses para português
    meses_pt = {
        'JANUARY': 'JANEIRO', 'FEBRUARY': 'FEVEREIRO', 'MARCH': 'MARCO', 'APRIL': 'ABRIL',
        'MAY': 'MAIO', 'JUNE': 'JUNHO', 'JULY': 'JULHO', 'AUGUST': 'AGOSTO',
        'SEPTEMBER': 'SETEMBRO', 'OCTOBER': 'OUTUBRO', 'NOVEMBER': 'NOVEMBRO', 'DECEMBER': 'DEZEMBRO'
    }
    nome_mes_pt = meses_pt.get(nome_mes, nome_mes)

    arquivo_saida = f"{pasta_saida}\\SIISP-{nome_mes_pt}-{ano}.xlsx"
    tabela_resultado.to_excel(arquivo_saida)

print("FIM DA EXECUÇÃO.")
