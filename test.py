import os
import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from datetime import datetime

def extrair_dados_excel(caminho_excel):
    # Carrega o arquivo Excel
    df = pd.read_excel(caminho_excel)

    # Filtra as colunas necessárias
    colunas_necessarias = ['nome', 'valor', 'contato', 'descrição']
    df_filtrado = df[colunas_necessarias]

    return df_filtrado

def criar_recibo_pdf(linha, caminho_pdf):
    # Configura o tamanho da página como letter em formato paisagem
    altura, largura = letter
    c = canvas.Canvas(caminho_pdf, pagesize=(largura, altura))

    # Configura a fonte
    c.setFont("Helvetica", 12)

    # Calcula a posição central na página
    x_centro = largura / 2
    y_centro = altura / 2

    # Escreve os dados do recibo no PDF
    c.drawString(x_centro, y_centro, f"Nome: {linha['nome']}")
    c.drawString(x_centro, y_centro - 15, f"Valor: {linha['valor']}")
    c.drawString(x_centro, y_centro - 30, f"Contato: {linha['contato']}")
    c.drawString(x_centro, y_centro - 45, f"Descrição: {linha['descrição']}")
    
    # Salva o arquivo PDF
    c.save()

def criar_pasta_data_atual():
    # Obtém a data atual
    data_atual = datetime.now().strftime("%Y-%m-%d")
    
    # Caminho para a pasta no Desktop com a data atual
    caminho_pasta = os.path.join(os.path.expanduser('~'), 'Desktop', f'Recibos_{data_atual}')
    
    # Cria a pasta se não existir
    os.makedirs(caminho_pasta, exist_ok=True)
    
    return caminho_pasta

if __name__ == "__main__":
    # Caminho real do seu arquivo Excel
    caminho_excel = '/home/pv-lds/Desktop/base.xlsx'

    # Extrai dados do Excel
    dados_recibo = extrair_dados_excel(caminho_excel)

    # Cria uma pasta com a data atual no Desktop
    pasta_recibos = criar_pasta_data_atual()

    # Itera sobre os dados e gera um PDF por recibo na pasta criada
    for indice, linha in dados_recibo.iterrows():
        # Adota o nome da pessoa como nome do arquivo PDF
        nome_arquivo_pdf = f"{linha['nome']}_recibo.pdf"
        
        # Caminho completo para o arquivo PDF na pasta
        caminho_arquivo_pdf = os.path.join(pasta_recibos, nome_arquivo_pdf)

        # Cria o PDF com os dados do recibo
        criar_recibo_pdf(linha, caminho_arquivo_pdf)

