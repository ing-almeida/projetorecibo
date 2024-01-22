import os
import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from datetime import datetime
from PIL import Image

def extrair_dados_excel(caminho_excel):
    # Carrega o arquivo Excel
    df = pd.read_excel(caminho_excel)

    # Filtra as colunas necessárias
    colunas_necessarias = ['nome', 'valor', 'contato', 'descrição']
    df_filtrado = df[colunas_necessarias]

    return df_filtrado

def criar_recibo_pdf(linha, caminho_pdf, caminho_imagem):
    # Configura o tamanho da página como letter em formato paisagem
    altura, largura = letter
    c = canvas.Canvas(caminho_pdf, pagesize=(largura, altura))

    # Adiciona a imagem de fundo
    imagem = Image.open(caminho_imagem)
    c.drawInlineImage(imagem, 0, 0, width=largura, height=altura)

    # Configura a fonte
    c.setFont("Helvetica", 25)

    # Define as coordenadas específicas (2.1 cm, 6.27 cm)
    x_nome = 2.1 * 28.35  # Convertendo cm para pontos (1 cm = 28.35 pontos)
    y_nome = altura - 7.4 * 28.35  # Convertendo cm para pontos e invertendo o eixo y

    # Escreve o "nome" nas coordenadas especificadas
    c.drawString(x_nome, y_nome, f"{linha['nome']}")
    c.drawString(x_nome, y_nome - 35, f"R$ {linha['valor']}")
    c.drawString(x_nome, y_nome - 70, f"{linha['contato']}")
    c.drawString(x_nome, y_nome - 105, f"{linha['descrição']}")
    
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
    
    # Caminho para a imagem de fundo
    caminho_imagem = '/home/pv-lds/Desktop/temp.png'

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

        # Cria o PDF com os dados do recibo, a imagem de fundo, e posiciona o "nome"
        criar_recibo_pdf(linha, caminho_arquivo_pdf, caminho_imagem)
