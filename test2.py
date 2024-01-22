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

def criar_recibo_pdf(linha, caminho_pdf, caminho_imagem, coordenadas, fonte="Helvetica", tamanho=25):
    # Configura o tamanho da página como letter em formato paisagem
    altura, largura = letter
    c = canvas.Canvas(caminho_pdf, pagesize=(largura, altura))

    # Adiciona a imagem de fundo
    imagem = Image.open(caminho_imagem)
    c.drawInlineImage(imagem, 0, 0, width=largura, height=altura)

    # Especifica a fonte padrão e tamanho
    c.setFont(fonte, tamanho)

    # Escreve os campos nas coordenadas personalizadas
    for campo, (x, y, fonte_campo, tamanho_campo) in coordenadas.items():
        # Define a fonte e tamanho específicos para cada campo
        c.setFont(fonte_campo, tamanho_campo)
        c.drawString(x, y, f"{linha[campo]}")
    
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

    # Define as coordenadas personalizadas para cada campo com fontes e tamanhos específicos
    coordenadas_personalizadas = {
        'nome': (70, 410, "Times-Bold", 30),
        'valor': (70, 375, "Helvetica-Oblique", 25),
        'contato': (70, 340, "Courier", 20),
        'descrição': (300, 200, "Helvetica", 28),
    }

    # Cria uma pasta com a data atual no Desktop
    pasta_recibos = criar_pasta_data_atual()

    # Itera sobre os dados e gera um PDF por recibo na pasta criada
    for indice, linha in dados_recibo.iterrows():
        # Adota o nome da pessoa como nome do arquivo PDF
        nome_arquivo_pdf = f"{linha['nome']}_recibo.pdf"
        
        # Caminho completo para o arquivo PDF na pasta
        caminho_arquivo_pdf = os.path.join(pasta_recibos, nome_arquivo_pdf)

        # Cria o PDF com os dados do recibo, a imagem de fundo, e posiciona os campos
        criar_recibo_pdf(linha, caminho_arquivo_pdf, caminho_imagem, coordenadas_personalizadas)
