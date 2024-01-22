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
    colunas_necessarias = ['nome', 'valor', 'contato', 'descrição', 'pagante','cnpj','contato2']
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
    for campo, (x, y, fonte_campo, tamanho_campo, prefixo) in coordenadas.items():
        # Define a fonte e tamanho específicos para cada campo
        c.setFont(fonte_campo, tamanho_campo)
        # Adiciona o prefixo e formata a data, se existir
        if campo == 'data':
            c.drawString(x, y, f"{prefixo}{datetime.now().strftime('%Y-%m-%d')}")
        else:
            c.drawString(x, y, f"{prefixo}{linha[campo]}")

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
    caminho_excel = '/home/pv-lds/Desktop/22.01/base.xlsx'
    
    # Caminho para a imagem de fundo
    caminho_imagem = '/home/pv-lds/Desktop/22.01/temp2.png'

    # Extrai dados do Excel
    dados_recibo = extrair_dados_excel(caminho_excel)

    # Define as coordenadas personalizadas para cada campo com fontes e tamanhos específicos
    coordenadas_personalizadas = {
        #esquerda
        'nome': (75, 400, "Times-Roman", 28, ""),
        'valor': (75, 355, "Times-Roman", 25, "R$ "),
        'contato': (75, 320, "Times-Roman", 20, ""),
        
        #centro
        'descrição': (70, 190, "Times-Roman", 28, "• "),
        'data': (340, 470, "Times-Roman", 22, ""),

        #direita
        'pagante': (430, 400, "Times-Roman", 24, ""),
        'cnpj': (430, 355, "Times-Roman", 25, "CNPJ: "),
        'contato2': (430, 320, "Times-Roman", 20, ""),
    }

    # Cria uma pasta com a data atual no Desktop
    pasta_recibos = criar_pasta_data_atual()

    # Itera sobre os dados e gera um PDF por recibo na pasta criada
    for indice, linha in dados_recibo.iterrows():
        # Adota o nome da pessoa como nome do arquivo PDF
        nome_arquivo_pdf = f"{linha['nome']}_recibo.pdf"
        
        # Caminho completo para o arquivo PDF na pasta
        caminho_arquivo_pdf = os.path.join(pasta_recibos, nome_arquivo_pdf)

        # Cria o PDF com os dados do recibo, a imagem de fundo, as coordenadas e os prefixos
        criar_recibo_pdf(linha, caminho_arquivo_pdf, caminho_imagem, coordenadas_personalizadas)

