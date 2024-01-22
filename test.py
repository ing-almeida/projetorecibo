import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A5

def extrair_dados_excel(caminho_excel):
    # Carrega o arquivo Excel
    df = pd.read_excel(caminho_excel)

    # Filtra as colunas necessárias
    colunas_necessarias = ['nome', 'valor', 'contato', 'descrição']
    df_filtrado = df[colunas_necessarias]

    return df_filtrado

def criar_recibo_pdf(dados_recibo, caminho_pdf):
    # Configura o tamanho da página como A5
    largura, altura = A5
    c = canvas.Canvas(caminho_pdf, pagesize=A5)

    # Configura a fonte
    c.setFont("Helvetica", 12)

    # Calcula a posição central na página
    x_centro = largura / 2
    y_centro = altura / 2

    # Escreve os dados do recibo no PDF
    for indice, linha in dados_recibo.iterrows():
        c.drawString(x_centro, y_centro, f"Nome: {linha['nome']}")
        c.drawString(x_centro, y_centro - 15, f"Valor: {linha['valor']}")
        c.drawString(x_centro, y_centro - 30, f"Contato: {linha['contato']}")
        c.drawString(x_centro, y_centro - 45, f"Descrição: {linha['descrição']}")
        c.showPage()  # Avança para a próxima página se houver mais de um recibo

    # Salva o arquivo PDF
    c.save()

if __name__ == "__main__":
    # Substitua 'caminho/do/seu/arquivo.xlsx' pelo caminho real do seu arquivo Excel
    caminho_excel = '/home/pv-lds/Desktop/base.xlsx'
    
    # Substitua 'recibo.pdf' pelo nome desejado para o arquivo PDF
    caminho_pdf = 'recibo.pdf'

    dados_recibo = extrair_dados_excel(caminho_excel)
    criar_recibo_pdf(dados_recibo, caminho_pdf)

