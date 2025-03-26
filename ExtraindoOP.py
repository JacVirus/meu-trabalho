import fitz  # PyMuPDF

# Caminho do arquivo PDF
pdf_path = r"C:\PYTHON\ORDENS DE PRODUÇÃO\799737.pdf"

# Abrir o PDF
doc = fitz.open(pdf_path)

# Extrair texto de todas as páginas
texto_completo = ""
for pagina in doc:
    texto_completo += pagina.get_text("text") + "\n"
#OP 797724 AMERICA BFS 1.0/1.2 GOLD WHITE - GAZTAMBIDE
# Exibir o texto extraído
# print(texto_completo)

import re
# Função para encontrar informações específicas no texto extraído
def extrair_informacoes(texto):
    dados = {}



    padroes = {
         "Roteiro": r'Roteiro:\s*\d+\s*-\s*\d+\s*(.+)',
         "Fórmula": r'Fórmula:\s*\d+\s*-\s*\d+\s*(.+)',
         "Pedido/Carga": r'Pedido /Carga:\s*(\d+\s*[|]\s*\d+)',
         "Cliente": r'Cliente:\s*(\d+)',
         "Produto": r'Produto:\s+\d+\s*-\s*([^\n]+)',
         "OP": r'OP\s+(\d+)\s+([A-Za-z0-9\s./-]+)\s+-\s+([A-Za-z\s]+)',
         "Partida":r'Partida:\s*(\w+)',
         "Totais":r'Totais:\s*(\d+)',
         "FORNECEDOR":r'FORNECEDOR:\s*([\w*\s./-][^\n]+)',
         "Faixa de Tamanho":r'FAIXA DE TAMANHO:\s*([\d*\s\w*./-][^/\n]+)',
         "Data do Primeiro Rebaixe":r'DATA PRIMEIRO REBAIXE:\s*(\d{2}/\d{2}/\d{4})',
         "Data do Primeiro Rebaixe":r'DATA DO PRIMEIRO REBAIXE:\s*(\d{2}/\d{2}/\d{4})',
         "Data do PRIMEIRO Rebaixe":r'DATA DO PRIMEIRO RBX:\s*(\d{2}/\d{2}/\d{4})', 
         "OP de Recurtimeto":r'OP DE RECURTIMENTO:\s*(\d+)',
         "OP de Semi": r'OP DE SEMI:\s*(\d+)',  
         "Quantidade de Peças":r'Qtde.:\s*(\d+)',      
        }

    # Aplicando os padrões ao texto extraído  5974 - OBS PCP : FORNECEDOR:
    for chave, padrao in padroes.items():
        match = re.search(padrao, texto)
        if match:
            dados[chave] = match.group(1) 
    return dados

   

# Extraindo as informações
info_extraidas = extrair_informacoes(texto_completo)

# Exibir os dados encontrados
for chave, valor in info_extraidas.items():
    print(f"{chave}: {valor}")

    

