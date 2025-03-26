
from pyxlsb import open_workbook
import pandas as pd
import re

# Caminho do arquivo XLSB
caminho_arquivo = r"C:\PYTHON\Diária Nova (1).xlsb"

# Abrir o arquivo e listar as planilhas disponíveis
xlsb = pd.ExcelFile(caminho_arquivo, engine="pyxlsb")
planilhas = xlsb.sheet_names

# Ler a primeira planilha
df = pd.read_excel(caminho_arquivo, sheet_name=planilhas[0], engine="pyxlsb")

# Verificando se existe uma coluna com "Partida" no nome
colunas = df.columns
coluna_partida = [col for col in colunas if "Partida" in col]

# Se encontrarmos uma coluna correspondente, filtramos os dados
if coluna_partida:
    coluna_partida = coluna_partida[0]  # Pegamos a primeira ocorrência
    partida_info = df[df[coluna_partida] == 64041]

    # Se encontrar a partida, extrai os dados necessários
    if not partida_info.empty:
        artigo = partida_info["Artigo de Semi"].values[0] if "Artigo de Semi" in partida_info else "Não encontrado"
        artigo_extraido = re.sub(r'^KL\.\d+\s+', '', artigo)
        cor = partida_info["Cor de Semi.1"].values[0] if "Cor de Semi.1" in partida_info else "Não encontrado"
        cliente = partida_info["Cliente.1"].values[0] if "Cliente.1" in partida_info else "Não encontrado"
        pedido = partida_info["Pedido"].values[0] if "Pedido" in partida_info else "Não encontrado"
        Classe_de_Semi = partida_info["Classe Semi"].values[0] if "Classe Semi" in partida_info else "Não encontrado"
        espessura = partida_info["Espessura Final"].values[0] if "Espessura Final" in partida_info else "Não encontrado"
        QTDE_em_Peças = partida_info["QTD Pedido (Peças)"].values[0] if "QTD Pedido (Peças)" in partida_info else "Campo sem informação provavelmente programado em Metros" 
        QTDE_em_Metros = partida_info["QTD Pedido (M²)"].values[0] if "QTD Pedido (M²)" in partida_info else "Campo sem informação provavelmente programado em Peças"
        Peças_da_Partida: int = partida_info["Peças"].values[0] if "Peças" in partida_info else 0
        Metros_da_Partida: int = partida_info["Metros"].values[0] if "Metros" in partida_info else 0
        Tipo_de_Produto = partida_info["Tipo Divisão"].values[0] if "Tipo Divisão" in partida_info else "Não encontrado"

    else:
        artigo, cor, cliente, pedido = "Partida não encontrada", "-", "-"
else:
    artigo, cor, cliente, pedido = "Coluna 'Partida' não encontrada", "-", "-"

# Exibir os resultados
print(f"Artigo Semi: {artigo_extraido}")
print(f"Cor de Semi: {cor}")
print(f"Cliente: {cliente}")
print(f"Pedido: {pedido}")
print(f"Classe de Semi: {Classe_de_Semi}")
print(f"Espessura: {espessura}")
print(f"QTDE em Peças total do pedido: {QTDE_em_Peças} PÇs" if not pd.isna(QTDE_em_Peças) else " Campo sem informação, provavelmente programado em Metros")
print(f"QTDE em Metros total do pedido: {QTDE_em_Metros} M²" if not pd.isna(QTDE_em_Metros) else " Campo sem informação, provavelmente programado em Peças")
print(f"Peças da Partida: {Peças_da_Partida} PÇs" if Peças_da_Partida else "Partida não encontrada")
print(f"Metros da Partida: {Metros_da_Partida} M²" if Metros_da_Partida else "Partida não encontrada")
print(f"Tipo de Produto: {Tipo_de_Produto}")



