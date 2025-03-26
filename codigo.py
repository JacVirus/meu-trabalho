import fitz  # PyMuPDF
from pyxlsb import open_workbook
import pandas as pd
import re

def extrair_dados_excel(caminho_arquivo, partida_numero=64041):
    """Extrai dados de um arquivo XLSB conforme ExtraindoExcel.py"""
    try:
        xlsb = pd.ExcelFile(caminho_arquivo, engine="pyxlsb")
        planilhas = xlsb.sheet_names
        df = pd.read_excel(caminho_arquivo, sheet_name=planilhas[0], engine="pyxlsb")
        
        colunas = df.columns
        coluna_partida = [col for col in colunas if "Partida" in col]
        
        if coluna_partida:
            coluna_partida = coluna_partida[0]
            partida_info = df[df[coluna_partida] == partida_numero]
            
            if not partida_info.empty:
                dados = {
                    "Artigo Semi": partida_info["Artigo de Semi"].values[0] if "Artigo de Semi" in partida_info else "Não encontrado",
                    "Artigo Semi Limpo": re.sub(r'^KL\.\d+\s+', '', 
                        partida_info["Artigo de Semi"].values[0] if "Artigo de Semi" in partida_info else "Não encontrado"),
                    "Cor de Semi": partida_info["Cor de Semi.1"].values[0] if "Cor de Semi.1" in partida_info else "Não encontrado",
                    "Cliente": partida_info["Cliente.1"].values[0] if "Cliente.1" in partida_info else "Não encontrado",
                    "Pedido": partida_info["Pedido"].values[0] if "Pedido" in partida_info else "Não encontrado",
                    "Classe de Semi": partida_info["Classe Semi"].values[0] if "Classe Semi" in partida_info else "Não encontrado",
                    "Espessura": partida_info["Espessura Final"].values[0] if "Espessura Final" in partida_info else "Não encontrado",
                    "QTDE Peças": partida_info["QTD Pedido (Peças)"].values[0] if "QTD Pedido (Peças)" in partida_info else None,
                    "QTDE Metros": partida_info["QTD Pedido (M²)"].values[0] if "QTD Pedido (M²)" in partida_info else None,
                    "Peças Partida": partida_info["Peças"].values[0] if "Peças" in partida_info else 0,
                    "Metros Partida": partida_info["Metros"].values[0] if "Metros" in partida_info else 0,
                    "Tipo Produto": partida_info["Tipo Divisão"].values[0] if "Tipo Divisão" in partida_info else "Não encontrado",
                    "N° Partida": partida_info["Partida"].values[0] if "Partida" in partida_info else "Não encontrado"
                }
                
                # Preparar lista de palavras para destacar no PDF
                palavras_destaque = [
                    dados["Artigo Semi Limpo"],
                    dados["Cor de Semi"],
                    dados["Cliente"],
                    str(dados["Pedido"]),
                    dados["Espessura"],
                    f"{dados['QTDE Peças']}" if dados['QTDE Peças'] else "Campo sem informação, provavelmente programado em Peças",
                    f"{dados['Peças Partida']}" if dados['Peças Partida'] else "0 PÇs",
                    f"{dados['Metros Partida']}" if dados['Metros Partida'] else "0 M²",
                    f"{dados['Classe de Semi']}" if dados['Classe de Semi'] else " ",
                    f"{dados['Tipo Produto']}" if dados['Tipo Produto'] else " ",
                    f"{dados['N° Partida']}" if dados['N° Partida'] else " "
                ]
                dados["palavras_para_destacar"] = [p for p in palavras_destaque if p and str(p) != "nan"]
                
                return dados
            else:
                return {"Erro": f"Partida {partida_numero} não encontrada no arquivo Excel"}
        else:
            return {"Erro": "Coluna 'Partida' não encontrada no arquivo Excel"}
    except Exception as e:
        return {"Erro": f"Falha ao processar arquivo Excel: {str(e)}"}

def extrair_dados_pdf(caminho_arquivo):
    """Extrai dados de um arquivo PDF conforme ExtraindoOP.py"""
    try:
        doc = fitz.open(caminho_arquivo)
        texto_completo = ""
        for pagina in doc:
            texto_completo += pagina.get_text("text") + "\n"
        
        padroes = {
            "Roteiro": r'Roteiro:\s*\d+\s*-\s*\d+\s*(.+)',
            "Fórmula": r'Fórmula:\s*\d+\s*-\s*\d+\s*(.+)',
            "Pedido/Carga": r'Pedido /Carga:\s*(\d+\s*[|]\s*\d+)',
            "Cliente": r'Cliente:\s*(\d+)',
            "Produto": r'Produto:\s+\d+\s*-\s*([^\n]+)',
            "OP": r'OP\s+(\d+)\s+([A-Za-z0-9\s./-]+)\s+-\s+([A-Za-z\s]+)',
            "Partida": r'Partida:\s*(\w+)',
            "Totais": r'Totais:\s*(\d+)',
            "FORNECEDOR": r'FORNECEDOR:\s*([\w*\s./-][^\n]+)',
            "Faixa de Tamanho": r'FAIXA DE TAMANHO:\s*([\d*\s\w*./-][^/\n]+)',
            "Data do Primeiro Rebaixe": r'(DATA PRIMEIRO REBAIXE|DATA DO PRIMEIRO REBAIXE|DATA DO PRIMEIRO RBX):\s*(\d{2}/\d{2}/\d{4})',
            "OP de Recurtimeto": r'OP DE RECURTIMENTO:\s*(\d+)',
            "OP de Semi": r'OP DE SEMI:\s*(\d+)',
            "Quantidade de Peças": r'Qtde\.:\s*(\d+)',
        }
        
        dados = {}
        for chave, padrao in padroes.items():
            match = re.search(padrao, texto_completo)
            if match:
                dados[chave] = match.group(1) if len(match.groups()) == 1 else match.group(2)
        
        return dados if dados else {"Erro": "Nenhum dado encontrado no PDF"}
    except Exception as e:
        return {"Erro": f"Falha ao processar arquivo PDF: {str(e)}"}

def destacar_pdf(pdf_original, pdf_destacado, palavras_para_destacar):
    """Destaca palavras no PDF conforme destacando.PY"""
    try:
        doc = fitz.open(pdf_original)
        
        for pagina in doc:
            for palavra in palavras_para_destacar:
                textos = pagina.search_for(str(palavra))
                for posicao in textos:
                    anotacao = pagina.add_highlight_annot(posicao)
                    anotacao.set_colors(stroke=(0, 1, 0))  # Cor verde (RGB: 0,1,0)
                    anotacao.update()
        
        doc.save(pdf_destacado)
        doc.close()
        return True
    except Exception as e:
        print(f"Erro ao destacar PDF: {str(e)}")
        return False

def main():
    # Configurações
    caminho_excel = r"C:\PYTHON\Diária Nova (1).xlsb"
    pdf_original = r"C:\PYTHON\ORDENS DE PRODUÇÃO\799752.pdf"
    pdf_destacado = "799752_corrigido.pdf"
    partida_numero = 64041  # Pode ser ajustado conforme necessário
    
    # 1. Extrair dados do Excel
    print("\nExtraindo dados do Excel...")
    dados_excel = extrair_dados_excel(caminho_excel, partida_numero)
    
    if "Erro" in dados_excel:
        print(f"Erro: {dados_excel['Erro']}")
    else:
        for chave, valor in dados_excel.items():
            if chave != "palavras_para_destacar":
                print(f"{chave}: {valor}")
    
    # 2. Extrair dados do PDF
    print("\nExtraindo dados do PDF...")
    dados_pdf = extrair_dados_pdf(pdf_original)
    
    if "Erro" in dados_pdf:
        print(f"Erro: {dados_pdf['Erro']}")
    else:
        for chave, valor in dados_pdf.items():
            print(f"{chave}: {valor}")
    
    # 3. Destacar informações no PDF
    if "palavras_para_destacar" in dados_excel:
        print("\nDestacando informações no PDF...")
        palavras_destaque = dados_excel["palavras_para_destacar"]
        print("Palavras a destacar:", palavras_destaque)
        
        if destacar_pdf(pdf_original, pdf_destacado, palavras_destaque):
            print(f"PDF com destaques salvo como: {pdf_destacado}")
        else:
            print("Falha ao destacar PDF")
    
    # # 4. Combinar dados (opcional)
    # dados_combinados = {**dados_excel, **dados_pdf}
    # if "palavras_para_destacar" in dados_combinados:
    #     del dados_combinados["palavras_para_destacar"]
    
    # print("\nDados combinados:")
    # for chave, valor in dados_combinados.items():
    #     print(f"{chave}: {valor}")

if __name__ == "__main__":
    main()


