import fitz  # PyMuPDF

# Caminho do PDF original e do novo PDF com destaques
pdf_original = r"C:\PYTHON\ORDENS DE PRODUÇÃO\799737.pdf"
pdf_destacado = "799737_corrigido.pdf"

# Lista das palavras extraídas do Excel para destacar no PDF
palavras_para_destacar = [
    "KL.2 Senna FF",
    "Flint Grey 41159 (21)",
    "Kamali Group",
    "930430",
    "Silver",
    "1.1/1.3",
    "3000.0 PÇs",
    "Campo sem informação, provavelmente programado em Peças",
    "419 PÇs",
    "1573.509 M²"
]

# Abrir o PDF
doc = fitz.open(pdf_original)

# Percorrer todas as páginas do PDF
for pagina in doc:
    for palavra in palavras_para_destacar:
        # Procurar todas as ocorrências da palavra
        textos = pagina.search_for(palavra)

        # Grifar cada ocorrência encontrada em verde
        for posicao in textos:
            anotacao = pagina.add_highlight_annot(posicao)
            anotacao.set_colors(stroke=(0, 1, 0))  # Cor verde (RGB: 0,1,0)
            anotacao.update()

# Salvar o PDF com as palavras destacadas
doc.save(pdf_destacado)
doc.close()

print(f"PDF salvo como: {pdf_destacado}")




