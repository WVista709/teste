import openpyxl

def gerar_excel(qtd_linhas, qtd_colunas, nome_arquivo='produtos.xlsx'):
    # Cria uma nova planilha
    wb = openpyxl.Workbook()
    ws = wb.active

    # Preenche as células com texto exemplo
    for linha in range(1, qtd_linhas + 1):
        for coluna in range(1, qtd_colunas + 1):
            ws.cell(row=linha, column=coluna, value=f"Linha{linha}-Coluna{coluna}")

    # Salva o arquivo
    wb.save(nome_arquivo)
    print(f"Arquivo '{nome_arquivo}' criado com {qtd_linhas} linhas e {qtd_colunas} colunas.")

# Exemplo de uso
gerar_excel(1000, 22)

#4 exceis de 1000 linhas com 22 colunas - 14s a 14,92s
#6 exceis de 100 linhas com 22 colunas - 