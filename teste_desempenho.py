import openpyxl
import os
import time
import csv
import conversor
import check_bruto

def criar_pasta_unica(pasta_base):
    pasta = pasta_base
    contador = 1
    while os.path.exists(pasta):
        pasta = f"{pasta_base}_{contador}"
        contador += 1
    os.makedirs(pasta)
    return pasta

def gerar_excel(qtd_linhas, qtd_colunas, nome_arquivo):
    wb = openpyxl.Workbook()
    ws = wb.active

    for linha in range(1, qtd_linhas + 1):
        for coluna in range(1, qtd_colunas + 1):
            ws.cell(row=linha, column=coluna, value=f"L{linha}-C{coluna}")

    wb.save(nome_arquivo)
    print(f"Arquivo '{nome_arquivo}' criado com {qtd_linhas} linhas e {qtd_colunas} colunas.")

def gerar_varios_arquivos(qtd_arquivos, qtd_linhas, qtd_colunas, pasta_destino, prefixo_nome="arquivo"):
    if not os.path.exists(pasta_destino):
        os.makedirs(pasta_destino)

    arquivos_gerados = []
    for i in range(1, qtd_arquivos + 1):
        nome_arquivo = os.path.join(pasta_destino, f"{prefixo_nome}_{i}.xlsx")
        gerar_excel(qtd_linhas, qtd_colunas, nome_arquivo)
        arquivos_gerados.append(nome_arquivo)

    return arquivos_gerados

def teste_desempenho_csv(arq_selecionados, caminho_final, modo, caminho_csv, execucao_num=None, linhas=None, colunas=None):
    etapas = [
        ("Agrupando arquivos", lambda: conversor.agrupar_excels_em_um(arq_selecionados, caminho_final)),
        ("Processando SEFAZ", lambda: check_bruto.sefaz(caminho_final, modo)),
        ("Processando ALTERDATA", lambda: check_bruto.alterdata(caminho_final, modo)),
        ("Processando PRODUTO", lambda: check_bruto.produto(caminho_final)),
        ("Verificando COMPRAS", lambda: check_bruto.Check.check_compras(caminho_final, modo)),
        ("Verificando VENDAS", lambda: check_bruto.Check.check_vendas(caminho_final, modo)),
    ]

    with open(caminho_csv, mode='a', newline='', encoding='utf-8') as csvfile:
        writer = csv.writer(csvfile)
        if csvfile.tell() == 0:
            writer.writerow(["Execucao", "Modo", "Linhas", "Colunas", "Processo", "Tempo_segundos"])

        for descricao, func in etapas:
            inicio = time.time()
            func()
            fim = time.time()
            duracao = fim - inicio
            print(f"{descricao}: {duracao:.2f} segundos")
            writer.writerow([execucao_num, modo, linhas, colunas, descricao, f"{duracao:.2f}"])

if __name__ == "__main__":
    qtd_arquivos = 6
    colunas = 22
    prefixo = "produtos"
    pasta_teste = "teste"

    if not os.path.exists(pasta_teste):
        os.makedirs(pasta_teste)

    # Lista com os diferentes tamanhos de linhas para testar
    tamanhos_linhas = [1000, 10000, 100000]

    for linhas in tamanhos_linhas:
        diretorio_base = f"resultado_teste de {linhas} linhas e {colunas} colunas"
        diretorio_destino = criar_pasta_unica(os.path.join(pasta_teste, diretorio_base))

        caminho_csv = os.path.join(diretorio_destino, "resultados_desempenho.csv")

        for i in range(1, 51): 
            print(f"\n=== Execução {i} para {linhas} linhas ===\n")

            # Gera os arquivos de teste dentro da pasta da execução atual
            arquivos_gerados = gerar_varios_arquivos(qtd_arquivos, linhas, colunas, diretorio_destino, prefixo)

            # Teste para modo "Conferência de Nota"
            arquivos_conferencia = {
                "COMPRAS SEFAZ": arquivos_gerados[0],
                "COMPRAS ALTERDATA": arquivos_gerados[1],
                "VENDAS SEFAZ": arquivos_gerados[2],
                "VENDAS ALTERDATA": arquivos_gerados[3],
            }
            caminho_final_conf = os.path.join(diretorio_destino, f"resultado_conferencia_exec_{i}.xlsx")

            print("Iniciando teste para modo 'Conferência de Nota'...\n")
            teste_desempenho_csv(arquivos_conferencia, caminho_final_conf, "Conferência de Nota", caminho_csv, execucao_num=i, linhas=linhas, colunas=colunas)

            # Teste para modo "Check"
            arquivos_check = {
                "COMPRAS SEFAZ": arquivos_gerados[0],
                "COMPRAS ALTERDATA": arquivos_gerados[1],
                "COMPRAS PRODUTOS": arquivos_gerados[4],
                "VENDAS SEFAZ": arquivos_gerados[2],
                "VENDAS ALTERDATA": arquivos_gerados[3],
                "VENDAS PRODUTOS": arquivos_gerados[5],
            }
            caminho_final_check = os.path.join(diretorio_destino, f"resultado_check_exec_{i}.xlsx")

            print("Iniciando teste para modo 'Check'...\n")
            teste_desempenho_csv(arquivos_check, caminho_final_check, "Check", caminho_csv, execucao_num=i, linhas=linhas, colunas=colunas)