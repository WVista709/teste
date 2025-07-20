import pandas as pd
import os
from openpyxl import Workbook
from tkinter import Tk

def agrupar_excels_em_um(arquivos_selecionados, diretorio_destino, nome_arquivo_final):
    """
    Agrupa múltiplos arquivos Excel/CSV/XLS em um único arquivo Excel,
    onde cada arquivo original vira uma aba separada.
    """
    Tk().withdraw()
    
    # Cria um novo workbook
    wb = Workbook()
    # Remove a planilha padrão
    wb.remove(wb.active)
    
    for nome_aba, caminho_arquivo in arquivos_selecionados.items():
        extensao = os.path.splitext(caminho_arquivo)[1].lower()
        nome_arquivo = os.path.basename(caminho_arquivo)
        print(f"Processando: {nome_arquivo} -> Aba: {nome_aba}")
        
        try:
            # Lê o arquivo baseado na extensão
            if extensao == '.csv':
                try:
                    df = pd.read_csv(caminho_arquivo, delimiter=';', skiprows=1, encoding='utf-8')
                except UnicodeDecodeError:
                    df = pd.read_csv(caminho_arquivo, delimiter=';', skiprows=1, encoding='latin1')
            elif extensao == '.xls':
                df = pd.read_excel(caminho_arquivo, engine='xlrd')
            elif extensao == '.xlsx':
                df = pd.read_excel(caminho_arquivo, engine='openpyxl')
            else:
                print(f"Formato não suportado: {nome_arquivo}. Pulando.")
                continue
            
            # Cria uma nova aba no workbook
            ws = wb.create_sheet(title=nome_aba)
            
            # Escreve os dados na aba
            # Primeiro, escreve os cabeçalhos
            for col_num, column_title in enumerate(df.columns, 1):
                ws.cell(row=1, column=col_num, value=column_title)
            
            # Depois, escreve os dados
            for row_num, row_data in enumerate(df.values, 2):
                for col_num, cell_value in enumerate(row_data, 1):
                    ws.cell(row=row_num, column=col_num, value=cell_value)
            
            print(f"Aba '{nome_aba}' criada com {len(df)} linhas e {len(df.columns)} colunas")
            
        except Exception as e:
            print(f"Erro ao processar {nome_arquivo}: {str(e)}")
            # Cria uma aba com mensagem de erro
            ws = wb.create_sheet(title=f"{nome_aba}_ERRO")
            ws.cell(row=1, column=1, value=f"Erro ao processar: {str(e)}")
    
    if not wb.sheetnames:
        wb.create_sheet("CHECK")
        print("Nenhuma aba foi criada. Criando aba CHECK padrão.")
        
    # Salva o arquivo final
    if not nome_arquivo_final.endswith('.xlsx'):
        nome_arquivo_final += '.xlsx'
    
    caminho_final = os.path.join(diretorio_destino, nome_arquivo_final)
    wb.save(caminho_final)
    print(f"Arquivo agrupado salvo em: {caminho_final}")
    
    return caminho_final

def detectar_skiprows(caminho_arquivo):
    """Detecta se precisa pular a primeira linha (sep=;)"""
    try:
        with open(caminho_arquivo, encoding='utf-8') as f:
            primeira_linha = f.readline().strip().lower()
            return 1 if primeira_linha.startswith('sep=') else 0
    except:
        return 0

# Mantém a função antiga para compatibilidade
def converter_varios_para_xlsx(caminhos_arquivos, diretorio_destino):
    """Função original para converter arquivos individuais"""
    Tk().withdraw()
    erros = []
    
    for caminho_arquivo in caminhos_arquivos:
        extensao = os.path.splitext(caminho_arquivo)[1].lower()
        nome_arquivo = os.path.basename(caminho_arquivo)
        print(f"Convertendo: {nome_arquivo}")

        try:
            if extensao == '.csv':
                skiprows = detectar_skiprows(caminho_arquivo)
                try:
                    df = pd.read_csv(caminho_arquivo, delimiter=';', skiprows=skiprows, encoding='utf-8-sig')
                except UnicodeDecodeError:
                    df = pd.read_csv(caminho_arquivo, delimiter=';', skiprows=skiprows, encoding='latin1')
            elif extensao == '.xls':
                df = pd.read_excel(caminho_arquivo, engine='xlrd')
            elif extensao == '.xlsx':
                print(f"Arquivo {nome_arquivo} já está em .xlsx. Copiando para o destino.")
                import shutil
                nome_saida = os.path.splitext(nome_arquivo)[0] + '.xlsx'
                caminho_saida = os.path.join(diretorio_destino, nome_saida)
                shutil.copy2(caminho_arquivo, caminho_saida)
                print(f'Arquivo copiado para: {caminho_saida}')
                continue
            else:
                erros.append(f"Formato não suportado: {nome_arquivo}")
                continue

            for col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='ignore')

            nome_saida = os.path.splitext(nome_arquivo)[0] + '.xlsx'
            caminho_saida = os.path.join(diretorio_destino, nome_saida)
            df.to_excel(caminho_saida, index=False, engine='openpyxl')
            print(f'Arquivo salvo em: {caminho_saida}')
            
        except Exception as e:
            erros.append(f"Erro ao converter {nome_arquivo}: {str(e)}")
    
    return erros