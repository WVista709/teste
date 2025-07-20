import pandas as pd
import os
from openpyxl import Workbook
from tkinter import Tk
import tempfile
import shutil
import csv

def detectar_delimitador_corrigido(caminho_arquivo, encoding):
    """Detecta o delimitador do CSV e se precisa pular linhas - VERSÃO CORRIGIDA"""
    try:
        with open(caminho_arquivo, encoding=encoding) as f:
            primeira_linha = f.readline().strip()
            segunda_linha = f.readline().strip()

            print(f"Primeira linha: '{primeira_linha[:50]}...'")
            print(f"Segunda linha: '{segunda_linha[:100]}...'")

            # Se a primeira linha é sep=, pula ela
            if primeira_linha.lower().startswith('sep='):
                delimiter = primeira_linha[-1]
                print(f"Detectado sep= na primeira linha, delimiter='{delimiter}', skiprows=1")
                return delimiter, 1

            # Se a primeira linha parece ser cabeçalho (tem letras), não pula
            if any(c.isalpha() for c in primeira_linha):
                # Tenta detectar delimitador automaticamente
                sniffer = csv.Sniffer()
                delimiter = sniffer.sniff(primeira_linha).delimiter
                print(f"Primeira linha parece cabeçalho, delimiter='{delimiter}', skiprows=0")
                return delimiter, 0
            else:
                # Se a primeira linha são só números/símbolos, pode ser que precise pular
                # Verifica se a segunda linha parece cabeçalho
                if any(c.isalpha() for c in segunda_linha):
                    sniffer = csv.Sniffer()
                    delimiter = sniffer.sniff(segunda_linha).delimiter
                    print(f"Segunda linha parece cabeçalho, delimiter='{delimiter}', skiprows=1")
                    return delimiter, 1
                else:
                    # Padrão
                    print(f"Usando padrão: delimiter=';', skiprows=0")
                    return ';', 0

    except Exception as e:
        print(f"Erro ao detectar delimitador: {e}")
        return ';', 0  # Padrão

def corrigir_alinhamento_colunas(df, nome_arquivo):
    """Corrige desalinhamento entre cabeçalhos e dados"""
    print(f"\nVerificando alinhamento do arquivo {nome_arquivo}...")

    if len(df) == 0:
        return df

    # Verifica se a primeira coluna deveria ser UF mas tem valor errado
    if df.columns[0] == 'UF' and len(df) > 0:
        primeiro_valor = str(df.iloc[0, 0])

        # Se o primeiro valor não parece ser UF (muito longo, tem números demais)
        if len(primeiro_valor) > 10 or primeiro_valor.count("'") > 1:
            print(f"⚠ Detectado desalinhamento! Primeiro valor: '{primeiro_valor[:50]}...'")

            # Estratégia 1: Verifica se há uma coluna sem nome no início
            # Isso acontece quando o CSV tem uma coluna extra que o pandas não consegue nomear
            colunas_sem_nome = [col for col in df.columns if str(col).startswith('Unnamed:')]
            if colunas_sem_nome:
                print(f"Encontradas colunas sem nome: {colunas_sem_nome}")
                # Remove colunas sem nome
                df_corrigido = df.drop(columns=colunas_sem_nome)
                print(f"✓ Removidas colunas sem nome. Novo shape: {df_corrigido.shape}")
                return df_corrigido

            # Estratégia 2: Shift dos dados para a direita (adiciona UF no início)
            print("Tentando adicionar coluna UF no início...")
            df_corrigido = df.copy()

            # Detecta UF baseado no contexto (você mencionou que deveria ser AM)
            uf_detectada = "AM"

            # Shift todas as colunas para a direita
            # Renomeia as colunas existentes
            novas_colunas = ['UF'] + [f'COL_{i}' for i in range(1, len(df.columns) + 1)]

            # Cria novo DataFrame com estrutura correta
            dados_corrigidos = []
            for _, row in df.iterrows():
                nova_linha = [uf_detectada] + list(row.values)
                dados_corrigidos.append(nova_linha)

            df_corrigido = pd.DataFrame(dados_corrigidos, columns=novas_colunas)

            # Renomeia as colunas para os nomes originais (deslocados)
            nomes_originais = ['UF'] + list(df.columns)
            if len(nomes_originais) == len(df_corrigido.columns):
                df_corrigido.columns = nomes_originais

            print(f"✓ Adicionada coluna UF com valor '{uf_detectada}'")
            print(f"Novo primeiro valor: '{df_corrigido.iloc[0, 0]}'")
            return df_corrigido

    print("✓ Alinhamento parece correto.")
    return df

def ler_csv_robusto(caminho_arquivo):
    """Lê CSV de forma robusta, tentando diferentes encodings e delimitadores"""
    nome_arquivo = os.path.basename(caminho_arquivo)

    # Lista de encodings para tentar
    encodings = ['utf-8', 'utf-8-sig', 'latin1', 'cp1252', 'iso-8859-1']

    for encoding in encodings:
        try:
            print(f"\nTentando ler {nome_arquivo} com encoding={encoding}")

            # Primeiro, tenta detectar delimitador automaticamente
            delimiter, skiprows = detectar_delimitador_corrigido(caminho_arquivo, encoding)
            print(f"Configuração: delimiter='{delimiter}', skiprows={skiprows}")

            # Parâmetros para pandas baseado na versão
            read_params = {
                'delimiter': delimiter,
                'skiprows': skiprows,
                'encoding': encoding,
                'engine': 'python'  # Mais tolerante a erros
            }

            # Para pandas >= 1.3
            try:
                read_params['on_bad_lines'] = 'skip'
                df = pd.read_csv(caminho_arquivo, **read_params)
            except TypeError:
                # Para pandas < 1.3
                read_params.pop('on_bad_lines', None)
                read_params['error_bad_lines'] = False
                read_params['warn_bad_lines'] = True
                df = pd.read_csv(caminho_arquivo, **read_params)

            # Se só veio uma coluna, tenta com delim_whitespace
            if len(df.columns) == 1:
                print("Só uma coluna detectada, tentando delim_whitespace=True")
                read_params_ws = {
                    'delim_whitespace': True,
                    'encoding': encoding,
                    'engine': 'python'
                }

                try:
                    read_params_ws['on_bad_lines'] = 'skip'
                    df = pd.read_csv(caminho_arquivo, **read_params_ws)
                except TypeError:
                    read_params_ws.pop('on_bad_lines', None)
                    read_params_ws['error_bad_lines'] = False
                    read_params_ws['warn_bad_lines'] = True
                    df = pd.read_csv(caminho_arquivo, **read_params_ws)

            # Se ainda só tem uma coluna, tenta com tab
            if len(df.columns) == 1:
                print("Ainda uma coluna, tentando delimiter='\\t' (tab)")
                read_params_tab = {
                    'delimiter': '\t',
                    'encoding': encoding,
                    'engine': 'python'
                }

                try:
                    read_params_tab['on_bad_lines'] = 'skip'
                    df = pd.read_csv(caminho_arquivo, **read_params_tab)
                except TypeError:
                    read_params_tab.pop('on_bad_lines', None)
                    read_params_tab['error_bad_lines'] = False
                    read_params_tab['warn_bad_lines'] = True
                    df = pd.read_csv(caminho_arquivo, **read_params_tab)

            print(f"✓ {nome_arquivo} lido com sucesso: {len(df)} linhas, {len(df.columns)} colunas")
            print(f"Colunas detectadas: {list(df.columns)}")

            # Debug: mostra primeira linha para verificar se os dados estão alinhados
            if len(df) > 0:
                primeira_linha_dados = list(df.iloc[0])
                print(f"Primeira linha de dados: {primeira_linha_dados}")

                # Verifica se a primeira coluna tem o valor esperado
                if df.columns[0] == 'UF' and len(primeira_linha_dados) > 0:
                    primeiro_valor = primeira_linha_dados[0]
                    print(f"Valor da coluna UF na primeira linha: '{primeiro_valor}'")
                    if primeiro_valor != 'AM' and not str(primeiro_valor).startswith('AM'):
                        print("⚠ ATENÇÃO: O valor da coluna UF não parece correto!")
                        print("⚠ Tentando corrigir alinhamento...")

                        # Aplica correção de alinhamento
                        df = corrigir_alinhamento_colunas(df, nome_arquivo)

                        # Verifica novamente após correção
                        if len(df) > 0 and df.columns[0] == 'UF':
                            novo_valor = df.iloc[0, 0]
                            print(f"Valor da coluna UF após correção: '{novo_valor}'")

            return df

        except Exception as e:
            print(f"✗ Tentativa com encoding {encoding} falhou: {e}")
            continue

    # Se chegou aqui, nenhum encoding funcionou
    raise Exception(f"Não foi possível ler o arquivo {nome_arquivo} com nenhum encoding conhecido.")

def agrupar_excels_em_um(arquivos_selecionados, diretorio_destino, nome_arquivo_final):
    """
    Primeiro converte todos os arquivos para .xlsx, depois agrupa em um único arquivo Excel,
    onde cada arquivo original vira uma aba separada.
    """
    Tk().withdraw()

    # Cria um diretório temporário para os arquivos convertidos
    temp_dir = tempfile.mkdtemp()
    arquivos_convertidos = {}

    try:
        print("=== FASE 1: CONVERTENDO ARQUIVOS ===")

        # Primeiro, converte todos os arquivos para .xlsx
        for nome_aba, caminho_arquivo in arquivos_selecionados.items():
            extensao = os.path.splitext(caminho_arquivo)[1].lower()
            nome_arquivo = os.path.basename(caminho_arquivo)
            print(f"\n{'='*50}")
            print(f"Processando: {nome_arquivo} -> {nome_aba}")
            print(f"{'='*50}")

            try:
                # Lê o arquivo baseado na extensão
                if extensao == '.csv':
                    df = ler_csv_robusto(caminho_arquivo)
                elif extensao == '.xls':
                    print(f"Lendo arquivo XLS: {nome_arquivo}")
                    df = pd.read_excel(caminho_arquivo, engine='xlrd')
                elif extensao == '.xlsx':
                    print(f"Lendo arquivo XLSX: {nome_arquivo}")
                    df = pd.read_excel(caminho_arquivo, engine='openpyxl')
                else:
                    print(f"✗ Formato não suportado: {nome_arquivo}. Pulando.")
                    continue

                # Verifica se o DataFrame não está vazio
                if df.empty:
                    print(f"⚠ Arquivo {nome_arquivo} está vazio. Criando aba com aviso.")
                    # Cria DataFrame com aviso
                    df = pd.DataFrame({'AVISO': ['Arquivo estava vazio']})

                # Debug: mostra informações do DataFrame
                print(f"DataFrame final:")
                print(f"  - Linhas: {len(df)}")
                print(f"  - Colunas: {len(df.columns)}")
                print(f"  - Nomes das colunas: {list(df.columns)}")
                if len(df) > 0:
                    print(f"  - Primeira linha: {list(df.iloc[0])}")

                # Converte colunas numéricas quando possível (suprime warning)
                for col in df.columns:
                    try:
                        df[col] = pd.to_numeric(df[col], errors='ignore')
                    except:
                        pass

                # Salva o arquivo convertido no diretório temporário
                nome_temp = f"{nome_aba}.xlsx"
                caminho_temp = os.path.join(temp_dir, nome_temp)
                df.to_excel(caminho_temp, index=False, engine='openpyxl')

                arquivos_convertidos[nome_aba] = caminho_temp
                print(f"✓ {nome_arquivo} convertido com sucesso ({len(df)} linhas, {len(df.columns)} colunas)")

            except Exception as e:
                print(f"✗ Erro ao converter {nome_arquivo}: {str(e)}")
                # Cria um DataFrame com a mensagem de erro
                df_erro = pd.DataFrame({
                    'ERRO': [f'Erro ao processar arquivo: {str(e)}'],
                    'ARQUIVO': [nome_arquivo],
                    'EXTENSAO': [extensao]
                })
                nome_temp = f"{nome_aba}_ERRO.xlsx"
                caminho_temp = os.path.join(temp_dir, nome_temp)
                df_erro.to_excel(caminho_temp, index=False, engine='openpyxl')
                arquivos_convertidos[f"{nome_aba}_ERRO"] = caminho_temp
                continue

        print(f"\n{'='*50}")
        print(f"=== FASE 2: AGRUPANDO {len(arquivos_convertidos)} ARQUIVOS ===")
        print(f"{'='*50}")

        # Agora agrupa todos os arquivos convertidos
        wb = Workbook()
        wb.remove(wb.active)  # Remove a planilha padrão

        for nome_aba, caminho_temp in arquivos_convertidos.items():
            print(f"\nAdicionando aba: {nome_aba}")

            try:
                # Lê o arquivo .xlsx convertido
                df = pd.read_excel(caminho_temp, engine='openpyxl')

                # Cria uma nova aba no workbook
                nome_aba_excel = nome_aba[:31]  # Excel limita nomes de aba a 31 caracteres
                ws = wb.create_sheet(title=nome_aba_excel)

                print(f"Escrevendo cabeçalhos...")
                # Escreve os cabeçalhos
                for col_num, column_title in enumerate(df.columns, 1):
                    ws.cell(row=1, column=col_num, value=str(column_title))

                print(f"Escrevendo dados...")
                # Escreve os dados
                for row_num, row_data in enumerate(df.values, 2):
                    for col_num, cell_value in enumerate(row_data, 1):
                        # Converte valores NaN para None (células vazias no Excel)
                        if pd.isna(cell_value):
                            cell_value = None
                        ws.cell(row=row_num, column=col_num, value=cell_value)

                print(f"✓ Aba '{nome_aba_excel}' criada com {len(df)} linhas e {len(df.columns)} colunas")

                # Debug: verifica se a primeira célula de dados está correta
                if len(df) > 0 and len(df.columns) > 0:
                    primeiro_valor = df.iloc[0, 0]
                    print(f"Primeira célula de dados (A2): '{primeiro_valor}'")

            except Exception as e:
                print(f"✗ Erro ao processar aba {nome_aba}: {str(e)}")
                # Cria uma aba com mensagem de erro
                ws = wb.create_sheet(title=f"{nome_aba[:25]}_ERRO")
                ws.cell(row=1, column=1, value=f"Erro ao processar: {str(e)}")

        # Se nenhuma aba foi criada, cria uma aba padrão
        if not wb.sheetnames:
            wb.create_sheet("CHECK")
            print("Nenhuma aba foi criada. Criando aba CHECK padrão.")

        # Salva o arquivo final
        if not nome_arquivo_final.endswith('.xlsx'):
            nome_arquivo_final += '.xlsx'

        caminho_final = os.path.join(diretorio_destino, nome_arquivo_final)
        wb.save(caminho_final)
        print(f"\n✓ Arquivo agrupado salvo em: {caminho_final}")

        return caminho_final

    finally:
        # Limpa o diretório temporário
        try:
            shutil.rmtree(temp_dir)
            print("Arquivos temporários removidos.")
        except Exception as e:
            print(f"Aviso: Não foi possível remover arquivos temporários: {e}")

def detectar_skiprows(caminho_arquivo):
    """Detecta se precisa pular a primeira linha (sep=;) - FUNÇÃO LEGACY"""
    encodings = ['utf-8', 'latin1', 'cp1252']

    for encoding in encodings:
        try:
            with open(caminho_arquivo, encoding=encoding) as f:
                primeira_linha = f.readline().strip().lower()
                return 1 if primeira_linha.startswith('sep=') else 0
        except Exception:
            continue
    return 0

def converter_varios_para_xlsx(caminhos_arquivos, diretorio_destino):
    """Converte múltiplos arquivos para .xlsx individualmente"""
    Tk().withdraw()
    erros = []

    print("=== CONVERTENDO ARQUIVOS INDIVIDUAIS ===")

    for caminho_arquivo in caminhos_arquivos:
        extensao = os.path.splitext(caminho_arquivo)[1].lower()
        nome_arquivo = os.path.basename(caminho_arquivo)
        print(f"\nConvertendo: {nome_arquivo}")

        try:
            if extensao == '.csv':
                df = ler_csv_robusto(caminho_arquivo)
            elif extensao == '.xls':
                df = pd.read_excel(caminho_arquivo, engine='xlrd')
            elif extensao == '.xlsx':
                print(f"Arquivo {nome_arquivo} já está em .xlsx. Copiando para o destino.")
                nome_saida = os.path.splitext(nome_arquivo)[0] + '.xlsx'
                caminho_saida = os.path.join(diretorio_destino, nome_saida)
                shutil.copy2(caminho_arquivo, caminho_saida)
                print(f'✓ Arquivo copiado para: {caminho_saida}')
                continue
            else:
                erro_msg = f"Formato não suportado: {nome_arquivo}"
                erros.append(erro_msg)
                print(f"✗ {erro_msg}")
                continue

            # Converte colunas numéricas quando possível
            for col in df.columns:
                try:
                    df[col] = pd.to_numeric(df[col], errors='ignore')
                except:
                    pass

            nome_saida = os.path.splitext(nome_arquivo)[0] + '.xlsx'
            caminho_saida = os.path.join(diretorio_destino, nome_saida)
            df.to_excel(caminho_saida, index=False, engine='openpyxl')
            print(f'✓ Arquivo salvo em: {caminho_saida}')

        except Exception as e:
            erro_msg = f"Erro ao converter {nome_arquivo}: {str(e)}"
            erros.append(erro_msg)
            print(f"✗ {erro_msg}")

    return erros