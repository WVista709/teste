import tkinter as tk
from tkinter import filedialog, messagebox
import conversor
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import os

class InterfaceAgrupamentoExcel:
    def __init__(self, root):
        self.root = root
        self.root.title("Interface de Agrupamento de Excel")
        self.root.geometry("600x550")
        self.arquivos_selecionados = {}
        self.diretorio_destino = None
        self.nome_arquivo_final = None

        # Defina aqui os grupos e os nomes dos arquivos
        self.grupos = [
            ("COMPRAS", ["COMPRAS SEFAZ", "COMPRAS ALTERDATA", "COMPRAS PRODUTOS"]),
            ("VENDAS", ["VENDAS SEFAZ", "VENDAS ALTERDATA", "VENDAS PRODUTOS"]),
        ]

        self.labels = {}

        y = 20
        for grupo_nome, abas in self.grupos:
            frame = tk.LabelFrame(root, text=grupo_nome, padx=10, pady=10)
            frame.place(x=20, y=y)
            self.criar_botoes_arquivo(frame, abas)
            y += 150

        # Seção para escolher diretório de destino
        tk.Label(root, text="Diretório de destino:", font=("Arial", 10, "bold")).place(x=20, y=y+30)
        self.botao_diretorio = tk.Button(root, text="Escolher diretório de destino", command=self.selecionar_diretorio, width=25)
        self.botao_diretorio.place(x=20, y=y+55)
        
        self.label_diretorio = tk.Label(root, text="Nenhum diretório selecionado", anchor="w", width=50, relief="sunken")
        self.label_diretorio.place(x=20, y=y+85)

        # Campo para nome do arquivo final
        tk.Label(root, text="Nome do arquivo final:", font=("Arial", 10, "bold")).place(x=20, y=y+115)
        self.entry_nome_arquivo = tk.Entry(root, width=30)
        self.entry_nome_arquivo.place(x=20, y=y+140)
        self.entry_nome_arquivo.insert(0, "planilhas_agrupadas")  # Nome padrão

        # Botão confirmar
        self.botao_confirmar = tk.Button(root, text="Agrupar em um Excel", command=self.confirmar, state="disabled",bg="blue", fg="white", font=("Arial", 10, "bold"))
        self.botao_confirmar.place(x=200, y=y+170)

    def criar_botoes_arquivo(self, frame, abas):
        for i, aba in enumerate(abas):
            btn = tk.Button(frame, text=aba, width=20, command=lambda a=aba: self.selecionar_arquivo(a))
            btn.grid(row=i, column=0, padx=5, pady=2)
            lbl = tk.Label(frame, text="Nada selecionado", anchor="w", width=40, relief="sunken")
            lbl.grid(row=i, column=1, padx=5, pady=2)
            self.labels[aba] = lbl

    def selecionar_arquivo(self, aba_nome):
        caminho = filedialog.askopenfilename(
            title=f'Escolha o arquivo {aba_nome}',
            filetypes=[("Arquivos Excel", "*.xlsx *.csv *.xls"), ("Todos os arquivos", "*.*")]
        )
        if caminho:
            self.arquivos_selecionados[aba_nome] = caminho
            nome_arquivo = caminho.split('/')[-1]
            self.labels[aba_nome].config(text=nome_arquivo)

    def selecionar_diretorio(self):
        caminho = filedialog.askdirectory(title="Escolha o diretório para salvar o arquivo agrupado")
        if caminho:
            self.diretorio_destino = caminho
            self.label_diretorio.config(text=caminho)
            self.botao_confirmar.config(state="normal")  # Habilita o botão

    def confirmar(self):
        if not self.diretorio_destino:
            messagebox.showwarning("Atenção", "Escolha o diretório de destino antes de confirmar.")
            return

        nome_arquivo = self.entry_nome_arquivo.get().strip()
        if not nome_arquivo:
            nome_arquivo = "planilhas_agrupadas"

        caminho_final = os.path.join(self.diretorio_destino, f"{nome_arquivo}.xlsx")

        self.botao_confirmar.config(state="disabled", text="Agrupando...")
        self.root.update()

        try:
            # Passa o dicionário de arquivos e o caminho final para a função
            conversor.agrupar_excels_em_um(self.arquivos_selecionados, caminho_final)

            # Executa as funções de pós-processamento
            sefaz(caminho_final)
            alterdata(caminho_final)
            produto(caminho_final)

            messagebox.showinfo("Sucesso", f"Arquivos agrupados com sucesso!\n\nArquivo criado:\n{caminho_final}")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro durante o agrupamento: {str(e)}")
        finally:
            self.botao_confirmar.config(state="normal", text="Agrupar em um Excel")

def CelulaValorMesclada(caminho, aba, coluna_inicio, coluna_fim, valor, linha=1):
    wb = load_workbook(caminho)
    if aba not in wb.sheetnames:
        print(f"Aba '{aba}' não encontrada.")
        return

    ws = wb[aba]
    letra_inicio = get_column_letter(coluna_inicio)
    letra_fim = get_column_letter(coluna_fim)
    celula_inicio = f"{letra_inicio}{linha}"
    celula_fim = f"{letra_fim}{linha}"
    ws.merge_cells(f"{celula_inicio}:{celula_fim}")
    ws[celula_inicio] = valor
    wb.save(caminho)

def alterdata(caminho_final):
    abas = ["COMPRAS ALTERDATA" , "VENDAS ALTERDATA"]
    wb = load_workbook(caminho_final)
    abas_existentes = wb.sheetnames

    for aba in abas:
        if aba not in abas_existentes:
            continue
            
        colunas, linhas = contar_colunas_linhas_preenchidas(caminho_final, aba)
        CelulaValor(caminho_final, aba, colunas + 1, "SEFAZ", linha=1)
        CelulaValor(caminho_final, aba, colunas + 2, "PRODUTOS", linha=1)

        # Define a aba de referência
        if aba == "COMPRAS ALTERDATA":
            aba_sefaz = "COMPRAS SEFAZ"
            aba_produto = "COMPRAS PRODUTOS"
        else:  # VENDAS ALTERDATA
            aba_sefaz = "VENDAS SEFAZ"
            aba_produto = "VENDAS PRODUTOS"

        # Só insere fórmula se a aba de referência existir
        if aba_sefaz in abas_existentes:
            for linha in range(2, linhas + 1):
                formula = f'=IFERROR(VLOOKUP(B{linha},\'{aba_sefaz}\'!C:C,1,0),"ERRO")'
                CelulaValor(caminho_final, aba, colunas + 1, formula, linha=linha)
        else:
            print(f"Aba de referência '{aba_sefaz}' não existe. Fórmulas não inseridas em '{aba}'.")

        # Só insere fórmula se a aba de referência existir
        if aba_produto in abas_existentes:
            for linha in range(2, linhas + 1):
                formula = f'=IFERROR(VLOOKUP(B{linha},\'{aba_produto}\'!B:B,1,0),"ERRO")'
                CelulaValor(caminho_final, aba, colunas + 2, formula, linha=linha)
        else:
            print(f"Aba de referência '{aba_produto}' não existe. Fórmulas não inseridas em '{aba}'.")

def sefaz(caminho_final):
    abas = ["COMPRAS SEFAZ", "VENDAS SEFAZ"]

    # Carrega o workbook uma vez para checar as abas existentes
    wb = load_workbook(caminho_final)
    abas_existentes = wb.sheetnames

    for aba in abas:
        if aba not in abas_existentes:
            continue
            
        colunas, linhas = contar_colunas_linhas_preenchidas(caminho_final, aba)

        # Cabeçalhos
        CelulaValor(caminho_final, aba, colunas + 1, "CANCELADAS", linha=1)
        CelulaValor(caminho_final, aba, colunas + 2, "ALTERDATA", linha=1)
        CelulaValor(caminho_final, aba, colunas + 3, "PRODUTOS", linha=1)

        for linha in range(2, linhas + 1):
            formula = f'=IF(N{linha}="AUTORIZADA", "NÃO", "SIM")'
            CelulaValor(caminho_final, aba, colunas + 1, formula, linha=linha)

        # Define a aba de referência
        if aba == "COMPRAS SEFAZ":
            aba_alterdata = "COMPRAS ALTERDATA"
            aba_produto = "COMPRAS PRODUTOS"
        else:  # VENDAS SEFAZ
            aba_alterdata = "VENDAS ALTERDATA"
            aba_produto = "VENDAS PRODUTOS"

        # Só insere fórmula se a aba de referência existir
        if aba_alterdata in abas_existentes:
            for linha in range(2, linhas + 1):
                formula = f'=IFERROR(VLOOKUP(C{linha},\'{aba_alterdata}\'!B:B,1,0),"ERRO")'
                CelulaValor(caminho_final, aba, colunas + 2, formula, linha=linha)
        else:
            print(f"Aba de referência '{aba_alterdata}' não existe. Fórmulas não inseridas em '{aba}'.")

        if aba_produto in abas_existentes:
            for linha in range(2, linhas + 1):
                formula = f'=IFERROR(VLOOKUP(C{linha},\'{aba_produto}\'!B:B,1,0),"ERRO")'
                CelulaValor(caminho_final, aba, colunas + 3, formula, linha=linha)    
        else:
            print(f"Aba de referência '{aba_produto}' não existe. Fórmulas não inseridas em '{aba}'.")

def produto(caminho_final):
    abas = ["COMPRAS PRODUTOS" , "VENDAS PRODUTOS"]
    wb = load_workbook(caminho_final)
    abas_existentes = wb.sheetnames

    for aba in abas:
        if aba not in abas_existentes:
            continue
            
        colunas, linhas = contar_colunas_linhas_preenchidas(caminho_final, aba)
        CelulaValor(caminho_final, aba, colunas + 1, "SEFAZ", linha=1)
        CelulaValor(caminho_final, aba, colunas + 2, "ALTERDATA", linha=1)

        # Define a aba de referência
        if aba == "COMPRAS PRODUTOS":
            aba_sefaz = "COMPRAS SEFAZ"
            aba_produto = "COMPRAS ALTERDATA"
        else:  # VENDAS ALTERDATA
            aba_sefaz = "VENDAS SEFAZ"
            aba_produto = "VENDAS ALTERDATA"

        # Só insere fórmula se a aba de referência existir
        if aba_sefaz in abas_existentes:
            for linha in range(2, linhas + 1):
                formula = f'=IFERROR(VLOOKUP(B{linha},\'{aba_sefaz}\'!C:C,1,0),"ERRO")'
                CelulaValor(caminho_final, aba, colunas + 1, formula, linha=linha)
        else:
            print(f"Aba de referência '{aba_sefaz}' não existe. Fórmulas não inseridas em '{aba}'.")

        # Só insere fórmula se a aba de referência existir
        if aba_produto in abas_existentes:
            for linha in range(2, linhas + 1):
                formula = f'=IFERROR(VLOOKUP(B{linha},\'{aba_produto}\'!B:B,1,0),"ERRO")'
                CelulaValor(caminho_final, aba, colunas + 2, formula, linha=linha)
        else:
            print(f"Aba de referência '{aba_produto}' não existe. Fórmulas não inseridas em '{aba}'.")

def contar_colunas_linhas_preenchidas(caminho_arquivo, nome_aba):
    wb = load_workbook(caminho_arquivo, data_only=True)
    if nome_aba not in wb.sheetnames:
        return 0, 0
    ws = wb[nome_aba]
    primeira_linha = next(ws.iter_rows(min_row=1, max_row=1, values_only=True))
    colunas_preenchidas = sum(1 for cell in primeira_linha if cell not in (None, ""))
    linhas_preenchidas = sum(
        1 for row in ws.iter_rows(values_only=True)
        if any(cell not in (None, "") for cell in row)
    )
    return colunas_preenchidas, linhas_preenchidas

def CelulaValor(caminho, aba, coluna_num, valor, linha=1):
    from openpyxl import load_workbook
    wb = load_workbook(caminho)
    if aba not in wb.sheetnames:
        print(f"Aba '{aba}' não encontrada.")
        return

    ws = wb[aba]
    letra_coluna = get_column_letter(coluna_num)
    ws[f"{letra_coluna}{linha}"] = valor
    wb.save(caminho)

if __name__ == "__main__":
    root = tk.Tk()
    app = InterfaceAgrupamentoExcel(root)
    root.mainloop()