import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os, time, threading, check_bruto, conversor, shutil, sys 

class InterfaceAgrupamentoExcel:
    def __init__(self, root):
        self.root = root
        self.root.title("Interface de Agrupamento de Excel")
        self.root.geometry("625x600")
        self.arquivos_selecionados = {}
        self.diretorio_destino = None
        self.nome_arquivo_final = None

        # Definição dos modos e seus grupos/abas
        self.modos = {
            "Conferência de Nota": [
                ("COMPRAS", ["COMPRAS SEFAZ", "COMPRAS ALTERDATA"]),
                ("VENDAS", ["VENDAS SEFAZ", "VENDAS ALTERDATA"]),
            ],
            "Check": [
                ("COMPRAS", ["COMPRAS SEFAZ", "COMPRAS ALTERDATA", "COMPRAS PRODUTOS"]),
                ("VENDAS", ["VENDAS SEFAZ", "VENDAS ALTERDATA", "VENDAS PRODUTOS"]),
            ]
        }

        self.labels = {}
        self.frames_grupos = []  # Para guardar frames e limpar depois

        # Botões para escolher modo
        self.modo_atual = None
        self.frame_modos = tk.Frame(root)
        self.frame_modos.place(x=20, y=10)

        self.botao_conferencia = tk.Button(self.frame_modos, text="Conferência de Nota", width=20, command=lambda: self.selecionar_modo("Conferência de Nota"))
        self.botao_conferencia.grid(row=0, column=0, padx=5)

        self.botao_check = tk.Button(self.frame_modos, text="Check", width=20, command=lambda: self.selecionar_modo("Check"))
        self.botao_check.grid(row=0, column=1, padx=5)

        # Frame onde os grupos e abas serão exibidos
        self.frame_arquivos = tk.Frame(root)
        self.frame_arquivos.place(x=20, y=50)

        # Seção para escolher diretório de destino
        self.label_diretorio_texto = tk.Label(root, text="Diretório de destino:", font=("Arial", 10, "bold"))
        self.label_diretorio_texto.place(x=20, y=325)
        self.botao_diretorio = tk.Button(root, text="Escolher diretório de destino", command=self.selecionar_diretorio, width=25)
        self.botao_diretorio.place(x=20, y=350)
        
        self.label_diretorio = tk.Label(root, text="Nenhum diretório selecionado", anchor="w", width=50, relief="sunken")
        self.label_diretorio.place(x=20, y=380)

        # Barra de progresso
        self.progress = ttk.Progressbar(root, orient="horizontal", length=560, mode="determinate")
        self.progress.place(x=20, y=420)

        # Label para mostrar a etapa atual e tempo decorrido
        self.label_status = tk.Label(root, text="Aguardando ação...", font=("Arial", 10, "italic"))
        self.label_status.place(x=20, y=450)

        # Campo para nome do arquivo final
        tk.Label(root, text="Nome do arquivo final:", font=("Arial", 10, "bold")).place(x=20, y=480)
        self.entry_nome_arquivo = tk.Entry(root, width=30)
        self.entry_nome_arquivo.place(x=20, y=500)
        self.entry_nome_arquivo.insert(0, "planilhas_agrupadas")  # Nome padrão

        # Botão confirmar
        self.botao_confirmar = tk.Button(root, text="Agrupar em um Excel", command=self.confirmar, state="disabled", bg="blue", fg="white", font=("Arial", 10, "bold"))
        self.botao_confirmar.place(x=200, y=525)

        # Variáveis para controle do tempo e estado
        self.tempo_inicio = None
        self.tempo_decorrido = 0
        self.processo_rodando = False
        self.etapa_atual = ""

        # Inicializa com modo Conferência de Nota
        self.selecionar_modo("Conferência de Nota")

    def selecionar_modo(self, modo):
        if self.processo_rodando:
            messagebox.showwarning("Atenção", "Processo em andamento. Aguarde terminar para mudar o modo.")
            return

        self.modo_atual = modo
        self.arquivos_selecionados.clear()
        self.labels.clear()

        # Limpa frames antigos
        for f in self.frames_grupos:
            f.destroy()
        self.frames_grupos.clear()

        # Cria os frames e botões para o modo selecionado
        grupos = self.modos[modo]
        for grupo_nome, abas in grupos:
            frame = tk.LabelFrame(self.frame_arquivos, text=grupo_nome, padx=10, pady=10)
            frame.pack(fill="x", pady=5)
            self.criar_botoes_arquivo(frame, abas)
            self.frames_grupos.append(frame)

        # Limpa labels e diretório selecionado
        self.label_diretorio.config(text="Nenhum diretório selecionado")
        self.diretorio_destino = None
        self.botao_confirmar.config(state="disabled")

        # Reseta nome do arquivo
        self.entry_nome_arquivo.delete(0, tk.END)
        self.entry_nome_arquivo.insert(0, "planilhas_agrupadas")

        self.label_status.config(text=f"Modo selecionado: {modo}. Aguardando ação...")

        self.atualizar_estilo_botoes()  # Atualiza o estilo dos botões de modo

    def atualizar_estilo_botoes(self):
        cor_padrao = self.root.cget("bg")
        if self.modo_atual == "Conferência de Nota":
            self.botao_conferencia.config(bg="blue", fg="white")
            self.botao_check.config(bg=cor_padrao, fg="black")
        else:
            self.botao_check.config(bg="blue", fg="white")
            self.botao_conferencia.config(bg=cor_padrao, fg="black")

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
            nome_arquivo = os.path.basename(caminho)
            self.labels[aba_nome].config(text=nome_arquivo)

    def selecionar_diretorio(self):
        caminho = filedialog.askdirectory(title="Escolha o diretório para salvar o arquivo agrupado")
        if caminho:
            self.diretorio_destino = caminho
            self.label_diretorio.config(text=caminho)
            self.botao_confirmar.config(state="normal")

    def atualizar_tempo(self):
        if self.processo_rodando:
            self.tempo_decorrido = time.time() - self.tempo_inicio
            texto = f"{self.etapa_atual} - Tempo decorrido: {self.tempo_decorrido:.1f} s"
            self.label_status.config(text=texto)
            self.root.after(500, self.atualizar_tempo)

    def confirmar(self):
        if not self.diretorio_destino:
            messagebox.showwarning("Atenção", "Escolha o diretório de destino antes de confirmar.")
            return

        nome_arquivo = self.entry_nome_arquivo.get().strip()

        # Criando os arquivos e obtendo o caminho final do Excel
        caminho_final = self.criando_arquivos(nome_arquivo, self.modo_atual)

        if caminho_final is None:
            # Erro já mostrado, apenas retorna para não continuar
            return

        self.botao_confirmar.config(state="disabled", text="Agrupando...")
        self.progress['value'] = 0
        self.processo_rodando = True
        self.tempo_inicio = time.time()
        self.etapa_atual = "Iniciando agrupamento dos arquivos..."
        self.root.update()

        self.atualizar_tempo()

        thread = threading.Thread(target=self.processar_etapas, args=(caminho_final,))
        thread.start()

    def criando_arquivos(self, nome_arquivo, modo_atual):
        if not nome_arquivo:
            nome_arquivo = "planilhas_agrupadas"

        # Pasta base dentro do diretório destino
        caminho_pasta = os.path.join(self.diretorio_destino, nome_arquivo)

        # Para o modo "Check", cria subpasta "excel"
        if modo_atual == "Check":
            caminho_pasta_excel = os.path.join(caminho_pasta, "excel")
            if not os.path.exists(caminho_pasta_excel):
                os.makedirs(caminho_pasta_excel)
                print("Pasta criada:", caminho_pasta_excel)
            else:
                print("Pasta já existe:", caminho_pasta_excel)
            caminho_final = os.path.join(caminho_pasta_excel, f"{nome_arquivo}.xlsx")

            # Copiando o Power BI para a pasta base (não dentro do excel)
            power_bi_origem = resource_path("powerBI.pbix")  # Ajuste o caminho se necessário

            if not os.path.isfile(power_bi_origem):
                messagebox.showerror("Erro", f"Arquivo '{power_bi_origem}' não encontrado. Verifique o caminho.")
                return None
            else:
                nome_power_bi_destino = f"power bi {nome_arquivo}.pbix"
                caminho_power_bi_destino = os.path.join(caminho_pasta, nome_power_bi_destino)
                shutil.copy(power_bi_origem, caminho_power_bi_destino)
        else:
            # Para outros modos, cria só a pasta base
            if not os.path.exists(caminho_pasta):
                os.makedirs(caminho_pasta)
                print("Pasta criada:", caminho_pasta)
            else:
                print("Pasta já existe:", caminho_pasta)
            # Salva o Excel direto na pasta base
            caminho_final = os.path.join(caminho_pasta, f"{nome_arquivo}.xlsx")

        return caminho_final

    def processar_etapas(self, caminho_final):
        modo = self.modo_atual
        try:
            etapas = [
                ("Agrupando arquivos", lambda: conversor.agrupar_excels_em_um(self.arquivos_selecionados, caminho_final)),
                ("Processando SEFAZ", lambda: check_bruto.sefaz(caminho_final, modo)),
                ("Processando ALTERDATA", lambda: check_bruto.alterdata(caminho_final, modo)),
                ("Processando PRODUTO", lambda: check_bruto.produto(caminho_final)),
                ("Verificando COMPRAS", lambda: check_bruto.Check.check_compras(caminho_final, modo)),
                ("Verificando VENDAS", lambda: check_bruto.Check.check_vendas(caminho_final, modo)),
            ]
            passo = 100 / len(etapas)

            for i, (descricao, func) in enumerate(etapas):
                self.etapa_atual = descricao
                func()
                self.progress['value'] = (i + 1) * passo
                self.root.update()

            elapsed_time = time.time() - self.tempo_inicio
            elapsed_str = f"{elapsed_time:.2f} segundos"

            def finalizar():
                self.label_status.config(text=f"Concluído em {elapsed_str}. Processo finalizado.")
                self.botao_confirmar.config(state="normal", text="Agrupar em um Excel")
                self.progress['value'] = 0
                self.processo_rodando = False

                # Limpar arquivos selecionados e diretório
                self.arquivos_selecionados.clear()

                # Resetar labels dos arquivos para "Nada selecionado"
                for aba, lbl in self.labels.items():
                    lbl.config(text="Nada selecionado")

                # Resetar nome do arquivo final para padrão
                self.entry_nome_arquivo.delete(0, tk.END)
                self.entry_nome_arquivo.insert(0, "planilhas_agrupadas")

            self.root.after(0, finalizar)

        except Exception as e:
            def erro(exc):
                messagebox.showerror("Erro", f"Erro durante o agrupamento: {str(exc)}")
                self.botao_confirmar.config(state="normal", text="Agrupar em um Excel")
                self.progress['value'] = 0
                self.processo_rodando = False
            self.root.after(0, erro, e)

def resource_path(relative_path):
    """Obtem o caminho absoluto para o recurso, funciona para dev e PyInstaller."""
    try:
        # PyInstaller cria uma pasta temporária e define essa variável
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

if __name__ == "__main__":
    root = tk.Tk()
    app = InterfaceAgrupamentoExcel(root)
    root.mainloop()