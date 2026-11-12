"""Interface gráfica para adicionar o setor às glosas."""
import os
import queue
import threading
import traceback
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText
import tkinter as tk

TITULO = "Adicionar Setor às Glosas"
TIPOS_EXCEL = [("Planilhas Excel", "*.xlsx"), ("Todos os arquivos", "*.*")]


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(TITULO)
        self.minsize(640, 440)
        self.geometry("720x480")

        self.entrada = tk.StringVar()
        self.saida = tk.StringVar()
        self.status = tk.StringVar(value="Selecione a planilha de entrada.")
        self.fila = queue.Queue()      # mensagens da thread de processamento
        self.rodando = False

        self._montar_tela()
        self.protocol("WM_DELETE_WINDOW", self._fechar)
        self.after(100, self._ler_fila)

    # ---------------- Tela ----------------
    def _montar_tela(self):
        quadro = ttk.Frame(self, padding=16)
        quadro.pack(fill="both", expand=True)
        quadro.columnconfigure(1, weight=1)

        ttk.Label(quadro, text=TITULO, font=("Segoe UI", 14, "bold")).grid(
            row=0, column=0, columnspan=3, sticky="w", pady=(0, 4))
        ttk.Label(
            quadro,
            text="Filtra as linhas glosadas do relatório e adiciona a coluna Setor consultando o Tasy.",
            foreground="#555",
        ).grid(row=1, column=0, columnspan=3, sticky="w", pady=(0, 14))

        ttk.Label(quadro, text="1. Planilha de entrada:").grid(row=2, column=0, sticky="w")
        self.campo_entrada = ttk.Entry(quadro, textvariable=self.entrada)
        self.campo_entrada.grid(row=2, column=1, sticky="ew", padx=8, pady=4)
        self.botao_entrada = ttk.Button(quadro, text="Procurar...", command=self._escolher_entrada)
        self.botao_entrada.grid(row=2, column=2, pady=4)

        ttk.Label(quadro, text="2. Salvar resultado em:").grid(row=3, column=0, sticky="w")
        self.campo_saida = ttk.Entry(quadro, textvariable=self.saida)
        self.campo_saida.grid(row=3, column=1, sticky="ew", padx=8, pady=4)
        self.botao_saida = ttk.Button(quadro, text="Procurar...", command=self._escolher_saida)
        self.botao_saida.grid(row=3, column=2, pady=4)

        self.botao_processar = ttk.Button(quadro, text="3. Processar", command=self._iniciar)
        self.botao_processar.grid(row=4, column=0, columnspan=3, pady=(12, 8), ipadx=20, ipady=4)

        self.barra = ttk.Progressbar(quadro, maximum=100)
        self.barra.grid(row=5, column=0, columnspan=3, sticky="ew")
        ttk.Label(quadro, textvariable=self.status).grid(
            row=6, column=0, columnspan=3, sticky="w", pady=(4, 8))

        self.log = ScrolledText(quadro, height=10, state="disabled", font=("Consolas", 9))
        self.log.grid(row=7, column=0, columnspan=3, sticky="nsew")
        quadro.rowconfigure(7, weight=1)

    # ---------------- Seleção de arquivos ----------------
    def _escolher_entrada(self):
        caminho = filedialog.askopenfilename(title="Selecione o relatório de pagamento e glosa",
                                             filetypes=TIPOS_EXCEL)
        if not caminho:
            return
        self.entrada.set(os.path.normpath(caminho))
        if not self.saida.get():  # sugere o nome de saída na mesma pasta
            from .main import nome_saida_padrao
            self.saida.set(os.path.normpath(nome_saida_padrao(Path(caminho))))
        self.status.set("Confira o local de saída e clique em Processar.")

    def _escolher_saida(self):
        atual = Path(self.saida.get() or self.entrada.get() or Path.home())
        caminho = filedialog.asksaveasfilename(
            title="Salvar planilha com setores como",
            defaultextension=".xlsx",
            filetypes=TIPOS_EXCEL,
            initialdir=atual.parent if atual.suffix else atual,
            initialfile=atual.name if atual.suffix else "",
        )
        if caminho:
            self.saida.set(os.path.normpath(caminho))

    # ---------------- Processamento ----------------
    def _validar(self):
        entrada = Path(self.entrada.get().strip())
        saida_txt = self.saida.get().strip()
        if not self.entrada.get().strip() or not entrada.is_file():
            return "Selecione uma planilha de entrada válida."
        if not saida_txt:
            return "Escolha onde salvar o resultado."
        saida = Path(saida_txt)
        if saida.suffix.lower() != ".xlsx":
            self.saida.set(str(saida.with_suffix(".xlsx")))
            saida = saida.with_suffix(".xlsx")
        if saida.resolve() == entrada.resolve():
            return "O arquivo de saída não pode ser o mesmo da entrada."
        if not saida.parent.is_dir():
            return f"A pasta de saída não existe:\n{saida.parent}"
        if saida.exists() and not messagebox.askyesno(
                TITULO, f"O arquivo {saida.name} já existe.\nDeseja substituí-lo?"):
            return ""
        return None

    def _iniciar(self):
        erro = self._validar()
        if erro is not None:
            if erro:
                messagebox.showwarning(TITULO, erro)
            return

        self._limpar_log()
        self.barra["value"] = 0
        self._travar(True)
        self.status.set("Processando... aguarde.")
        threading.Thread(
            target=self._trabalho,
            args=(self.entrada.get().strip(), self.saida.get().strip()),
            daemon=True,
        ).start()

    def _trabalho(self, entrada, saida):
        """Roda fora da tela para ela não congelar; conversa com a tela pela fila."""
        try:
            from .main import processar  # importa aqui: conecta ao Oracle só quando precisa
            resumo = processar(
                entrada, saida,
                log=lambda msg: self.fila.put(("log", msg)),
                progresso=lambda pct: self.fila.put(("progresso", pct)),
            )
            self.fila.put(("fim", resumo))
        except Exception as e:  # mostra qualquer erro de forma amigável
            self.fila.put(("log", traceback.format_exc()))
            self.fila.put(("erro", str(e) or e.__class__.__name__))

    def _ler_fila(self):
        try:
            while True:
                tipo, dado = self.fila.get_nowait()
                if tipo == "log":
                    self._escrever_log(dado)
                elif tipo == "progresso":
                    self.barra["value"] = dado
                elif tipo == "fim":
                    self._concluir(dado)
                elif tipo == "erro":
                    self._travar(False)
                    self.status.set("Ocorreu um erro. Veja os detalhes abaixo.")
                    messagebox.showerror(TITULO, dado)
        except queue.Empty:
            pass
        self.after(100, self._ler_fila)

    def _concluir(self, resumo):
        self._travar(False)
        self.status.set("Concluído!")
        texto = (f"Processamento concluído.\n\n"
                 f"Linhas glosadas: {resumo['linhas']}\n"
                 f"Guias consultadas: {resumo['guias']}\n"
                 f"Sem setor encontrado: {resumo['nao_encontrado']}\n")
        if resumo["erros"]:
            texto += f"Linhas com erro na consulta: {resumo['erros']}\n"
        texto += "\nDeseja abrir a planilha gerada?"
        if messagebox.askyesno(TITULO, texto) and hasattr(os, "startfile"):
            os.startfile(resumo["saida"])

    # ---------------- Utilidades ----------------
    def _travar(self, travar):
        self.rodando = travar
        estado = "disabled" if travar else "normal"
        for w in (self.campo_entrada, self.campo_saida, self.botao_entrada,
                  self.botao_saida, self.botao_processar):
            w.configure(state=estado)

    def _escrever_log(self, msg):
        self.log.configure(state="normal")
        self.log.insert("end", msg.rstrip() + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def _limpar_log(self):
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")

    def _fechar(self):
        if self.rodando and not messagebox.askyesno(
                TITULO, "O processamento ainda está em andamento.\nDeseja sair mesmo assim?"):
            return
        self.destroy()


def main():
    App().mainloop()


if __name__ == "__main__":
    main()