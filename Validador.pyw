"""
Janela do Validador de Legibilidade de Documentos.
Dê dois cliques neste arquivo para abrir (não precisa do VS Code).
Precisa estar na mesma pasta que validador_legibilidade.py.
"""

import json
import os
import queue
import sys
import threading
import traceback
from pathlib import Path

# pythonw (duplo clique) não tem console: evita erro ao tentar imprimir
if sys.stdout is None:
    sys.stdout = open(os.devnull, "w", encoding="utf-8")
if sys.stderr is None:
    sys.stderr = open(os.devnull, "w", encoding="utf-8")

PASTA_PROGRAMA = Path(__file__).resolve().parent

try:
    import tkinter as tk
    from tkinter import filedialog, messagebox, simpledialog, ttk
except Exception:
    (PASTA_PROGRAMA / "erro_validador.txt").write_text(
        "O Python foi instalado sem o componente de janelas (tcl/tk).\n"
        "Rode o instalador do Python de novo, escolha Modify e marque 'tcl/tk and IDLE'.\n\n"
        + traceback.format_exc(), encoding="utf-8")
    raise
ARQUIVO_ENV = PASTA_PROGRAMA / ".env"
ARQUIVO_PREFS = PASTA_PROGRAMA / "_ultima_pasta.json"
sys.path.insert(0, str(PASTA_PROGRAMA))

def _instalar_bibliotecas_e_reabrir(faltando: str) -> None:
    """Instala as bibliotecas no MESMO Python que está abrindo a janela e reabre o programa."""
    import subprocess
    raiz = tk.Tk()
    raiz.withdraw()
    if sys.version_info >= (3, 15):
        messagebox.showerror(
            "Validador de Documentos",
            f"Este programa foi aberto com o Python {sys.version.split()[0]}, que é novo demais: "
            "algumas bibliotecas ainda não funcionam nele.\n\n"
            "Dê dois cliques em INSTALAR.bat (na pasta do programa). Ele escolhe uma versão "
            "estável do Python e recria o ícone da Área de Trabalho.")
        sys.exit(1)
    if not messagebox.askyesno(
            "Validador de Documentos",
            f"Falta instalar algumas bibliotecas neste computador (ex.: {faltando}).\n\n"
            "Instalar agora? Pode levar alguns minutos."):
        sys.exit(1)
    exe = Path(sys.executable)
    python_console = exe.with_name("python.exe") if exe.name.lower() == "pythonw.exe" else exe
    requisitos = PASTA_PROGRAMA / "requirements.txt"
    comando = [str(python_console), "-m", "pip", "install", "--disable-pip-version-check"]
    comando += ["-r", str(requisitos)] if requisitos.exists() else [
        "pymupdf", "pillow", "opencv-python", "numpy", "pytesseract",
        "anthropic", "pandas", "openpyxl", "python-dotenv"]
    aviso = tk.Toplevel(raiz)
    aviso.title("Validador de Documentos")
    tk.Label(aviso, text="Instalando as bibliotecas…\nIsso pode levar alguns minutos. Não feche o programa.",
             padx=30, pady=20).pack()
    aviso.update()
    sem_janela = getattr(subprocess, "CREATE_NO_WINDOW", 0)  # só existe no Windows
    resultado = subprocess.run(comando, capture_output=True, text=True, errors="replace",
                               creationflags=sem_janela)
    aviso.destroy()
    log = PASTA_PROGRAMA / "log_instalacao.txt"
    log.write_text(f"Comando: {' '.join(comando)}\n\n{resultado.stdout}\n{resultado.stderr}",
                   encoding="utf-8")
    if resultado.returncode != 0:
        ultimas = "\n".join((resultado.stderr or resultado.stdout).strip().splitlines()[-6:])
        messagebox.showerror(
            "Validador de Documentos",
            "A instalação das bibliotecas falhou.\n\n"
            f"Detalhes (salvos em {log.name}):\n{ultimas}")
        sys.exit(1)
    messagebox.showinfo("Validador de Documentos", "Bibliotecas instaladas! O programa vai abrir agora.")
    subprocess.Popen([str(exe), str(Path(__file__).resolve())], cwd=str(PASTA_PROGRAMA))
    sys.exit(0)


try:
    import validador_legibilidade as v
except ModuleNotFoundError as erro:
    if (PASTA_PROGRAMA / "validador_legibilidade.py").exists():
        _instalar_bibliotecas_e_reabrir(erro.name or "biblioteca")
    raiz = tk.Tk()
    raiz.withdraw()
    messagebox.showerror("Validador de Documentos",
                         "Não encontrei o arquivo validador_legibilidade.py.\n"
                         "Ele precisa estar na mesma pasta deste programa.")
    sys.exit(1)
except Exception:
    raiz = tk.Tk()
    raiz.withdraw()
    messagebox.showerror(
        "Validador de Documentos",
        "Não foi possível carregar o programa.\n\n" + traceback.format_exc(limit=1))
    sys.exit(1)

COR = {v.LEGIVEL: "#1E7B34", v.ILEGIVEL: "#B42318", v.REVISAR: "#B7791F"}


def salvar_chave_no_env(chave: str) -> None:
    """Grava/atualiza ANTHROPIC_API_KEY no .env mantendo as outras linhas."""
    linhas = []
    if ARQUIVO_ENV.exists():
        linhas = [l for l in ARQUIVO_ENV.read_text(encoding="utf-8").splitlines()
                  if not l.strip().startswith("ANTHROPIC_API_KEY")]
    linhas.insert(0, f"ANTHROPIC_API_KEY={chave}")
    ARQUIVO_ENV.write_text("\n".join(linhas) + "\n", encoding="utf-8")
    os.environ["ANTHROPIC_API_KEY"] = chave


def chave_configurada() -> bool:
    chave = os.getenv("ANTHROPIC_API_KEY", "").strip()
    return bool(chave) and chave != "cole-sua-chave-aqui"


class App:
    def __init__(self, janela: tk.Tk):
        self.janela = janela
        self.fila: queue.Queue = queue.Queue()
        self.planilha = None
        self.rodando = False

        janela.title("Validador de Documentos")
        janela.geometry("820x680")
        janela.minsize(620, 460)

        principal = ttk.Frame(janela, padding=16)
        principal.pack(fill="both", expand=True)

        ttk.Label(principal, text="Validador de Legibilidade de Documentos",
                  font=("Segoe UI", 14, "bold")).pack(anchor="w")
        ttk.Label(principal, text="Escolha a pasta com os documentos (uma subpasta por colaborador) "
                  "e clique em Analisar.").pack(anchor="w", pady=(2, 12))

        prefs = self._ler_prefs()
        ttk.Label(principal, text="1. Pasta com os documentos").pack(anchor="w")
        linha_pasta = ttk.Frame(principal)
        linha_pasta.pack(fill="x")
        self.pasta = tk.StringVar(value=prefs.get("pasta", ""))
        ttk.Entry(linha_pasta, textvariable=self.pasta).pack(side="left", fill="x", expand=True)
        ttk.Button(linha_pasta, text="Escolher pasta…", command=self.escolher_pasta).pack(
            side="left", padx=(8, 0))

        ttk.Label(principal, text="2. Planilha de colaboradores (NOME e FUNÇÃO) — opcional: "
                  "confere o checklist de admissão, NRs e RACs").pack(anchor="w", pady=(10, 0))
        linha_colab = ttk.Frame(principal)
        linha_colab.pack(fill="x")
        self.colaboradores = tk.StringVar(value=prefs.get("colaboradores", ""))
        ttk.Entry(linha_colab, textvariable=self.colaboradores).pack(side="left", fill="x", expand=True)
        ttk.Button(linha_colab, text="Escolher planilha…", command=self.escolher_planilha).pack(
            side="left", padx=(8, 0))
        ttk.Button(linha_colab, text="Modelo", command=self.criar_modelo).pack(side="left", padx=(8, 0))

        linha_botoes = ttk.Frame(principal)
        linha_botoes.pack(fill="x", pady=12)
        self.bt_analisar = ttk.Button(linha_botoes, text="▶  Analisar", command=self.iniciar)
        self.bt_analisar.pack(side="left")
        self.bt_planilha = ttk.Button(linha_botoes, text="Abrir planilha",
                                      command=self.abrir_planilha, state="disabled")
        self.bt_planilha.pack(side="left", padx=8)
        ttk.Button(linha_botoes, text="Chave da IA", command=self.pedir_chave).pack(side="right")
        ttk.Button(linha_botoes, text="Regras das funções", command=self.abrir_regras).pack(
            side="right", padx=8)

        self.barra = ttk.Progressbar(principal, mode="determinate")
        self.barra.pack(fill="x")
        self.texto_progresso = ttk.Label(principal, text="Pronto.")
        self.texto_progresso.pack(anchor="w", pady=(4, 8))

        placar = ttk.Frame(principal)
        placar.pack(fill="x", pady=(0, 8))
        self.contadores = {}
        for status in (v.LEGIVEL, v.ILEGIVEL, v.REVISAR):
            caixa = tk.Frame(placar, bg=COR[status], padx=14, pady=6)
            caixa.pack(side="left", padx=(0, 8))
            numero = tk.Label(caixa, text="0", bg=COR[status], fg="white",
                              font=("Segoe UI", 16, "bold"))
            numero.pack()
            tk.Label(caixa, text=status, bg=COR[status], fg="white",
                     font=("Segoe UI", 9, "bold")).pack()
            self.contadores[status] = numero

        quadro_log = ttk.Frame(principal)
        quadro_log.pack(fill="both", expand=True)
        self.log = tk.Text(quadro_log, height=12, font=("Consolas", 9), wrap="none",
                           state="disabled")
        rolagem = ttk.Scrollbar(quadro_log, command=self.log.yview)
        self.log.configure(yscrollcommand=rolagem.set)
        self.log.pack(side="left", fill="both", expand=True)
        rolagem.pack(side="right", fill="y")
        for status, cor in COR.items():
            self.log.tag_configure(status, foreground=cor)

        janela.protocol("WM_DELETE_WINDOW", self.fechar)
        janela.after(150, self.ler_fila)

    # ---------- pasta ----------
    def _ler_prefs(self) -> dict:
        try:
            return json.loads(ARQUIVO_PREFS.read_text(encoding="utf-8"))
        except Exception:
            return {}

    def _guardar_prefs(self) -> None:
        try:
            ARQUIVO_PREFS.write_text(json.dumps({
                "pasta": self.pasta.get().strip(),
                "colaboradores": self.colaboradores.get().strip()}), encoding="utf-8")
        except Exception:
            pass

    def escolher_planilha(self) -> None:
        arquivo = filedialog.askopenfilename(
            title="Planilha de colaboradores (NOME e FUNÇÃO)",
            filetypes=[("Planilhas", "*.xlsx *.xls *.csv"), ("Todos os arquivos", "*.*")])
        if arquivo:
            self.colaboradores.set(str(Path(arquivo)))

    def criar_modelo(self) -> None:
        import checklist_mobilizacao as ck
        destino = filedialog.asksaveasfilename(
            title="Salvar modelo de planilha", defaultextension=".xlsx",
            initialfile="colaboradores.xlsx", filetypes=[("Excel", "*.xlsx")])
        if destino:
            ck.criar_modelo_colaboradores(Path(destino))
            self.colaboradores.set(str(Path(destino)))
            self._abrir_arquivo(Path(destino))

    def abrir_regras(self) -> None:
        import checklist_mobilizacao as ck
        if not ck.ARQUIVO_REGRAS.exists():
            ck.salvar_regras(ck.ler_regras())
        self._abrir_arquivo(ck.ARQUIVO_REGRAS)

    def _abrir_arquivo(self, caminho: Path) -> None:
        if sys.platform.startswith("win"):
            os.startfile(caminho)
        else:
            messagebox.showinfo("Arquivo", str(caminho))

    def escolher_pasta(self) -> None:
        pasta = filedialog.askdirectory(title="Pasta com os documentos",
                                        initialdir=self.pasta.get() or None)
        if pasta:
            self.pasta.set(str(Path(pasta)))

    # ---------- chave ----------
    def pedir_chave(self) -> bool:
        chave = simpledialog.askstring(
            "Chave da IA", "Cole aqui a sua chave da API da Anthropic\n"
            "(console.anthropic.com → API Keys):", show="*", parent=self.janela)
        if chave and chave.strip():
            salvar_chave_no_env(chave.strip())
            v._cliente_ia = None  # recria o cliente com a chave nova
            messagebox.showinfo("Chave da IA", "Chave salva.")
            return True
        return False

    # ---------- análise ----------
    def iniciar(self) -> None:
        if self.rodando:
            return
        pasta = Path(self.pasta.get().strip().strip('"'))
        if not self.pasta.get().strip() or not pasta.is_dir():
            messagebox.showwarning("Validador", "Escolha uma pasta válida com os documentos.")
            return
        texto_colab = self.colaboradores.get().strip().strip('"')
        planilha_colab = Path(texto_colab) if texto_colab else None
        if planilha_colab is not None and not planilha_colab.is_file():
            messagebox.showwarning("Validador", "A planilha de colaboradores não foi encontrada.\n"
                                   "Escolha de novo ou deixe o campo em branco.")
            return
        if not chave_configurada() and not self.pedir_chave():
            return

        self._guardar_prefs()
        self.rodando = True
        self.planilha = None
        self.bt_analisar.configure(state="disabled")
        self.bt_planilha.configure(state="disabled")
        self.barra.configure(value=0, maximum=1)
        for numero in self.contadores.values():
            numero.configure(text="0")
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")
        self.texto_progresso.configure(text="Procurando documentos…")

        threading.Thread(target=self._trabalhar, args=(pasta, planilha_colab), daemon=True).start()

    def _trabalhar(self, pasta: Path, planilha_colab: Path | None = None) -> None:
        """Roda em segundo plano; só conversa com a janela pela fila."""
        try:
            destino, contagem = v.executar_analise(
                pasta, usar_ia=True, planilha_colaboradores=planilha_colab,
                log=lambda texto: self.fila.put(("log", texto)),
                progresso=lambda n, total, r: self.fila.put(("progresso", (n, total, r))))
            self.fila.put(("fim", (destino, contagem)))
        except RuntimeError as erro:
            self.fila.put(("erro", str(erro)))
        except Exception:
            self.fila.put(("erro", traceback.format_exc(limit=3)))

    def ler_fila(self) -> None:
        try:
            while True:
                tipo, dado = self.fila.get_nowait()
                if tipo == "log":
                    self._escrever(dado)
                elif tipo == "progresso":
                    n, total, r = dado
                    self.barra.configure(maximum=total, value=n)
                    self.texto_progresso.configure(text=f"Analisando… {n} de {total}")
                    rotulo = self.contadores.get(r["Status"])
                    if rotulo:
                        rotulo.configure(text=str(int(rotulo.cget("text")) + 1))
                elif tipo == "fim":
                    self._terminar(*dado)
                elif tipo == "erro":
                    self._falhar(dado)
        except queue.Empty:
            pass
        self.janela.after(150, self.ler_fila)

    def _escrever(self, texto: str) -> None:
        self.log.configure(state="normal")
        tag = next((s for s in COR if texto.split("] ", 1)[-1].startswith(s)), None)
        self.log.insert("end", texto + "\n", tag or ())
        self.log.see("end")
        self.log.configure(state="disabled")

    def _terminar(self, destino, contagem) -> None:
        self.rodando = False
        self.bt_analisar.configure(state="normal")
        if destino is None:
            self.texto_progresso.configure(text="Nenhum PDF ou imagem encontrado na pasta.")
            return
        self.planilha = destino
        self.bt_planilha.configure(state="normal")
        self.texto_progresso.configure(text=f"Concluído! Planilha: {destino.name}")
        if messagebox.askyesno(
                "Análise concluída",
                f"Legíveis: {contagem[v.LEGIVEL]}\nIlegíveis: {contagem[v.ILEGIVEL]}\n"
                f"Revisar: {contagem[v.REVISAR]}\n\nAbrir a planilha agora?"):
            self.abrir_planilha()

    def _falhar(self, mensagem: str) -> None:
        self.rodando = False
        self.bt_analisar.configure(state="normal")
        self.texto_progresso.configure(text="A análise parou por um erro.")
        self._escrever(f"ERRO: {mensagem}")
        messagebox.showerror("Validador", mensagem)

    def abrir_planilha(self) -> None:
        if self.planilha and Path(self.planilha).exists():
            if sys.platform.startswith("win"):
                os.startfile(self.planilha)  # abre no Excel
            else:
                messagebox.showinfo("Planilha", str(self.planilha))

    def fechar(self) -> None:
        if self.rodando and not messagebox.askyesno(
                "Validador", "A análise ainda está rodando. Fechar mesmo assim?"):
            return
        self.janela.destroy()


def main() -> None:
    janela = tk.Tk()
    try:
        ttk.Style().theme_use("vista" if sys.platform.startswith("win") else "clam")
    except tk.TclError:
        pass
    App(janela)
    janela.mainloop()


if __name__ == "__main__":
    try:
        main()
    except Exception:
        # Sem console (duplo clique), o erro sumiria: grava num arquivo ao lado do programa
        (PASTA_PROGRAMA / "erro_validador.txt").write_text(traceback.format_exc(), encoding="utf-8")
        raise
