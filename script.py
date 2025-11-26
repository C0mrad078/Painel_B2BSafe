import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import subprocess
import sys

# -------------------- CORES E ESTILO ------------------------
BG_PRINCIPAL = "#111827"   # fundo geral
BG_FRAME = "#1F2933"       # fundo dos blocos
FG_TEXTO = "#E5E7EB"       # texto principal
FG_SECUNDARIO = "#9CA3AF"  # texto secundário
ACCENT_GREEN = "#22C55E"   # botão de ação
ACCENT_BLUE = "#3B82F6"    # botão secundário
ACCENT_ORANGE = "#F97316"  # botão de pasta
BORDER_COR = "#374151"     # bordas / contornos
INPUT_BG = "#020617"       # campos de texto
INPUT_FG = "#F9FAFB"       # texto dos campos


# -------------------- FUNÇÕES ------------------------

def abrir_pasta():
    if caminho_arquivo_saida.get() != "":
        pasta = os.path.dirname(caminho_arquivo_saida.get())
        try:
            if sys.platform == "darwin":      # macOS
                subprocess.Popen(["open", pasta])
            elif os.name == "nt":             # Windows
                subprocess.Popen(f'explorer "{pasta}"')
            else:                             # Linux / outros
                subprocess.Popen(["xdg-open", pasta])
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível abrir a pasta.\n\n{e}")
    else:
        messagebox.showwarning("Aviso", "Nenhum arquivo foi gerado ainda.")


def selecionar_pasta_saida():
    pasta = filedialog.askdirectory(title="Selecione a pasta onde salvar o arquivo")
    if pasta:
        pasta_saida.set(pasta)
    else:
        messagebox.showinfo("Aviso", "Nenhuma pasta selecionada.")


def carregar_colunas():
    try:
        arquivo = caminho_arquivo.get()
        if not arquivo:
            messagebox.showwarning("Aviso", "Selecione um arquivo primeiro.")
            return

        if arquivo.endswith(".xlsx"):
            df = pd.read_excel(arquivo)
        elif arquivo.endswith(".csv"):
            df = pd.read_csv(arquivo, sep=";", encoding="utf-8")
        else:
            messagebox.showerror("Erro", "Selecione um arquivo CSV ou XLSX.")
            return

        colunas = list(df.columns)
        combo_colA["values"] = colunas
        combo_colB["values"] = colunas
        messagebox.showinfo("OK", "Colunas carregadas com sucesso!")
    except Exception as e:
        messagebox.showerror("Erro", f"Não foi possível carregar colunas.\n\n{e}")


def executar_comparacao():
    try:
        progress["value"] = 0
        janela.update_idletasks()

        opcao = combo_opcao.get()
        if opcao == "":
            messagebox.showwarning("Aviso", "Selecione o tipo de comparação.")
            return

        if caminho_arquivo.get() == "":
            messagebox.showwarning("Aviso", "Selecione um arquivo.")
            return

        if pasta_saida.get() == "":
            messagebox.showwarning("Aviso", "Selecione onde salvar o arquivo final.")
            return

        arquivo = caminho_arquivo.get()
        pasta_destino = pasta_saida.get()
        ext = arquivo.split(".")[-1].lower()
        progress["value"] = 10
        janela.update_idletasks()

        # ---------------- Carregar arquivo ----------------
        if ext == "xlsx":
            df = pd.read_excel(arquivo)
        elif ext == "csv":
            df = pd.read_csv(arquivo, sep=";", encoding="utf-8")
        else:
            messagebox.showerror("Erro", "Formato não suportado. Use CSV ou XLSX.")
            return

        progress["value"] = 30
        janela.update_idletasks()

        colA = combo_colA.get()
        colB = combo_colB.get()
        if colA == "" or colB == "":
            messagebox.showerror("Erro", "Selecione as colunas para comparação.")
            return

        # ---------------- Comparação ---------------------
        if opcao == "O que tem na A e não tem na B":
            resultado = df[~df[colA].isin(df[colB])][colA]
            tipo = f"{colA}NAO_ESTA_EM{colB}"
            coluna_base = colA
            outra_coluna = colB
        else:
            resultado = df[~df[colB].isin(df[colA])][colB]
            tipo = f"{colB}NAO_ESTA_EM{colA}"
            coluna_base = colB
            outra_coluna = colA

        progress["value"] = 50
        janela.update_idletasks()

        df[tipo] = ""
        df.loc[~df[coluna_base].isin(df[outra_coluna]), tipo] = df[coluna_base]

        # ---------------- Salvar arquivo -----------------
        nome_arquivo_saida = f"resultado_{tipo}.xlsx"
        arquivo_saida = os.path.join(pasta_destino, nome_arquivo_saida)
        df.to_excel(arquivo_saida, index=False)
        caminho_arquivo_saida.set(arquivo_saida)
        progress["value"] = 70
        janela.update_idletasks()

        # ---------------- Pintar de amarelo ---------------
        wb = load_workbook(arquivo_saida)
        ws = wb.active
        fill_amarelo = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        col_base_idx = df.columns.get_loc(coluna_base) + 1
        col_res_idx = df.columns.get_loc(tipo) + 1

        for row in range(2, ws.max_row + 1):
            cell = ws.cell(row=row, column=col_base_idx)
            cell_res = ws.cell(row=row, column=col_res_idx)
            if cell_res.value not in ("", None):
                cell.fill = fill_amarelo
                cell_res.fill = fill_amarelo

        wb.save(arquivo_saida)
        progress["value"] = 90
        janela.update_idletasks()

        # ---------------- Relatório -----------------------
        relatorio = f"""
PROCESSO COMPLETO

Arquivo analisado: {arquivo}
Arquivo gerado em: {arquivo_saida}

Tipo de comparação: {tipo}

Linhas analisadas: {len(df)}
Itens encontrados: {len(resultado)}

Lista dos itens encontrados:
{resultado.to_list()}
"""
        txt_relatorio.delete("1.0", tk.END)
        txt_relatorio.insert(tk.END, relatorio)
        progress["value"] = 100
        janela.update_idletasks()
        messagebox.showinfo("Concluído", "Comparação finalizada com sucesso!")

    except Exception as e:
        messagebox.showerror("Erro", str(e))


# -------------------- INTERFACE ------------------------
janela = tk.Tk()
janela.title("Suite de Ferramentas - PROCV B2B")
janela.geometry("900x780")
janela.configure(bg=BG_PRINCIPAL)

# Fontes
fonte_label = ("Segoe UI", 10)
fonte_entry = ("Segoe UI", 10)
fonte_titulo = ("Segoe UI", 16, "bold")

# ttk Style
style = ttk.Style()
try:
    style.theme_use("clam")  # tema neutro, respeita cores
except:
    pass

style.configure("TLabel", background=BG_PRINCIPAL, foreground=FG_TEXTO, font=fonte_label)

style.configure("Frame.TLabelframe",
                background=BG_FRAME,
                foreground=FG_TEXTO,
                bordercolor=BORDER_COR)
style.configure("Frame.TLabelframe.Label",
                background=BG_FRAME,
                foreground=FG_TEXTO,
                font=("Segoe UI", 11, "bold"))

style.configure(
    "Custom.Horizontal.TProgressbar",
    troughcolor=BG_FRAME,
    background=ACCENT_GREEN,
    bordercolor=BG_FRAME,
    lightcolor=ACCENT_GREEN,
    darkcolor=ACCENT_GREEN
)

style.configure("TCombobox",
                fieldbackground=INPUT_BG,
                background=INPUT_BG,
                foreground=INPUT_FG,
                arrowcolor=INPUT_FG,
                bordercolor=BORDER_COR)
style.map("TCombobox",
          fieldbackground=[("readonly", INPUT_BG)],
          foreground=[("readonly", INPUT_FG)])

# Estilos para botões ttk
style.configure(
    "Accent.TButton",
    font=("Segoe UI", 11, "bold"),
    foreground="white",
    background=ACCENT_GREEN,
    borderwidth=0,
    focuscolor=BG_PRINCIPAL
)
style.map(
    "Accent.TButton",
    background=[("active", "#16A34A")],
    foreground=[("active", "white")]
)

style.configure(
    "Primary.TButton",
    font=("Segoe UI", 10, "bold"),
    foreground="white",
    background=ACCENT_BLUE,
    borderwidth=0,
    focuscolor=BG_PRINCIPAL
)
style.map(
    "Primary.TButton",
    background=[("active", "#1D4ED8")],
    foreground=[("active", "white")]
)

style.configure(
    "Warn.TButton",
    font=("Segoe UI", 10, "bold"),
    foreground="white",
    background=ACCENT_ORANGE,
    borderwidth=0,
    focuscolor=BG_PRINCIPAL
)
style.map(
    "Warn.TButton",
    background=[("active", "#EA580C")],
    foreground=[("active", "white")]
)


# ---------- Notebook (Abas) ----------
notebook = ttk.Notebook(janela)
notebook.pack(fill="both", expand=True, padx=8, pady=8)

frame_home = tk.Frame(notebook, bg=BG_PRINCIPAL)
frame_procv = tk.Frame(notebook, bg=BG_PRINCIPAL)

notebook.add(frame_home, text="Início")
notebook.add(frame_procv, text="PROCV B2B")


# ---------- ABA INÍCIO ----------
lbl_home_title = tk.Label(
    frame_home,
    text="Bem-vindo à sua suíte de ferramentas ⚙️",
    bg=BG_PRINCIPAL,
    fg=FG_TEXTO,
    font=fonte_titulo
)
lbl_home_title.pack(pady=(40, 10))

lbl_home_sub = tk.Label(
    frame_home,
    text=(
        "Aqui você pode centralizar várias automações e utilitários.\n"
        "Por enquanto, você tem disponível o módulo PROCV B2B para comparar colunas\n"
        "de planilhas e destacar itens que estão em uma lista e não estão em outra."
    ),
    bg=BG_PRINCIPAL,
    fg=FG_SECUNDARIO,
    font=("Segoe UI", 11),
    justify="center"
)
lbl_home_sub.pack(pady=(0, 30))


def ir_para_procv():
    notebook.select(frame_procv)


btn_ir_procv = ttk.Button(
    frame_home,
    text="Abrir módulo PROCV B2B",
    style="Accent.TButton",
    command=ir_para_procv
)
btn_ir_procv.pack(pady=10)

lbl_tip = tk.Label(
    frame_home,
    text="Dica: no futuro você pode adicionar aqui outros módulos (ex: limpeza de dados, relatórios, etc.).",
    bg=BG_PRINCIPAL,
    fg=FG_SECUNDARIO,
    font=("Segoe UI", 9),
    justify="center"
)
lbl_tip.pack(pady=(30, 10))


# ---------- ABA PROCV B2B ----------

# Variáveis
caminho_arquivo = tk.StringVar()
caminho_arquivo_saida = tk.StringVar()
pasta_saida = tk.StringVar()

# Título
lbl_titulo = tk.Label(
    frame_procv,
    text="PROCV B2B - Comparador de Colunas",
    bg=BG_PRINCIPAL,
    fg=FG_TEXTO,
    font=fonte_titulo
)
lbl_titulo.pack(pady=(10, 2))

lbl_sub = tk.Label(
    frame_procv,
    text="Compare colunas de arquivos CSV/XLSX e gere um Excel com os itens exclusivos, já destacados.",
    bg=BG_PRINCIPAL,
    fg=FG_SECUNDARIO,
    font=("Segoe UI", 10),
)
lbl_sub.pack(pady=(0, 10))

# ---------- Frame arquivo --------------
frame_arquivo = ttk.Labelframe(
    frame_procv,
    text="Arquivo de entrada",
    style="Frame.TLabelframe",
    padding=10
)
frame_arquivo.pack(padx=12, pady=6, fill="x")


entry_arquivo = tk.Entry(
    frame_arquivo,
    textvariable=caminho_arquivo,
    font=fonte_entry,
    width=65,
    bg=INPUT_BG,
    fg=INPUT_FG,
    insertbackground=INPUT_FG,
    bd=1,
    relief="solid",
    highlightthickness=0
)
entry_arquivo.pack(side=tk.LEFT, padx=5, pady=3)


def selecionar_arquivo():
    arquivo = filedialog.askopenfilename(
        title="Selecione um arquivo",
        filetypes=[("Excel e CSV", "*.xlsx *.csv"), ("Todos os arquivos", "*.*")]
    )
    if arquivo:
        caminho_arquivo.set(arquivo)


btn_sel_arquivo = ttk.Button(
    frame_arquivo,
    text="Selecionar Arquivo",
    style="Primary.TButton",
    command=selecionar_arquivo
)
btn_sel_arquivo.pack(side=tk.LEFT, padx=5)

btn_carregar_cols = ttk.Button(
    frame_arquivo,
    text="Carregar Colunas",
    style="Primary.TButton",
    command=carregar_colunas
)
btn_carregar_cols.pack(side=tk.LEFT, padx=5)

# ---------- Frame pasta de saída -----------
frame_pasta = ttk.Labelframe(
    frame_procv,
    text="Pasta de saída",
    style="Frame.TLabelframe",
    padding=10
)
frame_pasta.pack(padx=12, pady=6, fill="x")

entry_pasta = tk.Entry(
    frame_pasta,
    textvariable=pasta_saida,
    font=fonte_entry,
    width=65,
    bg=INPUT_BG,
    fg=INPUT_FG,
    insertbackground=INPUT_FG,
    bd=1,
    relief="solid",
    highlightthickness=0
)
entry_pasta.pack(side=tk.LEFT, padx=5, pady=3)

btn_sel_pasta = ttk.Button(
    frame_pasta,
    text="Selecionar Pasta",
    style="Warn.TButton",
    command=selecionar_pasta_saida
)
btn_sel_pasta.pack(side=tk.LEFT, padx=5)

# ---------- Frame colunas --------------
frame_colunas = ttk.Labelframe(
    frame_procv,
    text="Colunas para comparação",
    style="Frame.TLabelframe",
    padding=10
)
frame_colunas.pack(padx=12, pady=6, fill="x")

lbl_colA = tk.Label(frame_colunas, text="Coluna A:", bg=BG_FRAME, fg=FG_TEXTO, font=fonte_label)
lbl_colA.grid(row=0, column=0, padx=5, pady=3, sticky="w")

combo_colA = ttk.Combobox(frame_colunas, width=30, state="readonly")
combo_colA.grid(row=0, column=1, padx=5, pady=3)

lbl_colB = tk.Label(frame_colunas, text="Coluna B:", bg=BG_FRAME, fg=FG_TEXTO, font=fonte_label)
lbl_colB.grid(row=1, column=0, padx=5, pady=3, sticky="w")

combo_colB = ttk.Combobox(frame_colunas, width=30, state="readonly")
combo_colB.grid(row=1, column=1, padx=5, pady=3)

# ---------- Tipo de comparação --------------
lbl_tipo = tk.Label(
    frame_procv,
    text="Tipo de comparação:",
    bg=BG_PRINCIPAL,
    fg=FG_TEXTO,
    font=fonte_label
)
lbl_tipo.pack(pady=(8, 3))

combo_opcao = ttk.Combobox(
    frame_procv,
    width=40,
    values=[
        "O que tem na A e não tem na B",
        "O que tem na B e não tem na A"
    ],
    state="readonly"
)
combo_opcao.pack(pady=(0, 8))

# ---------- Botão Executar ----------------
btn_exec = ttk.Button(
    frame_procv,
    text="Executar Comparação",
    style="Accent.TButton",
    command=executar_comparacao
)
btn_exec.pack(pady=12)

# ---------- Barra de progresso ------------
progress = ttk.Progressbar(
    frame_procv,
    length=720,
    mode="determinate",
    style="Custom.Horizontal.TProgressbar"
)
progress.pack(pady=8)

# ---------- Relatório ---------------------
frame_relatorio = ttk.Labelframe(
    frame_procv,
    text="Relatório",
    style="Frame.TLabelframe",
    padding=10
)
frame_relatorio.pack(padx=12, pady=8, fill="both", expand=True)

txt_relatorio = tk.Text(
    frame_relatorio,
    width=95,
    height=15,
    font=("Consolas", 10),
    bg=INPUT_BG,
    fg=INPUT_FG,
    insertbackground=INPUT_FG,
    bd=1,
    relief="solid",
    highlightthickness=1,
    highlightbackground=BORDER_COR,
    highlightcolor=BORDER_COR
)
txt_relatorio.pack(fill="both", expand=True)

# ---------- Botão abrir pasta -------------
btn_pasta = ttk.Button(
    frame_procv,
    text="Abrir pasta do arquivo gerado",
    style="Primary.TButton",
    command=abrir_pasta
)
btn_pasta.pack(pady=(4, 12))

janela.mainloop()
