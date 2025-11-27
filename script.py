import os
import sys
import re
import math
import subprocess
import json
import time
import shutil

import pandas as pd
from typing import List, Set

from openpyxl import load_workbook, Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import PatternFill, Font

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from sqlalchemy import create_engine, text as sql_text
from sqlalchemy.engine import URL
import sqlalchemy

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

# -------------------- CONSTANTES LIMPEZA --------------------
PHONE_MIN_LEN = 8
CHUNK_SIZE = 19999

# -------------------- CONFIG BANCO -------------------------
DB_CONFIG_FILE = "db_config.json"
db_engine = None
db_connected = False

# =======================================================================
#           FUNÇÕES COMPARTILHADAS / UTILITÁRIOS (LÓGICA)
# =======================================================================

def normalize_col_name(name: str) -> str:
    return re.sub(r"[^0-9a-zA-Z]+", "", str(name)).strip().lower()

def read_table(path: str) -> pd.DataFrame:
    ext = os.path.splitext(path)[1].lower()
    if ext in ['.xls', '.xlsx']:
        return pd.read_excel(path, dtype=str)
    elif ext in ['.csv', '.txt']:
        return pd.read_csv(path, dtype=str, sep=None, engine='python')
    else:
        raise ValueError('Formato não suportado: ' + ext)

def keep_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    Mantém colunas principais, mas tenta reconhecer nomes parecidos.
    Se não reconhecer nada, mantém TODAS as colunas originais para
    evitar zerar a base.
    """
    wanted = [
        'Cnpj',
        'Razao Social',
        'Telefones',
        'E-mail',
        'Codigo da Atividade Principal',
        'Codigos das Atividades Secundarias'
    ]
    col_map = {}
    normals = {normalize_col_name(c): c for c in df.columns}

    # 1) Igualdade exata
    for w in wanted:
        norm_w = normalize_col_name(w)
        for k, orig in normals.items():
            if norm_w == k:
                col_map[orig] = w
                break

    # 2) Heurísticas

    # CNPJ
    if 'Cnpj' not in col_map.values():
        for c in df.columns:
            if 'cnpj' in normalize_col_name(c):
                if c not in col_map:
                    col_map[c] = 'Cnpj'
                    break

    # Razão Social
    if 'Razao Social' not in col_map.values():
        for c in df.columns:
            cn = normalize_col_name(c)
            if 'razao' in cn or 'nomeempresa' in cn or 'nomedaempresa' in cn:
                if c not in col_map:
                    col_map[c] = 'Razao Social'
                    break

    # Telefones
    if 'Telefones' not in col_map.values():
        for c in df.columns:
            cn = normalize_col_name(c)
            if 'tel' in cn or 'fone' in cn or 'telefone' in cn:
                if c not in col_map:
                    col_map[c] = 'Telefones'
                    break

    # E-mail
    if 'E-mail' not in col_map.values():
        for c in df.columns:
            cn = normalize_col_name(c)
            if 'email' in cn or 'eemail' in cn:
                if c not in col_map:
                    col_map[c] = 'E-mail'
                    break

    # Código Atividade Principal
    if 'Codigo da Atividade Principal' not in col_map.values():
        for c in df.columns:
            cn = normalize_col_name(c)
            if ('atividadeprincipal' in cn or 'cnaeprincipal' in cn or
                ('principal' in cn and ('cnae' in cn or 'atividade' in cn))):
                if c not in col_map:
                    col_map[c] = 'Codigo da Atividade Principal'
                    break

    # Códigos Atividades Secundárias
    if 'Codigos das Atividades Secundarias' not in col_map.values():
        for c in df.columns:
            cn = normalize_col_name(c)
            if ('atividade' in cn and 'secundaria' in cn) or 'cnaesecundaria' in cn:
                if c not in col_map:
                    col_map[c] = 'Codigos das Atividades Secundarias'
                    break

    # 3) Monta o DF
    if col_map:
        df2 = df[list(col_map.keys())].copy()
        df2.columns = list(col_map.values())
    else:
        # Não reconheceu nada → mantém tudo
        df2 = df.copy()

    # 4) Garante colunas padrão
    for w in wanted:
        if w not in df2.columns:
            df2[w] = ''

    return df2

def split_telefones_field(field: str):
    if pd.isna(field):
        return [None, None]
    s = str(field)
    parts = re.split(r'[;,/\|\s]+', s)
    phones = [re.sub(r'\D', '', p) for p in parts if re.sub(r'\D', '', p)]
    return (phones + [None, None])[:2]

def is_invalid_phone(num: str) -> bool:
    if not num:
        return True
    digits = re.sub(r'\D', '', str(num))
    if len(digits) < PHONE_MIN_LEN:
        return True
    if len(set(digits)) == 1:
        return True
    return False

def normalize_phone(num: str) -> str:
    return re.sub(r'\D', '', str(num or ''))

def clean_razao_social(s: str) -> str:
    """
    Remove da Razao Social:
    0-9 - = ' " , . ; [ ] : { } ! @ # $ % & ( ) _ +
    Mantém letras e espaços.
    """
    if pd.isna(s):
        return ''
    pattern = r"[0-9\-\=\'\",.;\[\]:\{\}!@#$%&\(\)_\+]"
    return re.sub(pattern, "", str(s))

def save_to_excel(df: pd.DataFrame, path: str):
    wb = Workbook()
    ws = wb.active
    ws.title = 'Empresas'
    for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), 1):
        ws.append(row)
        if r_idx == 1:
            for cell in ws[r_idx]:
                cell.font = Font(bold=True)
    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        ws.column_dimensions[col[0].column_letter].width = max_len + 2
    wb.save(path)

def normalize_cnpj(c):
    digits = re.sub(r'\D', '', str(c or ''))
    if len(digits) > 14:
        digits = digits[-14:]
    return digits.zfill(14) if digits else None

def pick_col(normals: dict, candidates: list):
    for cand in candidates:
        key = normalize_col_name(cand)
        if key in normals:
            return normals[key]
        for norm, orig in normals.items():
            if key in norm:
                return orig
    return None

# =======================================================================
#           FUNÇÕES DA INTERFACE PROCV B2B
# =======================================================================

def abrir_pasta():
    if caminho_arquivo_saida.get() != "":
        pasta = os.path.dirname(caminho_arquivo_saida.get())
        try:
            if sys.platform == "darwin":
                subprocess.Popen(["open", pasta])
            elif os.name == "nt":
                subprocess.Popen(f'explorer "{pasta}"')
            else:
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

        nome_arquivo_saida = f"resultado_{tipo}.xlsx"
        arquivo_saida = os.path.join(pasta_destino, nome_arquivo_saida)
        df.to_excel(arquivo_saida, index=False)
        caminho_arquivo_saida.set(arquivo_saida)
        progress["value"] = 70
        janela.update_idletasks()

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

# =======================================================================
#           FUNÇÕES DA LIMPEZA (Casa de dados)
# =======================================================================

def log_limpeza(msg: str):
    txt_log_limpeza.insert(tk.END, msg + "\n")
    txt_log_limpeza.see(tk.END)
    janela.update_idletasks()

def abrir_pasta_limpeza():
    pasta = out_dir_limpeza.get()
    if not pasta:
        messagebox.showwarning("Aviso", "Nenhuma pasta de saída selecionada.")
        return
    try:
        if sys.platform == "darwin":
            subprocess.Popen(["open", pasta])
        elif os.name == "nt":
            subprocess.Popen(f'explorer "{pasta}"')
        else:
            subprocess.Popen(["xdg-open", pasta])
    except Exception as e:
        messagebox.showerror("Erro", f"Não foi possível abrir a pasta.\n\n{e}")

def executar_limpeza_empresas():
    try:
        txt_log_limpeza.delete("1.0", tk.END)
        progress_limpeza["value"] = 0
        janela.update_idletasks()

        in_path = base_empresas_path.get().strip()
        if not in_path:
            messagebox.showwarning("Aviso", 'Selecione o arquivo "empresas bruto".')
            return

        out_dir = out_dir_limpeza.get().strip() or os.path.dirname(in_path)
        os.makedirs(out_dir, exist_ok=True)

        log_limpeza('=== Automação: limpeza e filtragem de empresas ===\n')

        log_limpeza('1) Lendo arquivo base...')
        df_raw = read_table(in_path)
        log_limpeza(f'✅ Arquivo lido com {len(df_raw)} linhas e {len(df_raw.columns)} colunas.')
        progress_limpeza["value"] = 10
        janela.update_idletasks()

        log_limpeza('2) Mantendo colunas relevantes...')
        df = keep_columns(df_raw)
        log_limpeza(f'➡️ Colunas mantidas: {list(df.columns)}')
        progress_limpeza["value"] = 25
        janela.update_idletasks()

        modo = clean_mode_var.get()
        if modo == "Lemit":
            log_limpeza('3) Modo Lemit: removendo caracteres especiais da "Razao Social"...')
            df['Razao Social'] = df['Razao Social'].astype(str).apply(clean_razao_social)
        else:
            log_limpeza(f'3) Modo {modo}: mantendo caracteres especiais na "Razao Social"...')
            df['Razao Social'] = df['Razao Social'].astype(str)
        progress_limpeza["value"] = 35
        janela.update_idletasks()

        log_limpeza('4) Separando e limpando telefones...')
        df['Telefone1'], df['Telefone2'] = zip(*df['Telefones'].map(split_telefones_field))
        df['Telefone1'] = df['Telefone1'].apply(normalize_phone)
        df['Telefone2'] = df['Telefone2'].apply(normalize_phone)
        progress_limpeza["value"] = 50
        janela.update_idletasks()

        log_limpeza('5) Removendo linhas com telefones inválidos...')
        before = len(df)
        df = df[~((df['Telefone1'].apply(is_invalid_phone)) & (df['Telefone2'].apply(is_invalid_phone)))]
        log_limpeza(f'⚠️ {before - len(df)} linhas removidas por telefones inválidos.')
        progress_limpeza["value"] = 60
        janela.update_idletasks()

        block_path_local = blocklist_path.get().strip()
        if block_path_local:
            log_limpeza('\n6) Aplicando BLOCK LIST C6...')
            block = read_table(block_path_local)
            tel_col = next((c for c in block.columns if 'tel' in normalize_col_name(c)), None)
            if tel_col:
                block_set = set(block[tel_col].astype(str).map(normalize_phone))
                before = len(df)
                df = df[~(df['Telefone1'].isin(block_set) | df['Telefone2'].isin(block_set))]
                log_limpeza(f'⚠️ {before - len(df)} linhas removidas (Blocklist C6).')
            else:
                log_limpeza('⚠️ Nenhuma coluna de telefone identificada na BLOCK LIST C6.')
        progress_limpeza["value"] = 70
        janela.update_idletasks()

        block_b2b_p = blocklist_b2b_path.get().strip()
        if block_b2b_p:
            log_limpeza('\n7) Aplicando BLOCK LIST B2B...')
            block_b2b = read_table(block_b2b_p)
            tel_col = next((c for c in block_b2b.columns if 'tel' in normalize_col_name(c)), None)
            if tel_col:
                block_b2b_set = set(block_b2b[tel_col].astype(str).map(normalize_phone))
                before = len(df)
                df = df[~(df['Telefone1'].isin(block_b2b_set) | df['Telefone2'].isin(block_b2b_set))]
                log_limpeza(f'⚠️ {before - len(df)} linhas removidas (Blocklist B2B).')
            else:
                log_limpeza('⚠️ Nenhuma coluna de telefone identificada na BLOCK LIST B2B.')
        progress_limpeza["value"] = 75
        janela.update_idletasks()

        dnc_p = dnc_path.get().strip()
        if dnc_p:
            log_limpeza('\n8) Aplicando NÃO PERTURBE...')
            dnc = read_table(dnc_p)
            tel_col = next((c for c in dnc.columns if 'tel' in normalize_col_name(c)), None)
            if tel_col:
                dnc_set = set(dnc[tel_col].astype(str).map(normalize_phone))
                before = len(df)
                df = df[~(df['Telefone1'].isin(dnc_set) | df['Telefone2'].isin(dnc_set))]
                log_limpeza(f'⚠️ {before - len(df)} linhas removidas (Não Perturbe).')
            else:
                log_limpeza('⚠️ Nenhuma coluna de telefone identificada no NÃO PERTURBE.')
        progress_limpeza["value"] = 80
        janela.update_idletasks()

        cnais_p = cnais_path.get().strip()
        if cnais_p:
            log_limpeza('\n9) Filtrando CNAEs desejados...')
            cnais_table = read_table(cnais_p)
            code_col = next(
                (c for c in cnais_table.columns
                 if any(x in normalize_col_name(c) for x in ['cnae', 'cnai', 'codigo'])),
                cnais_table.columns[0]
            )
            desired = set(
                cnais_table[code_col].astype(str).str.strip().dropna().unique()
            )

            def extract_codes(s):
                if pd.isna(s):
                    return set()
                return set(re.split(r'[;,/\|\s]+', str(s)))

            def keep_row(row):
                codes = extract_codes(row['Codigo da Atividade Principal']) | \
                        extract_codes(row['Codigos das Atividades Secundarias'])
                return not codes.isdisjoint(desired)

            before = len(df)
            df = df[df.apply(keep_row, axis=1)]
            log_limpeza(f'⚠️ {before - len(df)} linhas removidas (CNAEs não desejados).')
        progress_limpeza["value"] = 85
        janela.update_idletasks()

        final_cols = ['Razao Social', 'Telefone1', 'Telefone2', 'Cnpj', 'E-mail']
        for c in final_cols:
            if c not in df.columns:
                df[c] = ''
        df_final = df[final_cols].copy()

        total = len(df_final)
        if total == 0:
            log_limpeza('\n⚠️ Nenhuma linha restante após aplicação de todos os filtros.')
            progress_limpeza["value"] = 100
            messagebox.showinfo("Concluído", "Processo concluído, mas não restaram linhas após os filtros.")
            return

        parts = math.ceil(total / CHUNK_SIZE)
        log_limpeza(f'\n10) Salvando {parts} arquivo(s) Excel com até {CHUNK_SIZE} linhas cada...')

        for i in range(parts):
            start = i * CHUNK_SIZE
            end = min(start + CHUNK_SIZE, total)
            part = df_final.iloc[start:end]
            path = os.path.join(out_dir, f'empresas_limpa_part{i+1}.xlsx')
            save_to_excel(part, path)
            log_limpeza(f'✅ Parte {i+1} salva: {path} ({len(part)} linhas)')
            progress_limpeza["value"] = 85 + ((i + 1) / parts) * 15
            janela.update_idletasks()

        log_limpeza('\n🎉 Execução concluída com sucesso!')
        log_limpeza(f'Total de linhas finais: {len(df_final)}')
        progress_limpeza["value"] = 100
        janela.update_idletasks()
        messagebox.showinfo("Concluído", "Limpeza de empresas finalizada com sucesso!")

    except Exception as e:
        log_limpeza(f'\n❌ Erro fatal: {e}')
        messagebox.showerror("Erro", f"Ocorreu um erro durante a execução.\n\n{e}")

# =======================================================================
#           FUNÇÕES ROBÔ C6
# =======================================================================

def log_robo(msg: str):
    txt_log_robo.insert(tk.END, msg + "\n")
    txt_log_robo.see(tk.END)
    janela.update_idletasks()

def selecionar_arquivos_robo():
    paths = filedialog.askopenfilenames(
        title="Selecione as planilhas (até 20 mil linhas cada)",
        filetypes=[("Planilhas", "*.xlsx *.xls *.csv *.txt"), ("Todos os arquivos", "*.*")]
    )
    if paths:
        robo_arquivos.clear()
        robo_arquivos.extend(paths)
        lbl_arquivos_robo.config(
            text=f"{len(robo_arquivos)} arquivo(s) selecionado(s)."
        )

def selecionar_bat_robo():
    path = filedialog.askopenfilename(
        title="Selecione o arquivo .BAT",
        filetypes=[("Arquivos BAT", "*.bat"), ("Todos os arquivos", "*.*")]
    )
    if path:
        robo_bat_path.set(path)
        lbl_bat_robo.config(text=f".BAT selecionado: {path}")
        # se não tiver pasta de resultado definida, tentar detectar "resultado" na mesma pasta
        bat_dir = os.path.dirname(path)
        resultado_sugerido = os.path.join(bat_dir, "resultado")
        if os.path.isdir(resultado_sugerido) and not robo_resultado_dir.get():
            robo_resultado_dir.set(resultado_sugerido)
            lbl_pasta_resultado.config(text=f"Pasta de resultados: {resultado_sugerido}")

def selecionar_resultado_robo():
    path = filedialog.askdirectory(
        title="Selecione a pasta de resultados do .BAT"
    )
    if path:
        robo_resultado_dir.set(path)
        lbl_pasta_resultado.config(text=f"Pasta de resultados: {path}")

def executar_robo_c6():
    try:
        txt_log_robo.delete("1.0", tk.END)
        progress_robo["value"] = 0
        janela.update_idletasks()

        if not robo_arquivos:
            messagebox.showwarning("Aviso", "Selecione as planilhas de entrada primeiro.")
            return

        bat_path = robo_bat_path.get().strip()
        if not bat_path or not os.path.isfile(bat_path):
            messagebox.showwarning("Aviso", "Selecione um arquivo .BAT válido.")
            return

        resultado_dir = robo_resultado_dir.get().strip()
        if not resultado_dir or not os.path.isdir(resultado_dir):
            messagebox.showwarning(
                "Aviso",
                "Selecione uma pasta de resultados válida do .BAT (pode ser a pasta 'resultado')."
            )
            return

        modo = robo_modo_var.get()
        if modo not in ["Lemit", "Simples"]:
            messagebox.showwarning(
                "Aviso",
                "Selecione se o arquivo é para Lemit ou Simples."
            )
            return

        bat_dir = os.path.dirname(bat_path)
        log_robo("=== Robô C6 iniciado ===")
        log_robo(f"Arquivos selecionados: {len(robo_arquivos)}")
        log_robo(f"Caminho do .BAT: {bat_path}")
        log_robo(f"Pasta de resultados do .BAT: {resultado_dir}")
        log_robo(f"Modo de tratamento final: {modo}")
        log_robo("")

        total_arquivos = len(robo_arquivos)
        # limpar pasta de resultados antes de começar, para evitar misturar execuções antigas
        log_robo("Limpando pasta de resultados antes de iniciar...")
        for f in os.listdir(resultado_dir):
            full = os.path.join(resultado_dir, f)
            if os.path.isfile(full) and any(full.lower().endswith(ext) for ext in [".xlsx", ".xls", ".csv", ".txt"]):
                os.remove(full)
        log_robo("Pasta de resultados limpa.\n")

        for idx, arquivo in enumerate(robo_arquivos, start=1):
            progress_robo["value"] = (idx - 1) / total_arquivos * 40
            janela.update_idletasks()

            log_robo(f"[{idx}/{total_arquivos}] Preparando arquivo: {arquivo}")
            try:
                # copiar arquivo para a pasta do BAT
                base_name = os.path.basename(arquivo)
                dest_path = os.path.join(bat_dir, base_name)
                shutil.copy2(arquivo, dest_path)
                log_robo(f"→ Copiado para pasta do .BAT: {dest_path}")
            except Exception as e:
                log_robo(f"❌ Erro ao copiar arquivo para pasta do .BAT: {e}")
                continue

            # executa o BAT
            try:
                log_robo("→ Executando .BAT...")
                # shell=True para Windows interpretar .bat corretamente
                proc = subprocess.Popen(
                    bat_path,
                    cwd=bat_dir,
                    stdin=subprocess.PIPE,
                    shell=True
                )
                # envia ENTER ao final
                proc.communicate(input=b"\n")
                log_robo("→ Execução do .BAT concluída.")
            except Exception as e:
                log_robo(f"❌ Erro ao executar .BAT: {e}")
                continue

            # aguarda 6 minutos antes do próximo
            if idx < total_arquivos:
                log_robo("⏱ Aguardando 6 minutos antes do próximo arquivo...")
                janela.update_idletasks()
                time.sleep(6 * 60)  # 6 minutos
                log_robo("✔ Intervalo concluído.\n")
            else:
                log_robo("Último arquivo processado.\n")

            progress_robo["value"] = idx / total_arquivos * 60
            janela.update_idletasks()

        # juntar todos os resultados
        log_robo("Lendo arquivos de resultados gerados pelo .BAT...")
        result_files = []
        for f in os.listdir(resultado_dir):
            full = os.path.join(resultado_dir, f)
            if os.path.isfile(full) and full.lower().endswith((".xlsx", ".xls", ".csv", ".txt")):
                result_files.append(full)

        if not result_files:
            log_robo("⚠️ Nenhum arquivo de resultado encontrado na pasta informada.")
            messagebox.showwarning(
                "Aviso",
                "Nenhum arquivo de resultado foi encontrado na pasta de resultados."
            )
            progress_robo["value"] = 100
            return

        log_robo(f"Encontrados {len(result_files)} arquivo(s) de resultado.")
        dfs = []
        for fpath in result_files:
            try:
                log_robo(f"Lendo resultado: {fpath}")
                dfs.append(read_table(fpath))
            except Exception as e:
                log_robo(f"❌ Erro ao ler resultado {fpath}: {e}")

        if not dfs:
            log_robo("⚠️ Não foi possível ler nenhum arquivo de resultado.")
            progress_robo["value"] = 100
            return

        df_total = pd.concat(dfs, ignore_index=True)
        log_robo(f"Total de linhas combinadas (antes da filtragem): {len(df_total)}")

        # ===================================================================
        #      FILTRO: remover "Nao disponivel" e manter só "Novo cliente"
        # ===================================================================
        log_robo("Aplicando filtro: remover linhas com 'Nao disponivel' e manter apenas 'Novo cliente'...")
        df_str = df_total.astype(str)

        mask_nao_disponivel = df_str.apply(
            lambda col: col.str.contains("Nao disponivel", case=False, na=False)
        ).any(axis=1)

        mask_novo_cliente = df_str.apply(
            lambda col: col.str.contains("Novo cliente", case=False, na=False)
        ).any(axis=1)

        antes = len(df_total)
        df_filtrado = df_total[~mask_nao_disponivel & mask_novo_cliente].copy()
        removidas = antes - len(df_filtrado)
        log_robo(f"Linhas removidas pelo filtro: {removidas}")
        log_robo(f"Linhas finais após filtro: {len(df_filtrado)}")

        if df_filtrado.empty:
            log_robo("⚠️ Nenhuma linha restante após aplicar o filtro de 'Novo cliente' e 'Nao disponivel'.")
            messagebox.showinfo(
                "Concluído",
                "Robô C6 finalizado, mas nenhuma linha restou após o filtro de 'Novo cliente'."
            )
            progress_robo["value"] = 100
            return

        progress_robo["value"] = 80
        janela.update_idletasks()

        # ===================================================================
        #      TRATAMENTO FINAL POR MODO (DEPOIS DO FILTRO)
        # ===================================================================
        if modo == "Lemit":
            log_robo("Modo Lemit: gerando 1 planilha única com o resultado final...")
            out_path = os.path.join(resultado_dir, "robo_c6_final_LEMIT.xlsx")
            save_to_excel(df_filtrado, out_path)
            log_robo(f"✅ Arquivo final gerado: {out_path}")

        else:  # Simples
            log_robo("Modo Simples: separando resultado em planilhas de 5.000 linhas...")
            total = len(df_filtrado)
            chunk_size = 5000
            parts = math.ceil(total / chunk_size)
            log_robo(f"Total de linhas: {total} → {parts} arquivo(s) de até {chunk_size} linhas.")

            for i in range(parts):
                start = i * chunk_size
                end = min(start + chunk_size, total)
                part = df_filtrado.iloc[start:end]
                out_path = os.path.join(resultado_dir, f"robo_c6_SIMPLES_part{i+1}.xlsx")
                save_to_excel(part, out_path)
                log_robo(f"✅ Parte {i+1} salva: {out_path} ({len(part)} linhas)")

        progress_robo["value"] = 100
        janela.update_idletasks()
        log_robo("\n🎉 Robô C6 concluído com sucesso!")
        messagebox.showinfo("Concluído", "Robô C6 finalizado com sucesso!")

    except Exception as e:
        log_robo(f"\n❌ Erro fatal no Robô C6: {e}")
        messagebox.showerror("Erro", f"Ocorreu um erro durante o Robô C6.\n\n{e}")

# =======================================================================
#           INTERFACE GRÁFICA (TKINTER)
# =======================================================================

janela = tk.Tk()
janela.title("Suite de Ferramentas - PROCV B2B")
janela.geometry("1350x780")
janela.configure(bg=BG_PRINCIPAL)

fonte_label = ("Segoe UI", 10)
fonte_entry = ("Segoe UI", 10)
fonte_titulo = ("Segoe UI", 16, "bold")

style = ttk.Style()
try:
    style.theme_use("clam")
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

notebook = ttk.Notebook(janela)
notebook.pack(fill="both", expand=True, padx=8, pady=8)

frame_home = tk.Frame(notebook, bg=BG_PRINCIPAL)
frame_procv = tk.Frame(notebook, bg=BG_PRINCIPAL)
frame_limpeza = tk.Frame(notebook, bg=BG_PRINCIPAL)
frame_robo = tk.Frame(notebook, bg=BG_PRINCIPAL)
frame_bd = tk.Frame(notebook, bg=BG_PRINCIPAL)
frame_conexao = tk.Frame(notebook, bg=BG_PRINCIPAL)

notebook.add(frame_home, text="Início")
notebook.add(frame_procv, text="PROCV B2B")
notebook.add(frame_limpeza, text="Limpeza Casa de dados")
notebook.add(frame_robo, text="Robô C6")
notebook.add(frame_bd, text="Banco de Dados")
notebook.add(frame_conexao, text="Conexão BD")

# =======================================================================
#                           ABA INÍCIO
# =======================================================================

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
        "Módulos disponíveis:\n"
        "• PROCV B2B (comparação de colunas)\n"
        "• Limpeza Casa de dados (tratamento de base, blocklist, CNAEs, etc.)\n"
        "• Robô C6 (automação de .BAT + consolidação de resultados)\n"
        "• Banco de Dados (importar arquivos Excel/CSV para tabelas + visualizar)\n"
        "• Conexão BD (configuração e conexão com banco MySQL/PostgreSQL)"
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
btn_ir_procv.pack(pady=8)

def ir_para_limpeza():
    notebook.select(frame_limpeza)

btn_ir_limpeza = ttk.Button(
    frame_home,
    text="Abrir módulo Limpeza Casa de dados",
    style="Primary.TButton",
    command=ir_para_limpeza
)
btn_ir_limpeza.pack(pady=8)

def ir_para_robo():
    notebook.select(frame_robo)

btn_ir_robo = ttk.Button(
    frame_home,
    text="Abrir módulo Robô C6",
    style="Warn.TButton",
    command=ir_para_robo
)
btn_ir_robo.pack(pady=8)

def ir_para_bd():
    notebook.select(frame_bd)

btn_ir_bd = ttk.Button(
    frame_home,
    text="Abrir módulo Banco de Dados",
    style="Primary.TButton",
    command=ir_para_bd
)
btn_ir_bd.pack(pady=8)

def ir_para_conexao():
    notebook.select(frame_conexao)

btn_ir_conexao = ttk.Button(
    frame_home,
    text="Abrir módulo Conexão BD",
    style="Primary.TButton",
    command=ir_para_conexao
)
btn_ir_conexao.pack(pady=8)

lbl_tip = tk.Label(
    frame_home,
    text="Dica: no futuro você pode adicionar aqui outros módulos (ex: relatórios, APIs, etc.).",
    bg=BG_PRINCIPAL,
    fg=FG_SECUNDARIO,
    font=("Segoe UI", 9),
    justify="center"
)
lbl_tip.pack(pady=(30, 10))

# =======================================================================
#                       ABA PROCV B2B
# =======================================================================

caminho_arquivo = tk.StringVar()
caminho_arquivo_saida = tk.StringVar()
pasta_saida = tk.StringVar()

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

btn_exec = ttk.Button(
    frame_procv,
    text="Executar Comparação",
    style="Accent.TButton",
    command=executar_comparacao
)
btn_exec.pack(pady=12)

progress = ttk.Progressbar(
    frame_procv,
    length=720,
    mode="determinate",
    style="Custom.Horizontal.TProgressbar"
)
progress.pack(pady=8)

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

btn_pasta = ttk.Button(
    frame_procv,
    text="Abrir pasta do arquivo gerado",
    style="Primary.TButton",
    command=abrir_pasta
)
btn_pasta.pack(pady=(4, 12))

# =======================================================================
#                       ABA LIMPEZA CASA DE DADOS
# =======================================================================

base_empresas_path = tk.StringVar()
blocklist_path = tk.StringVar()
blocklist_b2b_path = tk.StringVar()
dnc_path = tk.StringVar()
cnais_path = tk.StringVar()
out_dir_limpeza = tk.StringVar()
clean_mode_var = tk.StringVar(value="Simples")

lbl_limpeza_title = tk.Label(
    frame_limpeza,
    text="Limpeza Casa de dados",
    bg=BG_PRINCIPAL,
    fg=FG_TEXTO,
    font=fonte_titulo
)
lbl_limpeza_title.pack(pady=(10, 2))

frame_limpeza_main = tk.Frame(frame_limpeza, bg=BG_PRINCIPAL)
frame_limpeza_main.pack(fill="both", expand=True, padx=8, pady=8)

frame_limpeza_left = tk.Frame(frame_limpeza_main, bg=BG_PRINCIPAL)
frame_limpeza_left.pack(side=tk.LEFT, fill="both", expand=True, padx=(0, 6))

frame_limpeza_right = tk.Frame(frame_limpeza_main, bg=BG_PRINCIPAL)
frame_limpeza_right.pack(side=tk.LEFT, fill="both", expand=True, padx=(6, 0))

btn_iniciar_automacao = ttk.Button(
    frame_limpeza_left,
    text="Iniciar automação de limpeza",
    style="Accent.TButton",
    command=executar_limpeza_empresas
)
btn_iniciar_automacao.pack(pady=(0, 10))

frame_modo = ttk.Labelframe(
    frame_limpeza_left,
    text="Modo de limpeza da Razão Social",
    style="Frame.TLabelframe",
    padding=10
)
frame_modo.pack(padx=0, pady=6, fill="x")

lbl_modo = tk.Label(
    frame_modo,
    text="Selecione o modo:",
    bg=BG_FRAME,
    fg=FG_TEXTO,
    font=fonte_label
)
lbl_modo.pack(side=tk.LEFT, padx=(5, 5))

combo_modo = ttk.Combobox(
    frame_modo,
    textvariable=clean_mode_var,
    state="readonly",
    width=20,
    values=["Simples", "Lemit", "Callix", "Tallos"]
)
combo_modo.pack(side=tk.LEFT, padx=5)
combo_modo.current(0)

frame_base = ttk.Labelframe(
    frame_limpeza_left,
    text='1) Arquivo "empresas bruto"',
    style="Frame.TLabelframe",
    padding=10
)
frame_base.pack(padx=0, pady=6, fill="x")

entry_base = tk.Entry(
    frame_base,
    textvariable=base_empresas_path,
    font=fonte_entry,
    width=50,
    bg=INPUT_BG,
    fg=INPUT_FG,
    insertbackground=INPUT_FG,
    bd=1,
    relief="solid",
    highlightthickness=0
)
entry_base.pack(side=tk.LEFT, padx=5, pady=3)

def selecionar_base_empresas():
    path = filedialog.askopenfilename(
        title='Selecione o arquivo "empresas bruto"',
        filetypes=[("Excel/CSV/TXT", "*.xlsx *.xls *.csv *.txt"), ("Todos os arquivos", "*.*")]
    )
    if path:
        base_empresas_path.set(path)

btn_sel_base = ttk.Button(
    frame_base,
    text="Selecionar",
    style="Primary.TButton",
    command=selecionar_base_empresas
)
btn_sel_base.pack(side=tk.LEFT, padx=5)

frame_block = ttk.Labelframe(
    frame_limpeza_left,
    text='5) Arquivo BLOCK LIST C6 (opcional)',
    style="Frame.TLabelframe",
    padding=10
)
frame_block.pack(padx=0, pady=6, fill="x")

entry_block = tk.Entry(
    frame_block,
    textvariable=blocklist_path,
    font=fonte_entry,
    width=50,
    bg=INPUT_BG,
    fg=INPUT_FG,
    insertbackground=INPUT_FG,
    bd=1,
    relief="solid",
    highlightthickness=0
)
entry_block.pack(side=tk.LEFT, padx=5, pady=3)

def selecionar_blocklist():
    path = filedialog.askopenfilename(
        title='Selecione o arquivo de BLOCK LIST C6 (opcional)',
        filetypes=[("Excel/CSV/TXT", "*.xlsx *.xls *.csv *.txt"), ("Todos os arquivos", "*.*")]
    )
    if path:
        blocklist_path.set(path)

btn_sel_block = ttk.Button(
    frame_block,
    text="Selecionar",
    style="Primary.TButton",
    command=selecionar_blocklist
)
btn_sel_block.pack(side=tk.LEFT, padx=5)

frame_block_b2b = ttk.Labelframe(
    frame_limpeza_left,
    text='5b) Arquivo BLOCK LIST B2B (opcional)',
    style="Frame.TLabelframe",
    padding=10
)
frame_block_b2b.pack(padx=0, pady=6, fill="x")

entry_block_b2b = tk.Entry(
    frame_block_b2b,
    textvariable=blocklist_b2b_path,
    font=fonte_entry,
    width=50,
    bg=INPUT_BG,
    fg=INPUT_FG,
    insertbackground=INPUT_FG,
    bd=1,
    relief="solid",
    highlightthickness=0
)
entry_block_b2b.pack(side=tk.LEFT, padx=5, pady=3)

def selecionar_blocklist_b2b():
    path = filedialog.askopenfilename(
        title='Selecione o arquivo de BLOCK LIST B2B (opcional)',
        filetypes=[("Excel/CSV/TXT", "*.xlsx *.xls *.csv *.txt"), ("Todos os arquivos", "*.*")]
    )
    if path:
        blocklist_b2b_path.set(path)

btn_sel_block_b2b = ttk.Button(
    frame_block_b2b,
    text="Selecionar",
    style="Primary.TButton",
    command=selecionar_blocklist_b2b
)
btn_sel_block_b2b.pack(side=tk.LEFT, padx=5)

frame_dnc = ttk.Labelframe(
    frame_limpeza_left,
    text='6) Arquivo NÃO PERTURBE (opcional)',
    style="Frame.TLabelframe",
    padding=10
)
frame_dnc.pack(padx=0, pady=6, fill="x")

entry_dnc = tk.Entry(
    frame_dnc,
    textvariable=dnc_path,
    font=fonte_entry,
    width=50,
    bg=INPUT_BG,
    fg=INPUT_FG,
    insertbackground=INPUT_FG,
    bd=1,
    relief="solid",
    highlightthickness=0
)
entry_dnc.pack(side=tk.LEFT, padx=5, pady=3)

def selecionar_dnc():
    path = filedialog.askopenfilename(
        title='Selecione o arquivo NÃO PERTURBE (opcional)',
        filetypes=[("Excel/CSV/TXT", "*.xlsx *.xls *.csv *.txt"), ("Todos os arquivos", "*.*")]
    )
    if path:
        dnc_path.set(path)

btn_sel_dnc = ttk.Button(
    frame_dnc,
    text="Selecionar",
    style="Primary.TButton",
    command=selecionar_dnc
)
btn_sel_dnc.pack(side=tk.LEFT, padx=5)

frame_cnais = ttk.Labelframe(
    frame_limpeza_left,
    text='7) Arquivo de CNAEs desejados (opcional)',
    style="Frame.TLabelframe",
    padding=10
)
frame_cnais.pack(padx=0, pady=6, fill="x")

entry_cnais = tk.Entry(
    frame_cnais,
    textvariable=cnais_path,
    font=fonte_entry,
    width=50,
    bg=INPUT_BG,
    fg=INPUT_FG,
    insertbackground=INPUT_FG,
    bd=1,
    relief="solid",
    highlightthickness=0
)
entry_cnais.pack(side=tk.LEFT, padx=5, pady=3)

def selecionar_cnais():
    path = filedialog.askopenfilename(
        title='Selecione o arquivo com CNAEs desejados (opcional)',
        filetypes=[("Excel/CSV/TXT", "*.xlsx *.xls *.csv *.txt"), ("Todos os arquivos", "*.*")]
    )
    if path:
        cnais_path.set(path)

btn_sel_cnais = ttk.Button(
    frame_cnais,
    text="Selecionar",
    style="Primary.TButton",
    command=selecionar_cnais
)
btn_sel_cnais.pack(side=tk.LEFT, padx=5)

frame_out = ttk.Labelframe(
    frame_limpeza_left,
    text='10) Diretório para salvar os arquivos finais',
    style="Frame.TLabelframe",
    padding=10
)
frame_out.pack(padx=0, pady=6, fill="x")

entry_out = tk.Entry(
    frame_out,
    textvariable=out_dir_limpeza,
    font=fonte_entry,
    width=50,
    bg=INPUT_BG,
    fg=INPUT_FG,
    insertbackground=INPUT_FG,
    bd=1,
    relief="solid",
    highlightthickness=0
)
entry_out.pack(side=tk.LEFT, padx=5, pady=3)

def selecionar_out_dir():
    path = filedialog.askdirectory(
        title='Selecione a pasta para salvar os arquivos finais'
    )
    if path:
        out_dir_limpeza.set(path)

btn_sel_out = ttk.Button(
    frame_out,
    text="Selecionar Pasta",
    style="Warn.TButton",
    command=selecionar_out_dir
)
btn_sel_out.pack(side=tk.LEFT, padx=5)

frame_btns_limpeza = tk.Frame(frame_limpeza_left, bg=BG_PRINCIPAL)
frame_btns_limpeza.pack(pady=(10, 0))

btn_abrir_pasta_limpeza = ttk.Button(
    frame_btns_limpeza,
    text="Abrir pasta de saída",
    style="Primary.TButton",
    command=abrir_pasta_limpeza
)
btn_abrir_pasta_limpeza.pack(side=tk.LEFT, padx=6)

progress_limpeza = ttk.Progressbar(
    frame_limpeza_right,
    length=400,
    mode="determinate",
    style="Custom.Horizontal.TProgressbar"
)
progress_limpeza.pack(pady=(0, 8), padx=4, fill="x")

frame_log_limpeza = ttk.Labelframe(
    frame_limpeza_right,
    text="Log / Relatório",
    style="Frame.TLabelframe",
    padding=10
)
frame_log_limpeza.pack(padx=4, pady=4, fill="both", expand=True)

txt_log_limpeza = tk.Text(
    frame_log_limpeza,
    width=60,
    height=25,
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
txt_log_limpeza.pack(fill="both", expand=True)

# =======================================================================
#                       ABA ROBÔ C6
# =======================================================================

robo_arquivos: List[str] = []
robo_bat_path = tk.StringVar()
robo_resultado_dir = tk.StringVar()
robo_modo_var = tk.StringVar(value="Simples")

lbl_robo_title = tk.Label(
    frame_robo,
    text="Robô C6",
    bg=BG_PRINCIPAL,
    fg=FG_TEXTO,
    font=fonte_titulo
)
lbl_robo_title.pack(pady=(10, 2))

lbl_robo_sub = tk.Label(
    frame_robo,
    text=(
        "Automatiza a execução de um .BAT do C6 para múltiplas planilhas.\n"
        "• Lê várias planilhas (até 20 mil linhas cada)\n"
        "• Envia uma por vez para a pasta do .BAT\n"
        "• Executa o .BAT, envia ENTER ao final e espera 6 minutos\n"
        "• Consolida resultados, filtra somente 'Novo cliente' (removendo 'Nao disponivel')\n"
        "• Gera saída em modo Lemit (1 arquivo) ou Simples (arquivos de 5.000 linhas)"
    ),
    bg=BG_PRINCIPAL,
    fg=FG_SECUNDARIO,
    font=("Segoe UI", 10),
    justify="center"
)
lbl_robo_sub.pack(pady=(0, 10))

frame_robo_main = tk.Frame(frame_robo, bg=BG_PRINCIPAL)
frame_robo_main.pack(fill="both", expand=True, padx=8, pady=8)

frame_robo_left = tk.Frame(frame_robo_main, bg=BG_PRINCIPAL)
frame_robo_left.pack(side=tk.LEFT, fill="both", expand=True, padx=(0, 6))

frame_robo_right = tk.Frame(frame_robo_main, bg=BG_PRINCIPAL)
frame_robo_right.pack(side=tk.LEFT, fill="both", expand=True, padx=(6, 0))

# 1) Arquivos de entrada
frame_robo_files = ttk.Labelframe(
    frame_robo_left,
    text="1) Planilhas de entrada (até 20 mil linhas cada)",
    style="Frame.TLabelframe",
    padding=10
)
frame_robo_files.pack(fill="x", pady=6)

lbl_arquivos_robo = tk.Label(
    frame_robo_files,
    text="Nenhum arquivo selecionado.",
    bg=BG_FRAME,
    fg=FG_TEXTO,
    font=fonte_label,
    justify="left"
)
lbl_arquivos_robo.pack(anchor="w", pady=(0, 4))

btn_sel_arquivos_robo = ttk.Button(
    frame_robo_files,
    text="Selecionar planilhas",
    style="Primary.TButton",
    command=selecionar_arquivos_robo
)
btn_sel_arquivos_robo.pack(anchor="w", pady=4)

# 2) Caminho do BAT
frame_robo_bat = ttk.Labelframe(
    frame_robo_left,
    text="2) Arquivo .BAT do C6",
    style="Frame.TLabelframe",
    padding=10
)
frame_robo_bat.pack(fill="x", pady=6)

lbl_bat_robo = tk.Label(
    frame_robo_bat,
    text="Nenhum .BAT selecionado.",
    bg=BG_FRAME,
    fg=FG_TEXTO,
    font=fonte_label,
    justify="left"
)
lbl_bat_robo.pack(anchor="w", pady=(0, 4))

btn_sel_bat_robo = ttk.Button(
    frame_robo_bat,
    text="Selecionar .BAT",
    style="Warn.TButton",
    command=selecionar_bat_robo
)
btn_sel_bat_robo.pack(anchor="w", pady=4)

# 3) Pasta de resultados do BAT
frame_robo_resultado = ttk.Labelframe(
    frame_robo_left,
    text="3) Pasta de resultados do .BAT",
    style="Frame.TLabelframe",
    padding=10
)
frame_robo_resultado.pack(fill="x", pady=6)

lbl_pasta_resultado = tk.Label(
    frame_robo_resultado,
    text=(
        "Nenhuma pasta de resultados selecionada.\n"
        "(se não selecionar, será tentado usar 'pasta_do_bat/resultado' se existir)"
    ),
    bg=BG_FRAME,
    fg=FG_TEXTO,
    font=fonte_label,
    wraplength=350,
    justify="left"
)
lbl_pasta_resultado.pack(pady=(0, 5), anchor="w")

btn_sel_resultado_robo = ttk.Button(
    frame_robo_resultado,
    text="Selecionar pasta resultado",
    style="Primary.TButton",
    command=selecionar_resultado_robo
)
btn_sel_resultado_robo.pack(pady=5, anchor="w")

# 4) Modo de tratamento final
frame_robo_modo = ttk.Labelframe(
    frame_robo_left,
    text="4) Modo de tratamento final (Lemit / Simples)",
    style="Frame.TLabelframe",
    padding=10
)
frame_robo_modo.pack(fill="x", pady=6)

lbl_robo_modo = tk.Label(
    frame_robo_modo,
    text="Arquivo é para:",
    bg=BG_FRAME,
    fg=FG_TEXTO,
    font=fonte_label
)
lbl_robo_modo.pack(side=tk.LEFT, padx=(5, 5))

combo_robo_modo = ttk.Combobox(
    frame_robo_modo,
    textvariable=robo_modo_var,
    state="readonly",
    width=20,
    values=["Lemit", "Simples"]
)
combo_robo_modo.pack(side=tk.LEFT, padx=5)
combo_robo_modo.current(1)  # Simples

# Botão executar
btn_executar_robo = ttk.Button(
    frame_robo_left,
    text="Executar Robô C6",
    style="Accent.TButton",
    command=executar_robo_c6
)
btn_executar_robo.pack(pady=(10, 5), anchor="w")

# Barra de progresso Robô
progress_robo = ttk.Progressbar(
    frame_robo_right,
    length=400,
    mode="determinate",
    style="Custom.Horizontal.TProgressbar"
)
progress_robo.pack(pady=(0, 8), padx=4, fill="x")

frame_log_robo = ttk.Labelframe(
    frame_robo_right,
    text="Log do Robô C6",
    style="Frame.TLabelframe",
    padding=10
)
frame_log_robo.pack(padx=4, pady=4, fill="both", expand=True)

txt_log_robo = tk.Text(
    frame_log_robo,
    width=60,
    height=25,
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
txt_log_robo.pack(fill="both", expand=True)

# =======================================================================
#                      ABA BANCO DE DADOS (IMPORT + VIEW)
# =======================================================================

import_tabela_var = tk.StringVar(value="empresas")
import_arquivo_path = tk.StringVar()
view_tabela_var = tk.StringVar(value="empresas")

def log_bd(msg: str):
    txt_log_bd.insert(tk.END, msg + "\n")
    txt_log_bd.see(tk.END)
    janela.update_idletasks()

frame_bd_main = tk.Frame(frame_bd, bg=BG_PRINCIPAL)
frame_bd_main.pack(fill="both", expand=True, padx=8, pady=8)

frame_bd_left = tk.Frame(frame_bd_main, bg=BG_PRINCIPAL)
frame_bd_left.pack(side=tk.LEFT, fill="y", expand=False, padx=(0, 6))

frame_bd_right = tk.Frame(frame_bd_main, bg=BG_PRINCIPAL)
frame_bd_right.pack(side=tk.LEFT, fill="both", expand=True, padx=(6, 0))

lbl_bd_title = tk.Label(
    frame_bd_left,
    text="Importar dados para o banco",
    bg=BG_PRINCIPAL,
    fg=FG_TEXTO,
    font=fonte_titulo
)
lbl_bd_title.pack(pady=(0, 10))

frame_import_cfg = ttk.Labelframe(
    frame_bd_left,
    text="Configuração de importação",
    style="Frame.TLabelframe",
    padding=10
)
frame_import_cfg.pack(fill="x", pady=6)

lbl_tab_dest = tk.Label(
    frame_import_cfg,
    text="Tabela destino:",
    bg=BG_FRAME,
    fg=FG_TEXTO,
    font=fonte_label
)
lbl_tab_dest.grid(row=0, column=0, padx=5, pady=5, sticky="w")

combo_tab_dest = ttk.Combobox(
    frame_import_cfg,
    textvariable=import_tabela_var,
    state="readonly",
    width=25,
    values=[
        "empresas",
        "block_list_c6",
        "block_list_b2b",
        "nao_perturbe",
        "cnais_aceitos",
        "lemit_relatorio"
    ]
)
combo_tab_dest.grid(row=0, column=1, padx=5, pady=5)

frame_import_file = ttk.Labelframe(
    frame_bd_left,
    text="Arquivo bruto (Excel/CSV/TXT)",
    style="Frame.TLabelframe",
    padding=10
)
frame_import_file.pack(fill="x", pady=6)

entry_import_file = tk.Entry(
    frame_import_file,
    textvariable=import_arquivo_path,
    font=fonte_entry,
    width=40,
    bg=INPUT_BG,
    fg=INPUT_FG,
    insertbackground=INPUT_FG,
    bd=1,
    relief="solid",
    highlightthickness=0
)
entry_import_file.pack(side=tk.LEFT, padx=5, pady=3)

def selecionar_arquivo_import():
    path = filedialog.askopenfilename(
        title='Selecione o arquivo bruto',
        filetypes=[("Excel/CSV/TXT", "*.xlsx *.xls *.csv *.txt"), ("Todos os arquivos", "*.*")]
    )
    if path:
        import_arquivo_path.set(path)

btn_sel_import_file = ttk.Button(
    frame_import_file,
    text="Selecionar",
    style="Primary.TButton",
    command=selecionar_arquivo_import
)
btn_sel_import_file.pack(side=tk.LEFT, padx=5)

btn_importar = ttk.Button(
    frame_bd_left,
    text="Importar arquivo para tabela",
    style="Accent.TButton",
    command=lambda: importar_arquivo_para_tabela()
)
btn_importar.pack(pady=(10, 5))

# --------- IMPORTAÇÃO PARA TABELAS ---------

def importar_arquivo_para_tabela():
    global db_engine, db_connected

    if not db_connected or db_engine is None:
        messagebox.showwarning("Aviso", "Banco de dados não está conectado. Vá na aba 'Conexão BD' e conecte primeiro.")
        return

    tabela = import_tabela_var.get()
    path = import_arquivo_path.get().strip()

    if not tabela:
        messagebox.showwarning("Aviso", "Selecione a tabela destino.")
        return
    if not path:
        messagebox.showwarning("Aviso", "Selecione um arquivo bruto para importar.")
        return

    try:
        log_bd(f"📂 Lendo arquivo bruto: {path}")
        df_raw = read_table(path)
        log_bd(f"✅ Arquivo lido com {len(df_raw)} linhas e {len(df_raw.columns)} colunas.")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao ler arquivo.\n\n{e}")
        log_bd(f"❌ Erro ao ler arquivo: {e}")
        return

    normals = {normalize_col_name(c): c for c in df_raw.columns}

    try:
        if tabela == "empresas":
            log_bd("🧩 Preparando dados para tabela EMPRESAS...")

            cnpj_col = pick_col(normals, ["cnpj"])
            razao_col = pick_col(normals, ["razao social", "razao_social", "razao"])
            situacao_col = pick_col(normals, ["situacao cadastral", "situacao"])
            uf_col = pick_col(normals, ["uf", "estado"])
            data_abertura_col = pick_col(normals, ["data abertura", "dt abertura", "abertura"])
            telefones_col = pick_col(normals, ["telefones", "telefone"])
            tel1_col = pick_col(normals, ["telefone1", "tel1"])
            tel2_col = pick_col(normals, ["telefone2", "tel2"])
            email_col = pick_col(normals, ["email", "e-mail"])
            capital_col = pick_col(normals, ["capital social"])
            socios_col = pick_col(normals, ["socios", "sócios"])
            ultimo_uso_col = pick_col(normals, ["ultimo uso", "ultimo_uso", "ultimo contato"])
            plataforma_col = pick_col(normals, ["plataforma usada", "plataforma", "origem"])

            df_emp = pd.DataFrame()

            if cnpj_col:
                df_emp["cnpj"] = df_raw[cnpj_col].apply(normalize_cnpj)
            else:
                df_emp["cnpj"] = None

            if razao_col:
                df_emp["razao_social"] = df_raw[razao_col].astype(str)
            else:
                df_emp["razao_social"] = None

            df_emp["situacao_cadastral"] = df_raw[situacao_col].astype(str) if situacao_col else None
            df_emp["uf"] = df_raw[uf_col].astype(str).str[:2].str.upper() if uf_col else None

            if data_abertura_col:
                dt = pd.to_datetime(df_raw[data_abertura_col], errors="coerce", dayfirst=True)
                df_emp["data_abertura"] = dt.dt.date
            else:
                df_emp["data_abertura"] = None

            tel1 = None
            tel2 = None
            if telefones_col:
                t1, t2 = zip(*df_raw[telefones_col].map(split_telefones_field))
                tel1 = pd.Series(t1).apply(normalize_phone)
                tel2 = pd.Series(t2).apply(normalize_phone)
            if tel1_col and tel1 is None:
                tel1 = df_raw[tel1_col].apply(normalize_phone)
            if tel2_col and tel2 is None:
                tel2 = df_raw[tel2_col].apply(normalize_phone)

            df_emp["telefone1"] = tel1 if tel1 is not None else None
            df_emp["telefone2"] = tel2 if tel2 is not None else None

            df_emp["email"] = df_raw[email_col].astype(str) if email_col else None

            if capital_col:
                cap_series = df_raw[capital_col].astype(str).str.replace(".", "", regex=False).str.replace(",", ".", regex=False)
                df_emp["capital_social"] = pd.to_numeric(cap_series, errors="coerce")
            else:
                df_emp["capital_social"] = None

            df_emp["socios"] = df_raw[socios_col].astype(str) if socios_col else None

            if ultimo_uso_col:
                dt_ult = pd.to_datetime(df_raw[ultimo_uso_col], errors="coerce", dayfirst=True)
                df_emp["ultimo_uso"] = dt_ult
            else:
                df_emp["ultimo_uso"] = None

            df_emp["plataforma_usada"] = df_raw[plataforma_col].astype(str) if plataforma_col else None

            before = len(df_emp)
            df_emp = df_emp[df_emp["cnpj"].notna() & (df_emp["cnpj"] != "")]
            log_bd(f"➡️ Removidas {before - len(df_emp)} linhas sem CNPJ.")
            log_bd(f"📥 Preparado {len(df_emp)} registros para INSERT em 'empresas'.")

            df_emp.to_sql("empresas", db_engine, if_exists="append", index=False, chunksize=5000, method="multi")
            log_bd("✅ Importação para 'empresas' concluída.")

        elif tabela in ["block_list_c6", "block_list_b2b", "nao_perturbe"]:
            log_bd(f"🧩 Preparando dados para tabela {tabela.upper()}...")

            tel_col = None
            for c in df_raw.columns:
                if "tel" in normalize_col_name(c):
                    tel_col = c
                    break

            if not tel_col:
                raise ValueError("Não foi encontrada nenhuma coluna de telefone no arquivo.")

            df_tel = pd.DataFrame()
            df_tel["telefone"] = df_raw[tel_col].apply(normalize_phone)
            before = len(df_tel)
            df_tel = df_tel[df_tel["telefone"] != ""]
            df_tel = df_tel.drop_duplicates()
            log_bd(f"➡️ {before - len(df_tel)} linhas removidas (vazias/duplicadas).")

            df_tel.to_sql(tabela, db_engine, if_exists="append", index=False, chunksize=10000, method="multi")
            log_bd(f"✅ Importação para '{tabela}' concluída ({len(df_tel)} registros).")

        elif tabela == "cnais_aceitos":
            log_bd("🧩 Preparando dados para tabela CNAIS_ACEITOS...")

            cnai_col = None
            for c in df_raw.columns:
                norm = normalize_col_name(c)
                if any(k in norm for k in ["cnae", "cnai", "codigo"]):
                    cnai_col = c
                    break

            if not cnai_col:
                raise ValueError("Não foi encontrada nenhuma coluna de CNAI/CNAE no arquivo.")

            df_cnai = pd.DataFrame()
            df_cnai["cnai"] = df_raw[cnai_col].astype(str).str.strip()
            before = len(df_cnai)
            df_cnai = df_cnai[df_cnai["cnai"] != ""]
            df_cnai = df_cnai.drop_duplicates()
            log_bd(f"➡️ {before - len(df_cnai)} linhas removidas (vazias/duplicadas).")

            df_cnai.to_sql("cnais_aceitos", db_engine, if_exists="append", index=False, chunksize=10000, method="multi")
            log_bd(f"✅ Importação para 'cnais_aceitos' concluída ({len(df_cnai)} registros).")

        elif tabela == "lemit_relatorio":
            log_bd("🧩 Preparando dados para tabela LEMIT_RELATORIO...")

            contato_col = pick_col(normals, ["contato", "nome contato", "responsavel"])
            telefone_col = pick_col(normals, ["telefone", "tel", "telefone contato"])
            desc_col = pick_col(normals, ["descricao interacao", "descrição interacao", "descricao", "observacao"])
            cnpj_col = pick_col(normals, ["cnpj"])
            email_col = pick_col(normals, ["email", "e-mail"])
            data_abertura_col = pick_col(normals, ["data abertura", "dt abertura", "abertura"])
            uf_col = pick_col(normals, ["uf", "estado"])

            df_lr = pd.DataFrame()

            df_lr["contato"] = df_raw[contato_col].astype(str) if contato_col else None
            df_lr["telefone"] = df_raw[telefone_col].apply(normalize_phone) if telefone_col else None
            df_lr["descricao_interacao"] = df_raw[desc_col].astype(str) if desc_col else None
            df_lr["cnpj"] = df_raw[cnpj_col].apply(normalize_cnpj) if cnpj_col else None
            df_lr["email"] = df_raw[email_col].astype(str) if email_col else None

            if data_abertura_col:
                dt = pd.to_datetime(df_raw[data_abertura_col], errors="coerce", dayfirst=True)
                df_lr["data_abertura"] = dt.dt.date
            else:
                df_lr["data_abertura"] = None

            df_lr["uf"] = df_raw[uf_col].astype(str).str[:2].str.upper() if uf_col else None

            before = len(df_lr)
            df_lr = df_lr[ df_lr["cnpj"].notna() | df_lr["telefone"].notna() ]
            log_bd(f"➡️ {before - len(df_lr)} linhas removidas (sem CNPJ e sem telefone).")

            df_lr.to_sql("lemit_relatorio", db_engine, if_exists="append", index=False, chunksize=5000, method="multi")
            log_bd(f"✅ Importação para 'lemit_relatorio' concluída ({len(df_lr)} registros).")

        else:
            raise ValueError(f"Tabela '{tabela}' não tratada na importação.")

        messagebox.showinfo("Concluído", f"Importação para '{tabela}' finalizada com sucesso.")

    except Exception as e:
        log_bd(f"❌ Erro durante importação: {e}")
        messagebox.showerror("Erro", f"Erro durante importação.\n\n{e}")

# --------- VISUALIZAÇÃO ---------

lbl_view_title = tk.Label(
    frame_bd_right,
    text="Visualizar dados das tabelas",
    bg=BG_PRINCIPAL,
    fg=FG_TEXTO,
    font=fonte_titulo
)
lbl_view_title.pack(pady=(0, 10))

frame_view_top = tk.Frame(frame_bd_right, bg=BG_PRINCIPAL)
frame_view_top.pack(fill="x", pady=(0, 5), padx=4)

lbl_tab_view = tk.Label(
    frame_view_top,
    text="Tabela:",
    bg=BG_PRINCIPAL,
    fg=FG_TEXTO,
    font=fonte_label
)
lbl_tab_view.pack(side=tk.LEFT, padx=(0, 5))

combo_tab_view = ttk.Combobox(
    frame_view_top,
    textvariable=view_tabela_var,
    state="readonly",
    width=25,
    values=[
        "empresas",
        "block_list_c6",
        "block_list_b2b",
        "nao_perturbe",
        "cnais_aceitos",
        "lemit_relatorio"
    ]
)
combo_tab_view.pack(side=tk.LEFT, padx=(0, 5))
combo_tab_view.current(0)

def visualizar_tabela():
    global db_engine, db_connected

    if not db_connected or db_engine is None:
        messagebox.showwarning("Aviso", "Banco de dados não está conectado. Vá na aba 'Conexão BD' e conecte primeiro.")
        return

    tabela = view_tabela_var.get()
    if not tabela:
        messagebox.showwarning("Aviso", "Selecione uma tabela para visualizar.")
        return

    try:
        query = f"SELECT * FROM {tabela} LIMIT 200"
        df = pd.read_sql_query(query, db_engine)
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao consultar tabela.\n\n{e}")
        return

    tree_bd.delete(*tree_bd.get_children())
    tree_bd["columns"] = list(df.columns)
    tree_bd["show"] = "headings"

    for col in df.columns:
        tree_bd.heading(col, text=col)
        tree_bd.column(col, width=120, anchor="w")

    for _, row in df.iterrows():
        values = [str(row[col]) if row[col] is not None else "" for col in df.columns]
        tree_bd.insert("", tk.END, values=values)

    log_bd(f"👁️ Visualizando tabela '{tabela}' (até 200 linhas).")

btn_view = ttk.Button(
    frame_view_top,
    text="Atualizar visualização",
    style="Primary.TButton",
    command=visualizar_tabela
)
btn_view.pack(side=tk.LEFT, padx=(5, 0))

frame_tree = tk.Frame(frame_bd_right, bg=BG_PRINCIPAL)
frame_tree.pack(fill="both", expand=True, padx=4, pady=(0, 4))

tree_bd = ttk.Treeview(frame_tree)
tree_bd.pack(side=tk.LEFT, fill="both", expand=True)

scroll_y = ttk.Scrollbar(frame_tree, orient="vertical", command=tree_bd.yview)
scroll_y.pack(side=tk.RIGHT, fill="y")
tree_bd.configure(yscrollcommand=scroll_y.set)

frame_log_bd = ttk.Labelframe(
    frame_bd_right,
    text="Log do Banco de Dados",
    style="Frame.TLabelframe",
    padding=10
)
frame_log_bd.pack(fill="x", padx=4, pady=(4, 4))

txt_log_bd = tk.Text(
    frame_log_bd,
    width=60,
    height=8,
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
txt_log_bd.pack(fill="both", expand=True)

# =======================================================================
#                        ABA CONEXÃO BANCO DE DADOS
# =======================================================================

db_tipo_var = tk.StringVar(value="MySQL")
db_host_var = tk.StringVar()
db_port_var = tk.StringVar(value="3306")
db_user_var = tk.StringVar()
db_pass_var = tk.StringVar()
db_name_var = tk.StringVar()

lbl_conexao_title = tk.Label(
    frame_conexao,
    text="Conexão com Banco de Dados",
    bg=BG_PRINCIPAL,
    fg=FG_TEXTO,
    font=fonte_titulo
)
lbl_conexao_title.pack(pady=(10, 10))

frame_conexao_main = tk.Frame(frame_conexao, bg=BG_PRINCIPAL)
frame_conexao_main.pack(fill="both", expand=True, padx=10, pady=10)

frame_conexao_left = tk.Frame(frame_conexao_main, bg=BG_PRINCIPAL)
frame_conexao_left.pack(side=tk.LEFT, fill="both", expand=True, padx=(0, 6))

frame_conexao_right = tk.Frame(frame_conexao_main, bg=BG_PRINCIPAL)
frame_conexao_right.pack(side=tk.LEFT, fill="both", expand=True, padx=(6, 0))

frame_conn = ttk.Labelframe(
    frame_conexao_left,
    text="Configurações de Conexão",
    style="Frame.TLabelframe",
    padding=10
)
frame_conn.pack(padx=0, pady=6, fill="x")

lbl_tipo_db = tk.Label(frame_conn, text="Banco:", bg=BG_FRAME, fg=FG_TEXTO, font=fonte_label)
lbl_tipo_db.grid(row=0, column=0, padx=5, pady=5, sticky="w")

combo_tipo_db = ttk.Combobox(
    frame_conn,
    textvariable=db_tipo_var,
    state="readonly",
    width=20,
    values=["MySQL", "PostgreSQL"]
)
combo_tipo_db.grid(row=0, column=1, padx=5, pady=5)

lbl_host = tk.Label(frame_conn, text="Host:", bg=BG_FRAME, fg=FG_TEXTO, font=fonte_label)
lbl_host.grid(row=1, column=0, padx=5, pady=5, sticky="w")

entry_host = tk.Entry(
    frame_conn, textvariable=db_host_var,
    bg=INPUT_BG, fg=INPUT_FG, insertbackground=INPUT_FG,
    width=30
)
entry_host.grid(row=1, column=1, padx=5, pady=5)

lbl_port = tk.Label(frame_conn, text="Porta:", bg=BG_FRAME, fg=FG_TEXTO, font=fonte_label)
lbl_port.grid(row=2, column=0, padx=5, pady=5, sticky="w")

entry_port = tk.Entry(
    frame_conn, textvariable=db_port_var,
    bg=INPUT_BG, fg=INPUT_FG, insertbackground=INPUT_FG,
    width=30
)
entry_port.grid(row=2, column=1, padx=5, pady=5)

lbl_user = tk.Label(frame_conn, text="Usuário:", bg=BG_FRAME, fg=FG_TEXTO, font=fonte_label)
lbl_user.grid(row=3, column=0, padx=5, pady=5, sticky="w")

entry_user = tk.Entry(
    frame_conn, textvariable=db_user_var,
    bg=INPUT_BG, fg=INPUT_FG, insertbackground=INPUT_FG,
    width=30
)
entry_user.grid(row=3, column=1, padx=5, pady=5)

lbl_pass = tk.Label(frame_conn, text="Senha:", bg=BG_FRAME, fg=FG_TEXTO, font=fonte_label)
lbl_pass.grid(row=4, column=0, padx=5, pady=5, sticky="w")

entry_pass = tk.Entry(
    frame_conn, textvariable=db_pass_var,
    bg=INPUT_BG, fg=INPUT_FG, insertbackground=INPUT_FG,
    width=30, show="*"
)
entry_pass.grid(row=4, column=1, padx=5, pady=5)

lbl_name = tk.Label(frame_conn, text="Banco:", bg=BG_FRAME, fg=FG_TEXTO, font=fonte_label)
lbl_name.grid(row=5, column=0, padx=5, pady=5, sticky="w")

entry_name = tk.Entry(
    frame_conn, textvariable=db_name_var,
    bg=INPUT_BG, fg=INPUT_FG, insertbackground=INPUT_FG,
    width=30
)
entry_name.grid(row=5, column=1, padx=5, pady=5)

frame_log_conexao = ttk.Labelframe(
    frame_conexao_right,
    text="Log de Conexão",
    style="Frame.TLabelframe",
    padding=10
)
frame_log_conexao.pack(fill="both", expand=True, padx=5, pady=5)

txt_log_conexao = tk.Text(
    frame_log_conexao,
    width=60,
    height=20,
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
txt_log_conexao.pack(fill="both", expand=True)

def conectar_bd():
    global db_engine, db_connected

    txt_log_conexao.delete("1.0", tk.END)

    tipo = db_tipo_var.get()
    host = db_host_var.get().strip()
    port = db_port_var.get().strip()
    user = db_user_var.get().strip()
    password = db_pass_var.get().strip()
    database = db_name_var.get().strip()

    if not (host and port and user and database):
        messagebox.showwarning("Aviso", "Preencha Host, Porta, Usuário e Banco.")
        return

    try:
        if tipo == "PostgreSQL":
            url = URL.create(
                drivername="postgresql+psycopg2",
                username=user,
                password=password,
                host=host,
                port=int(port),
                database=database
            )
        elif tipo == "MySQL":
            url = URL.create(
                drivername="mysql+pymysql",
                username=user,
                password=password,
                host=host,
                port=int(port),
                database=database
            )
        else:
            raise ValueError("Tipo de banco não suportado.")

        engine = create_engine(url)

        with engine.connect() as conn:
            conn.execute(sql_text("SELECT 1"))

        db_engine = engine
        db_connected = True

        cfg = {
            "tipo": tipo,
            "host": host,
            "port": port,
            "user": user,
            "password": password,
            "database": database,
        }
        with open(DB_CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)

        txt_log_conexao.insert(tk.END, "✅ Conectado com sucesso!\n")
        txt_log_conexao.insert(tk.END, f"Configuração salva em {DB_CONFIG_FILE}\n")
        messagebox.showinfo("OK", "Conectado e configuração salva com sucesso!")

    except Exception as e:
        db_engine = None
        db_connected = False
        txt_log_conexao.insert(tk.END, f"❌ Erro ao conectar:\n{e}")
        messagebox.showerror("Erro", f"Falha ao conectar.\n\n{e}")

def auto_conectar_bd():
    global db_engine, db_connected

    if not os.path.exists(DB_CONFIG_FILE):
        return

    try:
        with open(DB_CONFIG_FILE, "r", encoding="utf-8") as f:
            cfg = json.load(f)

        db_tipo_var.set(cfg.get("tipo", "MySQL"))
        db_host_var.set(cfg.get("host", ""))
        db_port_var.set(cfg.get("port", "3306"))
        db_user_var.set(cfg.get("user", ""))
        db_pass_var.set(cfg.get("password", ""))
        db_name_var.set(cfg.get("database", ""))

        tipo = db_tipo_var.get()
        host = db_host_var.get().strip()
        port = db_port_var.get().strip()
        user = db_user_var.get().strip()
        password = db_pass_var.get().strip()
        database = db_name_var.get().strip()

        if not (host and port and user and database):
            return

        if tipo == "PostgreSQL":
            url = URL.create(
                drivername="postgresql+psycopg2",
                username=user,
                password=password,
                host=host,
                port=int(port),
                database=database
            )
        elif tipo == "MySQL":
            url = URL.create(
                drivername="mysql+pymysql",
                username=user,
                password=password,
                host=host,
                port=int(port),
                database=database
            )
        else:
            return

        engine = create_engine(url)
        with engine.connect() as conn:
            conn.execute(sql_text("SELECT 1"))

        db_engine = engine
        db_connected = True
        txt_log_conexao.insert(tk.END, "✅ Conectado automaticamente usando configuração salva.\n")

    except Exception as e:
        db_engine = None
        db_connected = False
        txt_log_conexao.insert(tk.END, f"⚠️ Não foi possível conectar automaticamente:\n{e}\n")

btn_conectar = ttk.Button(
    frame_conn,
    text="Conectar & Salvar",
    style="Accent.TButton",
    command=conectar_bd
)
btn_conectar.grid(row=6, column=0, columnspan=2, pady=12)

# =======================================================================
#                           MAINLOOP
# =======================================================================

janela.after(500, auto_conectar_bd)
janela.mainloop()
