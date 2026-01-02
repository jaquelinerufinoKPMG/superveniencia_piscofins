import os
import re
import shutil

from dotenv import find_dotenv
from environs import Env
import pandas as pd
from pathlib import Path
from tqdm import tqdm

from src.utils.normalize_text import DocumentFormatter

env = Env()
env.read_env(find_dotenv())

INPUT_DIR = env.str("INPUT_DIR")
OUTPUT_DIR = env.str("OUTPUT_DIR")

PATH_BASE = r"data\Input\DOC_BASE_2014.xlsx"
PATH_TELA_PRETA = r"data\Input\Anexo A"

def _split_por_contrato(texto: str) -> list[str]:
    texto = texto.replace("\r\n", "\n").replace("\r", "\n")
    # divide preservando o '@.../....' como parte do bloco
    parts = re.split(r"(?=^@\d+/\d+\s*$)", texto, flags=re.MULTILINE)
    return [p.strip() for p in parts if p.strip()]

def _read_txt_file(path_txt: str) -> str:
    try:
        with open(path_txt, 'r', encoding='utf-8') as f:
            texto = f.read()
    except UnicodeDecodeError:
        with open(path_txt, 'r', encoding='latin1') as f:
            texto = f.read()
    return texto

def _Tela_Contrato_L7RR(bloco: str):
    # EMP/CONTRATO: 1/0006306
    m = re.search(r"EMP/CONTRATO:\s*(?P<emp_contrato>\d+/\d+)", bloco)
    emp_contrato = m.group("emp_contrato") if m else None

    # CLIENTE: ... CPF : ...
    m = re.search(r"CLIENTE:\s*(?P<cliente>.+?)\s+CPF\s*:\s*(?P<cpf>[0-9\.\-]+)", bloco)
    cliente = m.group("cliente").strip() if m else None
    cpf = m.group("cpf").strip() if m else None

    # DT.CONTRAT.INICIAL: 27/07/05
    m = re.search(r"DT\.CONTRAT\.INICIAL:\s*(\d{2}/\d{2}/\d{2})", bloco)
    dt_contrato_inicial = m.group(1) if m else None

    # QTD CPA:  60
    m = re.search(r"QTD\s+CPA:\s*(\d+)", bloco)
    qtd_cpa = m.group(1) if m else None

    return [emp_contrato, cliente, dt_contrato_inicial, qtd_cpa, cpf]

def extracao_detalhes_contrato(texto: str) -> pd.DataFrame:
    blocos = _split_por_contrato(texto)
    rows = []

    for blk in blocos:
        # só tenta parsear blocos que realmente pareçam ser essa tela
        if "EMP/CONTRATO:" not in blk.upper():
            continue

        emp_contrato, cliente, dt_contrato_inicial, qtd_cpa, cpf = _Tela_Contrato_L7RR(blk)
        rows.append({
            "contrato": emp_contrato,
            "cliente": cliente,
            "dt_contrato_inicial": dt_contrato_inicial,
            "qtd_cpa": qtd_cpa,
            "cpf": cpf,
        })

    return pd.DataFrame(rows)

def _f4_Tela_Consulta_de_bens(bloco: str):
    m = re.search(r"DT\.LIQUIDACAO:\s*(\d{2}/\d{2}/\d{2})", bloco)
    dt_liquidacao = m.group(1) if m else None

    m = re.search(r"CONTRATO:\s*(\d+)", bloco)
    n_contrato = m.group(1) if m else None

    valores = re.findall(r"VL\.UNITARIO:\s*([\d\.]+,\d{2})", bloco)
    valor_bem = sum(float(v.replace(".", "").replace(",", ".")) for v in valores) if valores else 0.0

    return [n_contrato, dt_liquidacao, round(valor_bem, 2)]

def extracao_consulta_bens(texto: str) -> pd.DataFrame:
    blocos = _split_por_contrato(texto)
    rows = []

    for blk in blocos:
        if "CONSULTA BEM" not in blk.upper() and "CONTRATO:" not in blk.upper():
            continue

        n_contrato, dt_liquidacao, valor_bem = _f4_Tela_Consulta_de_bens(blk)
        if n_contrato:  # evita lixo
            rows.append({
                "contrato": n_contrato,
                "dt_liquidacao": dt_liquidacao,
                "valor_bem": valor_bem
            })

    return pd.DataFrame(rows)

def process_doc_base(path_base: str, sheet_name: str = "Base_2014"):
    file_name = rf"{OUTPUT_DIR}\DOC_BASE_Reprocessado.xlsx"

    df_base_completa = pd.read_excel(path_base)
    df_base_completa = df_base_completa[["nº do contrato","razão social/nome completo do arrendatário","CNPJ/CPF do arrendatário","valor do bem","data do contrato","data de liquidação do contrato","nº de parcelas contratadas"]]
    df_base_completa.columns =["nº do contrato","Razão Social_Nome Completo do Arrendatário","CNPJ_CPF do Arrendatário","Valor do Bem","data do contrato","Liquidação","N° de parcelas contratadas"]

    df_base_completa.to_excel(file_name, sheet_name=sheet_name, index=False)

def merge_tela_preta(path_bens: str, path_contratos: str) -> pd.DataFrame:    
    df_resultados = extracao_consulta_bens(_read_txt_file(path_bens))
    df_contratos = extracao_detalhes_contrato(_read_txt_file(path_contratos))
    df_contratos["contrato"] = df_contratos["contrato"].str.split("/").str[-1]


    df_merged = pd.merge(df_resultados, df_contratos, how='left', on='contrato')
    
    df_merged = df_merged.drop_duplicates()
    return df_merged

TIPOS_COLUNAS = {
    'nº do contrato': str,   # Forçar coluna como string
    'Razão Social_Nome Completo do Arrendatário': str,   # Forçar coluna como string
    'CNPJ_CPF do Arrendatário': str,   # Forçar coluna como string
    'Nº de Parcelas Contratadas': str, # Forçar coluna como número
    'Valor do Bem': float, # Forçar coluna como número
}

DATA_COLUMNS = [
    'data do contrato',
    "Liquidação"
]

STATUS_COLS = ["Contrato","Cliente - Status","CNPJ/CPF - Status","Valor - Status","Data do Contrato - Status","Data de Liquidação - Status",'Quantidade de Parcelas - Status']

def export_status_file(path_base: str, file_name: str, path_bens: str, path_contratos: str):
    df_merged = merge_tela_preta(path_bens, path_contratos)

    df_base = pd.read_excel(path_base, dtype=TIPOS_COLUNAS, parse_dates=DATA_COLUMNS)
    rename_df_base = {'nº do contrato':"contrato_base", 'Razão Social_Nome Completo do Arrendatário': "cliente",
       'CNPJ_CPF do Arrendatário': "cpf", 'Valor do Bem': "valor_bem", 'data do contrato': "data_contrato",
       'Liquidação': "dt_liquidacao", 'N° de parcelas contratadas': "parcelas_contratadas"}
    df_base = df_base.rename(columns=rename_df_base)

    df_base['parcelas_contratadas'] = pd.to_numeric(df_base['parcelas_contratadas'], errors='coerce').fillna(0).astype(int)
    df_base['cpf'] = df_base['cpf'].apply(DocumentFormatter.format_documents)
    df_base['cliente'] = df_base['cliente'].apply(DocumentFormatter.to_pascal_case)
    df_base.groupby('contrato').count().reset_index().sort_values('cpf',ascending=False)
    df_base = DocumentFormatter.format_date_columns(df_base, DATA_COLUMNS)
    df_base['contrato'] = df_base['contrato'].astype(str) 

    df_merged['data_contrato'] = df_merged['data_contrato'].apply(DocumentFormatter.correct_year)
    df_merged['dt_liquidacao'] = df_merged['dt_liquidacao'].apply(DocumentFormatter.correct_year)

    df_merged_comparativo = pd.merge(
        df_merged,
        df_base,
        left_on="contrato",
        right_on="contrato_base",
        how="left",
        suffixes=('_extraido', '_base')
    )

    def compara_str(a, b):
        if pd.isnull(a) and pd.isnull(b):
            return "Conferido"
        if pd.isnull(a) or pd.isnull(b):
            return "Divergente"
        return "Conferido" if str(a).strip().upper() == str(b).strip().upper() else "Divergente"

    def compara_num(a, b, tol=0.01):
        try:
            if pd.isnull(a) and pd.isnull(b):
                return "Conferido"
            if pd.isnull(a) or pd.isnull(b):
                return "Divergente"
            return "Conferido" if abs(float(a) - float(b)) < tol else "Divergente"
        except Exception:
            return "Divergente"

    def compara_data(a, b):
        if pd.isnull(a) and pd.isnull(b):
            return "Conferido"
        if pd.isnull(a) or pd.isnull(b):
            return "Divergente"
        # Normaliza para dd/mm/aaaa
        try:
            a_fmt = pd.to_datetime(a, dayfirst=True, errors='coerce')
            b_fmt = pd.to_datetime(b, dayfirst=True, errors='coerce')
            if pd.isnull(a_fmt) or pd.isnull(b_fmt):
                return "Divergente"
            return "Conferido" if a_fmt == b_fmt else "Divergente"
        except Exception:
            return "Divergente"

    df_merged_comparativo["Contrato"] = df_merged_comparativo.apply(
        lambda x: compara_str(x["contrato"], x["contrato_base"]), axis=1)
    df_merged_comparativo["Cliente - Status"] = df_merged_comparativo.apply(
        lambda x: compara_str(x["cliente_base"], x["cliente_extraido"]), axis=1)
    df_merged_comparativo["CNPJ/CPF - Status"] = df_merged_comparativo.apply(
        lambda x: compara_str(
            ''.join(filter(str.isdigit, str(x["cpf_base"]))),
            ''.join(filter(str.isdigit, str(x["cpf_extraido"])))
        ), axis=1)
    df_merged_comparativo["Valor - Status"] = df_merged_comparativo.apply(
        lambda x: compara_num(x["valor_bem_extraido"], x["valor_bem_base"]), axis=1)
    df_merged_comparativo["Data do Contrato - Status"] = df_merged_comparativo.apply(
        lambda x: compara_data(x["data_contrato"], x["dt_contrato_inicial"]), axis=1)
    df_merged_comparativo["Data de Liquidação - Status"] = df_merged_comparativo.apply(
        lambda x: compara_data(x["dt_liquidacao_extraido"], x["dt_liquidacao_base"]), axis=1)
    df_merged_comparativo["Quantidade de Parcelas - Status"] = df_merged_comparativo.apply(
        lambda x: compara_num(x["qtd_cpa"], x["parcelas_contratadas"]), axis=1)
    df_merged_comparativo["CNPJ/CPF - Status"] = df_merged_comparativo.apply(lambda x: "Conferido" if x["Cliente - Status"] == "Conferido" else x["CNPJ/CPF - Status"], axis=1)
    df_merged_comparativo["Cliente - Status"] = df_merged_comparativo.apply(lambda x: "Conferido" if x["CNPJ/CPF - Status"] == "Conferido" else x["Cliente - Status"], axis=1)

    df_merged_comparativo["Contrato"] = df_merged_comparativo["contrato"]
    
    df_merged_comparativo[STATUS_COLS].to_excel(rf"{OUTPUT_DIR}/{file_name}.xlsx", index=False)