from __future__ import annotations

import os
from pathlib import Path

import pandas as pd
from dotenv import load_dotenv
from tqdm.auto import tqdm

from src.anexo_c.process_dashboard import process_dashboard

load_dotenv()

DASHBOARD_CSV = Path(r"data\input\Diligencia_Contratos.csv")
CONTRATOS_ERRO_XLSX = Path(r"docs\contratos_com_erro.xlsx")
NAO_ENCONTRADOS_TXT = Path(r"numeros_extraidos.txt")

TEMPLATE_PATH = Path(r"docs\template_csll.xlsx")
OUTPUT_DIR = Path(r"data\output")

TAX_FILTERS = {
    "IRPJ": ["RAIR", "Exclusão", "Adição"],
    "CS": ["RAIR", "Exclusão", "Adição"],
}

CONFIGS = [
    (
        {"IRPJ": ["RAIR", "Exclusão", "Adição"], "CS": ["RAIR", "Exclusão", "Adição"]},
        "IRPJ",
        "Resultado antes do IR",
    ),
    ({"IRPJ": ["Adição"], "CS": ["RAIR", "Exclusão", "Adição"]}, "IRPJ", "Adição"),
    ({"IRPJ": ["Exclusão"], "CS": ["RAIR", "Exclusão", "Adição"]}, "IRPJ", "Exclusão"),
    (
        {"IRPJ": ["RAIR", "Exclusão", "Adição"], "CS": ["RAIR", "Exclusão", "Adição"]},
        "CSLL",
        "Resultado antes do CSLL",
    ),
    ({"IRPJ": ["RAIR", "Exclusão", "Adição"], "CS": ["Adição"]}, "CSLL", "Adição"),
    ({"IRPJ": ["RAIR", "Exclusão", "Adição"], "CS": ["Exclusão"]}, "CSLL", "Exclusão"),
]

def load_contratos_selecionados(path: Path) -> set[int]:
    """
    Lê a planilha docs/contratos_com_erro.xlsx (coluna 'arquivo')
    no formato C0001234.xlsx e retorna set de int (1234).
    """
    s = pd.read_excel(path)["arquivo"].astype(str)

    # remove extensão e o "C"
    s = (
        s.str.replace(".xlsx", "", regex=False)
         .str.replace("C", "", regex=False)
         .str.strip()
    )

    # mantém só números válidos
    s = pd.to_numeric(s, errors="coerce").dropna().astype("int64")
    return set(s.tolist())

def load_nao_encontrados(path: Path) -> set[int]:
    """
    Lê numeros_extraidos.txt (1 número por linha) e retorna set[int].
    Ignora linhas inválidas.
    """
    if not path.exists():
        return set()

    out: set[int] = set()
    with path.open("r", encoding="utf-8") as f:
        for line in f:
            t = line.strip()
            if not t:
                continue
            if t.isdigit():
                out.add(int(t))
    return out

dashboard = pd.read_csv(
    DASHBOARD_CSV,
    sep=";",
    encoding="latin1",
    dtype={"NumContrato": "int64"},
)

contratos_selecionados = load_contratos_selecionados(CONTRATOS_ERRO_XLSX)
nao_encontrados = load_nao_encontrados(NAO_ENCONTRADOS_TXT)

mask = dashboard["NumContrato"].isin(contratos_selecionados)
if nao_encontrados:
    mask &= ~dashboard["NumContrato"].isin(nao_encontrados)

dash_small = dashboard.loc[mask].copy()

del dashboard

cls = process_dashboard(TAX_FILTERS, str(TEMPLATE_PATH))
cls.get_contract_numbers(str(OUTPUT_DIR))

final_parts: list[pd.DataFrame] = []

n_contracts = int(dash_small["NumContrato"].nunique())

for contrato, df_contrato in tqdm(
    dash_small.groupby("NumContrato", sort=False),
    total=n_contracts,
    unit=" contrato",
):
    # replica anos 1x
    df_rep = cls.replicate_years(df_contrato, tax_cols=["IRPJ", "CS"])

    # monta blocos e concatena 1x
    blocos = [
        cls.group_revenues(df_rep, filtros, secao, desc)
        for filtros, secao, desc in CONFIGS
    ]
    final = pd.concat(blocos, ignore_index=True)

    final = cls.calcula_csll(final)
    final["NumContrato"] = contrato

    final_parts.append(final)

final_final = pd.concat(final_parts, ignore_index=True)

final_final.to_csv("final.csv", index=False)