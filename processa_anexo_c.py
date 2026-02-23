# %% [markdown]
# # Processamento do Anexo C

# %%
import pandas as pd
from dotenv import load_dotenv
import os
import re
from src.anexo_c.process_dashboard import process_dashboard
from src.anexo_c.export_pdf import export_pdf
from tqdm.auto import tqdm

# %% [markdown]
# ## Configurações

# %%
load_dotenv()

DASHBOARD_CSV = r"data\input\Diligencia_Contratos.csv"
TEMPLATE_PATH = r"docs\template_csll.xlsx"
TAX_FILTERS = {
    "IRPJ": ["RAIR", "Exclusão", "Adição"],
    "CS": ["RAIR", "Exclusão", "Adição"],
}

INPUT_DIR = r"data\input"
OUTPUT_DIR = r"data\output"
INVALID_DIR = r"data\invalid"

VALUE_COLS = ["ValorDebito", "ValorCredito", "Movimentacao"]
GROUP_COLS = ["Conta_Nome", "Cosif_Nome"]

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

# %% [markdown]
# ## Geração do Excel

# %%
dashboard = pd.read_csv(DASHBOARD_CSV, sep=';', encoding='latin1')

# %%
contratos_to_process = set()
if os.path.isfile("processar_kpmg.txt"):
    with open("processar_kpmg.txt", "r") as f:
        contratos_to_process = set(line.strip() for line in f)

# %%
cls = process_dashboard(TAX_FILTERS, TEMPLATE_PATH)
cls.get_contract_numbers(OUTPUT_DIR)

# %%
nao_encontrados_3 = set()
if os.path.isfile("numeros_extraidos.txt"):
    with open("numeros_extraidos.txt", "r") as f:
        nao_encontrados_3 = set(line.strip() for line in f)

# Filtre os contratos removendo aqueles que estão na lista de não encontrados
contratos = [
    int(contrato)
    for contrato in contratos_to_process
    if str(contrato) not in nao_encontrados_3
]

for contrato in tqdm(contratos, unit="contrato"):
    tqdm.desc = f"processando contrato: {contrato}"
    dashboard_filtrado = dashboard[dashboard["NumContrato"] == contrato]

    dashboard_filtro = cls.replicate_years(
        dashboard_filtrado,  tax_cols=["IRPJ", "CS"]
    )
    final = pd.DataFrame()

    # para cada filtro, gera o bloco e concatena
    for filtros, secao, desc in CONFIGS:
        bloco = cls.group_revenues(dashboard_filtro, filtros, secao, desc)
        final = pd.concat([final, bloco], ignore_index=True)

    final = cls.calcula_csll(final)

    df = final.copy()

    num_str = str(contrato)
    num_str = num_str.zfill(7)      # completa com zeros à esquerda
    output_path = os.path.join(OUTPUT_DIR, f"C{num_str}.xlsx")

    cls.atualizar_template_pivot(
        template_path=TEMPLATE_PATH, output_path=output_path, df=df, contrato=contrato
    )

# %% [markdown]
# ## Exportação do PDF

# %%
#export_pdf().process_file()

# %% [markdown]
# # ( •̀ ω •́ )y


