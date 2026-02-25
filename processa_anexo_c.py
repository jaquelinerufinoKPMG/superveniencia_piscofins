import os
import pandas as pd
from dotenv import load_dotenv
from tqdm.auto import tqdm

from src.anexo_c.process_dashboard import process_dashboard
load_dotenv()

FINAL_CSV = r"C:\dev\superveniencia_piscofins\data\input\final_20260224.csv"  # <-- vem do pipeline anterior
TEMPLATE_PATH = r"C:\dev\superveniencia_piscofins\docs\template_csll.xlsx"

TAX_FILTERS = {
    "IRPJ": ["RAIR", "Exclusão", "Adição"],
    "CS": ["RAIR", "Exclusão", "Adição"],
}

OUTPUT_DIR = r"C:\dev\superveniencia_piscofins\data\output"
final_all = pd.read_csv(FINAL_CSV, dtype={"NumContrato": "int64"})

cls = process_dashboard(TAX_FILTERS, TEMPLATE_PATH)
cls.get_contract_numbers(OUTPUT_DIR, r"C:\dev\superveniencia_piscofins\numeros_extraidos.txt")


contratos_to_process = set()
if os.path.isfile(r"C:\dev\superveniencia_piscofins\contratos_para_processar_jaq.txt"):
    with open(r"C:\dev\superveniencia_piscofins\contratos_para_processar_jaq.txt", "r", encoding="utf-8") as f:
        contratos_to_process = {int(line.strip()) for line in f if line.strip().isdigit()}

nao_encontrados_3 = set()
if os.path.isfile(r"C:\dev\superveniencia_piscofins\numeros_extraidos.txt"):
    with open(r"C:\dev\superveniencia_piscofins\numeros_extraidos.txt", "r", encoding="utf-8") as f:
        nao_encontrados_3 = {int(line.strip()) for line in f if line.strip().isdigit()}

# Filtre os contratos removendo aqueles que estão na lista de não encontrados
contratos = sorted(list(contratos_to_process - nao_encontrados_3))
for contrato in tqdm(contratos, unit="contrato"):
    tqdm.desc = f"processando contrato: {contrato}"

    # pega só o que é desse contrato no final.csv
    df = final_all[final_all["NumContrato"] == contrato]

    # se não existir no CSV, pula
    if df.empty:
        continue

    num_str = str(contrato).zfill(7)  # completa com zeros à esquerda
    output_path = os.path.join(OUTPUT_DIR, f"C{num_str}.xlsx")

    cls.atualizar_template_pivot(
        template_path=TEMPLATE_PATH,
        output_path=output_path,
        df=df,
        contrato=contrato
    )
