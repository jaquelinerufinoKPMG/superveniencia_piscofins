from tqdm import tqdm
import pandas as pd
import os
import pyautogui
import time
from src.anexo_c.dashboard_tools import processDashboard
from openpyxl.utils import get_column_letter
from dot

# %%
dashboard_file = r"C:\dev\Itau\Superveniência\superveniencia_piscofins\Input\Anexo C\Dashboard.xlsx"
template_file = r"C:\dev\Itau\Superveniência\superveniencia_piscofins\template.xlsx"
dir_output = r"C:\dev\Itau\Superveniência\superveniencia_piscofins\Output"
contracts_file = r"C:\dev\Itau\Superveniência\superveniencia_piscofins\Input\Anexo C\Contratos_2014.xlsx"
file_not_found = r"C:\dev\Itau\Superveniência\superveniencia_piscofins\Output\file_not_found.txt"
worksheet_name="PIS_COFINS_ANUAL"
buttons_dir = r"botoes"


# %%
automacao = processDashboard(dashboard_file=dashboard_file,
                            template_file=template_file,
                            dir_output=dir_output,
                            worksheet_name=worksheet_name,
                            file_not_found=file_not_found,
                            buttons_dir=buttons_dir,)

# %%
import os
import re

# 1) Defina aqui o caminho da pasta e o arquivo de saída
pasta = dir_output
arquivo_saida = 'numeros_extraidos.txt'

# 2) Regex para capturar só dígitos
padrao = re.compile(r'\d+')

# 3) Lista pra acumular as sequências numéricas
numeros = []

for nome_arquivo in os.listdir(pasta):
    achados = padrao.findall(nome_arquivo)  # e.g. ['123','456']
    if achados:
        numeros.append(''.join(achados))     # => '123456'

# 4) Grava no TXT, uma sequência por linha
with open(arquivo_saida, 'w', encoding='utf-8') as f:
    for seq in numeros:
        f.write(seq + '\n')

print(f"✅ Extraídos {len(numeros)} itens e salvos em '{arquivo_saida}'")

data = pd.read_excel(contracts_file)
contratos_validos = data["Contrato"].values.tolist()

nao_encontrados_3 = set()
if os.path.isfile("numeros_extraidos.txt"):
    with open("numeros_extraidos.txt", "r") as f:
        nao_encontrados_3 = set(line.strip() for line in f)

# Filtre os contratos removendo aqueles que estão na lista de não encontrados
contratos = [contrato for contrato in contratos_validos if str(contrato) not in nao_encontrados_3]


# %% [markdown]
# ## Roda

# %%
automacao.process_contracts(contratos,0)


