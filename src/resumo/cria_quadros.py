import numpy as np
import pandas as pd
from pathlib import Path

def preparar_conta_grafica(conta_grafica: str | Path) -> pd.DataFrame:
    COLUMNS = [
        "DATA GERACAO",
        "CONTADOR DEBITO",
        "NUM CONTRATO DEBITO",
        "CONTA DEBITO",
        "NOME CONTA DEBITO",
        "VALOR DEBITO",
        "CONTADOR CREDITO",
        "NUM CONTRATO CREDITO",
        "CONTA CREDITO",
        "NOME CONTA CREDITO",
        "VALOR CREDITO",
        "COSIF DEBITO",
        "NOME COSIF DEBITO",
        "COSIF CREDITO",
        "NOME COSIF CREDITO",
    ]

    DEBITO = [
        "DATA GERACAO",
        "ANO",
        "CONTADOR DEBITO",
        "NUM CONTRATO DEBITO",
        "CONTA DEBITO",
        "NOME CONTA DEBITO",
        "VALOR DEBITO",
        "COSIF DEBITO",
        "NOME COSIF DEBITO",
    ]
    CREDITO = [
        "DATA GERACAO",
        "ANO",
        "CONTADOR CREDITO",
        "NUM CONTRATO CREDITO",
        "CONTA CREDITO",
        "NOME CONTA CREDITO",
        "VALOR CREDITO",
        "COSIF CREDITO",
        "NOME COSIF CREDITO",
    ]
    RENAME_COLS_DEBITO = [
        "Data",
        "Ano",
        "Contador",
        "Num Contrato",
        "Conta",
        "Nome Conta",
        "Valor Debito",
        "COSIF",
        "Nome COSIF",
    ]
    RENAME_COLS_CREDITO = [
        "Data",
        "Ano",
        "Contador",
        "Num Contrato",
        "Conta",
        "Nome Conta",
        "Valor Credito",
        "COSIF",
        "Nome COSIF",
    ]

    conta_grafica = Path(conta_grafica)

    df = pd.read_csv(conta_grafica, sep=";", encoding="utf-8", dtype=str)

    df.rename(columns=dict(zip(df.columns, COLUMNS)), inplace=True)
    df["VALOR DEBITO"] = (
        df["VALOR DEBITO"].astype(str).str.replace(",", ".", regex=False).astype(float)
    )

    df["VALOR CREDITO"] = (
        df["VALOR CREDITO"].astype(str).str.replace(",", ".", regex=False).astype(float)
    )

    col = (
        df["DATA GERACAO"]
        .astype(str)
        .str.strip()
        .replace({"99999999": None, "00000000": None, "": None})
    )

    df["DATA GERACAO"] = pd.to_datetime(col, format="%Y%m%d", errors="coerce")

    df["ANO"] = df["DATA GERACAO"].dt.year.astype("Int64").astype(str)

    df_credito = df[CREDITO].copy()
    df_debito = df[DEBITO].copy()

    df_credito.rename(columns=dict(zip(CREDITO, RENAME_COLS_CREDITO)), inplace=True)
    df_debito.rename(columns=dict(zip(DEBITO, RENAME_COLS_DEBITO)), inplace=True)

    df_credito["Valor Debito"] = 0.0
    df_credito["Tipo"] = "C"
    df_debito["Valor Credito"] = 0.0
    df_debito["Tipo"] = "D"

    df_final = pd.concat([df_credito, df_debito], ignore_index=True)
    df_final["COSIF - Filtro"] = np.where(
        df_final["COSIF"].isna(), np.nan, df_final["COSIF"].str[:5]
    )
    df_final["AnoMes"] = (
        df_final["Data"].dt.to_period("M").astype(str).str.replace("-", "", regex=False)
    )
    df_final["Período de Apuração"] = "31/12/" + df_final["Ano"]
    df_final["COSIF - Nivel 1"] = np.where(
        df_final["COSIF"].isna(),
        np.nan,
        (df_final["COSIF"].astype(str).str[:3].str.ljust(14, "0")),
    )
    df_final["COSIF - Nivel 2"] = np.where(
        df_final["COSIF"].isna(),
        np.nan,
        (df_final["COSIF"].astype(str).str[:7].str.ljust(14, "0")),
    )
    df_final["Conta + Descrição"] = np.where(
        df_final["Conta"].isna(),
        np.nan,
        df_final["Conta"] + " - " + df_final["Nome Conta"],
    )
    df_final["Valor Líquido"] = np.where(
        df_final["Tipo"] == "D", -df_final["Valor Debito"], df_final["Valor Credito"]
    )
    cosif = df_final["COSIF - Filtro"].astype(str)

    df_final["COSIF Apresentação"] = cosif.str.replace(
        r"(\d)(\d)(\d)(\d{2}).*", r"\1.\2.\3.\4", regex=True
    )

    return df_final

def cria_quadro_1(df: pd.DataFrame, pasta_saida: str | Path) -> None:
    COSIFS = {
        "23210": "Valor do ativo contabilizado",
        "17110": "Contraprestação de arrendamento a receber contabilizada",
        "17510": "Valores residuais a realizar contabilizado",
    }

    resultados = pd.DataFrame(columns=["Num Contrato", "COSIF", "Valor", "Descrição"])

    for cosif, descricao in COSIFS.items():
        recorte_df = (
            df[df["COSIF - Filtro"] == cosif]
            .sort_values(by=["Contador"])
            .groupby(["Num Contrato", "COSIF Apresentação"], as_index=False)
            .first()[["Num Contrato", "COSIF Apresentação", "Valor Debito"]]
        )
        recorte_df["Descrição"] = descricao
        recorte_df.rename(
            columns={"Valor Debito": "Valor", "COSIF Apresentação": "COSIF"},
            inplace=True,
        )
        resultados = pd.concat([resultados, recorte_df], ignore_index=True)

    pasta_saida = Path(pasta_saida)
    pasta_saida.mkdir(parents=True, exist_ok=True)

    resultados.sort_values(by=["Num Contrato", "COSIF"], inplace=True)
    resultados.to_csv(pasta_saida / "quadro_1.csv", index=False, encoding="utf-8")
    
    print("Quadro 1 criado com sucesso!")