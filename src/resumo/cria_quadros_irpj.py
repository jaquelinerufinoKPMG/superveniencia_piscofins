from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd

CONTA_GRAFICA_COLUMNS = [
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

LANCAMENTO_CONFIG = {
    "debito": {
        "columns": [
            "DATA GERACAO",
            "DATA GERACAO ALTERADA",
            "ANO",
            "CONTADOR DEBITO",
            "NUM CONTRATO DEBITO",
            "CONTA DEBITO",
            "NOME CONTA DEBITO",
            "VALOR DEBITO",
            "COSIF DEBITO",
            "NOME COSIF DEBITO",
        ],
        "rename": {
            "DATA GERACAO": "Data Original",
            "DATA GERACAO ALTERADA": "Data",
            "ANO": "Ano",
            "CONTADOR DEBITO": "Contador",
            "NUM CONTRATO DEBITO": "Num Contrato",
            "CONTA DEBITO": "Conta",
            "NOME CONTA DEBITO": "Nome Conta",
            "VALOR DEBITO": "Valor Debito",
            "COSIF DEBITO": "COSIF",
            "NOME COSIF DEBITO": "Nome COSIF",
        },
        "defaults": {"Valor Credito": 0.0, "Tipo": "D"},
    },
    "credito": {
        "columns": [
            "DATA GERACAO",
            "DATA GERACAO ALTERADA",
            "ANO",
            "CONTADOR CREDITO",
            "NUM CONTRATO CREDITO",
            "CONTA CREDITO",
            "NOME CONTA CREDITO",
            "VALOR CREDITO",
            "COSIF CREDITO",
            "NOME COSIF CREDITO",
        ],
        "rename": {
            "DATA GERACAO": "Data Original",
            "DATA GERACAO ALTERADA": "Data",
            "ANO": "Ano",
            "CONTADOR CREDITO": "Contador",
            "NUM CONTRATO CREDITO": "Num Contrato",
            "CONTA CREDITO": "Conta",
            "NOME CONTA CREDITO": "Nome Conta",
            "VALOR CREDITO": "Valor Credito",
            "COSIF CREDITO": "COSIF",
            "NOME COSIF CREDITO": "Nome COSIF",
        },
        "defaults": {"Valor Debito": 0.0, "Tipo": "C"},
    },
}

COSIFS_QUADRO_1 = {
    "23210": "Valor do ativo contabilizado",
    "17110": "Contraprestação de arrendamento a receber contabilizada",
    "17510": "Valores residuais a realizar contabilizado",
}
CONTA_QUADRO_2 = "71300010030016"
COSIF_DESCONTOS_CONCEDIDOS = "81999006"
COSIF_QUADRO_2 = "71210001"
COSIF_QUADRO_2_DESCRICAO = "RENDAS DE ARRENDAM.FINANC.- RECURSOS INTERNOS"
TEXT_COLUMNS_CONTA_GRAFICA = [
    "Data Original",
    "Ano",
    "Contador",
    "Num Contrato",
    "Conta",
    "Nome Conta",
    "COSIF",
    "Nome COSIF",
    "Tipo",
    "COSIF - Filtro",
    "AnoMes",
    "Período de Apuração",
    "COSIF - Nivel 1",
    "COSIF - Nivel 2",
    "Conta + Descrição",
    "COSIF Apresentação",
]


def _remover_colunas_extras_vazias(
    df: pd.DataFrame, expected_columns: list[str]
) -> pd.DataFrame:
    expected_len = len(expected_columns)
    actual_len = df.shape[1]

    if actual_len < expected_len:
        raise ValueError(
            f"Arquivo da conta gráfica possui {actual_len} colunas; "
            f"esperado no mínimo {expected_len}."
        )

    if actual_len == expected_len:
        return df

    extra_columns = df.iloc[:, expected_len:]
    extra_not_empty = extra_columns.replace(r"^\s*$", pd.NA, regex=True).dropna(
        axis=1, how="all"
    )

    if not extra_not_empty.empty:
        raise ValueError(
            f"Arquivo da conta gráfica possui {actual_len} colunas; "
            f"esperado {expected_len}. Há colunas extras não vazias no layout."
        )

    return df.iloc[:, :expected_len].copy()


def _parse_valor_monetario(series: pd.Series) -> pd.Series:
    return pd.to_numeric(
        series.fillna("").str.strip().str.replace(",", ".", regex=False),
        errors="coerce",
    )


def _parse_data_geracao(series: pd.Series, data_final: datetime) -> pd.Series:
    data_final_formatada = data_final.strftime("%Y%m%d")
    datas_normalizadas = (
        series.fillna("")
        .str.strip()
        .replace(
            {
                "99999999": data_final_formatada,
                "00000000": data_final_formatada,
                "": pd.NA,
            }
        )
    )
    return pd.to_datetime(datas_normalizadas, format="%Y%m%d", errors="coerce")


def _montar_lancamentos(df: pd.DataFrame, lado: str) -> pd.DataFrame:
    config = LANCAMENTO_CONFIG[lado]
    return (
        df[config["columns"]]
        .rename(columns=config["rename"])
        .assign(**config["defaults"])
    )


def _adicionar_colunas_derivadas(df: pd.DataFrame) -> pd.DataFrame:
    resultado = df.copy()
    resultado["COSIF - Filtro"] = resultado["COSIF"].str[:5]
    resultado["AnoMes"] = resultado["Data"].dt.strftime("%Y%m")
    resultado["Período de Apuração"] = "31/12/" + resultado["Ano"]
    resultado["COSIF - Nivel 1"] = resultado["COSIF"].str[:3].str.ljust(14, "0")
    resultado["COSIF - Nivel 2"] = resultado["COSIF"].str[:7].str.ljust(14, "0")
    resultado["Conta + Descrição"] = (
        resultado["Conta"] + " - " + resultado["Nome Conta"]
    )
    resultado["Valor Líquido"] = np.where(
        resultado["Tipo"].eq("D"),
        -resultado["Valor Debito"],
        resultado["Valor Credito"],
    )
    resultado["COSIF Apresentação"] = resultado["COSIF - Filtro"].str.replace(
        r"(\d)(\d)(\d)(\d{2}).*",
        r"\1.\2.\3.\4",
        regex=True,
    )
    return resultado


def _forcar_colunas_textuais(df: pd.DataFrame, columns: list[str]) -> pd.DataFrame:
    resultado = df.copy()
    for column in columns:
        if column in resultado.columns:
            resultado[column] = resultado[column].astype("string")
    return resultado


def _garantir_pasta_saida(pasta_saida: str | Path) -> Path:
    pasta = Path(pasta_saida)
    pasta.mkdir(parents=True, exist_ok=True)
    return pasta


def _normalizar_ano_para_expansao(grupo: pd.DataFrame) -> pd.DataFrame:
    anos = pd.to_numeric(grupo["Ano"], errors="coerce")
    if anos.isna().any():
        raise ValueError("A coluna 'Ano' contém valores não numéricos.")
    return grupo.assign(Ano=anos.astype(int))


def expandir(grupo: pd.DataFrame) -> pd.DataFrame:
    grupo_normalizado = _normalizar_ano_para_expansao(grupo)
    anos = range(
        grupo_normalizado["Ano"].min(),
        grupo_normalizado["Ano"].max() + 1,
    )
    return (
        grupo_normalizado.set_index("Ano")
        .reindex(anos, fill_value=0)
        .rename_axis("Ano")
        .reset_index()
        .assign(Contrato=grupo_normalizado["Contrato"].iloc[0])
    )


def _expandir_por_contrato(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df.copy()

    return (
        df.groupby("Contrato", group_keys=False, sort=False)
        .apply(expandir)
        .reset_index(drop=True)
    )


def _ordenar_ano_com_total_no_final(df: pd.DataFrame) -> pd.DataFrame:
    return (
        df.assign(
            Ano_sort=lambda current: pd.to_numeric(
                current["Ano"].replace("Total", pd.NA), errors="coerce"
            )
            .fillna(9999)
            .astype(int)
        )
        .sort_values(["Contrato", "Ano_sort"])
        .drop(columns="Ano_sort")
    )


def _adicionar_total_por_contrato(
    df: pd.DataFrame, value_columns: list[str]
) -> pd.DataFrame:
    totais = df.groupby("Contrato", as_index=False)[value_columns].sum()
    totais["Ano"] = "Total"
    final = pd.concat([df, totais], ignore_index=True, sort=False)
    ordered_columns = ["Contrato", "Ano"] + [
        column for column in df.columns if column not in {"Contrato", "Ano"}
    ]
    return _ordenar_ano_com_total_no_final(final[ordered_columns])


def _salvar_quadro(df: pd.DataFrame, pasta_saida: str | Path, nome_arquivo: str) -> None:
    pasta = _garantir_pasta_saida(pasta_saida)
    df.to_csv(pasta / nome_arquivo, index=False, encoding="utf-8")


def _linhas_anuais(df: pd.DataFrame) -> pd.DataFrame:
    return df[df["Ano"].astype(str) != "Total"].copy()


def preparar_conta_grafica(
    conta_grafica: str | Path, data_final: datetime, salvar_parquet: bool = True
) -> pd.DataFrame:
    """Lê a conta gráfica, normaliza lançamentos de débito/crédito e retorna o dataframe final."""
    conta_grafica_path = Path(conta_grafica)

    df = pd.read_csv(
        conta_grafica_path,
        sep=";",
        encoding="utf-8",
        dtype=str,
        header=None,
        low_memory=False,
    )
    df = _remover_colunas_extras_vazias(df, CONTA_GRAFICA_COLUMNS)
    df.columns = CONTA_GRAFICA_COLUMNS

    df["VALOR DEBITO"] = _parse_valor_monetario(df["VALOR DEBITO"])
    df["VALOR CREDITO"] = _parse_valor_monetario(df["VALOR CREDITO"])
    df["DATA GERACAO ALTERADA"] = _parse_data_geracao(df["DATA GERACAO"], data_final)
    df["ANO"] = df["DATA GERACAO ALTERADA"].dt.year.astype("Int64").astype("string")

    df_debito = _montar_lancamentos(df, "debito")
    df_credito = _montar_lancamentos(df, "credito")

    df_final = pd.concat([df_credito, df_debito], ignore_index=True, copy=False)
    df_final = _adicionar_colunas_derivadas(df_final)
    df_final = _forcar_colunas_textuais(df_final, TEXT_COLUMNS_CONTA_GRAFICA)

    if salvar_parquet:
        df_final.to_parquet(conta_grafica_path.with_suffix(".parquet"), index=False)

    return df_final


def cria_quadro_1(conta_grafica: pd.DataFrame, pasta_saida: str | Path) -> pd.DataFrame:
    """Gera o quadro 1 com saldos iniciais por contrato e COSIF de apresentação."""
    resultados = []

    for cosif, descricao in COSIFS_QUADRO_1.items():
        recorte_df = (
            conta_grafica[conta_grafica["COSIF - Filtro"] == cosif]
            .sort_values(by=["Contador"])
            .groupby(["Num Contrato", "COSIF Apresentação"], as_index=False)
            .first()[["Num Contrato", "COSIF Apresentação", "Valor Debito"]]
            .rename(
                columns={"Valor Debito": "Valor", "COSIF Apresentação": "COSIF"}
            )
            .assign(Descrição=descricao)
        )
        resultados.append(recorte_df)

    if resultados:
        df_final = pd.concat(resultados, ignore_index=True)
    else:
        df_final = pd.DataFrame(
            columns=["Num Contrato", "COSIF", "Valor", "Descrição"]
        )
    df_final = df_final.sort_values(by=["Num Contrato", "COSIF"]).reset_index(drop=True)

    _salvar_quadro(df_final, pasta_saida, "quadro_1.csv")
    return df_final


def cria_quadro_2(conta_grafica: pd.DataFrame, pasta_saida: str | Path) -> pd.DataFrame:
    """Gera o quadro 2 com a superveniência/insuficiência anual e total por contrato."""
    quadro_2 = conta_grafica[conta_grafica["Conta"] == CONTA_QUADRO_2].copy()
    quadro_2 = quadro_2.rename(
        columns={"Valor Líquido": "Valor Contabilizado", "Num Contrato": "Contrato"}
    )

    quadro_2_group = (
        quadro_2.groupby(["Contrato", "Ano"], as_index=False)["Valor Contabilizado"].sum()
    )
    quadro_2_group["Ano"] = pd.to_numeric(quadro_2_group["Ano"], errors="coerce")
    quadro_2_group = quadro_2_group.dropna(subset=["Ano"]).assign(
        Ano=lambda df: df["Ano"].astype(int)
    )

    quadro_2_expandido = _expandir_por_contrato(quadro_2_group).round(2)
    quadro_2_expandido["Contrato"] = (
        pd.to_numeric(quadro_2_expandido["Contrato"], errors="coerce")
        .astype("Int64")
        .astype("string")
        .str.zfill(7)
    )
    quadro_2_expandido["COSIF"] = COSIF_QUADRO_2
    quadro_2_expandido["COSIF - Descrição"] = COSIF_QUADRO_2_DESCRICAO

    df_final = _adicionar_total_por_contrato(
        quadro_2_expandido, ["Valor Contabilizado"]
    )
    _salvar_quadro(df_final, pasta_saida, "quadro_2.csv")
    return df_final


def cria_quadro_3(
    conta_grafica: pd.DataFrame, pasta_saida: str | Path, contas_path: str | Path
) -> pd.DataFrame:
    """Gera o quadro 3 com receitas, despesas, descontos e LAIR por contrato/ano."""
    contas = pd.read_csv(contas_path, sep=";", dtype={"Conta": str})
    contas_filtro = contas["Conta"].drop_duplicates().to_list()

    quadro_3 = conta_grafica[conta_grafica["Conta"].isin(contas_filtro)].copy()
    quadro_3["COSIF"] = quadro_3["COSIF"].astype(str)
    quadro_3["Contrato"] = quadro_3["Num Contrato"].astype(str)
    quadro_3["Ano"] = pd.to_numeric(quadro_3["Ano"], errors="coerce")
    quadro_3 = quadro_3.dropna(subset=["Ano"]).assign(Ano=lambda df: df["Ano"].astype(int))

    quadro_3["Receita de Contraprestação - Inclui Superveniência(A)"] = quadro_3[
        "Valor Líquido"
    ].where(quadro_3["Conta"].str.startswith("7"), 0)
    quadro_3["Despesa de Depreciação - Inclui Insuficiência(B)"] = quadro_3[
        "Valor Líquido"
    ].where(
        quadro_3["Conta"].str.startswith("8")
        & ~quadro_3["COSIF"].str.contains(COSIF_DESCONTOS_CONCEDIDOS, na=False),
        0,
    )
    quadro_3["Descontos Concedidos(C)"] = quadro_3["Valor Líquido"].where(
        quadro_3["COSIF"].str.contains(COSIF_DESCONTOS_CONCEDIDOS, na=False),
        0,
    )
    quadro_3["LAIR"] = (
        quadro_3["Receita de Contraprestação - Inclui Superveniência(A)"]
        + quadro_3["Despesa de Depreciação - Inclui Insuficiência(B)"]
        + quadro_3["Descontos Concedidos(C)"]
    )

    value_columns = [
        "Receita de Contraprestação - Inclui Superveniência(A)",
        "Despesa de Depreciação - Inclui Insuficiência(B)",
        "Descontos Concedidos(C)",
        "LAIR",
    ]
    quadro_3_group = quadro_3.groupby(["Contrato", "Ano"], as_index=False)[
        value_columns
    ].sum()
    quadro_3_expandido = _expandir_por_contrato(quadro_3_group).round(2)

    df_final = _adicionar_total_por_contrato(quadro_3_expandido, value_columns)
    _salvar_quadro(df_final, pasta_saida, "quadro_3.csv")
    return df_final


def cria_quadro_4(
    quadro_2: pd.DataFrame, quadro_3: pd.DataFrame, pasta_saida: str | Path
) -> pd.DataFrame:
    """Gera o quadro 4 com a base de cálculo do IRPJ por contrato/ano."""
    quadro_2_anuais = _linhas_anuais(quadro_2)
    quadro_3_anuais = _linhas_anuais(quadro_3)

    quadro_4 = quadro_3_anuais.join(
        quadro_2_anuais.set_index(["Ano", "Contrato"]),
        on=["Ano", "Contrato"],
        rsuffix="_superveniencia",
        how="left",
    )
    quadro_4["Valor Contabilizado"] = quadro_4["Valor Contabilizado"].fillna(0.0)
    quadro_4["RESULTADO ANTES DA IRPJ"] = quadro_4["LAIR"]
    quadro_4["ADIÇÕES - Descontos Concedidos"] = (
        quadro_4["Descontos Concedidos(C)"] * -1
    )
    quadro_4[
        "ADIÇÕES/(Exclusões) - Superveniência/Insuficiência de Depreciação"
    ] = quadro_4["Valor Contabilizado"] * -1
    quadro_4["Base de Cálculo da IRPJ"] = quadro_4[
        [
            "RESULTADO ANTES DA IRPJ",
            "ADIÇÕES - Descontos Concedidos",
            "ADIÇÕES/(Exclusões) - Superveniência/Insuficiência de Depreciação",
        ]
    ].sum(axis=1)
    quadro_4 = quadro_4[
        [
            "Ano",
            "Contrato",
            "RESULTADO ANTES DA IRPJ",
            "ADIÇÕES - Descontos Concedidos",
            "ADIÇÕES/(Exclusões) - Superveniência/Insuficiência de Depreciação",
            "Base de Cálculo da IRPJ",
        ]
    ]

    value_columns = [
        "RESULTADO ANTES DA IRPJ",
        "ADIÇÕES - Descontos Concedidos",
        "ADIÇÕES/(Exclusões) - Superveniência/Insuficiência de Depreciação",
        "Base de Cálculo da IRPJ",
    ]
    df_final = _adicionar_total_por_contrato(quadro_4, value_columns)
    _salvar_quadro(df_final, pasta_saida, "quadro_4.csv")
    return df_final


def cria_quadro_5(quadro_3: pd.DataFrame, pasta_saida: str | Path) -> pd.DataFrame:
    """Gera o quadro 5 com a base de cálculo da CSLL por contrato/ano."""
    quadro_3_anuais = _linhas_anuais(quadro_3)

    quadro_5 = quadro_3_anuais.copy()
    quadro_5["RESULTADO ANTES DA CSLL"] = quadro_5["LAIR"]
    quadro_5["Adições - Descontos Concedidos"] = quadro_5["Descontos Concedidos(C)"] * -1
    quadro_5[
        "Adições/(Exclusões) - Superveniência/Insuficiência de Depreciação"
    ] = 0.0
    quadro_5["Base de Cálculo da CSLL"] = quadro_5[
        [
            "RESULTADO ANTES DA CSLL",
            "Adições - Descontos Concedidos",
            "Adições/(Exclusões) - Superveniência/Insuficiência de Depreciação",
        ]
    ].sum(axis=1)
    quadro_5 = quadro_5[
        [
            "Contrato",
            "Ano",
            "RESULTADO ANTES DA CSLL",
            "Adições - Descontos Concedidos",
            "Adições/(Exclusões) - Superveniência/Insuficiência de Depreciação",
            "Base de Cálculo da CSLL",
        ]
    ]

    value_columns = [
        "RESULTADO ANTES DA CSLL",
        "Adições - Descontos Concedidos",
        "Adições/(Exclusões) - Superveniência/Insuficiência de Depreciação",
        "Base de Cálculo da CSLL",
    ]
    quadro_5_expandido = _expandir_por_contrato(quadro_5).round(2)
    df_final = _adicionar_total_por_contrato(quadro_5_expandido, value_columns)
    _salvar_quadro(df_final, pasta_saida, "quadro_5.csv")
    return df_final


def cria_quadro_6(
    quadro_4: pd.DataFrame, quadro_5: pd.DataFrame, pasta_saida: str | Path
) -> pd.DataFrame:
    """Gera o quadro 6 com a conciliação entre bases de IRPJ e CSLL."""
    quadro_4_anuais = _linhas_anuais(quadro_4)
    quadro_5_anuais = _linhas_anuais(quadro_5)

    quadro_6 = quadro_5_anuais[["Ano", "Contrato", "Base de Cálculo da CSLL"]].join(
        quadro_4_anuais[
            [
                "Ano",
                "Contrato",
                "Base de Cálculo da IRPJ",
                "ADIÇÕES/(Exclusões) - Superveniência/Insuficiência de Depreciação",
            ]
        ].set_index(["Ano", "Contrato"]),
        on=["Ano", "Contrato"],
        how="inner",
    )
    quadro_6["DIFERENÇA"] = (
        quadro_6["Base de Cálculo da CSLL"] - quadro_6["Base de Cálculo da IRPJ"]
    )
    quadro_6["SUPERVENIÊNCIA/INSUFICIÊNCIA DE DEPRECIAÇÃO"] = (
        quadro_6["ADIÇÕES/(Exclusões) - Superveniência/Insuficiência de Depreciação"]
        * -1
    )
    quadro_6 = quadro_6.rename(
        columns={
            "Base de Cálculo da CSLL": "BASE DO CSLL",
            "Base de Cálculo da IRPJ": "BASE DO IRPJ",
        }
    )
    quadro_6 = quadro_6[
        [
            "Ano",
            "Contrato",
            "BASE DO IRPJ",
            "BASE DO CSLL",
            "SUPERVENIÊNCIA/INSUFICIÊNCIA DE DEPRECIAÇÃO",
            "DIFERENÇA",
        ]
    ]
    quadro_6_expandido = _expandir_por_contrato(quadro_6).round(2)

    value_columns = [
        "BASE DO IRPJ",
        "BASE DO CSLL",
        "SUPERVENIÊNCIA/INSUFICIÊNCIA DE DEPRECIAÇÃO",
        "DIFERENÇA",
    ]
    df_final = _adicionar_total_por_contrato(quadro_6_expandido, value_columns)
    _salvar_quadro(df_final, pasta_saida, "quadro_6.csv")
    return df_final
