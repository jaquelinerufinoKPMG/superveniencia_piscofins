from pathlib import Path
import pandas as pd


def processar_anexo_b(
    caminho_csv: str | Path,
    pasta_saida: str | Path,
    coluna_contrato: str = "NUM CONTRATO DEBITO",
    sep: str = ";",
    encoding: str = "utf-8",
) -> pd.DataFrame:

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

    caminho_csv = Path(caminho_csv)
    pasta_saida = Path(pasta_saida)
    pasta_saida.mkdir(parents=True, exist_ok=True)

    df = pd.read_csv(caminho_csv, sep=sep, encoding=encoding, dtype=str)

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

    exportar_contratos_para_excel(df, coluna_contrato, pasta_saida)


def exportar_contratos_para_excel(
    df: pd.DataFrame, coluna_contrato: str, pasta_saida: str | Path
) -> None:
    """
    Lê um CSV e gera um arquivo Excel por contrato.

    Parâmetros:
        caminho_csv: caminho do arquivo CSV de entrada
        pasta_saida: pasta onde os arquivos Excel serão salvos
        coluna_contrato: nome da coluna que identifica o contrato
        sep: separador do CSV
        encoding: encoding do CSV
    """
    pasta_saida = Path(pasta_saida)
    pasta_saida.mkdir(parents=True, exist_ok=True)

    df = df[df[coluna_contrato].notna()].copy()

    # Converte contrato para string e limpa espaços
    df[coluna_contrato] = df[coluna_contrato].astype(str).str.strip()

    print(f"Total de linhas: {len(df):,}")
    print(f"Total de contratos: {df[coluna_contrato].nunique():,}")

    for contrato, grupo in df.groupby(coluna_contrato, dropna=False):
        nome_arquivo = f"B{contrato}.xlsx"
        caminho_saida = pasta_saida / nome_arquivo

        grupo.to_excel(caminho_saida, index=False, sheet_name="Conta Gráfica")
        print(f"Salvo: {caminho_saida}")
