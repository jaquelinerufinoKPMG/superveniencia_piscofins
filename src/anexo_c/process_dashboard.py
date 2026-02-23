import pandas as pd
import numpy as np
import xlwings as xw
import re
import os
from itertools import product

AGRUPAMENTO_COLS = ["Ano", "Conta_Nome", "Cosif_Nome"]
AGRUPAMENTO_VALUES = ["ValorDebito", "ValorCredito", "Movimentacao"]
DESC_COL = "Descrição"
ANOMES_COL = "AnoMes"
COL_VALOR = "Movimentacao"
ANO_COL = "Ano"


class process_dashboard:
    def __init__(self, tax_cols: dict[str, list[str]], template_path: str):
        self.template_path = template_path
        self.tax_cols = tax_cols
        pass

    def group_revenues(
        self,
        df: pd.DataFrame,
        tax_filters: dict[str, list],
        pivot_section: str,
        description: str,
        section_col: str = "Tributo",  # nome da coluna que identifica IRPJ/CS
        ) -> pd.DataFrame:
        """
        tax_filters:
        {"PIS": [...]}
        ou {"IRPJ": [...], "CS": [...]}

        pivot_section:
        "IRPJ" ou "CS" (carimbo para a pivot)

        description:
        texto da linha/bloco, ex: "Resultado antes do IR"
        """

        cols = list(tax_filters.keys())

        # valida colunas
        missing = [c for c in cols if c not in df.columns]
        if missing:
            raise KeyError(f"Colunas de filtro não existem no df: {missing}")

        # monta máscara
        if len(cols) == 1:
            col = cols[0]
            mask = df[col].isin(tax_filters[col])
        else:
            allowed_lists = [tax_filters[c] for c in cols]
            allowed_tuples = set(product(*allowed_lists))
            mask = df[cols].apply(tuple, axis=1).isin(allowed_tuples)

        sub = df.loc[mask].copy()

        # carimba seção e descrição
        sub[section_col] = pivot_section + " " + description
        sub[DESC_COL] = description

        # agrupa (IMPORTANTE: inclui section_col no groupby pra não misturar IRPJ/CS)
        group_keys = list(AGRUPAMENTO_COLS) + [section_col, DESC_COL]
        sub = sub.groupby(group_keys, as_index=False)[AGRUPAMENTO_VALUES].sum()

        return sub
    
    def replicate_years(
        self,
        df: pd.DataFrame,
        fill_value: float = 0,
        tax_cols: list[str] | str | None = None,
    ) -> pd.DataFrame:

        VALUE_COLS = ["ValorDebito", "ValorCredito", "Movimentacao"]
        GROUP_COLS = ["Conta_Nome", "Cosif_Nome"]

        df = df.copy()
        df.loc[:, ANO_COL] = df[ANOMES_COL] // 100

        # resolve tax_cols
        if tax_cols is None:
            tax_cols = getattr(self, "tax_cols", None) or getattr(self, "tax_col", None)
            if tax_cols is None:
                raise ValueError("Informe tax_cols (ex: ['IRPJ','CS']) ou defina tax_cols/tax_col.")
        if isinstance(tax_cols, str):
            tax_cols = [tax_cols]

        key_cols = tax_cols + GROUP_COLS

        # agrega original (inclui 9999)
        df_agg = df.groupby(key_cols + [ANO_COL], as_index=False)[VALUE_COLS].sum()

        # separa 9999 (fica no output, mas não entra na replicação)
        df_9999 = df_agg[df_agg[ANO_COL] == 9999].copy()
        df_real = df_agg[df_agg[ANO_COL] != 9999].copy()

        # se não tem anos reais, não tem o que "replicar"; devolve só o 9999
        if df_real.empty:
            return df_9999.reset_index(drop=True)

        # intervalo de anos reais (sem 9999)
        min_year = int(df_real[ANO_COL].min())
        max_year = int(df_real[ANO_COL].max())
        years = list(range(min_year, max_year + 1))

        # combos baseadas nos grupos que existem nos anos reais
        # (assim a gente não cria “linhas novas” só porque existe 9999)
        groups = df_real[key_cols].drop_duplicates().reset_index(drop=True)

        years_df = pd.DataFrame({ANO_COL: years})

        # cross join grupos × anos reais
        groups["key"] = 1
        years_df["key"] = 1
        full = groups.merge(years_df, on="key").drop("key", axis=1)

        # expande anos faltantes
        df_full_real = full.merge(df_real, on=key_cols + [ANO_COL], how="left")
        df_full_real[VALUE_COLS] = df_full_real[VALUE_COLS].fillna(fill_value)

        # junta de volta o 9999 (sem replicar)
        out = pd.concat([df_full_real, df_9999], ignore_index=True)

        return out

    def calcula_pis_cofins(self, df: pd.DataFrame) -> pd.DataFrame:

        DESCRICAO_TOTAL = "(01) Total das Receitas"
        DESCRICAO_EXCLUSAO = "(02) Exclusão"
        DESCRICAO_DEDUCAO = "(03) Dedução"
        TAXA_PIS = 0.0065
        TAXA_COFINS = 0.04

        # 1) soma por Ano e Descrição
        df_agg = df.groupby([ANO_COL, DESC_COL], as_index=False)[COL_VALOR].sum()

        # 2) pivot das descrições
        df_pivot = df_agg.pivot_table(
            index=[ANO_COL],
            columns=DESC_COL,
            values=COL_VALOR,
            aggfunc="sum",
            fill_value=0,
        ).reset_index()

        # 3) base de cálculo
        df_pivot["Base de Cálculo"] = (
            df_pivot.get(DESCRICAO_TOTAL, 0) - df_pivot.get(DESCRICAO_EXCLUSAO, 0)
        ) + df_pivot.get(DESCRICAO_DEDUCAO, 0)

        # 4) cálculos de PIS e Cofins
        df_pivot["Cálculo da Contribuição para o PIS  - Alíquota 0,65%"] = (
            df_pivot["Base de Cálculo"] * TAXA_PIS
        )
        df_pivot["Cálculo da Cofins  - Alíquota 4%"] = (
            df_pivot["Base de Cálculo"] * TAXA_COFINS
        )

        # 5) melt para formato longo
        df_long = df_pivot.melt(
            id_vars=[ANO_COL, "Base de Cálculo"],
            value_vars=[
                "Cálculo da Contribuição para o PIS  - Alíquota 0,65%",
                "Cálculo da Cofins  - Alíquota 4%",
            ],
            var_name="Cosif_Nome",
            value_name=COL_VALOR,
        )

        mask_pis = df_long["Cosif_Nome"].str.contains("PIS")
        df_long.loc[~mask_pis, "Base de Cálculo"] = np.nan

        # 6) colunas fixas para concatenar
        df_long["Conta_Nome"] = np.nan
        df_long["ValorDebito"] = np.nan
        df_long[DESC_COL] = "Base de Cálculo (01)-(02)-(03)"
        df_long = df_long.rename(columns={"Base de Cálculo": "ValorCredito"})

        # reordena pra ficar igual aos outros blocos
        final = df_long[
            [
                ANO_COL,
                "Conta_Nome",
                "Cosif_Nome",
                "ValorDebito",
                "ValorCredito",
                COL_VALOR,
                DESC_COL,
            ]
        ].copy()

        return final

    def calcula_csll(self, df: pd.DataFrame) -> pd.DataFrame:
        # máscara das linhas que precisam ser duplicadas
        mask = df[DESC_COL].astype(str).str.contains("Resultado antes do", na=False)

        # copia só essas linhas
        df_dup = df.loc[mask].copy()

        # altera a coluna Tributo na cópia
        df_dup["Tributo"] = df_dup["Tributo"].astype(str) + " - Total"

        # concatena de volta
        df = pd.concat([df, df_dup], ignore_index=True)

        mask_remover = (
            df[DESC_COL].astype(str).str.contains("Resultado antes do", na=False)
            & df["Conta_Nome"].astype(str).str.startswith("8", na=False)
            & ~df["Tributo"].str.contains(" - Total", na=False)
        )

        df = df[~mask_remover].copy()

        df_agg = df.groupby([ANO_COL, "Tributo"], as_index=False)[COL_VALOR].sum()

        df_pivot = df_agg.pivot_table(
            index=[ANO_COL],
            columns="Tributo",
            values=COL_VALOR,
            aggfunc="sum",
            fill_value=0,
        ).reset_index()

        total_row = df_pivot.drop(columns=[ANO_COL]).sum(numeric_only=True)
        total_row[ANO_COL] = "Grand Total"

        # 2) adiciona no final
        df_pivot = pd.concat([df_pivot, total_row.to_frame().T], ignore_index=True)

        df_pivot["LALUR"] = -df_pivot.get("IRPJ Adição", 0) - df_pivot.get(
            "IRPJ Exclusão", 0
        )
        df_pivot["Base de cálculo - IR"] = df_pivot.get(
            "IRPJ Resultado antes do IR - Total", 0
        ) + df_pivot.get("LALUR", 0)
        df_pivot["LACS"] = -df_pivot.get("CSLL Adição", 0) - df_pivot.get(
            "CSLL Exclusão", 0
        )
        df_pivot["Base de cálculo - CSLL"] = df_pivot.get(
            "CSLL Resultado antes do CSLL - Total", 0
        ) + df_pivot.get("LACS", 0)
        df_pivot["Diferença na Base de Cálculo entre IR e CS"] = df_pivot.get(
            "Base de cálculo - IR", 0
        ) + df_pivot.get("Base de cálculo - CSLL", 0)
        cols = df.columns
        cols = cols.drop(["Ano", "Tributo", "Movimentacao"]).tolist()

        df_long = df_pivot.melt(
            id_vars=[ANO_COL],
            value_vars=[
                "LALUR",
                "Base de cálculo - IR",
                "LACS",
                "Base de cálculo - CSLL",
                "Diferença na Base de Cálculo entre IR e CS",
            ],
            var_name="Tributo",
            value_name=COL_VALOR,
        )

        df_long[cols] = np.nan

        df_final = pd.concat([df, df_long])
        df_final["Descrição"] = np.where(
            df_final["Descrição"].isna(), df_final["Tributo"], df_final["Descrição"]
        )

        return df_final

    def x(self, file_path, file_name):
        pattern = re.compile(r"\d+")

        contracts = []

        for nome_arquivo in os.listdir(file_path):
            match = pattern.findall(nome_arquivo)
            if match:
                contracts.append("".join(match))

        with open(file_name, "w", encoding="utf-8") as f:
            for seq in contracts:
                f.write(seq + "\n")

    def atualizar_template_pivot(
        self,
        template_path: str,
        output_path: str,
        df,
        contrato,
        data_sheet: str = "Dados",
        pivot_sheet: str = "Pivot",
        table_index: int = 1,
        contract_cell: str = "C2",
        start_row_hide: int = 8,
        new_pivot_name: str = "PIS_COFINS_ANUAL"
    ):
        """
        Abre o template de Excel com PivotTables, injeta o DataFrame na Tabela,
        escreve o valor de contrato na célula especificada, dá Refresh em todas
        as PivotTables, oculta as linhas em branco (a partir de start_row_hide)
        mantendo as já ocultas e deixando uma linha em branco após cada PivotTable,
        esconde a aba de dados e renomeia a aba de pivô para 'Pivot_<contrato>',
        e salva em output_path.
        """
        app = xw.App(visible=False)
        try:
            wb = app.books.open(template_path)

            # 1) Atualiza a tabela de dados
            ws_data = wb.sheets[data_sheet]
            tbl = ws_data.api.ListObjects(table_index)
            hdr_row = tbl.HeaderRowRange.Row
            hdr_col = tbl.HeaderRowRange.Column

            # Limpa linhas antigas
            if tbl.DataBodyRange is not None:
                tbl.DataBodyRange.Clear()

            # Escreve só os valores (sem header e sem índice)
            start_row = hdr_row + 1
            start_col = hdr_col
            ws_data.range((start_row, start_col)).options(
                index=False, header=False
            ).value = df

            # Redimensiona a tabela
            last_data_row = start_row + df.shape[0] - 1
            last_col = hdr_col + tbl.HeaderRowRange.Columns.Count - 1
            new_range = ws_data.api.Range(
                ws_data.api.Cells(hdr_row, hdr_col),
                ws_data.api.Cells(last_data_row, last_col),
            )
            tbl.Resize(new_range)

            # 2) Escreve o valor do contrato e renomeia a aba de pivô
            ws_pivot = wb.sheets[pivot_sheet]
            ws_pivot.range(contract_cell).value = contrato
            new_pivot_name = "PIS_COFINS_ANUAL"
            ws_pivot.name = new_pivot_name

            # 3) Oculta a aba de dados
            ws_data.visible = False

            # 4) Refresh em todas as PivotTables
            wb.api.RefreshAll()

            # 5) Oculta linhas em branco na aba de pivôs
            used = ws_pivot.api.UsedRange
            first_row = start_row_hide
            last_row = used.Row + used.Rows.Count - 1
            first_col = used.Column
            last_col = used.Column + used.Columns.Count - 1

            hidden_initial = {
                r for r in range(first_row, last_row + 1) if ws_pivot.api.Rows(r).Hidden
            }

            separator_rows = set()
            pt_count = ws_pivot.api.PivotTables().Count
            for i in range(1, pt_count + 1):
                pt = ws_pivot.api.PivotTables().Item(i)
                rng = pt.TableRange1
                sep = rng.Row + rng.Rows.Count
                if first_row <= sep <= last_row:
                    separator_rows.add(sep)

            for r in range(first_row, last_row + 1):
                if r in hidden_initial:
                    continue
                if r in separator_rows:
                    ws_pivot.api.Rows(r).Hidden = False
                    continue
                vals = ws_pivot.range((r, first_col), (r, last_col)).value
                if all(v is None or v == "" for v in vals):
                    ws_pivot.api.Rows(r).Hidden = True
                else:
                    ws_pivot.api.Rows(r).Hidden = False

            # 6) Salva o resultado
            wb.save(output_path)
        finally:
            try:
                wb.close()
            except:
                pass
            app.quit()

    def get_contract_numbers(
        self, folder: str, file_name: str = "numeros_extraidos.txt"
    ):
        pattern = re.compile(r"\d+")
        contracts = []
        for file in os.listdir(folder):
            match = pattern.findall(file)  # e.g. ['123','456']
            if match:
                contracts.append("".join(match))  # => '123456'
        # Grava no TXT, uma sequência por linha
        with open(file_name, "w", encoding="utf-8") as f:
            for seq in contracts:
                f.write(seq + "\n")
        print(f"✅ Extraídos {len(contracts)} itens e salvos em '{file_name}'")
