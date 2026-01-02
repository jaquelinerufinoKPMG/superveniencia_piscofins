import pandas as pd
import numpy as np
import xlwings as xw
import re
import os
from itertools import product
from dotenv import find_dotenv
from environs import Env

env = Env()
env.read_env(find_dotenv())


class process_dashboard:
    def __init__(self, tax_col: str):
        self.tax_col = tax_col
        self.template_path = env("anexo_c_template_path")
        self.agrupamento_cols = env.list("anexo_c_agrupamento_cols")
        self.agrupamento_values = env.list("anexo_c_agrupamento_values")
        self.desc_col = env("anexo_c_agrupamento_desc_col")
        self.col_AnoMes = env("anexo_c_col_anoMes")
        self.descricao_total = env("anexo_c_desc_total")
        self.descricao_exclusao = env("anexo_c_desc_exclusao")
        self.descricao_deducao = env("anexo_c_desc_deducao")
        self.taxa_pis = env.float("anexo_c_taxa_pis")
        self.taxa_cofins = env.float("anexo_c_taxa_cofins")
        self.col_valor = env("anexo_c_desc_valor")
        self.col_ano = env("anexo_c_col_ano")
        pass

    def group_revenues(
        self, df: pd.DataFrame, revenue_filter: list, description: str
    ) -> pd.DataFrame:

        sub = df[df[self.tax_col].isin(revenue_filter)]

        sub = sub.groupby(self.agrupamento_cols, as_index=False)[
            self.agrupamento_values
        ].sum()

        sub[self.desc_col] = description

        return sub

    def replicate_years(
        self,
        df: pd.DataFrame,
        group_cols: list = None,
        value_cols: list = None,
        fill_value: float = 0,
    ) -> pd.DataFrame:
        # 1) extrai Ano de AnoMes
        df.loc[:, self.col_ano] = df[self.col_AnoMes] // 100

        # 2) checa parâmetros
        if group_cols is None or value_cols is None:
            raise ValueError(
                "Passe detalhes_cols, ex: ['Conta_Nome','Cosif_Nome'] "
                "e value_cols, ex: ['ValorDebito','ValorCredito','Movimentacao']"
            )

        # 3) intervalo de anos
        years = list(range(df[self.col_ano].min(), df[self.col_ano].max() + 1))

        # 4) pega todas as combinações únicas de PIS + detalhes
        groups = (
            df[[self.tax_col] + group_cols].drop_duplicates().reset_index(drop=True)
        )

        # 5) gera DataFrame de anos
        years_df = pd.DataFrame({self.col_ano: years})

        # 6) cross‑join combos × anos
        groups["key"] = 1
        years_df["key"] = 1
        full = groups.merge(years_df, on="key").drop("key", axis=1)

        # 7) agrega seu df original somando duplicatas
        df_agg = df.groupby(
            [self.tax_col] + group_cols + [self.col_ano], as_index=False
        )[value_cols].sum()

        # 8) faz o merge pra “expandir” os anos faltantes
        df_full = full.merge(
            df_agg, on=[self.tax_col] + group_cols + [self.col_ano], how="left"
        )

        # 9) preenche zeros (ou outro fill_value)
        for col in value_cols:
            df_full[col] = df_full[col].fillna(fill_value)

        return df_full

    def calcula_pis_cofins(self, df: pd.DataFrame) -> pd.DataFrame:
        # 1) soma por Ano e Descrição
        df_agg = df.groupby([self.col_ano, self.desc_col], as_index=False)[
            self.col_valor
        ].sum()

        # 2) pivot das descrições
        df_pivot = df_agg.pivot_table(
            index=[self.col_ano],
            columns=self.desc_col,
            values=self.col_valor,
            aggfunc="sum",
            fill_value=0,
        ).reset_index()

        # 3) base de cálculo
        df_pivot["Base de Cálculo"] = (
            df_pivot.get(self.descricao_total, 0)
            - df_pivot.get(self.descricao_exclusao, 0)
        ) + df_pivot.get(self.descricao_deducao, 0)

        # 4) cálculos de PIS e Cofins
        df_pivot["Cálculo da Contribuição para o PIS  - Alíquota 0,65%"] = (
            df_pivot["Base de Cálculo"] * self.taxa_pis
        )
        df_pivot["Cálculo da Cofins  - Alíquota 4%"] = (
            df_pivot["Base de Cálculo"] * self.taxa_cofins
        )

        # 5) melt para formato longo
        df_long = df_pivot.melt(
            id_vars=[self.col_ano, "Base de Cálculo"],
            value_vars=[
                "Cálculo da Contribuição para o PIS  - Alíquota 0,65%",
                "Cálculo da Cofins  - Alíquota 4%",
            ],
            var_name="Cosif_Nome",
            value_name=self.col_valor,
        )

        mask_pis = df_long["Cosif_Nome"].str.contains("PIS")
        df_long.loc[~mask_pis, "Base de Cálculo"] = np.nan

        # 6) colunas fixas para concatenar
        df_long["Conta_Nome"] = np.nan
        df_long["ValorDebito"] = np.nan
        df_long[self.desc_col] = "Base de Cálculo (01)-(02)-(03)"
        df_long = df_long.rename(columns={"Base de Cálculo": "ValorCredito"})

        # reordena pra ficar igual aos outros blocos
        final = df_long[
            [
                self.col_ano,
                "Conta_Nome",
                "Cosif_Nome",
                "ValorDebito",
                "ValorCredito",
                self.col_valor,
                self.desc_col,
            ]
        ].copy()

        return final

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