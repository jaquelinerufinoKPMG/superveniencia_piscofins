from datetime import datetime
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch

import pandas as pd

from src.resumo.cria_quadros_irpj import (
    cria_quadro_2,
    cria_quadro_3,
    cria_quadro_4,
    cria_quadro_5,
    cria_quadro_6,
    preparar_conta_grafica,
)
from src.resumo.cria_quadros_pis import preparar_conta_grafica as preparar_conta_grafica_pis
from src.resumo.cria_quadros_pis import cria_quadro_2 as cria_quadro_2_pis


class CriaQuadrosIrpjTestCase(unittest.TestCase):
    def _write_csv(self, path: Path, rows: list[list[str]]) -> None:
        path.write_text(
            "\n".join(";".join("" if value is None else str(value) for value in row) for row in rows),
            encoding="utf-8",
        )

    def test_preparar_conta_grafica_aceita_coluna_extra_vazia_e_normaliza_campos(self) -> None:
        with TemporaryDirectory() as tmp_dir:
            csv_path = Path(tmp_dir) / "conta_grafica.csv"
            rows = [
                [
                    "99999999",
                    "1",
                    "1234567",
                    "1000",
                    "Conta Debito",
                    "10,50",
                    "10",
                    "1234567",
                    "71300010030016",
                    "Conta Credito",
                    "100,25",
                    "23210000000000",
                    "COSIF Debito",
                    "71210001000000",
                    "COSIF Credito",
                    "",
                ],
                [
                    "00000000",
                    "2",
                    "1234567",
                    "2000",
                    "Conta Debito 2",
                    "abc",
                    "20",
                    "1234567",
                    "7001",
                    "Conta Credito 2",
                    "",
                    "17110000000000",
                    "COSIF Debito 2",
                    "81999006000000",
                    "COSIF Credito 2",
                    "",
                ],
            ]
            self._write_csv(csv_path, rows)

            df = preparar_conta_grafica(
                csv_path,
                data_final=datetime(2024, 12, 31),
                salvar_parquet=False,
            )

            self.assertIsInstance(df, pd.DataFrame)
            self.assertEqual(4, len(df))
            self.assertFalse(csv_path.with_suffix(".parquet").exists())
            self.assertIn("Valor Líquido", df.columns)
            self.assertIn("COSIF Apresentação", df.columns)
            self.assertEqual("string", str(df["COSIF"].dtype))
            self.assertEqual(
                {"2024"},
                set(df["Ano"].dropna().unique()),
            )

            debito_1050 = df[df["Valor Debito"] == 10.50].iloc[0]
            self.assertEqual(-10.50, debito_1050["Valor Líquido"])
            self.assertEqual("23210000000000", debito_1050["COSIF"])
            self.assertEqual("2.3.2.10", debito_1050["COSIF Apresentação"])
            self.assertEqual("202412", debito_1050["AnoMes"])

            debito_invalido = df[df["Contador"] == "2"].sort_values("Tipo").iloc[0]
            self.assertTrue(pd.isna(debito_invalido["Valor Debito"]))

    def test_preparar_conta_grafica_ignora_coluna_extra_nao_vazia(self) -> None:
        with TemporaryDirectory() as tmp_dir:
            csv_path = Path(tmp_dir) / "conta_grafica.csv"
            rows = [[
                "20220101",
                "1",
                "1234567",
                "1000",
                "Conta Debito",
                "10,50",
                "10",
                "1234567",
                "71300010030016",
                "Conta Credito",
                "100,25",
                "23210000000000",
                "COSIF Debito",
                "71210001000000",
                "COSIF Credito",
                "EXTRA",
            ]]
            self._write_csv(csv_path, rows)

            df = preparar_conta_grafica(
                csv_path,
                data_final=datetime(2024, 12, 31),
                salvar_parquet=False,
            )

            self.assertEqual(2, len(df))
            self.assertEqual("23210000000000", df[df["Tipo"] == "D"].iloc[0]["COSIF"])

    def test_preparar_conta_grafica_pis_ignora_linha_field_e_coluna_extra(self) -> None:
        with TemporaryDirectory() as tmp_dir:
            csv_path = Path(tmp_dir) / "conta_grafica_pis.csv"
            rows = [
                [
                    "Field",
                    "metadata",
                    "metadata",
                    "metadata",
                    "metadata",
                    "metadata",
                    "metadata",
                    "metadata",
                    "metadata",
                    "metadata",
                    "metadata",
                    "metadata",
                    "metadata",
                    "metadata",
                    "metadata",
                    "metadata",
                ],
                [
                    "20220101",
                    "1",
                    "1234567",
                    "1000",
                    "Conta Debito",
                    "10,50",
                    "10",
                    "1234567",
                    "71300010030016",
                    "Conta Credito",
                    "100,25",
                    "17110000000000",
                    "COSIF Debito",
                    "71210001000000",
                    "COSIF Credito",
                    "EXTRA",
                ],
            ]
            self._write_csv(csv_path, rows)

            df = preparar_conta_grafica_pis(
                csv_path,
                data_final=datetime(2024, 12, 31),
                salvar_parquet=False,
            )

            self.assertEqual(2, len(df))
            self.assertEqual("string", str(df["COSIF"].dtype))
            self.assertEqual("17110000000000", df[df["Tipo"] == "D"].iloc[0]["COSIF"])

    def test_quadros_2_a_6_expandem_anos_e_geram_totais(self) -> None:
        conta_grafica = pd.DataFrame(
            [
                {
                    "Conta": "7001",
                    "COSIF": "70000000000000",
                    "Num Contrato": "1234567",
                    "Ano": "2022",
                    "Valor Líquido": 100.0,
                },
                {
                    "Conta": "8001",
                    "COSIF": "80000000000000",
                    "Num Contrato": "1234567",
                    "Ano": "2022",
                    "Valor Líquido": -30.0,
                },
                {
                    "Conta": "8002",
                    "COSIF": "81999006000000",
                    "Num Contrato": "1234567",
                    "Ano": "2022",
                    "Valor Líquido": -5.0,
                },
                {
                    "Conta": "7001",
                    "COSIF": "70000000000000",
                    "Num Contrato": "1234567",
                    "Ano": "2024",
                    "Valor Líquido": 50.0,
                },
                {
                    "Conta": "71300010030016",
                    "COSIF": "71210001000000",
                    "Num Contrato": "7654321",
                    "Ano": "2022",
                    "Valor Líquido": 40.0,
                },
            ]
        )

        with TemporaryDirectory() as tmp_dir:
            pasta_saida = Path(tmp_dir) / "saida"
            contas_path = Path(tmp_dir) / "contas.csv"
            contas_path.write_text("Conta\n7001\n8001\n8002\n", encoding="utf-8")

            quadro_2 = cria_quadro_2(conta_grafica, pasta_saida)
            quadro_3 = cria_quadro_3(conta_grafica, pasta_saida, contas_path)
            quadro_4 = cria_quadro_4(quadro_2, quadro_3, pasta_saida)
            quadro_5 = cria_quadro_5(quadro_3, pasta_saida)
            quadro_6 = cria_quadro_6(quadro_4, quadro_5, pasta_saida)

            self.assertTrue((pasta_saida / "quadro_2.csv").exists())
            self.assertTrue((pasta_saida / "quadro_6.csv").exists())

            anos_quadro_3 = quadro_3["Ano"].tolist()
            self.assertEqual([2022, 2023, 2024, "Total"], anos_quadro_3)

            linha_2022_q3 = quadro_3[quadro_3["Ano"] == 2022].iloc[0]
            self.assertEqual(
                100.0, linha_2022_q3["Receita de Contraprestação - Inclui Superveniência(A)"]
            )
            self.assertEqual(
                -30.0, linha_2022_q3["Despesa de Depreciação - Inclui Insuficiência(B)"]
            )
            self.assertEqual(-5.0, linha_2022_q3["Descontos Concedidos(C)"])
            self.assertEqual(65.0, linha_2022_q3["LAIR"])

            linha_2023_q3 = quadro_3[quadro_3["Ano"] == 2023].iloc[0]
            self.assertEqual(0.0, linha_2023_q3["LAIR"])

            linha_total_q2 = quadro_2[quadro_2["Ano"] == "Total"].iloc[0]
            self.assertEqual("7654321", linha_total_q2["Contrato"])
            self.assertEqual(40.0, linha_total_q2["Valor Contabilizado"])

            linha_2022_q4 = quadro_4[quadro_4["Ano"] == 2022].iloc[0]
            self.assertEqual(65.0, linha_2022_q4["RESULTADO ANTES DA IRPJ"])
            self.assertEqual(5.0, linha_2022_q4["ADIÇÕES - Descontos Concedidos"])
            self.assertEqual(
                0.0,
                linha_2022_q4[
                    "ADIÇÕES/(Exclusões) - Superveniência/Insuficiência de Depreciação"
                ],
            )
            self.assertEqual(70.0, linha_2022_q4["Base de Cálculo da IRPJ"])

            linha_2022_q5 = quadro_5[quadro_5["Ano"] == 2022].iloc[0]
            self.assertEqual(70.0, linha_2022_q5["Base de Cálculo da CSLL"])

            linha_2022_q6 = quadro_6[quadro_6["Ano"] == 2022].iloc[0]
            self.assertEqual(70.0, linha_2022_q6["BASE DO IRPJ"])
            self.assertEqual(70.0, linha_2022_q6["BASE DO CSLL"])
            self.assertEqual(0.0, linha_2022_q6["DIFERENÇA"])

    def test_cria_quadro_2_le_parquet_com_colunas_minimas(self) -> None:
        parquet_path = Path("conta_grafica.parquet")
        mock_df = pd.DataFrame(
            [
                {
                    "Conta": "71300010030016",
                    "Valor Líquido": 40.0,
                    "Num Contrato": "7654321",
                    "Ano": "2022",
                }
            ]
        )

        with TemporaryDirectory() as tmp_dir, patch(
            "src.resumo.cria_quadros_irpj.pd.read_parquet", return_value=mock_df
        ) as mock_read_parquet:
            cria_quadro_2(parquet_path, Path(tmp_dir))

            mock_read_parquet.assert_called_once_with(
                parquet_path,
                columns=["Conta", "Valor Líquido", "Num Contrato", "Ano"],
            )

    def test_cria_quadro_2_envia_contrato_ao_ler_parquet(self) -> None:
        parquet_path = Path("conta_grafica.parquet")
        mock_df = pd.DataFrame(
            [
                {
                    "Conta": "71300010030016",
                    "Valor Líquido": 40.0,
                    "Num Contrato": "7654321",
                    "Ano": "2022",
                }
            ]
        )

        with TemporaryDirectory() as tmp_dir, patch(
            "src.resumo.cria_quadros_irpj.pd.read_parquet", return_value=mock_df
        ) as mock_read_parquet:
            quadro_2 = cria_quadro_2(parquet_path, Path(tmp_dir), contrato="7654321")

            mock_read_parquet.assert_called_once_with(
                parquet_path,
                columns=["Conta", "Valor Líquido", "Num Contrato", "Ano"],
                filters=[("Num Contrato", "in", ["7654321"])],
            )
            self.assertEqual(["7654321", "7654321"], quadro_2["Contrato"].tolist())

    def test_cria_quadro_2_pis_filtra_contrato_ao_ler_dataframe(self) -> None:
        conta_grafica = pd.DataFrame(
            [
                {
                    "Conta": "71300010030016",
                    "Valor Líquido": 40.0,
                    "Num Contrato": "7654321",
                    "Ano": "2022",
                },
                {
                    "Conta": "71300010030016",
                    "Valor Líquido": 15.0,
                    "Num Contrato": "1234567",
                    "Ano": "2022",
                },
            ]
        )

        with TemporaryDirectory() as tmp_dir:
            quadro_2 = cria_quadro_2_pis(
                conta_grafica, Path(tmp_dir), contrato="7654321"
            )

            self.assertEqual(["7654321", "7654321"], quadro_2["Contrato"].tolist())
            self.assertEqual(40.0, quadro_2.iloc[0]["Valor Contabilizado"])


if __name__ == "__main__":
    unittest.main()
