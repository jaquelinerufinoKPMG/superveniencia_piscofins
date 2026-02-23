from __future__ import annotations

from pathlib import Path
import os
import math
import pandas as pd
from concurrent.futures import ProcessPoolExecutor, as_completed
from tqdm.auto import tqdm

from src.anexo_c.process_dashboard import process_dashboard


FINAL_CSV = Path(r"input\final.csv")                 # CSV consolidado gerado antes
TEMPLATE_PATH = Path(r"docs\template_csll.xlsx")
OUTPUT_DIR = Path(r"data\output")
INVALID_DIR = Path(r"data\invalid")

TAX_FILTERS = {  # não deve ser usado nessa etapa, mas o construtor pode exigir
    "IRPJ": ["RAIR", "Exclusão", "Adição"],
    "CS": ["RAIR", "Exclusão", "Adição"],
}


def _safe_int(x) -> int | None:
    try:
        if pd.isna(x):
            return None
        return int(x)
    except Exception:
        return None


def build_one_excel(
    contrato: int,
    df: pd.DataFrame,
    template_path: str,
    output_dir: str,
) -> tuple[int, str]:
    """
    Roda em um processo separado.
    Cada processo cria sua própria instância do process_dashboard.
    """
    cls = process_dashboard(TAX_FILTERS, template_path)

    num_str = str(contrato).zfill(7)
    output_path = Path(output_dir) / f"C{num_str}.xlsx"

    # se sua função precisar do contrato em coluna, garante aqui
    # (não atrapalha se já existir)
    if "NumContrato" not in df.columns:
        df = df.copy()
        df["NumContrato"] = contrato

    cls.atualizar_template_pivot(
        template_path=template_path,
        output_path=str(output_path),
        df=df,
        contrato=contrato,
        new_pivot_name="IR_CS_ANUAL",
    )
    return contrato, str(output_path)


def main():
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    # 1) lê o final.csv (use dtype pra não virar string)
    df_all = pd.read_csv(FINAL_CSV, dtype={"NumContrato": "int64"})

    # 2) limpa contratos inválidos (paranoia saudável)
    df_all["NumContrato"] = df_all["NumContrato"].map(_safe_int)
    df_all = df_all.dropna(subset=["NumContrato"])
    df_all["NumContrato"] = df_all["NumContrato"].astype("int64")

    # 3) agrupa por contrato
    groups = list(df_all.groupby("NumContrato", sort=False))
    total = len(groups)

    # 4) ajusta nº de workers (pra não matar o disco)
    # regra prática: CPU/2 até CPU-1 costuma ser o melhor quando tem I/O
    cpu = os.cpu_count() or 4
    max_workers = max(2, min(cpu - 1, 8))  # teto 8 pra não virar triturador de HD
    # se você tiver SSD forte, pode subir esse teto pra 12/16

    errors = []

    with ProcessPoolExecutor(max_workers=max_workers) as ex:
        futures = [
            ex.submit(
                build_one_excel,
                int(contrato),
                df_contrato,
                str(TEMPLATE_PATH),
                str(OUTPUT_DIR),
            )
            for contrato, df_contrato in groups
        ]

        for fut in tqdm(as_completed(futures), total=total, unit="xlsx"):
            try:
                contrato, outpath = fut.result()
            except Exception as e:
                errors.append(str(e))

    if errors:
        # salva um logzinho simples
        Path(fr"{INVALID_DIR}\excel_errors.log").write_text("\n\n".join(errors), encoding="utf-8")
        print(f"⚠️ Terminei com {len(errors)} erros. Veja excel_errors.log")
    else:
        print("✅ Todos os XLSX foram gerados com sucesso.")


if __name__ == "__main__":
    main() 