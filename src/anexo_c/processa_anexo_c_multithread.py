from __future__ import annotations

from pathlib import Path
import os
import argparse
import pandas as pd
from concurrent.futures import ProcessPoolExecutor, as_completed
from tqdm.auto import tqdm

from src.anexo_c.process_dashboard import process_dashboard


FINAL_CSV = Path(r"data\input\final.csv")
TEMPLATE_PATH = Path(r"docs\template_csll.xlsx")
OUTPUT_DIR = Path(r"data\output")

TAX_FILTERS = {
    "IRPJ": ["RAIR", "Exclusão", "Adição"],
    "CS": ["RAIR", "Exclusão", "Adição"],
}


def load_contracts_from_txt(path: Path) -> set[int]:
    """
    Lê um .txt com 1 contrato por linha.
    Ignora linhas vazias/invalidas.
    """
    if not path.exists():
        return set()

    out: set[int] = set()
    with path.open("r", encoding="utf-8") as f:
        for line in f:
            t = line.strip()
            if not t:
                continue
            # aceita "0001234" também
            if t.isdigit():
                out.add(int(t))
    return out


def shard_contracts(contracts: list[int], shard_index: int, shard_total: int) -> list[int]:
    """
    Divide a lista em 'shard_total' partes, retorna a parte 'shard_index' (0-based).
    Ex: shard_total=2, shard_index=0 => metade A; shard_index=1 => metade B.
    """
    if shard_total <= 1:
        return contracts
    if not (0 <= shard_index < shard_total):
        raise ValueError("shard_index precisa estar entre 0 e shard_total-1")

    return [c for i, c in enumerate(contracts) if (i % shard_total) == shard_index]


def build_one_excel(
    contrato: int,
    df: pd.DataFrame,
    template_path: str,
    output_dir: str,
) -> tuple[int, str]:
    """
    Roda em um processo separado. Cada processo cria seu próprio cls.
    """
    cls = process_dashboard(TAX_FILTERS, template_path)

    num_str = str(contrato).zfill(7)
    output_path = Path(output_dir) / f"C{num_str}.xlsx"

    cls.atualizar_template_pivot(
        template_path=template_path,
        output_path=str(output_path),
        df=df,
        contrato=contrato,
    )
    return contrato, str(output_path)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--processar", default="processar_kpmg.txt", help="TXT com contratos a processar")
    parser.add_argument("--nao-encontrados", default="numeros_extraidos.txt", help="TXT com contratos não encontrados")
    parser.add_argument("--max-workers", type=int, default=None, help="Nº de processos (default: auto)")
    parser.add_argument("--shard-index", type=int, default=0, help="Índice do shard (0-based)")
    parser.add_argument("--shard-total", type=int, default=1, help="Total de shards (ex: 2 para dividir com alguém)")
    args = parser.parse_args()

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    # 1) lê as listas (seu trecho)
    contratos_to_process = load_contracts_from_txt(Path(args.processar))
    nao_encontrados_3 = load_contracts_from_txt(Path(args.nao_encontrados))

    contratos = sorted(list(contratos_to_process - nao_encontrados_3))
    if not contratos:
        print("Nada para processar (contratos vazios após filtro).")
        return

    # 2) divide o trabalho (pra você e a Fefa)
    contratos = shard_contracts(contratos, args.shard_index, args.shard_total)

    print(f"Contratos nesta máquina: {len(contratos)} (shard {args.shard_index+1}/{args.shard_total})")

    # 3) lê o final.csv e filtra só os contratos desta rodada
    df_all = pd.read_csv(FINAL_CSV, dtype={"NumContrato": "int64"})
    df_all = df_all[df_all["NumContrato"].isin(set(contratos))].copy()

    # 4) agrupa por contrato (só os que realmente existem no CSV)
    groups = list(df_all.groupby("NumContrato", sort=False))
    total = len(groups)
    if total == 0:
        print("Nenhum desses contratos foi encontrado no final.csv.")
        return

    # 5) decide workers (conservador pra não matar o disco)
    cpu = os.cpu_count() or 4
    max_workers = args.max_workers
    if max_workers is None:
        max_workers = max(2, min(cpu - 1, 8))  # ajuste se seu SSD aguentar mais

    errors: list[str] = []

    with ProcessPoolExecutor(max_workers=max_workers) as ex:
        futures = [
            ex.submit(build_one_excel, int(contrato), df_contrato, str(TEMPLATE_PATH), str(OUTPUT_DIR))
            for contrato, df_contrato in groups
        ]

        for fut in tqdm(as_completed(futures), total=total, unit="xlsx"):
            try:
                contrato, outpath = fut.result()
            except Exception as e:
                errors.append(str(e))

    if errors:
        Path(f"excel_errors_shard{args.shard_index}.log").write_text("\n\n".join(errors), encoding="utf-8")
        print(f"⚠️ Terminei com {len(errors)} erros. Veja excel_errors_shard{args.shard_index}.log")
    else:
        print("✅ Todos os XLSX deste shard foram gerados com sucesso.")


if __name__ == "__main__":
    main()