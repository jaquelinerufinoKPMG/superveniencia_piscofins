import os
import shutil
import threading
from pathlib import Path

import pandas as pd
from PyPDF2 import PdfReader
from tqdm import tqdm

import win32com.client as win32
import win32com

from dotenv import load_dotenv


# ================== CONFIG ==================
load_dotenv()

OUTPUT_DIRECTORY = os.getenv("output_dir")
INVALID_DIRECTORY = os.getenv("invalid_dir")
WORKSHEET_NAME = os.getenv("anexo_c_worksheet_name", "PIS_COFINS_ANUAL")
THREAD_LIMIT = 30
# ============================================


class export_pdf:
    def __init__(self):
        pass

    # ---------- VALIDAÇÃO DOS EXCEL ----------
    def _check_excel_files(self, excel_file, invalid_directory):
        p = Path(excel_file)

        if not p.is_file():
            return

        if p.suffix.lower() not in {".xlsx", ".xls", ".xlsm"}:
            return

        file_name = p.stem[1:].lstrip("0")

        try:
            with pd.ExcelFile(p, engine="openpyxl") as xls:
                df = pd.read_excel(
                    xls,
                    header=None,
                    dtype=str,
                    sheet_name=WORKSHEET_NAME
                )
        except Exception:
            Path(invalid_directory).mkdir(parents=True, exist_ok=True)
            shutil.move(str(p), str(Path(invalid_directory) / p.name))
            return

        try:
            if str(df.iloc[1, 2]) != file_name and str(df.iloc[0, 2]) != file_name:
                Path(invalid_directory).mkdir(parents=True, exist_ok=True)
                shutil.move(str(p), str(Path(invalid_directory) / p.name))
        except Exception:
            Path(invalid_directory).mkdir(parents=True, exist_ok=True)
            shutil.move(str(p), str(Path(invalid_directory) / p.name))

    # ---------- RENOMEIA ----------
    def _rename_files(self, directory_path):
        for file in os.listdir(directory_path):
            if "." not in file:
                continue

            name, ext = file.split(".", 1)
            if not name:
                continue

            letter = name[0]
            numbers = name[1:].zfill(7)
            new_name = f"{letter}{numbers}.{ext}"

            try:
                os.rename(
                    os.path.join(directory_path, file),
                    os.path.join(directory_path, new_name)
                )
            except Exception:
                pass

    # ---------- PDF JÁ EXISTE ----------
    def _check_pdf(self, input_file):
        output_file = os.path.splitext(input_file)[0] + ".pdf"
        if not os.path.isfile(output_file):
            return False

        try:
            with open(output_file, "rb") as f:
                return len(PdfReader(f).pages) == 1
        except Exception:
            return False

    # ---------- EXPORTA PDF (SEM MEXER EM VALORES) ----------
    def _export_pdf(self, input_file):
        excel = win32.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        # garante separadores do sistema (pt-BR)
        excel.UseSystemSeparators = True

        output_file = os.path.splitext(input_file)[0] + ".pdf"
        wb = None

        try:
            wb = excel.Workbooks.Open(os.path.abspath(input_file))
            ws = wb.Worksheets(WORKSHEET_NAME)

            last_column = ws.UsedRange.Columns.Count
            last_row = 1

            # última linha relevante (pela borda)
            for row in range(1, ws.UsedRange.Rows.Count + 1):
                borders = ws.Cells(row, 4).Borders
                for i in range(5, 13):
                    if borders(i).Color != 0.0 or borders(i).LineStyle != -4142:
                        last_row = row
                        break

            # última coluna relevante (pela borda)
            for col in range(1, ws.UsedRange.Columns.Count + 1):
                borders = ws.Cells(5, col).Borders
                for i in range(5, 13):
                    if borders(i).Color != 0.0 or borders(i).LineStyle != -4142:
                        last_column = col
                        break

            def col_letra(i):
                s = ""
                while i:
                    i, r = divmod(i - 1, 26)
                    s = chr(65 + r) + s
                return s

            rng = ws.Range(f"B1:{col_letra(last_column)}{last_row}")

            # ===== ÚNICO FIX APLICADO =====
            # NÃO altera valor, formato ou casas decimais
            rng.EntireColumn.AutoFit()
            # ==============================

            ws.PageSetup.PrintArea = rng.Address
            ws.DisplayPageBreaks = False
            ws.ResetAllPageBreaks()

            ps = ws.PageSetup
            ps.Zoom = False
            ps.FitToPagesWide = 1
            ps.FitToPagesTall = 1

            ps.Orientation = 1  # Portrait (como estava)
            ps.PaperSize = 9    # A4

            ps.LeftMargin = excel.InchesToPoints(0.15)
            ps.RightMargin = excel.InchesToPoints(0.15)
            ps.TopMargin = excel.InchesToPoints(0.15)
            ps.BottomMargin = excel.InchesToPoints(0.15)
            ps.CenterHorizontally = True

            ws.ExportAsFixedFormat(
                Type=0,
                Filename=os.path.abspath(output_file),
                IgnorePrintAreas=False,
                OpenAfterPublish=False
            )

        finally:
            if wb:
                wb.Close(False)
            excel.Quit()

    # ---------- PROCESSO PRINCIPAL ----------
    def process_file(self):
        try:
            shutil.rmtree(win32com.__gen_path__, ignore_errors=True)
        except Exception:
            pass

        os.makedirs(INVALID_DIRECTORY, exist_ok=True)

        arquivos = os.listdir(OUTPUT_DIRECTORY)
        threads = []

        print("Verificando arquivos inválidos...")
        for idx, arquivo in enumerate(tqdm(arquivos)):
            if arquivo.endswith(".pdf"):
                continue

            path = os.path.join(OUTPUT_DIRECTORY, arquivo)

            t = threading.Thread(
                target=self._check_excel_files,
                args=(path, INVALID_DIRECTORY)
            )
            t.start()
            threads.append(t)

            if len(threads) >= THREAD_LIMIT or idx == len(arquivos) - 1:
                for t in threads:
                    t.join()
                threads = []

        print("Renomeando arquivos...")
        self._rename_files(OUTPUT_DIRECTORY)

        print("Convertendo para PDF...")
        arquivos = sorted([
            f for f in os.listdir(OUTPUT_DIRECTORY)
            if f.endswith(".xlsx")
            and not self._check_pdf(os.path.join(OUTPUT_DIRECTORY, f))
        ])

        for arquivo in tqdm(arquivos):
            try:
                self._export_pdf(os.path.join(OUTPUT_DIRECTORY, arquivo))
            except Exception as e:
                print(f"Erro ao converter {arquivo}: {e}")

        print("✅ Processo finalizado.")


if __name__ == "__main__":
    export_pdf().process_file()
