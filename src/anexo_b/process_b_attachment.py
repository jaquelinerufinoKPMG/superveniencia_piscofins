from src.anexo_b.normalize_text import *
import glob
import io
import numpy as np
import os
import pandas as pd
import PyPDF2
import re
import shutil
import sys
import traceback
import threading
import warnings
import win32com.client as win32

from openpyxl import load_workbook, Workbook
from openpyxl.styles import Alignment, Border, Font, NamedStyle, PatternFill, Side
from openpyxl.utils import get_column_letter, range_boundaries
from openpyxl.cell.cell import MergedCell
from pandas import Timestamp
from pathlib import Path
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from tqdm import tqdm


class anexo_b:
    def __init__(self):
        pass

    def create_excel_model(
        self,
        file_name: str = None,
        header_color: str = "FFFF00",
        header_text: str = "Cabeçalho 1",
        header_list: list = None,
        alignments: list = None,
        fill_colors: list = None,
        merge_columns: list = None,
        font: str = "Calibri",
        font_size: int = 10,
    ) -> Workbook:
        # Garantir que os cabeçalhos foram fornecidos
        if header_list is None:
            raise ValueError("A lista de nomes para o cabeçalho 2 deve ser fornecida.")
        if fill_colors is None:
            fill_colors = ["FFFFFF"] * len(header_list)  # Branco padrão
        if alignments is None:
            alignments = ["center"] * len(header_list)  # Alinhamento central padrão
        if merge_columns is None:
            merge_columns = [1] * len(
                header_list
            )  # Cada cabeçalho ocupa uma célula por padrão

        # Validar tamanhos das listas
        if not (
            len(header_list)
            == len(fill_colors)
            == len(alignments)
            == len(merge_columns)
        ):
            raise ValueError(
                "As listas de alinhamentos, cores de preenchimento e colunas_merge devem ter o mesmo comprimento do cabeçalho 2."
            )

        # Criar um novo workbook e planilha
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.title = "Modelo"

        # Configurar Cabeçalho 1 (linha 1)
        worksheet.merge_cells(
            start_row=1, start_column=1, end_row=1, end_column=sum(merge_columns)
        )
        header_cell_1 = worksheet.cell(row=1, column=1, value=header_text)
        header_cell_1.fill = PatternFill(
            start_color=header_color, end_color=header_color, fill_type="solid"
        )
        header_cell_1.font = Font(
            bold=True, name=font, size=font_size, color="FFFFFF"
        )  # Texto branco obrigatório
        header_cell_1.alignment = Alignment(horizontal="center", vertical="center")

        # Linha 2 (linha oculta)
        worksheet.row_dimensions[2].hidden = True  # Ocultar a linha 2

        # Configurar Cabeçalho 2 (linha 3)
        col_idx = 1
        for name, fill_color, alignment, merge in zip(
            header_list, fill_colors, alignments, merge_columns
        ):
            # Mesclar as colunas, se necessário
            if merge > 1:
                worksheet.merge_cells(
                    start_row=3,
                    start_column=col_idx,
                    end_row=3,
                    end_column=col_idx + merge - 1,
                )

            cell = worksheet.cell(row=3, column=col_idx, value=name)
            cell.font = Font(
                color="FFFFFF", bold=True, name=font, size=font_size
            )  # Texto branco obrigatório
            cell.fill = PatternFill(
                start_color=fill_color, end_color=fill_color, fill_type="solid"
            )
            cell.alignment = Alignment(horizontal=alignment, vertical="center")

            # Ajustar largura das colunas mescladas
            for i in range(merge):
                worksheet.column_dimensions[get_column_letter(col_idx + i)].width = max(
                    len(name) + 2, 10
                )

            col_idx += merge  # Avançar para a próxima célula disponível

        # Salvar o arquivo
        return workbook

    def fill_excel_data(
        self,
        workbook: Workbook,
        worksheet_name: str,
        df: pd.DataFrame,
        text_colors: list = None,
        bold: list = None,
        cell_borders: list = None,
        cell_alignments: list = None,
        columns_to_merge: list = None,
        column_types: list = None,
    ):

        # Obter a planilha ativa pelo nome
        if worksheet_name not in workbook.sheetnames:
            raise ValueError(f"A planilha '{worksheet_name}' não existe no workbook.")
        worksheet = workbook[worksheet_name]

        # Mapear os cabeçalhos 2 com base na linha 3 (após a linha oculta)
        header_2 = [
            cell.value for cell in worksheet[3]
        ]  # Os valores da linha 3 são os cabeçalhos 2
        col_map = {
            nome: idx + 1 for idx, nome in enumerate(header_2)
        }  # Mapear nome para índice de coluna

        # Obter as informações de mesclagem do cabeçalho
        merge_map = {}
        for merged_range in worksheet.merged_cells.ranges:
            if merged_range.min_row == 3:  # Somente os merges do cabeçalho 2
                merge_map[merged_range.min_col] = merged_range.max_col

        # Validar se todos os nomes do DataFrame estão no cabeçalho 2
        for column in df.columns:
            if column not in col_map:
                raise ValueError(
                    f"A coluna '{column}' no DataFrame não está no cabeçalho 2 do template."
                )

        # Preencher os dados com formatação
        for row_idx, row_data in enumerate(
            df.itertuples(index=False), start=4
        ):  # Dados começam na linha 4

            total_row = False  # Flag para verificar se a linha contém "Total"
            for col_idx, col_name in enumerate(df.columns):
                cell_idx = col_map[col_name]
                value = row_data[col_idx]
                cell = worksheet.cell(row=row_idx, column=cell_idx)

                # Verificar se a célula contém "Total" (independente de maiúsculas/minúsculas)
                if isinstance(value, str) and "total" == value.lower():
                    total_row = True

                # Determinar o tipo e formatar o valor
                if column_types and len(column_types) > col_idx:
                    data_type = column_types[col_idx].lower()
                    if data_type == "número" and isinstance(value, (int, float)):
                        cell.value = value
                        cell.number_format = "#,##0.00"
                    elif data_type == "moeda" and isinstance(value, (int, float)):
                        cell.value = value
                        cell.number_format = "R$ #,##0.00"
                    else:  # Default: texto
                        cell.value = str(value)
                else:
                    cell.value = value

                # Replicar mesclagem do cabeçalho, se aplicável
                if cell_idx in merge_map:
                    start_col = cell_idx
                    end_col = merge_map[cell_idx]
                    worksheet.merge_cells(
                        start_row=row_idx,
                        start_column=start_col,
                        end_row=row_idx,
                        end_column=end_col,
                    )
                    # Aplicar 'Wrap Text' à célula mesclada
                    cell = worksheet.cell(row=row_idx, column=start_col)
                    cell.alignment = Alignment(
                        horizontal="center", vertical="center", wrap_text=True
                    )

                # Aplicar formatações, se fornecidas
                font_color = (
                    text_colors[col_idx]
                    if text_colors and len(text_colors) > col_idx
                    else "000000"
                )
                is_bold = bold[col_idx] if bold and len(bold) > col_idx else False

                # Alterar a cor do texto para vermelho se o valor começar com "-"
                if (isinstance(value, str) and value.startswith("-")) or (
                    (isinstance(value, float) or isinstance(value, int)) and value < 0
                ):
                    font_color = "FF0000"

                cell_font = Font(color=font_color, bold=is_bold)
                cell.font = cell_font

                # Aplicar bordas, se fornecidas
                if (
                    cell_borders
                    and len(cell_borders) > col_idx
                    and cell_borders[col_idx]
                ):
                    dotted_border = Border(
                        bottom=Side(style="dotted"),
                        top=Side(style="dotted"),
                        left=Side(style="dotted"),
                        right=Side(style="dotted"),
                    )
                    cell.border = dotted_border

                # Aplicar alinhamentos, se fornecidos
                alignment_style = Alignment(horizontal="center")
                if cell_alignments and len(cell_alignments) > col_idx:
                    cell_alignment = cell_alignments[col_idx].lower()
                    if cell_alignment in ["left", "right", "center"]:
                        alignment_style = Alignment(horizontal=cell_alignment)

                cell.alignment = alignment_style
            # Aplicar o estilo negrito em toda a linha, se "Total" estiver presente
            if total_row:
                for col_idx in range(1, len(df.columns) + 1):
                    cell = worksheet.cell(row=row_idx, column=col_idx)
                    cell.font = Font(bold=True, color=cell.font.color)

        # Mesclar células em colunas específicas para valores iguais
        if columns_to_merge:
            for column in columns_to_merge:
                col_idx = col_map[column]
                start_row = 4  # Dados começam na linha 4
                value_to_merge = None
                merge_start = None

                for row in range(start_row, start_row + len(df)):
                    cell = worksheet.cell(row=row, column=col_idx)
                    if cell.value == value_to_merge:
                        # Continuar mesclagem
                        continue
                    else:
                        # Finalizar mesclagem anterior
                        if merge_start and merge_start != row - 1:
                            worksheet.merge_cells(
                                start_row=merge_start,
                                start_column=col_idx,
                                end_row=row - 1,
                                end_column=col_idx,
                            )
                            merged_cell = worksheet.cell(
                                row=merge_start, column=col_idx
                            )
                            merged_cell.alignment = Alignment(
                                horizontal="center", vertical="center", wrap_text=True
                            )
                            # Aplicar 'Wrap Text' à célula mesclada
                            # Garantir que o tipo e formato da célula mesclada seja o mesmo
                            if column_types and len(column_types) > col_idx:
                                data_type = column_types[col_idx].lower()
                                if data_type == "número":
                                    merged_cell.number_format = "#,##0.00"
                                elif data_type == "moeda":
                                    merged_cell.number_format = "R$ #,##0.00"

                        # Iniciar nova mesclagem
                        value_to_merge = cell.value
                        merge_start = row

                # Finalizar a última mesclagem
                if merge_start and merge_start != (start_row + len(df) - 1):
                    worksheet.merge_cells(
                        start_row=merge_start,
                        start_column=col_idx,
                        end_row=start_row + len(df) - 1,
                        end_column=col_idx,
                    )
                    merged_cell = worksheet.cell(row=merge_start, column=col_idx)
                    merged_cell.alignment = Alignment(
                        horizontal="center", vertical="center", wrap_text=True
                    )
                    # Garantir que o tipo e formato da célula mesclada seja o mesmo
                    if column_types and len(column_types) > col_idx:
                        data_type = column_types[col_idx].lower()
                        if data_type == "número":
                            merged_cell.number_format = "#,##0.00"
                        elif data_type == "moeda":
                            merged_cell.number_format = "R$ #,##0.00"

    def consolidate_workbooks_to_xlsx(
        self,
        workbook_list,
        input_file_path: str,
        output_directory: str,
        output_file_name: str,
        starting_cell: str = "D9",
    ):

        base_workbook = load_workbook(input_file_path)
        base_worksheet = base_workbook.active

        # Determinar linha e coluna inicial a partir da célula inicial
        starting_column = starting_cell[0].upper()
        starting_row = int(starting_cell[1:])
        starting_column_index = ord(starting_column) - ord("A") + 1

        current_row = starting_row  # Começa pela linha inicial especificada

        # Unir workbooks
        for workbook in workbook_list:
            worksheet = (
                workbook.active
            )  # Assume que o conteúdo está na aba ativa do workbook

            # Determinar o maior índice de coluna usado no workbook atual
            column_count = worksheet.max_column

            # Copiar visibilidade de colunas
            for column_index in range(1, column_count + 1):
                column_letter = get_column_letter(column_index)
                base_worksheet.column_dimensions[column_letter].hidden = (
                    worksheet.column_dimensions[column_letter].hidden
                )

            # Copiar visibilidade de linhas
            for row_index in range(1, worksheet.max_row + 1):
                if worksheet.row_dimensions[row_index].hidden:
                    base_worksheet.row_dimensions[
                        current_row + row_index - 1
                    ].hidden = True

            # Copiar conteúdo e estilos de células
            for row in worksheet.iter_rows(
                min_row=1, max_row=worksheet.max_row, max_col=column_count
            ):
                for column_index, cell in enumerate(row, start=starting_column_index):
                    new_cell = base_worksheet.cell(row=current_row, column=column_index)
                    new_cell.value = cell.value

                    # Copiar estilo, recriando os objetos de estilo
                    if cell.has_style:
                        new_cell.font = Font(
                            name=cell.font.name,
                            size=cell.font.size,
                            bold=cell.font.bold,
                            italic=cell.font.italic,
                            underline=cell.font.underline,
                            strike=cell.font.strike,
                            color=cell.font.color,
                        )
                        new_cell.fill = PatternFill(
                            fill_type=cell.fill.fill_type,
                            start_color=cell.fill.start_color,
                            end_color=cell.fill.end_color,
                        )
                        new_cell.border = Border(
                            left=cell.border.left,
                            right=cell.border.right,
                            top=cell.border.top,
                            bottom=cell.border.bottom,
                        )
                        new_cell.alignment = Alignment(
                            horizontal=cell.alignment.horizontal,
                            vertical=cell.alignment.vertical,
                            wrap_text=cell.alignment.wrap_text,
                            shrink_to_fit=cell.alignment.shrink_to_fit,
                            indent=cell.alignment.indent,
                        )
                        new_cell.number_format = cell.number_format

                current_row += 1
            for column_index in range(1, column_count + 1):
                column_letter = get_column_letter(column_index)
                base_worksheet.column_dimensions[column_letter].hidden = (
                    worksheet.column_dimensions[column_letter].hidden
                )

            # Copiar visibilidade de linhas
            for row_index in range(1, worksheet.max_row + 1):
                if worksheet.row_dimensions[row_index].hidden:
                    base_worksheet.row_dimensions[
                        current_row + row_index - 1
                    ].hidden = True
            # Copiar células mescladas
            for merged_cell_range in worksheet.merged_cells.ranges:
                # Traduzir o intervalo para a nova posição
                min_col, min_row, max_col, max_row = range_boundaries(
                    str(merged_cell_range)
                )
                min_col += starting_column_index - 1  # Ajustar para a coluna inicial
                max_col += starting_column_index - 1
                min_row += (
                    current_row - worksheet.max_row - 1
                )  # Ajustar para a linha atual
                max_row += current_row - worksheet.max_row - 1
                new_range = f"{get_column_letter(min_col)}{min_row}:{get_column_letter(max_col)}{max_row}"
                base_worksheet.merge_cells(new_range)
                # cell = ws.cell(row=min_row, column=min_col)
                # cell.alignment = Alignment(horizontal="center", vertical="center",wrap_text=True)

            # Adicionar uma linha em branco como separador
            current_row += 1

        # Adicionar borda três linhas abaixo da última linha
        final_border_row = current_row + 2
        bottom_border = Border(bottom=Side(style="thin"))
        for column_index in range(4, 10):  # Colunas C a J (índices 3 a 10)
            cell = base_worksheet.cell(row=final_border_row, column=column_index)
            cell.border = bottom_border

        left_border_style = Border(left=Side(style="thin"), bottom=Side(style="thin"))
        cell = base_worksheet.cell(row=final_border_row, column=3)
        cell.border = left_border_style

        right_border = Border(right=Side(style="thin"), bottom=Side(style="thin"))
        cell = base_worksheet.cell(row=final_border_row, column=10)
        cell.border = right_border

        thin = Side(style="thin")

        # Loop de primeira até a linha_borda (incluindo ela)
        for row in range(starting_row, final_border_row):
            # Coluna C = 3 -> borda ESQUERDA
            cell_c = base_worksheet.cell(row=row, column=3)
            cell_c.border = Border(left=thin)

            # Coluna J = 10 -> borda DIREITA
            cell_j = base_worksheet.cell(row=row, column=10)
            cell_j.border = Border(right=thin)

        # Remover formatação de todas as células abaixo da linha com a borda
        for row in base_worksheet.iter_rows(
            min_row=final_border_row + 1,
            max_row=base_worksheet.max_row,
            max_col=base_worksheet.max_column,
        ):
            for cell in row:
                cell.value = None
                cell.font = Font()  # Font padrão (reset)
                cell.fill = PatternFill()  # Sem preenchimento
                cell.border = Border()  # Sem borda
                cell.alignment = Alignment()  # Alinhamento padrão

        # Salvar o arquivo final
        final_output_path = Path.joinpath(output_directory, output_file_name)
        base_workbook.save(final_output_path)
        # print(f"\nArquivo salvo em: {caminho_final}")
        return final_output_path

    def xlsx_to_pdf_one_page(self, input_file: str, output_path: str):

        os.makedirs(output_path, exist_ok=True)
        filename = os.path.splitext(os.path.basename(input_file))[0] + ".pdf"

        output_file = Path(output_path, filename)

        if os.path.isfile(output_file):
            return

        excel = win32.gencache.EnsureDispatch("Excel.Application")
        excel.Visible = False

        try:
            workbook = excel.Workbooks.Open(input_file)
        except Exception as e:
            print(f"Erro ao abrir o arquivo {input_file}: {e}")
            return

        worksheet = workbook.Worksheets(1)  # planilha COM
        worksheet.Columns.AutoFit()

        # área de impressão de C3 até última coluna/linha
        last_column = 12
        last_row = 1
        for row in range(1, worksheet.UsedRange.Rows.Count + 1):
            borders = worksheet.Cells(row, 4).Borders
            for border_id in range(5, 13):
                if (
                    borders(border_id).Color != 0.0
                    or borders(border_id).LineStyle != -4142
                ):
                    last_row = row
                    break

        for column_index in range(1, worksheet.UsedRange.Columns.Count + 1):
            borders = worksheet.Cells(5, column_index).Borders
            for border_id in range(5, 13):
                if (
                    borders(border_id).Color != 0.0
                    or borders(border_id).LineStyle != -4142
                ):
                    # last_column = col
                    break

        def column_index_to_letter(index):
            letter = ""
            while index > 0:
                index, rem = divmod(index - 1, 26)
                letter = chr(65 + rem) + letter
            return letter

        last_col_letter = column_index_to_letter(last_column)
        worksheet.PageSetup.PrintArea = worksheet.Range(
            f"C3:{last_col_letter}{last_row}"
        ).Address

        # configura impressão em 1 página
        worksheet.PageSetup.Zoom = False
        worksheet.PageSetup.FitToPagesWide = 1
        worksheet.PageSetup.FitToPagesTall = 1

        # **Auto-ajusta a largura de todas as colunas usadas**

        # exporta pra PDF
        workbook.ExportAsFixedFormat(0, str(output_file))

        workbook.Close(False)
        excel.Quit()
        # print(f"\nArquivo PDF gerado: {output_file}")
