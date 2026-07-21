from __future__ import annotations

import re
import sys
from contextlib import suppress
from pathlib import Path
from typing import Any

try:
    import xlwings as xw
except ImportError:
    xw = None

def write_excel_text(cell: Any, text: str) -> None:
    if not text.startswith("="):
        cell.api.NumberFormat = "@"
    cell.api.Value = text


def excel_rgb(color: Any) -> int | None:
    if color is None or color.type != "rgb" or not color.rgb:
        return None

    rgb = color.rgb[-6:]
    try:
        return int(rgb[4:6] + rgb[2:4] + rgb[:2], 16)
    except ValueError:
        return None


def apply_excel_inline_font(excel_font: Any, font: Any) -> None:
    for name, value in {
        "Name": font.rFont,
        "Size": font.sz,
        "Bold": font.b,
        "Italic": font.i,
    }.items():
        if value is not None:
            setattr(excel_font, name, value)

    if font.strike is not None:
        excel_font.Strikethrough = font.strike
    if font.vertAlign is not None:
        excel_font.Superscript = font.vertAlign == "superscript"
        excel_font.Subscript = font.vertAlign == "subscript"

    underline = {"single": 2, "double": -4119, "singleAccounting": 4, "doubleAccounting": 5}
    if font.u is not None:
        excel_font.Underline = underline.get(font.u, -4142)

    color = excel_rgb(font.color)
    if color is not None:
        excel_font.Color = color


def write_excel_rich_text(cell: Any, value: Any) -> None:
    text = str(value)
    write_excel_text(cell, text)

    position = 1
    for part in value:
        part_text = str(part)
        font = getattr(part, "font", None)
        if font is not None and part_text:
            excel_font = cell.api.GetCharacters(position, len(part_text)).Font
            apply_excel_inline_font(excel_font, font)
        position += len(part_text)


def apply_excel_fill(cell: Any, fill: Any) -> None:
    if fill.fill_type != "solid":
        return

    color = excel_rgb(fill.fgColor)
    if color is None:
        return

    cell.api.Interior.Pattern = 1
    cell.api.Interior.Color = color


def fill_template_with_excel(
    output_path: Path,
    sheet_name: str,
    start_row: int,
    source_row_numbers: list[int],
    rendered_rows: list[list[tuple[Any, Any | None]]],
) -> None:
    if sys.platform != "win32":
        raise RuntimeError("Заполнение .xlsm с надстройками выполняется только в Windows через Microsoft Excel.")
    if xw is None:
        raise RuntimeError("Для .xlsm установите xlwings: py -m pip install -r requirements.txt")

    try:
        app = xw.App(visible=False, add_book=False)
    except Exception as error:
        raise RuntimeError(f"Не удалось запустить Microsoft Excel: {error}") from error

    app.display_alerts = False
    app.screen_updating = False
    workbook = None
    source_sheet = None
    try:
        workbook = app.books.open(str(output_path), update_links=False, read_only=False)
        worksheet = workbook.sheets[sheet_name]

        worksheet.api.Copy(After=worksheet.api)
        source_sheet = workbook.sheets.active
        source_name = "__flashscore_template_source__"
        sheet_names = {sheet.name for sheet in workbook.sheets}
        suffix = 1
        while source_name in sheet_names:
            suffix += 1
            source_name = f"__flashscore_template_source_{suffix}__"
        source_sheet.name = source_name

        last_row = worksheet.used_range.last_cell.row
        if last_row >= start_row:
            worksheet.api.Rows(f"{start_row}:{last_row}").Delete()

        total_rows = len(rendered_rows)
        print(f"Excel: заполнение {total_rows} строк...", flush=True)
        for row_offset, rendered_row in enumerate(rendered_rows):
            target_row = start_row + row_offset
            source_row = source_row_numbers[row_offset % len(source_row_numbers)]
            source_range = source_sheet.range((source_row, 1), (source_row, len(rendered_row)))
            target_range = worksheet.range((target_row, 1), (target_row, len(rendered_row)))
            source_range.api.Copy(Destination=target_range.api)

            for column, (value, conditional_fill) in enumerate(rendered_row, start=1):
                target_cell = worksheet.cells(target_row, column)

                if hasattr(value, "_opt"):
                    write_excel_rich_text(target_cell, value)
                else:
                    rendered_value = "" if value is None else str(value)
                    write_excel_text(target_cell, rendered_value)

                if conditional_fill is not None:
                    apply_excel_fill(target_cell, conditional_fill)

            if (row_offset + 1) % 100 == 0 or row_offset + 1 == total_rows:
                print(f"Excel: заполнено {row_offset + 1}/{total_rows} строк.", flush=True)

        source_sheet.delete()
        source_sheet = None
        print("Excel: сохранение файла...", flush=True)
        workbook.save()
        print("Excel: файл сохранён.", flush=True)
    except Exception as error:
        raise RuntimeError(f"Не удалось заполнить .xlsm через Microsoft Excel: {error}") from error
    finally:
        try:
            if source_sheet is not None:
                with suppress(Exception):
                    source_sheet.delete()
        finally:
            try:
                if workbook is not None:
                    workbook.close()
            finally:
                app.quit()
