from __future__ import annotations

import json
import re
from copy import copy
from dataclasses import dataclass
from pathlib import Path
from shutil import copy2
from typing import Any

from openpyxl import load_workbook
from openpyxl.cell.cell import Cell
from openpyxl.cell.rich_text import CellRichText, InlineFont, TextBlock
from openpyxl.worksheet.worksheet import Worksheet

from flashscore_parser.table_io_xlwings import fill_template_with_excel

PLACEHOLDER_RE = re.compile(r"\{((?:[^{}]|\{[^{}]*\})+)\}")
CELL_REF = r"\$?[A-Z]{1,3}\$?\d+"
FORMAT_ATTR_RE = re.compile(
    rf"^\s*.+?\s*\?\s*{CELL_REF}\s*:\s*"
    rf"(?:.+?\s*\?\s*{CELL_REF}\s*:\s*)*{CELL_REF}\s*$"
)
FORMAT_RULE_RE = re.compile(
    rf"^\s*(?P<condition>.+?)\s*\?\s*"
    rf"(?P<cell>{CELL_REF})\s*:\s*(?P<remaining>.+)$"
)
CELL_REF_RE = re.compile(rf"^{CELL_REF}$")

@dataclass(frozen=True)
class XlsxTemplate:
    headers: list[str]
    row_templates: list[list[Any]]
    sheet_name: str
    header_row: int
    template_start_row: int

@dataclass
class Attributes:
    is_final: bool
    odds: tuple[float, float]
    rules: str


WORKBOOK_SUFFIXES = {".xlsx", ".xlsm"}


def ensure_workbook_path(path: str | Path) -> Path:
    workbook_path = Path(path)
    if workbook_path.suffix.lower() not in WORKBOOK_SUFFIXES:
        raise RuntimeError(f"Поддерживаются только .xlsx и .xlsm файлы: {workbook_path}")
    return workbook_path


def keep_vba(path: Path) -> bool:
    return path.suffix.lower() == ".xlsm"

def cell_text(value: Any) -> str:
    return "" if value is None else str(value)

def cell_has_placeholder(value: str) -> bool:
    return PLACEHOLDER_RE.search(cell_text(value)) is not None

def row_has_value(row: tuple[Cell, ...]) -> bool:
    return any(cell_has_placeholder(cell.value) for cell in row)

def is_special_name(name: str) -> bool:
    if name in {"repeat"}:
        return True
    return False
        
def is_dict_attr(name: str, attr: str):
    if name in {"over", "under", "asian-1", "asian-2", "european-1", "european-x", "european-2"}:
        return True if float(attr) % 0.25 == 0 else False 
    elif name == "correct":
        scores = attr.split(':')
        try:
            map(int, scores)
            return True
        except:
            return False
    elif name == "ht-ft":
        ht_ft = attr.split("/")
        if (isinstance(ht_ft[0], int) or ht_ft[0] == 'X') and \
            (isinstance(ht_ft[1], int) or ht_ft[1] == 'X'):
            return True
        else:
            return False

def is_format_attr(attr: str) -> re.Match[str] | None:
    return FORMAT_ATTR_RE.fullmatch(attr)


def split_attrs(attrs: str) -> list[str]:
    parts: list[str] = []
    start = 0
    nesting = 0

    for index, character in enumerate(attrs):
        if character == "{":
            nesting += 1
        elif character == "}":
            nesting = max(0, nesting - 1)
        elif character == ";" and nesting == 0:
            parts.append(attrs[start:index].strip())
            start = index + 1

    parts.append(attrs[start:].strip())
    return parts


def parse_format_attr(attr: str) -> tuple[list[tuple[str, str]], str]:
    remaining = attr.strip()
    rules: list[tuple[str, str]] = []

    while True:
        match = FORMAT_RULE_RE.fullmatch(remaining)
        if match is None:
            break
        rules.append((match.group("condition"), match.group("cell")))
        remaining = match.group("remaining").strip()

    if not rules or CELL_REF_RE.fullmatch(remaining) is None:
        raise RuntimeError(f"Некорректный атрибут форматирования: {attr}")
    return rules, remaining


def split_format_attr(placeholder: str) -> tuple[str, str | None]:
    key = placeholder.strip()
    if ":" not in key:
        return key, None

    name, attrs = key.split(":", 1)
    regular_attrs: list[str] = []
    format_attr: str | None = None

    for attr in split_attrs(attrs):
        if is_format_attr(attr):
            if format_attr is not None:
                raise RuntimeError(f"В плейсхолдере указан второй атрибут форматирования: {{{key}}}")
            format_attr = attr
        else:
            regular_attrs.append(attr)

    if format_attr is None:
        return key, None
    if not regular_attrs:
        return name, format_attr
    return f"{name}:{';'.join(regular_attrs)}", format_attr


def font_to_inline(font: Any) -> InlineFont:
    return InlineFont(
        rFont=font.name,
        charset=font.charset,
        family=font.family,
        b=font.b,
        i=font.i,
        strike=font.strike,
        outline=font.outline,
        shadow=font.shadow,
        condense=font.condense,
        extend=font.extend,
        color=copy(font.color),
        sz=font.sz,
        u=font.u,
        vertAlign=font.vertAlign,
        scheme=font.scheme,
    )


def resolve_format_cell(attr: str, values: dict[str, Any], worksheet: Worksheet) -> Cell:
    def replace_condition_placeholder(placeholder_match: re.Match[str]) -> str:
        value = resolve_placeholder(values, placeholder_match.group(1))
        return repr(value)

    rules, default_cell = parse_format_attr(attr)
    for raw_condition, cell_ref in rules:
        condition = PLACEHOLDER_RE.sub(replace_condition_placeholder, raw_condition)
        try:
            condition_result = bool(eval(condition, {"__builtins__": {}}, values))
        except (NameError, SyntaxError, TypeError, ValueError) as error:
            raise RuntimeError(f"Не удалось вычислить условие форматирования '{condition}': {error}") from error

        if condition_result:
            return worksheet[cell_ref.replace("$", "")]

    return worksheet[default_cell.replace("$", "")]


def resolve_format_attr(attr: str, values: dict[str, Any], worksheet: Worksheet) -> InlineFont:
    return font_to_inline(resolve_format_cell(attr, values, worksheet).font)
        
def is_odds_name(name: str, container: Any):
    if isinstance(container, tuple) and name in {"win1", "win2", "draw", "favorite", "outsider", "over", 
                                                 "under", "both-yes", "both-no", "double-1x", "double-12", 
                                                 "double-x2", "no-bet-1", "no-bet-2", "odd", "even"}:
        return True
    return False

def is_h2h_name(name: str):
    if name in {"home", "away", "h2h"}:
        return True
    return False

def resolve_placeholder(values: dict[str, Any], placeholder: str) -> Any:
    key = placeholder.strip()
    if key in values:
        value = values[key]
        if isinstance(value, dict) and not value:
            return ""
        if isinstance(value, (list, tuple)) and tuple(value) == (0.0, 0.0):
            return ""
        return value

    if ":" not in key:
        return ""

    name, attrs = key.split(":", 1)
    if is_special_name(name):
        odd = resolve_placeholder(values, attrs)
        repeat_values = values["repeat"].setdefault(key, {})
        repeat_values[odd] = repeat_values.get(odd, 0) + 1
        return repeat_values[odd]

    container = values.get(name)
    if isinstance(container, dict):
        attributes = Attributes(is_final=True, odds=(), rules="")
        split_attr = split_attrs(attrs)

        for attr in split_attr:
            if attr == "prev":
                attributes.is_final = False
            elif attr == "final":
                continue
            elif is_dict_attr(name, attr):
                attributes.odds = container.get(attr, tuple())
            else:
                return ""        
        if len(attributes.odds) < 2 or attributes.odds == (0.0, 0.0):
            return ""
        if attributes.is_final:
            return attributes.odds[1]
        else:
            return attributes.odds[0]
            
    if is_odds_name(name, container):
        attributes = Attributes(is_final=True, odds=(), rules="")
        split_attr = split_attrs(attrs)

        for attr in split_attr:
            if attr == "prev":
                attributes.is_final = False
            elif attr == "final":
                continue
            else:
                return ""
        if len(container) < 2 or tuple(container) == (0.0, 0.0):
            return ""
        if attributes.is_final:
            return container[1]
        else:
            return container[0]
    
    if is_h2h_name(name):
        split_attr = split_attrs(attrs)
        i = 0
        field = ""

        for attr in split_attr:
            if attr.isdigit():
                i = int(attr)
            else:
                field = attr
        
        h2h_dict = container[i-1].as_placeholder() if i <= len(container) else {field: "-"}
        return h2h_dict.get(field, "")

    return ""

def rich_parts(cell: Cell) -> list[tuple[str, InlineFont | None]]:
    if not isinstance(cell.value, CellRichText):
        return [(cell_text(cell.value), None)]

    parts: list[tuple[str, InlineFont | None]] = []
    for part in cell.value:
        if isinstance(part, TextBlock):
            parts.append((part.text, copy(part.font)))
        else:
            parts.append((str(part), None))
    return parts


def append_rich_text(result: CellRichText, font: InlineFont | None, text: Any) -> None:
    rendered_text = "" if text is None else str(text)
    if not rendered_text:
        return
    if font is None:
        result.append(rendered_text)
    else:
        result.append(TextBlock(copy(font), rendered_text))


def normalize_rich_text(result: CellRichText) -> CellRichText:
    normalized = CellRichText()
    pending_whitespace = ""

    for part in result:
        text = str(part)
        if text.isspace():
            pending_whitespace += text
            continue

        if pending_whitespace:
            if isinstance(part, TextBlock):
                part = TextBlock(copy(part.font), pending_whitespace + part.text)
            elif normalized and isinstance(normalized[-1], TextBlock):
                normalized[-1].text += pending_whitespace
            else:
                normalized.append(pending_whitespace)
            pending_whitespace = ""

        normalized.append(copy(part) if isinstance(part, TextBlock) else part)

    if pending_whitespace and normalized and isinstance(normalized[-1], TextBlock):
        normalized[-1].text += pending_whitespace

    normalized._opt()
    return normalized


def render_rich_cell(
    cell: Cell,
    values: dict[str, Any],
    worksheet: Worksheet,
) -> tuple[str | CellRichText, Any | None]:
    result = CellRichText()
    has_rich_text = isinstance(cell.value, CellRichText)
    has_format_attr = False
    conditional_fill: Any | None = None

    for text, original_font in rich_parts(cell):
        position = 0
        for match in PLACEHOLDER_RE.finditer(text):
            append_rich_text(result, original_font, text[position:match.start()])

            placeholder, format_attr = split_format_attr(match.group(1))
            rendered_value = resolve_placeholder(values, placeholder)
            placeholder_font = original_font
            if format_attr is not None:
                format_cell = resolve_format_cell(format_attr, values, worksheet)
                placeholder_font = font_to_inline(format_cell.font)
                if conditional_fill is None:
                    conditional_fill = copy(format_cell.fill)
                has_format_attr = True
            append_rich_text(result, placeholder_font, rendered_value)
            position = match.end()

        append_rich_text(result, original_font, text[position:])

    if has_rich_text or has_format_attr:
        result._opt()
        return normalize_rich_text(result), conditional_fill
    rendered_value = str(result)
    return ("" if rendered_value.isspace() else rendered_value), conditional_fill

def render_rich_rows(
    source_rows: list[tuple[Cell, ...]],
    data_rows: list[dict[str, Any]],
    worksheet: Worksheet,
) -> list[list[tuple[str | CellRichText, Any | None]]]:
    rendered: list[list[tuple[str | CellRichText, Any | None]]] = []
    repeats = [{} for _ in range(max(len(row) for row in source_rows))]

    for values in data_rows:
        for source_row in source_rows:
            row: list[tuple[str | CellRichText, Any | None]] = []
            for column, cell in enumerate(source_row):
                values["repeat"] = repeats[column]
                row.append(render_rich_cell(cell, values, worksheet))
            rendered.append(row)
    return rendered


def copy_cell_style(source: Cell, target: Cell) -> None:
    if source.has_style:
        target.font = copy(source.font)
        target.border = copy(source.border)
        target.alignment = copy(source.alignment)
        target.number_format = copy(source.number_format)
        target.protection = copy(source.protection)
    if source.hyperlink:
        target._hyperlink = copy(source.hyperlink)
    if source.comment:
        target.comment = copy(source.comment)


def template_rows(worksheet: Worksheet, start_row: int) -> list[tuple[Cell, ...]]:
    return [row for row in worksheet.iter_rows(min_row=start_row) if row_has_value(row)]


def clear_template_area(worksheet: Worksheet, start_row: int) -> None:
    if worksheet.max_row >= start_row:
        worksheet.delete_rows(start_row, worksheet.max_row - start_row + 1)


def read_xlsx_template(
    path: str | Path,
    sheet_name: str | None = None,
    header_row: int = 1,
) -> XlsxTemplate:
    template_path = ensure_workbook_path(path)
    workbook = load_workbook(template_path, rich_text=True, keep_vba=keep_vba(template_path))
    worksheet = workbook[sheet_name] if sheet_name else workbook.active

    if header_row < 1:
        raise RuntimeError("Номер строки заголовков должен быть больше 0.")

    sheet_title = worksheet.title
    headers = [cell_text(cell.value) for cell in worksheet[header_row]]
    row_templates: list[list[Any]] = []

    for row in worksheet.iter_rows(min_row=header_row + 1):
        if not row_has_value(row):
            break
        row_templates.append([cell.value for cell in row])

    workbook.close()

    if not any(headers):
        raise RuntimeError("В .xlsx шаблоне не найдена строка заголовков.")
    if not row_templates:
        raise RuntimeError("В .xlsx шаблоне не найдена строка с плейсхолдерами.")

    return XlsxTemplate(
        headers=headers,
        row_templates=row_templates,
        sheet_name=sheet_title,
        header_row=header_row,
        template_start_row=header_row + 1,
    )

def fill_xlsx_template(
    data_rows: list[dict[str, Any]],    
    template: XlsxTemplate,
    template_path: str | Path,
    output_path: str | Path
) -> None:
    template_path = ensure_workbook_path(template_path)
    output_path = ensure_workbook_path(output_path)
    if template_path.suffix.lower() != output_path.suffix.lower():
        raise RuntimeError("Расширение выходного файла должно совпадать с расширением шаблона.")
    if template_path.resolve() == output_path.resolve():
        raise RuntimeError("Выходной файл не должен совпадать с файлом шаблона.")

    copy2(template_path, output_path)
    workbook = load_workbook(output_path, rich_text=True, keep_vba=keep_vba(output_path))
    worksheet = workbook[template.sheet_name]
    start_row = template.header_row + 1
    source_rows = template_rows(worksheet, start_row)
    if not source_rows:
        workbook.close()
        raise RuntimeError("В .xlsx шаблоне не найдена строка с плейсхолдерами.")

    rendered_rows = render_rich_rows(source_rows, data_rows, worksheet)
    style_rows = [[cell for cell in row] for row in source_rows]

    if keep_vba(output_path):
        source_row_numbers = [row[0].row for row in source_rows]
        workbook.close()
        fill_template_with_excel(
            output_path,
            template.sheet_name,
            start_row,
            source_row_numbers,
            rendered_rows,
        )
        return

    clear_template_area(worksheet, start_row)

    for row_offset, rendered_row in enumerate(rendered_rows):
        target_row_index = start_row + row_offset
        style_row = style_rows[row_offset % len(style_rows)]
        for column_index, (value, conditional_fill) in enumerate(rendered_row, start=1):
            target_cell = worksheet.cell(row=target_row_index, column=column_index, value=value)
            if column_index <= len(style_row):
                source_cell = style_row[column_index - 1]
                copy_cell_style(source_cell, target_cell)
                if source_cell.data_type == "s" and isinstance(value, str) and value.startswith("="):
                    target_cell.data_type = "s"
            if conditional_fill is not None:
                target_cell.fill = conditional_fill

    workbook.save(output_path)
    workbook.close()
