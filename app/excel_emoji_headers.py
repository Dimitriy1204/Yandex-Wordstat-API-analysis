#!/usr/bin/env python3
"""
Разметка заголовков с «кружками» — как в yandex_analysis_exemple.xlsx.

В образце используется один символ 🔴 с разным цветом шрифта (не 🟢🟡🔴).
"""
from __future__ import annotations

import re

# Символ круга из файла-образца (окрашивается через font_color)
CIRCLE_CHAR = "\U0001f534"

# (текст, цвет шрифта RRGGBB без #) — скопировано из sharedStrings образца
RUNS_VOLATILITY_BLUE = [
    ("Волатильность\n", "FFFFFF"),
    (CIRCLE_CHAR, "00B050"),
    ("низкая ", "FFFFFF"),
    (CIRCLE_CHAR, "FFFF00"),
    ("средняя ", "FFFFFF"),
    (CIRCLE_CHAR, "FF0000"),
    ("высокая", "FFFFFF"),
]

RUNS_COMPETITION_BLUE = [
    ("Индекс конкуренции\n", "FFFFFF"),
    (CIRCLE_CHAR, "00B050"),
    ("<0.40 ", "FFFFFF"),
    (CIRCLE_CHAR, "FFFF00"),
    ("0.40-0.60 ", "FFFFFF"),
    (CIRCLE_CHAR, "FF0000"),
    (">0.60", "FFFFFF"),
]

# Методология на белом фоне — те же кружки, текст тёмный
RUNS_VOLATILITY_GUIDE = [
    (CIRCLE_CHAR, "00B050"),
    (" Низкая", "404040"),
    (CIRCLE_CHAR, "FFFF00"),
    (" Средняя", "404040"),
    (CIRCLE_CHAR, "FF0000"),
    (" Высокая", "404040"),
]

RUNS_COMPETITION_GUIDE = [
    ("Границы: ", "404040"),
    (CIRCLE_CHAR, "00B050"),
    (" < 0.40, ", "404040"),
    (CIRCLE_CHAR, "FFFF00"),
    (" 0.40–0.60, ", "404040"),
    (CIRCLE_CHAR, "FF0000"),
    (" > 0.60", "404040"),
]

_EMOJI_MARKERS = re.compile(r"(🟢|🟡|🔴)")
_TINT = {"🟢": "00B050", "🟡": "FFFF00", "🔴": "FF0000"}


def template_runs_for_header(text: str) -> list[tuple[str, str | None]] | None:
    """Готовые runs из образца или None."""
    if not text:
        return None
    t = text.lower()
    if "низкая" in t and "средняя" in t and "высокая" in t and "волатиль" in t:
        return list(RUNS_VOLATILITY_BLUE)
    if "конкуренц" in t and ("0.40" in t or "0,40" in t):
        return list(RUNS_COMPETITION_BLUE)
    return None


def template_runs_for_guide_line(text: str) -> list[tuple[str, str | None]] | None:
    if not text or not _EMOJI_MARKERS.search(text):
        return None
    if "низкая" in text.lower() and "волатиль" in text.lower():
        return _runs_volatility_guide_from_text(text)
    if "границы" in text.lower() and "0.40" in text:
        return _runs_competition_guide_from_text(text)
    if "0.8" in text and "сезон" in text.lower():
        return _runs_seasonality_guide_from_text(text)
    return _parse_emoji_line_guide(text)


def _runs_volatility_guide_from_text(text: str) -> list[tuple[str, str | None]]:
    """Строки методологии про волатильность с подстановкой порогов."""
    lines = text.split("\n")
    out: list[tuple[str, str | None]] = []
    for i, line in enumerate(lines):
        if i > 0:
            out.append(("\n", "404040"))
        if "низкая" in line.lower() or "средняя" in line.lower() or "высокая" in line.lower():
            out.extend(_parse_emoji_line_guide(line))
        else:
            out.append((line, "404040"))
    return out


def _runs_competition_guide_from_text(text: str) -> list[tuple[str, str | None]]:
    if "границы" in text.lower():
        idx = text.lower().find("границы")
        prefix = text[: idx + len("границы:")] if ":" in text[: idx + 10] else text.split("🟢")[0]
        out: list[tuple[str, str | None]] = [(prefix.rstrip() + " ", "404040")]
        out.extend(RUNS_COMPETITION_GUIDE[1:])
        return out
    return _parse_emoji_line_guide(text)


def _runs_seasonality_guide_from_text(text: str) -> list[tuple[str, str | None]]:
    """Методология сезонности: по одному кружку на строку — цвет по эмодзи."""
    lines = text.split("\n")
    out: list[tuple[str, str | None]] = []
    for i, line in enumerate(lines):
        if i > 0:
            out.append(("\n", "404040"))
        if _EMOJI_MARKERS.search(line):
            for seg in _EMOJI_MARKERS.split(line):
                if seg in _TINT:
                    out.append((CIRCLE_CHAR, _TINT[seg]))
                elif seg:
                    out.append((seg, "404040"))
        else:
            out.append((line, "404040"))
    return out


def _parse_emoji_line_guide(line: str) -> list[tuple[str, str | None]]:
    out: list[tuple[str, str | None]] = []
    for seg in _EMOJI_MARKERS.split(line):
        if seg in _TINT:
            out.append((CIRCLE_CHAR, _TINT[seg]))
        elif seg:
            out.append((seg, "404040"))
    return out


def parse_header_to_runs(text: str) -> list[tuple[str, str | None]]:
    """Парсит заголовок с 🟢🟡🔴 в runs как в образце (🔴 + tint)."""
    tmpl = template_runs_for_header(text)
    if tmpl:
        return tmpl
    if not _EMOJI_MARKERS.search(text):
        return [(text, "FFFFFF")]

    out: list[tuple[str, str | None]] = []
    pos = 0
    for m in _EMOJI_MARKERS.finditer(text):
        if m.start() > pos:
            out.append((text[pos:m.start()], "FFFFFF"))
        out.append((CIRCLE_CHAR, _TINT[m.group(1)]))
        pos = m.end()
    if pos < len(text):
        out.append((text[pos:], "FFFFFF"))
    return out


def write_rich_cell(ws, row: int, col: int, runs: list[tuple[str, str | None]],
                    wb, base_fmt: dict) -> None:
    """Пишет rich_string; у каждого фрагмента сохраняется base_fmt (в т.ч. bg_color)."""
    if not runs:
        ws.write(row, col, "", wb.add_format(base_fmt))
        return
    if len(runs) == 1 and not runs[0][1]:
        ws.write(row, col, runs[0][0], wb.add_format(base_fmt))
        return

    parts: list = []
    for frag, color in runs:
        if not frag:
            continue
        fmt = dict(base_fmt)
        if color:
            c = str(color).lstrip("#")
            if len(c) == 6:
                fmt["font_color"] = f"#{c}"
        parts.extend([wb.add_format(fmt), frag])
    cell_fmt = wb.add_format(base_fmt)
    if len(parts) >= 2:
        # xlsxwriter: формат ячейки (фон, рамка) — последний аргумент, не keyword
        ws.write_rich_string(row, col, *parts, cell_fmt)
    else:
        ws.write(row, col, "".join(t for t, _ in runs), cell_fmt)
