#!/usr/bin/env python3
"""
Генератор форматированного Excel-отчёта.
Форматирование строго по образцу yandex_analysis_exemple.xlsx.
"""

from __future__ import annotations

import os
import re
import logging
from typing import Optional
from datetime import datetime

import numpy as np
import pandas as pd
import xlsxwriter

from app.metrics import blue_verdict, growth_verdict
from app.excel_header_template import get_table_headers, get_table_header_runs
from app.prophet_forecast import prophet_status_message
from app.chart_data import build_chart_payload, prepare_chart_series_with_forecast, CHART_COLORS
from app.excel_emoji_headers import (
    parse_header_to_runs,
    template_runs_for_guide_line,
    template_runs_for_header,
    write_rich_cell,
)

logger = logging.getLogger("wordstat.excel")

# Цветовая схема как в шаблоне
COLOR_HEADER = "#4472C4"
COLOR_TITLE = "#2F5496"
COLOR_GREEN = "#A9D18E"
COLOR_YELLOW = "#FFD966"
COLOR_RED = "#F4B183"

# Единый шрифт для всего файла
FONT_NAME = "Calibri"
FONT_SIZE = 11
HDR_TABLE_HEIGHT = 30  # высота синей строки заголовков (по умолчанию)
HDR_TABLE_HEIGHT_TALL = 66  # листы «Данные» и «Сводка по запросам»

_HEADER_EMOJIS = ("🟢", "🟡", "🔴")

GROWTH_COL_TITLE = {
    "growth_momentum": "импульс — последние 6 мес.",
    "mean_growth": "средний рост — 12 мес.",
    "trend_slope": "наклон тренда — 12 мес.",
}
GROWTH_COL_HEADER = {
    "growth_momentum": "Импульс роста,\n%",
    "mean_growth": "Средний рост,\n%",
    "trend_slope": "Наклон тренда\n(показов/мес)",
}


def _base_fmt(extra: dict = None) -> dict:
    """Базовый словарь форматирования с единым шрифтом."""
    d = {"font_name": FONT_NAME, "font_size": FONT_SIZE}
    if extra:
        d.update(extra)
    return d


def _fmt(extra: dict = None) -> dict:
    return _base_fmt(extra)


def _fmt_bold(extra: dict = None) -> dict:
    return _base_fmt({"bold": True, **(extra or {})})


def _set_row(ws, row: int, text: str = "", min_height: int = 18) -> None:
    """Устанавливает высоту строки. Если текст длинный - выше."""
    if text and "\n" in text:
        lines = text.count("\n") + 1
        h = max(min_height, lines * 15)
        ws.set_row(row, h)
    else:
        ws.set_row(row, min_height)


def _hdr_dict_base() -> dict:
    return _fmt_bold({
        "bg_color": COLOR_HEADER,
        "font_color": "white",
        "border": 1,
        "align": "center",
        "text_wrap": True,
        "valign": "vcenter",
        "font_size": FONT_SIZE,
    })


def _write_header_cell(ws, row: int, col: int, text: str, wb, hdr_base: dict) -> None:
    """Заголовок: разметка кружков как в yandex_analysis_exemple.xlsx (🔴 + цвет шрифта)."""
    if any(e in text for e in _HEADER_EMOJIS):
        write_rich_cell(ws, row, col, parse_header_to_runs(text), wb, hdr_base)
    else:
        ws.write(row, col, text, wb.add_format(hdr_base))


def _write_verdict_cell(ws, row: int, col: int, text: str, wb, base_props: dict) -> None:
    """Вердикт целиком одним цветом (как в образце: #006400 / #404040 / #8B0000)."""
    ws.write(row, col, text or "", wb.add_format(base_props))


def _write_header_cell_runs(ws, row: int, col: int, runs: list, wb, hdr_base: dict) -> None:
    """Заголовок из шаблона; при кружках — разметка как в образце."""
    text = "".join(t for t, _ in runs) if runs else ""
    if not runs:
        return
    tmpl = template_runs_for_header(text)
    if tmpl:
        write_rich_cell(ws, row, col, tmpl, wb, hdr_base)
        return
    if any(e in text for e in _HEADER_EMOJIS):
        write_rich_cell(ws, row, col, parse_header_to_runs(text), wb, hdr_base)
        return
    if len(runs) > 1:
        write_rich_cell(ws, row, col, runs, wb, hdr_base)
        return
    _write_header_cell(ws, row, col, runs[0][0], wb, hdr_base)


def _col_by_header(headers: list[str], *needles: str,
                   exclude: tuple[str, ...] = ()) -> int | None:
    """Индекс колонки по подстрокам в заголовке (для сводки из образца)."""
    for i, h in enumerate(headers):
        hl = (h or "").lower()
        if exclude and any(ex in hl for ex in exclude):
            continue
        if all(n.lower() in hl for n in needles):
            return i
    return None


def _summary_col_map(headers: list[str]) -> dict[str, int | None]:
    return {
        "num": _col_by_header(headers, "п.п") or _col_by_header(headers, "№"),
        "query": _col_by_header(headers, "поисков"),
        "mean_freq": _col_by_header(headers, "средняя", exclude=("сумм", "ёмк", "емк")),
        "sum_freq": _col_by_header(headers, "суммар") or _col_by_header(headers, "ёмк")
        or _col_by_header(headers, "емк"),
        "std": _col_by_header(headers, "стандарт") or _col_by_header(headers, "отклон"),
        "mean_growth": _col_by_header(headers, "весь", "период")
        or _col_by_header(headers, "за весь"),
        "momentum": _col_by_header(headers, "6", "мес") or _col_by_header(headers, "импульс"),
        "growth_year": _col_by_header(headers, "год", exclude=("6", "наклон", "тренд")),
        "trend": _col_by_header(headers, "наклон") or _col_by_header(headers, "тренд"),
        "volatility": _col_by_header(headers, "волатиль"),
        "competition": _col_by_header(headers, "конкуренц"),
        "forecast": _col_by_header(headers, "прогноз"),
    }


def _write_header_row(
    ws, row: int, headers: list[str], wb, hdr_base: dict,
    header_runs: list | None = None,
    fixed_height: int | None = None,
) -> None:
    if header_runs and len(header_runs) == len(headers):
        for col, runs in enumerate(header_runs):
            _write_header_cell_runs(ws, row, col, runs, wb, hdr_base)
    else:
        for col, text in enumerate(headers):
            _write_header_cell(ws, row, col, text, wb, hdr_base)
    if fixed_height is not None:
        ws.set_row(row, fixed_height)
    else:
        line_count = max((h.count("\n") + 1) for h in headers)
        ws.set_row(row, min(36, max(HDR_TABLE_HEIGHT, line_count * 12 + 6)))


def _format_top_rows(df: pd.DataFrame, n: int, line_fn) -> str:
    """Форматирует топ-N строк для текстовых выводов на листе «О файле»."""
    if df.empty:
        return "Недостаточно данных."
    lines = [line_fn(df.iloc[i]) for i in range(min(n, len(df)))]
    if len(lines) == 1:
        return lines[0]
    head, *tail = lines
    if len(tail) == 1:
        return f"{head}\nТакже в топе: {tail[0]}"
    return f"{head}\nТакже в топе: {tail[0]} и {tail[1]}"


def _excel_series_name(query: str) -> str:
    """Имя ряда для Excel (ограничение длины)."""
    q = (query or "").strip()
    return (q[:28] + "…") if len(q) > 31 else q


def _write_chart_table(
    ws,
    start_row: int,
    queries: list[str],
    labels: list[str],
    series: dict[str, list],
    forecast_series: dict[str, list] | None,
    hdr_fmt,
    txt_fmt,
    num_fmt,
) -> tuple[int, int, list[tuple[int, str, bool]]]:
    """Таблица: колонка «Месяц» + факт и прогноз по каждому запросу."""
    ws.write(start_row, 0, "Месяц", hdr_fmt)
    col_specs: list[tuple[int, str, bool]] = []
    col = 1
    for q in queries:
        ws.write(start_row, col, _excel_series_name(q), hdr_fmt)
        col_specs.append((col, q, False))
        col += 1
        if forecast_series and q in forecast_series:
            fn = _excel_series_name(q)
            ws.write(start_row, col, (fn[:24] + " прогн.") if len(fn) > 24 else fn + " прогн.", hdr_fmt)
            col_specs.append((col, q, True))
            col += 1

    for ri, lbl in enumerate(labels):
        ws.write(start_row + 1 + ri, 0, lbl, txt_fmt)
        for ci, q, is_fc in col_specs:
            src = forecast_series if is_fc else series
            vals = (src or {}).get(q, [])
            v = vals[ri] if ri < len(vals) else None
            if v is None or (isinstance(v, float) and (pd.isna(v) or np.isinf(v))):
                ws.write(start_row + 1 + ri, ci, "", txt_fmt)
            else:
                ws.write(start_row + 1 + ri, ci, float(v), num_fmt)

    return start_row, len(labels), col_specs


def _insert_excel_line_chart(
    ws_chart,
    wb,
    data_sheet: str,
    block_row: int,
    n_data_rows: int,
    col_specs: list[tuple[int, str, bool]],
    title: str,
    position: str,
) -> None:
    """Вставляет линейную диаграмму Excel (факт + прогноз пунктиром)."""
    if n_data_rows < 1 or not col_specs:
        return
    chart = wb.add_chart({"type": "line"})
    chart.set_title({"name": title, "name_font": {"size": 12, "bold": True, "color": COLOR_TITLE}})
    chart.set_x_axis({"name": "Период", "major_gridlines": {"visible": False}})
    chart.set_y_axis({
        "name": "Показов в месяц",
        "major_gridlines": {"visible": True, "line": {"color": "#D9D9D9"}},
    })
    chart.set_legend({"position": "bottom", "font": {"size": 9}})
    chart.set_size({"width": 880, "height": 420})
    chart.set_style(10)

    first_r = block_row + 1
    last_r = block_row + n_data_rows
    q_color_idx: dict[str, int] = {}
    for col, q, is_forecast in col_specs:
        if q not in q_color_idx:
            q_color_idx[q] = len(q_color_idx)
        color = CHART_COLORS[q_color_idx[q] % len(CHART_COLORS)]
        line = {"width": 2.25, "color": color}
        if is_forecast:
            line["dash_type"] = "dash"
        chart.add_series({
            "name": [data_sheet, block_row, col],
            "categories": [data_sheet, first_r, 0, last_r, 0],
            "values": [data_sheet, first_r, col, last_r, col],
            "line": line,
            "marker": {
                "type": "circle",
                "size": 4 if is_forecast else 5,
                "border": {"color": color},
                "fill": {"color": color},
            },
        })
    ws_chart.insert_chart(position, chart)


def _write_charts_sheet(
    wb,
    ws_chart,
    analyzed: pd.DataFrame,
    summary: pd.DataFrame,
    top_n: int,
    date_range_str: str,
    title_fmt,
    hdr_fmt,
    txt_fmt,
    num_fmt,
    forecast_df: pd.DataFrame | None = None,
) -> None:
    """Лист «График»: две нативные диаграммы Excel (топ-N и все запросы) + прогноз."""
    data_sheet = "Данные графиков"
    ws_data = wb.add_worksheet(data_sheet)
    ws_data.hide()

    ordered = summary["query"].tolist() if not summary.empty else []
    top_queries = ordered[:top_n]
    all_queries = ordered if ordered else analyzed["query"].drop_duplicates().tolist()

    labels_top, series_top, fc_top = prepare_chart_series_with_forecast(
        analyzed, top_queries, forecast_df)
    labels_all, series_all, fc_all = prepare_chart_series_with_forecast(
        analyzed, all_queries, forecast_df)

    chart_hdr = wb.add_format(_fmt_bold({
        "bg_color": COLOR_HEADER, "font_color": "white", "border": 1, "align": "center",
    }))
    chart_num = wb.add_format(_fmt({"num_format": "# ##0", "border": 1}))

    row_top, n_top, spec_top = _write_chart_table(
        ws_data, 0, top_queries, labels_top, series_top, fc_top or None,
        chart_hdr, txt_fmt, chart_num)
    gap = n_top + 4
    row_all, n_all, spec_all = _write_chart_table(
        ws_data, gap, all_queries, labels_all, series_all, fc_all or None,
        chart_hdr, txt_fmt, chart_num)

    main_title = f"Динамика поискового спроса\n{date_range_str}"
    ws_chart.merge_range(0, 0, 0, 8, main_title, title_fmt)
    _set_row(ws_chart, 0, main_title, 36)

    sub_fmt = wb.add_format(_fmt_bold({"font_color": COLOR_TITLE, "font_size": 12}))
    fc_note = " (сплошная — факт, пунктир — прогноз Prophet)" if forecast_df is not None and not forecast_df.empty else ""
    ws_chart.write(2, 0, f"График 1. Топ-{top_n} запросов{fc_note}", sub_fmt)
    _insert_excel_line_chart(
        ws_chart, wb, data_sheet, row_top, n_top, spec_top,
        f"Топ-{top_n} запросов", "B4",
    )

    ws_chart.write(24, 0, f"График 2. Все сравниваемые запросы{fc_note}", sub_fmt)
    _insert_excel_line_chart(
        ws_chart, wb, data_sheet, row_all, n_all, spec_all,
        "Все запросы", "B26",
    )
    ws_chart.set_column(0, 0, 42)


def _write_ocean_block(ws, wb, df: pd.DataFrame, start_row: int,
                       block_title: str, date_range_str: str,
                       hdr_base: dict, title_fmt, txt_fmt,
                       int_fmt, num_fmt, comp_fmt,
                       v_fmt_good, v_fmt_bad, v_fmt_neutral,
                       growth_col: str, llm_verdicts: dict = None) -> int:
    """Пишет блок 'Голубые океаны' начиная с start_row."""
    nn_fmt = wb.add_format(_fmt({"border": 1, "text_wrap": True,
                                  "italic": True, "font_color": "#2F5496", "valign": "vcenter",
                                  "font_size": FONT_SIZE}))
    last_col = 6
    r = start_row
    ws.merge_range(r, 0, r, last_col, block_title, title_fmt)
    _set_row(ws, r, block_title, 32)
    r += 1
    cols = get_table_headers("ocean", growth_col)
    runs = get_table_header_runs("ocean", growth_col)
    _write_header_row(ws, r, cols, wb, hdr_base, runs)
    r += 1
    if df.empty:
        ws.merge_range(r, 0, r, last_col, "Нет данных для отображения", txt_fmt)
        _set_row(ws, r, min_height=22)
        return r + 2
    for ri, (_, row_data) in enumerate(df.iterrows()):
        q = row_data["query"]
        ws.write(r, 0, ri + 1, int_fmt)
        ws.write(r, 1, q, txt_fmt)
        score = round(row_data["competition_score"], 3)
        ws.write(r, 2, score, comp_fmt(score))
        growth_val = row_data.get(growth_col, 0)
        ws.write(r, 3, round(growth_val, 2) if pd.notna(growth_val) else 0.0, num_fmt)
        ws.write(r, 4, int(row_data["sum_frequency"]), int_fmt)
        bv = blue_verdict(row_data["competition_score"], growth_val if pd.notna(growth_val) else 0)
        is_good = "🟢" in bv
        is_bad = "🔴" in bv or "🟠" in bv
        v_props = _fmt({"border": 1, "text_wrap": True, "bold": True, "valign": "vcenter",
                        "font_size": FONT_SIZE,
                        "font_color": "#006400" if is_good else ("#8B0000" if is_bad else "#404040")})
        _write_verdict_cell(ws, r, 5, bv, wb, v_props)
        nn_verdict = ""
        if llm_verdicts and q in llm_verdicts:
            nn_verdict = llm_verdicts.get(q, "")
        ws.write(r, 6, nn_verdict or "", nn_fmt)
        _set_row(ws, r, max(bv, nn_verdict), 28)
        r += 1
    return r + 2


def _write_growing_block(ws, wb, df: pd.DataFrame, start_row: int,
                         block_title: str, date_range_str: str,
                         hdr_base: dict, title_fmt, txt_fmt,
                         int_fmt, num_fmt, comp_fmt, vol_fmt_h,
                         v_fmt_good, v_fmt_bad, v_fmt_neutral,
                         growth_col: str) -> int:
    """Пишет блок 'Растущие рынки' начиная с start_row."""
    last_col = 6
    r = start_row
    ws.merge_range(r, 0, r, last_col, block_title, title_fmt)
    _set_row(ws, r, block_title, 32)
    r += 1
    cols = get_table_headers("growing", growth_col)
    runs = get_table_header_runs("growing", growth_col)
    _write_header_row(ws, r, cols, wb, hdr_base, runs)
    r += 1
    if df.empty:
        ws.merge_range(r, 0, r, last_col, "Нет данных для отображения", txt_fmt)
        _set_row(ws, r, min_height=22)
        return r + 2
    for ri, (_, row_data) in enumerate(df.iterrows()):
        ws.write(r, 0, ri + 1, int_fmt)
        ws.write(r, 1, row_data["query"], txt_fmt)
        growth_val = row_data.get(growth_col, 0)
        ws.write(r, 2, round(growth_val, 2) if pd.notna(growth_val) else 0.0, num_fmt)
        ws.write(r, 3, int(row_data["sum_frequency"]), int_fmt)
        score = round(row_data["competition_score"], 3)
        ws.write(r, 4, score, comp_fmt(score))
        vol = round(row_data["mean_volatility"], 1)
        label = row_data["vol_label"]
        ws.write(r, 5, vol, vol_fmt_h(vol, label))
        gv = growth_verdict(growth_val if pd.notna(growth_val) else 0)
        is_growing = "🟢" in gv
        is_bad = "🔴" in gv
        v_props = _fmt({"border": 1, "text_wrap": True, "bold": True, "valign": "vcenter",
                        "font_size": FONT_SIZE,
                        "font_color": "#006400" if is_growing else ("#8B0000" if is_bad else "#404040")})
        _write_verdict_cell(ws, r, 6, gv, wb, v_props)
        _set_row(ws, r, gv, 28)
        r += 1
    return r + 2


def create_excel_report(raw_data: pd.DataFrame, date_from: str, date_to: str,
                        output_path: str, top_n: int = 10,
                        forecast_df: pd.DataFrame = None,
                        forecast_mean_map: dict[str, float] = None,
                        llm_recommendations: str = None,
                        llm_market_overview: str = None,
                        llm_expansion_note: str = None,
                        llm_verdicts: dict[str, str] = None,
                        all_queries: list[str] | None = None) -> str:
    """
    Создаёт форматированный Excel-отчёт.
    Форматирование строго по образцу yandex_analysis_exemple.xlsx.
    Возвращает (путь к файлу, данные для веб-графиков).
    """
    logger.info(f"Creating Excel report: {output_path}")

    from app.metrics import (
        calculate_metrics, compute_summary, ensure_queries_in_summary,
        make_blue_oceans, make_growing_markets
    )

    # ── Подготовка данных ──────────────────────────────────────
    analyzed = calculate_metrics(raw_data)
    summary = compute_summary(analyzed)
    if all_queries:
        summary = ensure_queries_in_summary(summary, all_queries)

    if forecast_mean_map:
        summary["forecast_mean"] = summary["query"].map(
            lambda q: forecast_mean_map.get(q, None)
        )
    else:
        summary["forecast_mean"] = None

    cols_data: list[str] = [
        "date", "query", "frequency", "freq_ma3", "growth_pct",
        "volatility", "seasonality_idx", "competition_score",
    ]
    sheet1_df = analyzed[cols_data].copy()
    sheet1_df["date"] = sheet1_df["date"].dt.strftime("%Y-%m-%d")

    try:
        df_dt = datetime.strptime(date_from[:10], "%Y-%m-%d")
        dt_dt = datetime.strptime(date_to[:10], "%Y-%m-%d")
        date_range_str = f"{df_dt.year}-{df_dt.month:02d} .. {dt_dt.year}-{dt_dt.month:02d}"
    except Exception:
        date_range_str = f"{date_from[:7]} .. {date_to[:7]}"

    # ── Таблицы "Голубые океаны" (3 варианта) ──────────────────
    blue_momentum = make_blue_oceans(summary, "growth_momentum")
    blue_mean = make_blue_oceans(summary, "mean_growth")
    blue_slope = make_blue_oceans(summary, "trend_slope")

    # ── Таблицы "Растущие рынки" (3 варианта) ──────────────────
    growing_momentum = make_growing_markets(summary, "growth_momentum")
    growing_mean = make_growing_markets(summary, "mean_growth")
    growing_slope = make_growing_markets(summary, "trend_slope")

    chart_payload = build_chart_payload(
        analyzed, summary, top_n, date_from, date_to, forecast_df=forecast_df)

    # ── Создание книги Excel ──────────────────────────────────
    wb = xlsxwriter.Workbook(output_path, {"nan_inf_to_errors": True})

    # Форматы с единым шрифтом Calibri
    hdr_base = _hdr_dict_base()
    hdr_fmt = wb.add_format(hdr_base)
    title_fmt = wb.add_format(_fmt_bold({"font_size": 14, "font_color": COLOR_TITLE,
                                          "bottom": 2, "bottom_color": COLOR_HEADER,
                                          "text_wrap": True}))
    conclusion_fmt = wb.add_format(_fmt_bold({"font_size": 14, "font_color": COLOR_TITLE,
                                               "bottom": 2, "bottom_color": COLOR_HEADER,
                                               "text_wrap": True}))
    guide_fmt = wb.add_format(_fmt({"text_wrap": True, "valign": "top", "font_size": FONT_SIZE}))
    guide_bold = wb.add_format(_fmt_bold({"text_wrap": True, "valign": "top", "font_size": FONT_SIZE}))
    num_fmt = wb.add_format(_fmt({"num_format": "# ##0.00", "border": 1, "font_size": FONT_SIZE}))
    int_fmt = wb.add_format(_fmt({"num_format": "# ##0", "border": 1, "font_size": FONT_SIZE}))
    pct_fmt = wb.add_format(_fmt({"num_format": "# ##0.00", "border": 1, "font_size": FONT_SIZE}))
    date_fmt = wb.add_format(_fmt({"num_format": "yyyy-mm-dd", "border": 1, "font_size": FONT_SIZE}))
    txt_fmt = wb.add_format(_fmt({"border": 1, "text_wrap": True, "valign": "vcenter", "font_size": FONT_SIZE}))
    green_fmt = wb.add_format(_fmt({"num_format": "0.000", "border": 1, "bg_color": COLOR_GREEN,
                                     "font_size": FONT_SIZE}))
    yellow_fmt = wb.add_format(_fmt({"num_format": "0.000", "border": 1, "bg_color": COLOR_YELLOW,
                                      "font_size": FONT_SIZE}))
    red_fmt = wb.add_format(_fmt({"num_format": "0.000", "border": 1, "bg_color": COLOR_RED,
                                   "font_size": FONT_SIZE}))
    green_vol_fmt = wb.add_format(_fmt({"num_format": "# ##0.0", "border": 1, "bg_color": COLOR_GREEN,
                                         "font_size": FONT_SIZE}))
    yellow_vol_fmt = wb.add_format(_fmt({"num_format": "# ##0.0", "border": 1, "bg_color": COLOR_YELLOW,
                                          "font_size": FONT_SIZE}))
    red_vol_fmt = wb.add_format(_fmt({"num_format": "# ##0.0", "border": 1, "bg_color": COLOR_RED,
                                       "font_size": FONT_SIZE}))
    seas_green_fmt = wb.add_format(_fmt({"num_format": "0.000", "border": 1, "bg_color": COLOR_GREEN,
                                          "font_size": FONT_SIZE}))
    seas_yellow_fmt = wb.add_format(_fmt({"num_format": "0.000", "border": 1, "bg_color": COLOR_YELLOW,
                                           "font_size": FONT_SIZE}))
    seas_red_fmt = wb.add_format(_fmt({"num_format": "0.000", "border": 1, "bg_color": COLOR_RED,
                                        "font_size": FONT_SIZE}))
    v_fmt_good = wb.add_format(_fmt({"border": 1, "text_wrap": True, "bold": True,
                                      "font_color": "#006400", "font_size": FONT_SIZE}))
    v_fmt_bad = wb.add_format(_fmt({"border": 1, "text_wrap": True, "bold": True,
                                     "font_color": "#8B0000", "font_size": FONT_SIZE}))
    v_fmt_neutral = wb.add_format(_fmt({"border": 1, "text_wrap": True, "bold": True,
                                         "font_color": "#404040", "font_size": FONT_SIZE}))
    # Формат для пустых строк-разделителей
    empty_fmt = wb.add_format(_fmt({"font_size": 4}))

    def comp_fmt(score: float):
        if score < 0.4:
            return green_fmt
        elif score > 0.6:
            return red_fmt
        else:
            return yellow_fmt

    def vol_fmt_h(v: float, label: str):
        if label == "Низкая":
            return green_vol_fmt
        elif label == "Высокая":
            return red_vol_fmt
        else:
            return yellow_vol_fmt

    def seas_fmt(v) -> object:
        if v is None or (isinstance(v, float) and (pd.isna(v) or np.isinf(v))):
            return txt_fmt
        if v > 1.2:
            return seas_red_fmt
        if v < 0.8:
            return seas_green_fmt
        return seas_yellow_fmt

    vol_values = summary["mean_volatility"]
    vol_lo = vol_values.quantile(0.33)
    vol_hi = vol_values.quantile(0.66)
    vol_lo_str = f"{vol_lo:,.1f}".replace(",", " ")
    vol_hi_str = f"{vol_hi:,.1f}".replace(",", " ")

    # ==========================================================
    # ЛИСТ 1: О файле
    # ==========================================================
    ws_info = wb.add_worksheet("О файле")
    ws_info.set_column(0, 0, 6)
    ws_info.set_column(1, 3, 70)  # Широкие колонки, чтобы текст не обрезался

    row = 0
    title_text = f"Анализ поискового спроса - Yandex Wordstat API\n{date_range_str}"
    ws_info.merge_range(row, 0, row, 3, title_text, title_fmt)
    _set_row(ws_info, row, title_text, 36)
    row += 2

    # Выводы
    conclusion_title = "ПРАКТИЧЕСКИЙ ВЫВОД ПО ВАШИМ ДАННЫМ"
    ws_info.merge_range(row, 0, row, 3, conclusion_title, conclusion_fmt)
    _set_row(ws_info, row, conclusion_title, 28)
    row += 2

    s = summary
    top_demand = s.nlargest(3, "sum_frequency")
    top_growth_mom = s.nlargest(3, "growth_momentum")
    low_comp_growth_mom = s[(s["competition_score"] < 0.4) & (s["growth_momentum"] > 0)] \
        .sort_values("growth_momentum", ascending=False)
    low_comp = s[s["competition_score"] < 0.4].sort_values("competition_score")

    conclusions = []
    conclusions.append(("Самая востребованная ниша (емкость рынка):",
        _format_top_rows(top_demand, 3,
            lambda r: f"{r['query']} - {r['sum_frequency']:,.0f} показов")))
    conclusions.append(("Самая перспективная растущая ниша (за последние 6 мес.):",
        _format_top_rows(top_growth_mom, 3,
            lambda r: f"{r['query']} - рост {r['growth_momentum']:+.2f}% в месяц")))
    if not low_comp_growth_mom.empty:
        best = low_comp_growth_mom.iloc[0]
        conclusions.append(("Лучшее для старта (низкая конкуренция + рост = ГОЛУБЫЕ ОКЕАНЫ):",
            f"{best['query']} - конкуренция {best['competition_score']:.2f}, "
            f"рост (momentum) {best['growth_momentum']:+.2f}%/мес"))
    else:
        conclusions.append(("Лучшее для старта (низкая конкуренция + рост = ГОЛУБЫЕ ОКЕАНЫ):",
            "Не найдено запросов, одновременно удовлетворяющих условиям."))
    if not low_comp.empty:
        lowest = low_comp.iloc[0]
        conclusions.append(("Самая низкая конкуренция (легкий вход):",
            f"{lowest['query']} - индекс конкуренции {lowest['competition_score']:.2f}"))

    for label, text in conclusions:
        ws_info.write(row, 0, ">", guide_bold)
        ws_info.merge_range(row, 1, row, 3, label, guide_bold)
        _set_row(ws_info, row, label, 22)
        row += 1
        ws_info.merge_range(row, 1, row, 3, text, guide_fmt)
        _set_row(ws_info, row, text, 22)
        row += 2

    # Блоки YandexGPT
    for block_title, block_text in (
        ("РАСШИРЕНИЕ ЗАПРОСОВ (YandexGPT)", llm_expansion_note),
        ("ОБЗОР РЫНКА (YandexGPT)", llm_market_overview),
        ("РЕКОМЕНДАЦИИ (YandexGPT)", llm_recommendations),
    ):
        if not block_text:
            continue
        ws_info.merge_range(row, 0, row, 3, block_title, conclusion_fmt)
        _set_row(ws_info, row, block_title, 28)
        row += 2
        ws_info.merge_range(row, 1, row, 3, block_text, guide_fmt)
        _set_row(ws_info, row, block_text, 40)
        row += 3

    # Описание листов
    sheets_title = "ОПИСАНИЕ ЛИСТОВ ФАЙЛА"
    ws_info.merge_range(row, 0, row, 3, sheets_title, conclusion_fmt)
    _set_row(ws_info, row, sheets_title, 28)
    row += 2

    sheets_desc = [
        ("Лист Данные", "Сырые данные по каждому запросу помесячно."),
        ("Лист Сводка по запросам",
         "Итоговые метрики по каждому запросу, включая все 4 метрики роста: "
         "средний за последние 6 мес., средний за последний год, "
         "по наклону линейного тренда, средний за весь период"),
        ("Лист График",
         f"Две редактируемые диаграммы Excel: топ-{top_n} запросов и все сравниваемые запросы."),
        ("Листы Голубые океаны (3 шт)",
         "3 варианта оценки низкоконкурентных ниш с ростом: "
         "по momentum (свежий тренд), по mean (весь период), "
         "по slope (линейный тренд)."),
        ("Листы Растущие рынки (3 шт)",
         "3 варианта оценки растущих рынков: "
         "по momentum (свежий тренд), по mean (весь период), "
         "по slope (линейный тренд)."),
    ]
    for s_title, desc in sheets_desc:
        ws_info.write(row, 1, s_title, guide_bold)
        _set_row(ws_info, row, s_title, 22)
        ws_info.merge_range(row + 1, 1, row + 1, 3, desc, guide_fmt)
        _set_row(ws_info, row + 1, desc, 22)
        row += 3

    # Методология
    meth_title = "МЕТОДОЛОГИЯ, ТЕРМИНЫ И ОПРЕДЕЛЕНИЯ"
    ws_info.merge_range(row, 0, row, 3, meth_title, conclusion_fmt)
    _set_row(ws_info, row, meth_title, 28)
    row += 2

    methodology = [
        ("Источник данных",
         f"Yandex Search API v2 (Wordstat GetDynamics).\n"
         f"Период: {date_range_str} ({date_from[:10]} .. {date_to[:10]}).\n"
         f"Регион: вся Россия."),
        ("Показы",
         "Абсолютное количество показов запроса в Яндексе за месяц.\n"
         "Чем выше - тем больше людей ищут этот товар/услугу."),
        ("Средний рост за последние 6 мес.",
         "Средний помесячный темп роста за последние 6 месяцев.\n"
         "Расчет: среднее growth_pct по последним 7 записям (6 переходов).\n"
         "Показывает свежий тренд, не подвержен давним скачкам."),
        ("Средний рост за последний год",
         "Средний помесячный темп роста за последние 12 месяцев.\n"
         "Расчет: среднее growth_pct по последним 13 записям (12 переходов).\n"
         "Показывает тренд за год."),
        ("Наклон линейного тренда",
         "Абсолютный прирост частотности в месяц (линейная регрессия).\n"
         "Расчет: numpy polyfit(x, frequency, 1), где x = 0,1,2...\n"
         "Показывает, на сколько показов в месяц в среднем меняется спрос."),
        ("Средний рост за весь период",
         "Средний помесячный темп роста за весь период наблюдений.\n"
         "Подвержен влиянию давних скачков — менее точен для текущего тренда."),
        ("Скользящее среднее за 3 мес.",
         "Сглаживает сезонные и случайные колебания.\n"
         "Показывает тренд чище, чем сырая частотность.\n"
         "Расчет: среднее значение за текущий и 2 предыдущих месяца."),
        ("Темп роста, %",
         "Помесячный прирост/падение частотности в процентах.\n"
         "Расчет: ((frequency_мес_N / frequency_мес_N-1) - 1) x 100.\n"
         "Положительное = спрос растет, отрицательное = падает."),
        ("Волатильность",
         f"Стандартное отклонение частотности за весь период.\n"
         f"Границы данного отчета:\n"
         f"  🟢 Низкая (≤ {vol_lo_str})\n"
         f"  🟡 Средняя ({vol_lo_str}–{vol_hi_str})\n"
         f"  🔴 Высокая (≥ {vol_hi_str})"),
        ("Индекс сезонности",
         "Отношение значения текущего месяца к среднегодовому.\n"
         "Расчет: frequency / средняя_частотность_за_год.\n"
         "  🔴 > 1.2 = пик сезона (спрос выше нормы на 20%+)\n"
         "  🟡 0.8–1.2 = норма (ровно как в среднем за год)\n"
         "  🟢 < 0.8 = провал (спрос ниже нормы на 20%+)\n"
         "Позволяет увидеть сезонные паттерны и планировать рекламу."),
        ("Индекс конкуренции, 0..1",
         "0.5 × норм.частотность + 0.5 × (1 − норм.волатильность).\n"
         f"Границы: 🟢 < 0.40, 🟡 0.40–0.60, 🔴 > 0.60"),
        ("Голубые океаны",
         "Составной индекс: 50% × (1 − Индекс конкуренции) + 50% × норм.growth.\n"
         "Чем выше — тем перспективнее ниша для входа."),
    ]
    for m_title, desc in methodology:
        ws_info.write(row, 1, m_title, guide_bold)
        _set_row(ws_info, row, m_title, 22)
        desc_row = row + 1
        guide_props = _fmt({"text_wrap": True, "valign": "top", "font_size": FONT_SIZE})
        if any(e in desc for e in _HEADER_EMOJIS):
            ws_info.merge_range(desc_row, 1, desc_row, 3, "", guide_fmt)
            runs = template_runs_for_guide_line(desc)
            if runs:
                write_rich_cell(ws_info, desc_row, 1, runs, wb, guide_props)
            else:
                ws_info.write(desc_row, 1, desc, wb.add_format(guide_props))
        else:
            ws_info.merge_range(desc_row, 1, desc_row, 3, desc, guide_fmt)
        _set_row(ws_info, desc_row, desc, 22)
        row += 4

    # ==========================================================
    # ЛИСТ 2: Данные
    # ==========================================================
    ws1 = wb.add_worksheet("Данные")
    data_title = f"Сырые данные: каждая строка - один месяц x один запрос\n{date_range_str}"
    ws1.merge_range(0, 0, 0, 7, data_title, title_fmt)
    _set_row(ws1, 0, data_title, 36)
    data_headers = get_table_headers("Данные")
    _write_header_row(ws1, 1, data_headers, wb, hdr_base, get_table_header_runs("Данные"),
                      fixed_height=HDR_TABLE_HEIGHT_TALL)
    ws1.set_column(0, 0, 14)
    ws1.set_column(1, 1, 42)
    ws1.set_column(2, 7, 26)

    for i, (_, row_data) in enumerate(sheet1_df.iterrows()):
        r = i + 2
        ws1.write(r, 0, row_data["date"], date_fmt)
        ws1.write(r, 1, row_data["query"], txt_fmt)
        ws1.write(r, 2, row_data["frequency"], int_fmt)
        ws1.write(r, 3, round(row_data["freq_ma3"], 2), num_fmt)
        gp = row_data["growth_pct"]
        if pd.notna(gp):
            ws1.write(r, 4, round(float(gp), 2), pct_fmt)
        else:
            ws1.write(r, 4, "—", txt_fmt)
        vol = row_data["volatility"]
        if pd.notna(vol):
            q = row_data["query"]
            q_label = summary.loc[summary["query"] == q, "vol_label"].values
            vlabel = q_label[0] if len(q_label) > 0 else "Средняя"
            ws1.write(r, 5, round(float(vol), 2), vol_fmt_h(float(vol), vlabel))
        else:
            ws1.write(r, 5, "—", txt_fmt)
        seas = row_data["seasonality_idx"]
        if pd.notna(seas):
            ws1.write(r, 6, round(float(seas), 3), seas_fmt(float(seas)))
        else:
            ws1.write(r, 6, "—", txt_fmt)
        score = row_data["competition_score"]
        if pd.notna(score):
            ws1.write(r, 7, round(float(score), 3), comp_fmt(float(score)))
        else:
            ws1.write(r, 7, "—", txt_fmt)
        _set_row(ws1, r)

    # ==========================================================
    # ЛИСТ 3: Сводка по запросам
    # ==========================================================
    ws2 = wb.add_worksheet("Сводка по запросам")
    sum_title = f"Итоговые метрики, от популярных к нишевым\n{date_range_str}"
    sum_headers = list(get_table_headers("Сводка по запросам"))
    if _col_by_header(sum_headers, "прогноз") is None:
        sum_headers.append("Прогноз сред.\nпоказов/мес\n(6 мес.)")
    sum_last_col = max(0, len(sum_headers) - 1)
    ws2.merge_range(0, 0, 0, sum_last_col, sum_title, title_fmt)
    _set_row(ws2, 0, sum_title, 36)
    sum_runs = get_table_header_runs("Сводка по запросам")
    if sum_runs and len(sum_runs) < len(sum_headers):
        sum_runs = list(sum_runs) + [[(sum_headers[-1], None)]]
    _write_header_row(ws2, 1, sum_headers, wb, hdr_base, sum_runs,
                      fixed_height=HDR_TABLE_HEIGHT_TALL)
    sc = _summary_col_map(sum_headers)
    for c in range(len(sum_headers)):
        if c == sc.get("query"):
            ws2.set_column(c, c, 42)
        elif c == sc.get("num"):
            ws2.set_column(c, c, 6)
        else:
            ws2.set_column(c, c, 22)

    for i, (_, row_data) in enumerate(summary.iterrows()):
        r = i + 2

        def wcol(key: str, *args, **kwargs) -> None:
            c = sc.get(key)
            if c is not None:
                ws2.write(r, c, *args, **kwargs)

        if sc.get("num") is not None:
            ws2.write(r, sc["num"], i + 1, int_fmt)
        wcol("query", row_data["query"], txt_fmt)
        wcol("mean_freq", round(row_data["mean_frequency"], 1), int_fmt)
        wcol("sum_freq", int(row_data["sum_frequency"]), int_fmt)
        wcol("std", round(row_data["std_frequency"], 1), num_fmt)
        wcol("mean_growth", round(row_data["mean_growth"], 2), num_fmt)
        wcol("momentum", round(row_data["growth_momentum"], 2), num_fmt)
        wcol("growth_year", round(row_data["growth_year"], 2), num_fmt)
        wcol("trend", round(row_data["trend_slope"], 2), num_fmt)
        vol = round(row_data["mean_volatility"], 1)
        label = row_data["vol_label"]
        wcol("volatility", vol, vol_fmt_h(vol, label))
        score = row_data["competition_score"]
        if sc.get("competition") is not None:
            if pd.notna(score):
                ws2.write(r, sc["competition"], round(float(score), 3),
                          comp_fmt(float(score)))
            else:
                ws2.write(r, sc["competition"], "—", txt_fmt)
        forecast_mean = row_data.get("forecast_mean", None)
        if sc.get("forecast") is not None:
            if pd.notna(forecast_mean) and forecast_mean is not None:
                ws2.write(r, sc["forecast"], round(float(forecast_mean), 0), int_fmt)
            else:
                ws2.write(r, sc["forecast"], "—", txt_fmt)
        _set_row(ws2, r)

    # ==========================================================
    # ЛИСТ 4: Прогноз
    # ==========================================================
    ws_forecast = wb.add_worksheet("Прогноз")
    f_title = f"Прогноз Prophet на 6 месяцев\n{date_range_str}"
    ws_forecast.merge_range(0, 0, 0, 4, f_title, title_fmt)
    _set_row(ws_forecast, 0, f_title, 36)

    if forecast_df is not None and not forecast_df.empty:
        f_headers = get_table_headers("Прогноз")
        _write_header_row(ws_forecast, 1, f_headers, wb, hdr_base, get_table_header_runs("Прогноз"))
        ws_forecast.set_column(0, 0, 42)
        ws_forecast.set_column(1, 1, 14)
        ws_forecast.set_column(2, 4, 22)
        ws_forecast.set_column(5, 5, 40)

        for ri, (_, row_data) in enumerate(forecast_df.iterrows()):
            r = ri + 2
            ws_forecast.write(r, 0, row_data["query"], txt_fmt)
            try:
                dt_val = row_data["date"]
                if hasattr(dt_val, "strftime"):
                    ws_forecast.write(r, 1, dt_val.strftime("%Y-%m-%d"), date_fmt)
                else:
                    ws_forecast.write(r, 1, str(dt_val)[:10], txt_fmt)
            except Exception:
                ws_forecast.write(r, 1, str(row_data["date"])[:10], txt_fmt)

            yhat = row_data["yhat"]
            yhat_l = row_data["yhat_lower"]
            yhat_u = row_data["yhat_upper"]
            warning = row_data.get("forecast_warning", "")

            if pd.notna(yhat) and yhat is not None:
                ws_forecast.write(r, 2, round(float(yhat), 0), int_fmt)
                ws_forecast.write(r, 3, round(float(yhat_l), 0) if pd.notna(yhat_l) else 0, int_fmt)
                ws_forecast.write(r, 4, round(float(yhat_u), 0) if pd.notna(yhat_u) else 0, int_fmt)
            else:
                ws_forecast.write(r, 2, "—", txt_fmt)
                ws_forecast.write(r, 3, "—", txt_fmt)
                ws_forecast.write(r, 4, "—", txt_fmt)

            ws_forecast.write(r, 5, warning if warning else "", txt_fmt)
    else:
        msg = prophet_status_message() or (
            "Прогноз не построен. Установите Prophet или проверьте данные."
        )
        ws_forecast.merge_range(
            1, 0, 3, 5, msg,
            wb.add_format(_fmt({"text_wrap": True, "font_size": FONT_SIZE,
                                "valign": "vcenter", "border": 1})),
        )

    # ==========================================================
    # ЛИСТ 5: График (нативные диаграммы Excel)
    # ==========================================================
    ws3 = wb.add_worksheet("График")
    _write_charts_sheet(
        wb, ws3, analyzed, summary, top_n, date_range_str,
        title_fmt, hdr_fmt, txt_fmt, num_fmt, forecast_df=forecast_df,
    )

    # ==========================================================
    # ЛИСТ 6: Голубые океаны
    # ==========================================================
    ws4 = wb.add_worksheet("Голубые океаны")
    ws4.set_column(0, 0, 6)
    ws4.set_column(1, 1, 42)
    ws4.set_column(2, 5, 28)
    ws4.set_column(6, 6, 54)
    ws4.set_column(7, 7, 54)

    next_row = _write_ocean_block(ws4, wb, blue_momentum, 0,
        f"Блок 1: Низкая конкуренция + рост ({GROWTH_COL_TITLE['growth_momentum']}) | {date_range_str}",
        date_range_str, hdr_base, title_fmt, txt_fmt, int_fmt, num_fmt, comp_fmt,
        v_fmt_good, v_fmt_bad, v_fmt_neutral, "growth_momentum", llm_verdicts)

    next_row = _write_ocean_block(ws4, wb, blue_mean, next_row,
        f"Блок 2: Низкая конкуренция + рост ({GROWTH_COL_TITLE['mean_growth']}) | {date_range_str}",
        date_range_str, hdr_base, title_fmt, txt_fmt, int_fmt, num_fmt, comp_fmt,
        v_fmt_good, v_fmt_bad, v_fmt_neutral, "mean_growth", llm_verdicts)

    _write_ocean_block(ws4, wb, blue_slope, next_row,
        f"Блок 3: Низкая конкуренция + рост ({GROWTH_COL_TITLE['trend_slope']}) | {date_range_str}",
        date_range_str, hdr_base, title_fmt, txt_fmt, int_fmt, num_fmt, comp_fmt,
        v_fmt_good, v_fmt_bad, v_fmt_neutral, "trend_slope", llm_verdicts)

    # ==========================================================
    # ЛИСТ 7: Растущие рынки
    # ==========================================================
    ws5 = wb.add_worksheet("Растущие рынки")
    ws5.set_column(0, 0, 6)
    ws5.set_column(1, 1, 42)
    ws5.set_column(2, 6, 28)
    ws5.set_column(7, 7, 44)

    next_row = _write_growing_block(ws5, wb, growing_momentum, 0,
        f"Блок 1: Растущие рынки ({GROWTH_COL_TITLE['growth_momentum']}) | {date_range_str}",
        date_range_str, hdr_base, title_fmt, txt_fmt, int_fmt, num_fmt, comp_fmt, vol_fmt_h,
        v_fmt_good, v_fmt_bad, v_fmt_neutral, "growth_momentum")

    next_row = _write_growing_block(ws5, wb, growing_mean, next_row,
        f"Блок 2: Растущие рынки ({GROWTH_COL_TITLE['mean_growth']}) | {date_range_str}",
        date_range_str, hdr_base, title_fmt, txt_fmt, int_fmt, num_fmt, comp_fmt, vol_fmt_h,
        v_fmt_good, v_fmt_bad, v_fmt_neutral, "mean_growth")

    _write_growing_block(ws5, wb, growing_slope, next_row,
        f"Блок 3: Растущие рынки ({GROWTH_COL_TITLE['trend_slope']}) | {date_range_str}",
        date_range_str, hdr_base, title_fmt, txt_fmt, int_fmt, num_fmt, comp_fmt, vol_fmt_h,
        v_fmt_good, v_fmt_bad, v_fmt_neutral, "trend_slope")

    # ── Закрытие книги ─────────────────────────────────────────
    wb.close()

    logger.info(f"Report saved: {output_path}")
    return output_path, chart_payload