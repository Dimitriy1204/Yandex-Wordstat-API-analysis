#!/usr/bin/env python3
"""Подготовка данных для Excel-диаграмм и веб-графиков."""
from __future__ import annotations

from datetime import datetime

import pandas as pd

# Палитра в духе Excel / Office
CHART_COLORS = [
    "#4472C4", "#ED7D31", "#A5A5A5", "#FFC000", "#5B9BD5",
    "#70AD47", "#264478", "#9E480E", "#636363", "#997300",
    "#C00000", "#548235", "#1F4E79", "#843C0C", "#375623",
]


def _truncate_label(text: str, max_len: int = 40) -> str:
    t = (text or "").strip()
    if len(t) <= max_len:
        return t
    return t[: max_len - 1] + "…"


def _all_timeline_dates(
    analyzed: pd.DataFrame,
    forecast_df: pd.DataFrame | None,
    queries: list[str],
) -> list[pd.Timestamp]:
    dates: list[pd.Timestamp] = []
    if not analyzed.empty:
        sub = analyzed[analyzed["query"].isin(queries)]
        if not sub.empty:
            dates.extend(pd.to_datetime(sub["date"]).tolist())
    if forecast_df is not None and not forecast_df.empty:
        sub = forecast_df[forecast_df["query"].isin(queries)]
        if not sub.empty and "date" in sub.columns:
            dates.extend(pd.to_datetime(sub["date"]).tolist())
    if not dates:
        return []
    return sorted(set(dates))


def prepare_chart_series_with_forecast(
    raw_data: pd.DataFrame,
    queries: list[str],
    forecast_df: pd.DataFrame | None = None,
) -> tuple[list[str], dict[str, list], dict[str, list] | None]:
    """
    Метки месяцев, факт и прогноз по каждому запросу.
    Прогноз: None на исторических месяцах, факт: None на будущих.
    """
    if not queries:
        return [], {}, None

    dates = _all_timeline_dates(raw_data, forecast_df, queries)
    if not dates:
        return [], {}, None

    labels = [pd.Timestamp(d).strftime("%Y-%m") for d in dates]

    hist = raw_data[raw_data["query"].isin(queries)].copy() if not raw_data.empty else raw_data
    if not hist.empty:
        hist["date"] = pd.to_datetime(hist["date"])

    fc: dict[str, list] | None = {} if forecast_df is not None and not forecast_df.empty else None
    if fc is not None:
        forecast_df = forecast_df[forecast_df["query"].isin(queries)].copy()
        forecast_df["date"] = pd.to_datetime(forecast_df["date"])

    series: dict[str, list] = {}
    for q in queries:
        fact_vals: list = []
        fc_vals: list = []
        q_hist = hist[hist["query"] == q].set_index("date")["frequency"] if not hist.empty else pd.Series(dtype=float)
        q_fc = None
        if fc is not None and not forecast_df.empty:
            qf = forecast_df[forecast_df["query"] == q]
            if not qf.empty and "yhat" in qf.columns:
                q_fc = qf.set_index("date")["yhat"]

        for d in dates:
            if d in q_hist.index:
                v = q_hist.loc[d]
                fact_vals.append(float(v) if pd.notna(v) else None)
            else:
                fact_vals.append(None)
            if q_fc is not None and d in q_fc.index:
                v = q_fc.loc[d]
                fc_vals.append(float(v) if pd.notna(v) else None)
            else:
                fc_vals.append(None)
        series[q] = fact_vals
        if fc is not None and any(v is not None for v in fc_vals):
            fc[q] = fc_vals

    if fc is not None and not any(any(v is not None for v in fc[q]) for q in fc):
        fc = None

    return labels, series, fc


def prepare_chart_series(
    raw_data: pd.DataFrame,
    queries: list[str],
) -> tuple[list[str], dict[str, list[float | None]]]:
    labels, series, _ = prepare_chart_series_with_forecast(raw_data, queries, None)
    return labels, series


def build_chart_payload(
    analyzed: pd.DataFrame,
    summary: pd.DataFrame,
    top_n: int,
    date_from: str,
    date_to: str,
    forecast_df: pd.DataFrame | None = None,
) -> dict:
    """JSON для /api/chart-data и внутренних диаграмм."""
    top_n = max(1, int(top_n or 10))
    ordered = summary["query"].tolist() if not summary.empty else []
    top_queries = ordered[:top_n]
    all_queries = ordered if ordered else analyzed["query"].drop_duplicates().tolist()

    labels_top, series_top, fc_top = prepare_chart_series_with_forecast(
        analyzed, top_queries, forecast_df)
    labels_all, series_all, fc_all = prepare_chart_series_with_forecast(
        analyzed, all_queries, forecast_df)

    def _datasets(
        queries: list[str],
        series_map: dict[str, list],
        forecast_map: dict[str, list] | None,
    ) -> list[dict]:
        out = []
        for i, q in enumerate(queries):
            color = CHART_COLORS[i % len(CHART_COLORS)]
            out.append({
                "label": _truncate_label(q, 60),
                "query": q,
                "data": series_map.get(q, []),
                "borderColor": color,
                "backgroundColor": color + "33",
                "borderWidth": 2,
                "pointRadius": 4,
                "pointHoverRadius": 7,
                "tension": 0.15,
                "fill": False,
                "forecast": False,
            })
            if forecast_map and q in forecast_map:
                out.append({
                    "label": _truncate_label(q, 50) + " (прогноз)",
                    "query": q,
                    "data": forecast_map.get(q, []),
                    "borderColor": color,
                    "backgroundColor": "transparent",
                    "borderWidth": 2,
                    "borderDash": [6, 4],
                    "pointRadius": 3,
                    "pointHoverRadius": 6,
                    "tension": 0.15,
                    "fill": False,
                    "forecast": True,
                })
        return out

    try:
        d0 = datetime.strptime(date_from[:10], "%Y-%m-%d").strftime("%d.%m.%Y")
        d1 = datetime.strptime(date_to[:10], "%Y-%m-%d").strftime("%d.%m.%Y")
        period = f"{d0} — {d1}"
    except ValueError:
        period = f"{date_from[:10]} — {date_to[:10]}"

    has_fc = forecast_df is not None and not forecast_df.empty

    return {
        "period": period,
        "top_n": top_n,
        "has_forecast": has_fc,
        "top_chart": {
            "title": f"Динамика частотности — топ-{top_n} запросов",
            "labels": labels_top,
            "datasets": _datasets(top_queries, series_top, fc_top),
        },
        "all_chart": {
            "title": "Динамика частотности — все сравниваемые запросы",
            "labels": labels_all,
            "datasets": _datasets(all_queries, series_all, fc_all),
        },
    }
