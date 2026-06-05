#!/usr/bin/env python3
"""Расчёт метрик поискового спроса."""
from __future__ import annotations
import logging
import numpy as np
import pandas as pd
from typing import Any
logger = logging.getLogger("wordstat.metrics")

_BASE_COLS = ("query", "date", "frequency")

def calculate_metrics(df: pd.DataFrame) -> pd.DataFrame:
    """Добавляет метрики: скользящее среднее, рост, волатильность, сезонность, конкуренцию."""
    cols = [c for c in _BASE_COLS if c in df.columns]
    if len(cols) < 3:
        raise ValueError("Ожидаются колонки query, date, frequency")
    df = df[cols].copy().sort_values(["query", "date"]).reset_index(drop=True)
    if df.empty:
        for col in ("freq_ma3", "growth_pct", "volatility", "seasonality_idx", "competition_score"):
            df[col] = pd.Series(dtype=float)
        return df
    # Скользящее среднее за 3 мес
    df["freq_ma3"] = df.groupby("query")["frequency"].transform(lambda x: x.rolling(3, min_periods=1).mean())
    # Помесячный рост %
    df["growth_pct"] = df.groupby("query")["frequency"].transform(lambda x: x.pct_change() * 100)
    # Волатильность (скользящее стандартное отклонение)
    df["volatility"] = df.groupby("query")["frequency"].transform(lambda x: x.rolling(12, min_periods=1).std())
    # Индекс сезонности
    yearly_mean = df.groupby(["query", df["date"].dt.year])["frequency"].transform("mean")
    df["seasonality_idx"] = df["frequency"] / yearly_mean.replace(0, np.nan)
    # Индекс конкуренции
    query_stats = df.groupby("query").agg(
        mean_freq=("frequency", "mean"),
        std_freq=("frequency", "std")
    ).reset_index()
    query_stats["std_freq"] = query_stats["std_freq"].fillna(0)
    max_mean = query_stats["mean_freq"].max()
    max_std = query_stats["std_freq"].max()
    if max_mean > 0: query_stats["norm_freq"] = query_stats["mean_freq"] / max_mean
    else: query_stats["norm_freq"] = 0
    if max_std > 0: query_stats["norm_vol"] = query_stats["std_freq"] / max_std
    else: query_stats["norm_vol"] = 0
    query_stats["competition_score"] = 0.5 * query_stats["norm_freq"] + 0.5 * (1 - query_stats["norm_vol"])
    df = df.merge(query_stats[["query", "competition_score"]], on="query", how="left")
    return df

def compute_summary(df: pd.DataFrame) -> pd.DataFrame:
    """Сводка по запросам: средние, суммы, рост, тренд."""
    if df.empty:
        return pd.DataFrame(columns=[
            "query", "mean_frequency", "sum_frequency", "std_frequency", "mean_growth",
            "mean_volatility", "competition_score", "trend_slope", "growth_momentum",
            "growth_year", "vol_label",
        ])
    if "competition_score" not in df.columns:
        df = calculate_metrics(df)
    summary = df.groupby("query").agg(
        mean_frequency=("frequency", "mean"),
        sum_frequency=("frequency", "sum"),
        std_frequency=("frequency", "std"),
        mean_growth=("growth_pct", "mean"),
        mean_volatility=("volatility", "mean"),
        competition_score=("competition_score", "first"),
    ).reset_index()
    summary["std_frequency"] = summary["std_frequency"].fillna(0)
    summary["mean_volatility"] = summary["mean_volatility"].fillna(0)
    # Тренд (наклон линейной регрессии)
    trend_data = []
    for q in df["query"].unique():
        qd = df[df["query"] == q].sort_values("date")
        x = np.arange(len(qd))
        y = qd["frequency"].values
        if len(x) > 1:
            slope = np.polyfit(x, y, 1)[0]
        else:
            slope = 0.0
        trend_data.append({"query": q, "trend_slope": slope})
    trend_df = pd.DataFrame(trend_data)
    summary = summary.merge(trend_df, on="query", how="left")
    # growth_momentum (средний рост за последние 6 мес = 7 записей)
    momentum_data = []
    for q in df["query"].unique():
        qd = df[df["query"] == q].dropna(subset=["growth_pct"]).sort_values("date")
        recent6 = qd.tail(7) if len(qd) >= 7 else qd
        mom = recent6["growth_pct"].mean() if not recent6.empty else 0.0
        momentum_data.append({"query": q, "growth_momentum": mom})
    momentum_df = pd.DataFrame(momentum_data)
    summary = summary.merge(momentum_df, on="query", how="left")
    # growth_year (средний рост за последние 12 мес = 13 записей)
    year_data = []
    for q in df["query"].unique():
        qd = df[df["query"] == q].dropna(subset=["growth_pct"]).sort_values("date")
        recent12 = qd.tail(13) if len(qd) >= 13 else qd
        yr = recent12["growth_pct"].mean() if not recent12.empty else 0.0
        year_data.append({"query": q, "growth_year": yr})
    year_df = pd.DataFrame(year_data)
    summary = summary.merge(year_df, on="query", how="left")
    summary = summary.fillna(0)
    # Волатильность метки
    vol_values = summary["mean_volatility"]
    vol_lo = vol_values.quantile(0.33)
    vol_hi = vol_values.quantile(0.66)
    def vol_label(v):
        if v <= vol_lo: return "Низкая"
        elif v >= vol_hi: return "Высокая"
        else: return "Средняя"
    summary["vol_label"] = summary["mean_volatility"].apply(vol_label)
    summary = summary.sort_values("sum_frequency", ascending=False).reset_index(drop=True)
    return summary


def ensure_queries_in_summary(summary: pd.DataFrame, queries: list[str]) -> pd.DataFrame:
    """Добавляет в сводку запросы без данных Wordstat (нулевые метрики)."""
    if not queries:
        return summary
    present: set[str] = set()
    if summary is not None and not summary.empty and "query" in summary.columns:
        present = {str(q).strip().lower() for q in summary["query"]}

    missing_rows: list[dict[str, Any]] = []
    for q in queries:
        q = str(q).strip()
        if not q or q.lower() in present:
            continue
        present.add(q.lower())
        missing_rows.append({
            "query": q,
            "mean_frequency": 0.0,
            "sum_frequency": 0.0,
            "std_frequency": 0.0,
            "mean_growth": 0.0,
            "mean_volatility": 0.0,
            "competition_score": 0.0,
            "trend_slope": 0.0,
            "growth_momentum": 0.0,
            "growth_year": 0.0,
            "vol_label": "Низкая",
        })

    if not missing_rows:
        return summary
    extra_df = pd.DataFrame(missing_rows)
    if summary is None or summary.empty:
        return extra_df
    return pd.concat([summary, extra_df], ignore_index=True).fillna(0)


def make_blue_oceans(summary: pd.DataFrame, growth_col: str) -> pd.DataFrame:
    """Голубые океаны: рейтинг по низкой конкуренции и росту (все запросы, как в шаблоне)."""
    df = summary.copy()
    g = df[growth_col]
    gmax = g.max() if len(g) else 0
    df["blue_score"] = (1 - df["competition_score"]) * 50 + (
        (g - g.min()) / (gmax - g.min() + 1e-9) * 50 if gmax > g.min() else 0
    )
    return df.sort_values("blue_score", ascending=False).reset_index(drop=True)

def make_growing_markets(summary: pd.DataFrame, growth_col: str) -> pd.DataFrame:
    """Растущие рынки: все запросы, отсортированные по метрике роста."""
    return summary.sort_values(growth_col, ascending=False).reset_index(drop=True)

def blue_verdict(competition, growth):
    if competition < 0.3 and growth > 5: return "🟢 Отличная ниша для входа"
    elif competition < 0.4 and growth > 2: return "🟢 Хорошая ниша"
    elif competition < 0.5 and growth > 0: return "🟡 Средняя"
    else: return "🔴 Высокая конкуренция или падение"

def growth_verdict(growth):
    if growth > 5: return "🟢 Быстрорастущий рынок"
    elif growth > 2: return "🟢 Растущий рынок"
    elif growth > 0: return "🟡 Слабый рост"
    else: return "🔴 Падающий рынок"