#!/usr/bin/env python3
"""Модуль прогнозирования временных рядов с помощью Prophet."""
from __future__ import annotations
import logging
import warnings
from typing import Optional
import numpy as np
import pandas as pd

logger = logging.getLogger("wordstat.prophet")
warnings.filterwarnings("ignore", category=FutureWarning)

_PROPhet_IMPORT_ERROR: str = ""


def _check_prophet_available() -> bool:
    global _PROPhet_IMPORT_ERROR
    try:
        from prophet import Prophet  # noqa: F401
        _PROPhet_IMPORT_ERROR = ""
        return True
    except Exception as e:
        _PROPhet_IMPORT_ERROR = str(e)
        logger.warning("Prophet недоступен: %s", e)
        return False


def _prepare_prophet_df(qdf: pd.DataFrame) -> pd.DataFrame:
    prophet_df = qdf[["date", "frequency"]].rename(columns={"date": "ds", "frequency": "y"})
    ds = prophet_df["ds"]
    if hasattr(ds.dtype, "tz") and ds.dtype.tz is not None:
        prophet_df["ds"] = ds.dt.tz_localize(None)
    else:
        prophet_df["ds"] = pd.to_datetime(ds).dt.tz_localize(None)
    return prophet_df


def forecast_queries(df: pd.DataFrame, periods: int = 6, min_data_points: int = 12) -> Optional[pd.DataFrame]:
    if not _check_prophet_available():
        return None
    from prophet import Prophet

    logger.info("Prophet: прогноз на %s месяцев для каждого запроса", periods)
    forecast_rows = []

    for query in df["query"].unique():
        qdf = df[df["query"] == query].copy().sort_values("date")
        if len(qdf) < min_data_points:
            last_date = qdf["date"].max() if not qdf.empty else pd.Timestamp.now()
            for p in range(1, periods + 1):
                forecast_rows.append({
                    "query": query,
                    "date": last_date + pd.DateOffset(months=p),
                    "yhat": None, "yhat_lower": None, "yhat_upper": None,
                    "forecast_warning": "Недостаточно данных для прогноза Prophet",
                })
            continue
        prophet_df = _prepare_prophet_df(qdf)
        try:
            model = Prophet(
                yearly_seasonality=True,
                weekly_seasonality=False,
                daily_seasonality=False,
                seasonality_mode="multiplicative",
                changepoint_prior_scale=0.05,
                seasonality_prior_scale=10.0,
            )
            model.fit(prophet_df)
            future = model.make_future_dataframe(periods=periods, freq="MS")
            forecast_tail = model.predict(future).tail(periods)
            for _, row in forecast_tail.iterrows():
                forecast_rows.append({
                    "query": query,
                    "date": row["ds"],
                    "yhat": row["yhat"],
                    "yhat_lower": row["yhat_lower"],
                    "yhat_upper": row["yhat_upper"],
                    "forecast_warning": "",
                })
        except Exception as e:
            logger.error("  '%s': ошибка Prophet: %s", query, e)
            last_date = qdf["date"].max()
            for p in range(1, periods + 1):
                forecast_rows.append({
                    "query": query,
                    "date": last_date + pd.DateOffset(months=p),
                    "yhat": None, "yhat_lower": None, "yhat_upper": None,
                    "forecast_warning": f"Ошибка Prophet: {e}",
                })

    return pd.DataFrame(forecast_rows) if forecast_rows else None


def compute_forecasted_mean(forecast_df: pd.DataFrame) -> dict[str, float]:
    if forecast_df is None or forecast_df.empty:
        return {}
    result = {}
    for query, grp in forecast_df.groupby("query"):
        values = grp["yhat"].dropna()
        result[query] = float(values.mean()) if not values.empty else None
    return result


def prophet_status_message() -> str:
    """Текст для листа «Прогноз», если Prophet не сработал."""
    if _check_prophet_available():
        return ""
    if _PROPhet_IMPORT_ERROR:
        return f"Прогноз не построен. Prophet недоступен: {_PROPhet_IMPORT_ERROR}"
    return "Прогноз не построен. Установите Prophet или проверьте данные."
