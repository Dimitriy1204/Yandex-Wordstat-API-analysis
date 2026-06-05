#!/usr/bin/env python3
"""Клиент YandexGPT (Foundation Models Completion API)."""
from __future__ import annotations

import logging
import re
from typing import Any, Optional

import pandas as pd
import requests

from app.config import config
from app.yandexgpt_prompts import (
    SYSTEM_MARKET_OVERVIEW,
    SYSTEM_QUERY_EXPANSION,
    SYSTEM_RECOMMENDATIONS,
    SYSTEM_VERDICT,
    parse_expanded_queries,
    summary_rows_from_df,
    user_market_overview,
    user_query_expansion,
    user_recommendations,
    user_verdict,
)

logger = logging.getLogger("wordstat.yagpt")


def _get_proxies() -> dict | None:
    if config.proxy_type == "none" or not config.proxy_host:
        return None
    p = f"{config.proxy_type}://{config.proxy_host}:{config.proxy_port}"
    if config.proxy_user and config.proxy_password:
        p = (
            f"{config.proxy_type}://{config.proxy_user}:{config.proxy_password}"
            f"@{config.proxy_host}:{config.proxy_port}"
        )
    return {"http": p, "https": p}


def get_credentials() -> tuple[str, str, str, str]:
    """
    API-ключ и folder: отдельные поля YandexGPT или общие с Wordstat.
    """
    api_key = (config.yagpt_api_key or config.yandex_api_key or "").strip()
    folder_id = (config.yagpt_folder_id or config.folder_id or "").strip()
    model = (config.yagpt_model or "yandexgpt-lite").strip()
    endpoint = (config.yagpt_endpoint or "").strip()
    return api_key, folder_id, model, endpoint


def is_configured() -> bool:
    api_key, folder_id, _, _ = get_credentials()
    return bool(api_key) and bool(folder_id)


def call_completion(
    system_text: str,
    user_text: str,
    *,
    temperature: float = 0.3,
    max_tokens: int = 2000,
    timeout: int = 60,
) -> Optional[str]:
    """Вызов YandexGPT с отдельными system и user сообщениями."""
    api_key, folder_id, model, endpoint = get_credentials()
    if not api_key or not folder_id:
        logger.warning("YandexGPT: не задан API-ключ или Folder ID")
        return None
    if not endpoint:
        endpoint = "https://llm.api.cloud.yandex.net/foundationModels/v1/completion"

    headers = {
        "Authorization": f"Api-Key {api_key}",
        "Content-Type": "application/json",
    }
    payload = {
        "modelUri": f"gpt://{folder_id}/{model}",
        "completionOptions": {
            "stream": False,
            "temperature": temperature,
            "maxTokens": str(max_tokens),
        },
        "messages": [
            {"role": "system", "text": system_text},
            {"role": "user", "text": user_text},
        ],
    }

    try:
        response = requests.post(
            endpoint,
            headers=headers,
            json=payload,
            timeout=timeout,
            proxies=_get_proxies(),
        )
        if response.status_code != 200:
            logger.error(
                "YandexGPT HTTP %s: %s",
                response.status_code,
                response.text[:500],
            )
            return None

        data = response.json()
        alternatives = data.get("result", {}).get("alternatives", [])
        if not alternatives:
            logger.warning("YandexGPT: пустой ответ (alternatives)")
            return None

        text = alternatives[0].get("message", {}).get("text", "").strip()
        if not text:
            logger.warning("YandexGPT: пустой текст в ответе")
            return None
        return text

    except requests.RequestException as e:
        logger.error("YandexGPT сеть: %s", e)
        return None
    except Exception as e:
        logger.error("YandexGPT: %s", e)
        return None


def test_connection() -> dict[str, Any]:
    """Проверка доступности API (короткий запрос)."""
    if not is_configured():
        return {
            "success": False,
            "message": "Укажите API-ключ и Folder ID (YandexGPT или Wordstat).",
        }
    text = call_completion(
        "Ты помощник. Отвечай одним словом.",
        "Напиши слово: ОК",
        temperature=0.1,
        max_tokens=16,
        timeout=30,
    )
    if text:
        _, _, model, _ = get_credentials()
        return {
            "success": True,
            "message": f"Связь с YandexGPT установлена (модель {model}).",
            "sample": text[:100],
        }
    return {
        "success": False,
        "message": "Нет ответа от API. Проверьте ключ, Folder ID и роль ai.languageModels.user.",
    }


def expand_queries(
    raw_queries: list[str],
    max_extra: int = 8,
) -> tuple[list[str], str]:
    """
    Расширяет список запросов через YandexGPT.
    Возвращает (новые_фразы, пояснение для отчёта).
    """
    if not raw_queries or not is_configured():
        return [], ""

    system = SYSTEM_QUERY_EXPANSION.format(max_extra=max_extra)
    user = user_query_expansion(raw_queries, max_extra)
    text = call_completion(system, user, temperature=0.4, max_tokens=800)
    if not text:
        return [], ""

    extra = parse_expanded_queries(text, raw_queries)
    extra = extra[:max_extra]
    if not extra:
        return [], ""

    note = "\n".join(extra)
    logger.info("YandexGPT: добавлено %s запросов", len(extra))
    return extra, note


def generate_market_overview(
    summary_df: pd.DataFrame,
    queries: list[str],
    period: str,
) -> Optional[str]:
    """Сравнительный обзор всех запросов."""
    if not is_configured() or not queries:
        return None
    rows = summary_rows_from_df(summary_df, queries[:15])
    if not rows:
        return None
    user = user_market_overview(rows, period)
    return call_completion(SYSTEM_MARKET_OVERVIEW, user, temperature=0.35, max_tokens=1500)


def generate_recommendations(
    summary_df: pd.DataFrame,
    top_queries: list[str],
    period: str = "",
    total_queries: int = 0,
) -> Optional[str]:
    """Рекомендации по голубым океанам."""
    if not is_configured() or not top_queries:
        return None

    rows = summary_rows_from_df(summary_df, top_queries[:5])
    if not rows:
        return None

    if not total_queries:
        total_queries = len(summary_df)

    user = user_recommendations(top_queries[:5], rows, period, total_queries)
    result = call_completion(SYSTEM_RECOMMENDATIONS, user, temperature=0.35, max_tokens=2000)
    if result:
        logger.info("YandexGPT: рекомендации получены (%s симв.)", len(result))
    return result


def generate_verdict(
    query: str,
    mean_freq: float,
    growth: float,
    competition: float,
    volatility: float,
    sum_freq: float = 0,
    growth_year: float = 0,
) -> Optional[str]:
    """Короткий вердикт по одной нише."""
    if not is_configured():
        return None

    user = user_verdict(
        query, mean_freq, growth, competition, volatility,
        sum_freq=sum_freq, growth_year=growth_year,
    )
    result = call_completion(SYSTEM_VERDICT, user, temperature=0.25, max_tokens=400)
    if result:
        result = re.sub(r"\s+", " ", result).strip()
        if len(result) > 320:
            result = result[:317] + "…"
    return result
