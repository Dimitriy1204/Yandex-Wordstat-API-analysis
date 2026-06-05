#!/usr/bin/env python3
"""Системные и пользовательские промпты для YandexGPT."""
from __future__ import annotations

from typing import Any

# ── Системные промпты ─────────────────────────────────────────────

SYSTEM_ANALYST = """Ты — ведущий аналитик поискового спроса и digital-маркетинга в России.
Работаешь с данными Yandex Wordstat: помесячная частотность, темпы роста, конкуренция, волатильность.
Пиши по-русски, деловым языком, без воды. Опирайся только на переданные цифры.
Не выдумывай данные. Если информации недостаточно — скажи об этом явно."""

SYSTEM_RECOMMENDATIONS = SYSTEM_ANALYST + """
Задача: сформировать практические рекомендации по входу в ниши (голубые океаны).
Структура ответа:
1) Краткий вывод (2–3 предложения).
2) По каждой нише из списка — 1–2 предложения: перспектива, риск, тактика.
3) Итог: 3–5 приоритетных действий для маркетинга/закупки рекламы.
Без markdown-заголовков уровня #, можно маркеры «•» или нумерацию."""

SYSTEM_VERDICT = SYSTEM_ANALYST + """
Задача: один короткий вердикт по нише для колонки «Вердикт нейросети» в Excel.
Формат: 1–2 предложения, до 280 символов. Укажи: стоит ли входить (да/осторожно/нет) и почему.
Без эмодзи в начале — их добавит отчёт отдельно."""

SYSTEM_QUERY_EXPANSION = """Ты — SEO-специалист по коммерческим запросам в Яндексе (Россия).
Предлагаешь дополнительные поисковые фразы для сравнительного анализа в Wordstat.
Правила:
— фразы на русском, релевантны исходной теме;
— коммерческий/информационный спрос B2B и B2C;
— без дубликатов исходных запросов;
— без кавычек и пояснений в ответе;
— каждая фраза с новой строки, не более {max_extra} фраз."""

SYSTEM_MARKET_OVERVIEW = SYSTEM_ANALYST + """
Задача: сравнительный обзор всех переданных запросов.
Выдели: лидера по спросу, самый быстрорастущий, самый стабильный, самый рискованный.
Дай 4–6 предложений с конкретными цифрами из таблицы."""


# ── Пользовательские промпты ──────────────────────────────────────

def user_query_expansion(raw_queries: list[str], max_extra: int) -> str:
    lines = "\n".join(f"— {q}" for q in raw_queries)
    return (
        f"Исходные запросы пользователя для анализа в Wordstat:\n{lines}\n\n"
        f"Предложи до {max_extra} дополнительных смежных запросов для сравнения "
        f"(синонимы, смежные товары/услуги, уточнения спроса).\n"
        "Ответ: только список фраз, по одной на строку."
    )


def _format_query_metrics(row: Any) -> str:
    return (
        f"средний спрос {row['mean_frequency']:.0f} показ./мес; "
        f"ёмкость рынка {row['sum_frequency']:.0f}; "
        f"рост (импульс 6 мес.) {row['growth_momentum']:+.2f}%; "
        f"рост за год {row.get('growth_year', row.get('growth_momentum', 0)):+.2f}%; "
        f"конкуренция {row['competition_score']:.2f} (0–1, ниже — лучше); "
        f"волатильность {row['mean_volatility']:.1f}"
    )


def user_recommendations(
    top_queries: list[str],
    summary_rows: list[dict],
    period: str,
    total_queries: int,
) -> str:
    blocks = []
    for item in summary_rows:
        blocks.append(f"• «{item['query']}»: {item['metrics']}")
    return (
        f"Период анализа: {period}.\n"
        f"Всего сравниваемых запросов: {total_queries}.\n"
        f"Топ ниш с низкой конкуренцией и ростом (голубые океаны, импульс 6 мес.):\n"
        + "\n".join(blocks)
        + "\n\nСформируй рекомендации по стратегии входа в эти ниши."
    )


def user_verdict(
    query: str,
    mean_freq: float,
    growth: float,
    competition: float,
    volatility: float,
    sum_freq: float = 0,
    growth_year: float = 0,
) -> str:
    return (
        f"Ниша: «{query}».\n"
        f"Средний спрос: {mean_freq:.0f} показ./мес.\n"
        f"Суммарная ёмкость: {sum_freq:.0f} показ. за период.\n"
        f"Импульс роста (6 мес.): {growth:+.2f}%/мес.\n"
        f"Рост за год: {growth_year:+.2f}%/мес.\n"
        f"Индекс конкуренции: {competition:.2f} (0 — низкая, 1 — высокая).\n"
        f"Волатильность спроса: {volatility:.1f}.\n"
        "Дай вердикт: входить в нишу или нет."
    )


def user_market_overview(summary_rows: list[dict], period: str) -> str:
    lines = []
    for item in summary_rows:
        lines.append(f"• «{item['query']}»: {item['metrics']}")
    return (
        f"Период: {period}.\n"
        "Сводка по всем запросам:\n"
        + "\n".join(lines)
        + "\n\nДай сравнительный обзор рынка по этим запросам."
    )


def summary_rows_from_df(summary_df, queries: list[str]) -> list[dict]:
    """Готовит строки метрик для промптов из DataFrame сводки."""
    rows = []
    for q in queries:
        part = summary_df[summary_df["query"] == q]
        if part.empty:
            continue
        row = part.iloc[0]
        rows.append({
            "query": q,
            "metrics": _format_query_metrics(row),
        })
    return rows


def parse_expanded_queries(llm_text: str, original: list[str]) -> list[str]:
    """Парсит ответ LLM со списком запросов."""
    if not llm_text:
        return []
    orig_lower = {o.strip().lower() for o in original}
    found: list[str] = []
    for line in llm_text.strip().splitlines():
        line = line.strip().lstrip("•-*0123456789.) ").strip('"\'')
        if not line or len(line) < 2:
            continue
        if line.lower() in orig_lower:
            continue
        if line not in found:
            found.append(line)
    return found
