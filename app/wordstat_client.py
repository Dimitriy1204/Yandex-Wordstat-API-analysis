#!/usr/bin/env python3
"""Клиент для Yandex Wordstat API v2. Сбор помесячной частотности."""
from __future__ import annotations
import os, re, time, logging
from typing import Any, Optional
from datetime import datetime
import requests, pandas as pd
from app.config import config
logger = logging.getLogger("wordstat.wordstat_client")
_OPERATOR_RE = re.compile(r'[+!]\S+|"[^"]*"|\[[^\]]*\]|-\S+')

def clean_query(query: str) -> str:
    q = re.sub(r'\s+-\S+', '', query)
    q = ' '.join(q.split())
    return q if q else query

def parse_raw_queries(raw_text: str) -> list[dict[str, str]]:
    parts = re.split(r'[\n,]+', raw_text)
    return [{"display": p.strip(), "api": clean_query(p.strip())} for p in parts if p.strip()]

def fetch_wordstat_dynamics(query: str, date_from: str, date_to: str, api_key: str, folder_id: str, endpoint: str = None) -> list[dict[str, Any]]:
    if endpoint is None: endpoint = config.api_endpoint
    headers = {"Authorization": f"Api-Key {api_key}", "Content-Type": "application/json"}
    payload = {"folderId": folder_id, "phrase": query, "period": "PERIOD_MONTHLY", "from_date": date_from, "to_date": date_to, "geo_ids": [225], "group_by": "TIME"}
    for attempt in range(1, 4):
        try:
            response = requests.post(endpoint, headers=headers, json=payload, timeout=60)
            if response.status_code in (401, 403): logger.error(f"Auth error {response.status_code}"); return []
            if response.status_code == 429: time.sleep(2 ** attempt); continue
            if response.status_code != 200:
                if attempt < 3: time.sleep(2 ** attempt); continue
                return []
            data = response.json()
            results = data.get("results", [])
            if not results: logger.warning(f"Empty data for '{query}'"); return []
            return results
        except (requests.ConnectionError, requests.Timeout) as e:
            if attempt < 3: time.sleep(2 ** attempt); continue
            return []
        except requests.RequestException as e: logger.error(f"Request error: {e}"); return []
    return []

def parse_dynamics(query: str, results: list[dict[str, Any]]) -> pd.DataFrame:
    rows = []
    for entry in results:
        date_str = entry.get("date")
        count_raw = entry.get("count")
        if not date_str or count_raw is None: continue
        try: frequency = int(str(count_raw))
        except (ValueError, TypeError): continue
        rows.append({"date": date_str, "query": query, "frequency": frequency})
    return pd.DataFrame(rows)

def collect_all_data(queries: list[str], date_from: str, date_to: str, api_key: str = None, folder_id: str = None, request_delay: float = None, progress_callback=None) -> pd.DataFrame:
    if api_key is None: api_key = config.yandex_api_key
    if folder_id is None: folder_id = config.folder_id
    if request_delay is None: request_delay = config.request_delay
    all_dfs = []
    for i, query in enumerate(queries, 1):
        if progress_callback: progress_callback(i, len(queries), query)
        results = fetch_wordstat_dynamics(query, date_from, date_to, api_key, folder_id)
        df = parse_dynamics(query, results)
        if not df.empty: all_dfs.append(df)
        if i < len(queries): time.sleep(request_delay)
    if not all_dfs: return pd.DataFrame()
    result = pd.concat(all_dfs, ignore_index=True)
    result["date"] = pd.to_datetime(result["date"]).dt.tz_localize(None)
    result = result.sort_values(["query", "date"]).reset_index(drop=True)
    return result

def last_day_of_month(year: int, month: int) -> int:
    if month == 12: next_month = datetime(year + 1, 1, 1)
    else: next_month = datetime(year, month + 1, 1)
    return (next_month - datetime(year, month, 1)).days

def parse_date_range(date_from_str: str, date_to_str: str) -> tuple[str, str]:
    today = datetime.now()
    if date_from_str.strip():
        try: parts = date_from_str.strip().split("-"); year, month = int(parts[0]), int(parts[1]); from_date = f"{year:04d}-{month:02d}-01T00:00:00Z"
        except: from_date = "2022-01-01T00:00:00Z"
    else: from_date = "2022-01-01T00:00:00Z"
    if date_to_str.strip():
        try: parts = date_to_str.strip().split("-"); year, month = int(parts[0]), int(parts[1]); to_date = f"{year:04d}-{month:02d}-{last_day_of_month(year, month):02d}T00:00:00Z"
        except:
            if today.month == 1: y, m = today.year - 1, 12
            else: y, m = today.year, today.month - 1
            to_date = f"{y:04d}-{m:02d}-{last_day_of_month(y, m):02d}T00:00:00Z"
    else:
        if today.month == 1: y, m = today.year - 1, 12
        else: y, m = today.year, today.month - 1
        to_date = f"{y:04d}-{m:02d}-{last_day_of_month(y, m):02d}T00:00:00Z"
    return from_date, to_date