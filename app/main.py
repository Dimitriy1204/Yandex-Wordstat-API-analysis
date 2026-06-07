#!/usr/bin/env python3
"""FastAPI приложение для анализа поискового спроса Yandex Wordstat."""
from __future__ import annotations
import os, sys, json, time, logging, traceback, asyncio
from typing import Optional, Any
from datetime import datetime
from contextlib import asynccontextmanager
import pandas as pd, uvicorn
from fastapi import FastAPI, Request, HTTPException
from fastapi.responses import HTMLResponse, FileResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from pydantic import BaseModel, field_validator

_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
if getattr(sys, 'frozen', False): _PROJECT_DIR = os.path.dirname(os.path.abspath(sys.executable))
else: _PROJECT_DIR = os.path.dirname(_SCRIPT_DIR)
_TEMPLATES_DIR = os.path.join(_SCRIPT_DIR, "templates")
_STATIC_DIR = os.path.join(_SCRIPT_DIR, "static")
_OUTPUT_DIR = _PROJECT_DIR
os.makedirs(_TEMPLATES_DIR, exist_ok=True); os.makedirs(_STATIC_DIR, exist_ok=True)

class _NullStream:
    def isatty(self): return False
    def write(self, msg): pass
    def flush(self): pass
    def close(self): pass
    def read(self, n=0): return ""
    def readline(self): return ""
    def readlines(self): return []
    def readable(self): return False
    def writable(self): return True
    def seekable(self): return False
    def closed(self): return False

if sys.stderr is None: sys.stderr = _NullStream()
if sys.stdout is None: sys.stdout = _NullStream()

_log_path = os.path.join(_PROJECT_DIR, "app.log")
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)-7s | %(name)s | %(message)s",
    datefmt="%H:%M:%S",
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler(_log_path, encoding="utf-8", mode="a"),
    ],
    force=True,
)
logger = logging.getLogger("wordstat.app")
logger.info("Лог-файл: %s", _log_path)

from app.config import config_manager
from app.wordstat_client import parse_raw_queries, collect_all_data, parse_date_range, clean_query
from app.metrics import calculate_metrics, compute_summary, make_blue_oceans, make_growing_markets
from app.prophet_forecast import forecast_queries, compute_forecasted_mean
from app.excel_generator import create_excel_report
from app.yandexgpt_client import (
    expand_queries,
    generate_market_overview,
    generate_recommendations,
    generate_verdict,
    is_configured as yagpt_configured,
    test_connection as yagpt_test_connection,
)
from app.senders import send_telegram, send_vk, send_email, test_smtp_connection

class AppState:
    def __init__(self):
        self.analysis_running = False; self.analysis_cancelled = False
        self.progress_current = 0; self.progress_total = 0; self.progress_message = ""
        self.last_result_path = ""; self.last_error = ""; self.last_send_results = []
        self.chart_data: dict | None = None
        self.llm_expansion_note: str = ""
app_state = AppState()

@asynccontextmanager
async def lifespan(app: FastAPI):
    logger.info("="*60); logger.info("Yandex Wordstat Agent запущен"); yield; logger.info("Остановлено")

app = FastAPI(title="Yandex Wordstat Analysis Agent", lifespan=lifespan)
app.mount("/static", StaticFiles(directory=_STATIC_DIR), name="static")
templates = Jinja2Templates(directory=_TEMPLATES_DIR)
templates.env.cache = None

class AnalysisRequest(BaseModel):
    queries: str; date_from: str; date_to: str
    send_telegram: bool = False; send_vk: bool = False; send_email: bool = False; use_llm: bool = False

class SendRequest(BaseModel):
    channel: str

class SettingsUpdate(BaseModel):
    yandex_api_key: Optional[str] = None; folder_id: Optional[str] = None
    yagpt_api_key: Optional[str] = None; yagpt_folder_id: Optional[str] = None; yagpt_model: Optional[str] = None
    tg_bot_token: Optional[str] = None; tg_chat_id: Optional[str] = None
    vk_token: Optional[str] = None; vk_group_id: Optional[str] = None; vk_peer_id: Optional[str] = None
    smtp_host: Optional[str] = None; smtp_port: Optional[int] = None
    smtp_user: Optional[str] = None; smtp_password: Optional[str] = None
    email_from: Optional[str] = None; email_to: Optional[str] = None; smtp_use_tls: Optional[bool] = None
    proxy_type: Optional[str] = None; proxy_host: Optional[str] = None; proxy_port: Optional[int] = None
    proxy_user: Optional[str] = None; proxy_password: Optional[str] = None
    top_n: Optional[int] = None; request_delay: Optional[float] = None

    model_config = {"protected_namespaces": ()}

    @field_validator("proxy_port", "smtp_port", "top_n", mode="before")
    @classmethod
    def _empty_int_to_none(cls, v):
        if v is None or v == "":
            return None
        return v

    @field_validator("request_delay", mode="before")
    @classmethod
    def _empty_float_to_none(cls, v):
        if v is None or v == "":
            return None
        return v

def update_progress(current, total, message):
    app_state.progress_current = current; app_state.progress_total = total; app_state.progress_message = message
    logger.info(f"Progress: [{current}/{total}] {message}")

@app.get("/", response_class=HTMLResponse)
async def index_page(request: Request):
    return templates.TemplateResponse("index.html", {"request": request, "configured": config_manager.is_configured(), "app_state": app_state})

@app.post("/api/settings")
async def save_settings(data: SettingsUpdate):
    cfg = data.model_dump(exclude_none=True)
    config_manager.update(cfg)
    success = config_manager.save()
    return {"success": success, "message": "Настройки сохранены" if success else "Ошибка сохранения"}

@app.get("/api/settings")
async def get_settings():
    cfg = config_manager.to_dict()
    cfg["yagpt_ready"] = yagpt_configured()
    return {"success": True, "config": cfg}

@app.post("/api/analyze")
async def start_analysis(data: AnalysisRequest):
    if app_state.analysis_running: return {"success": False, "message": "Анализ уже запущен"}
    if not config_manager.is_configured(): return {"success": False, "message": "Настройки API не заполнены"}
    app_state.analysis_running = True; app_state.analysis_cancelled = False
    app_state.last_error = ""; app_state.last_result_path = ""; app_state.last_send_results = []
    app_state.chart_data = None
    app_state.llm_expansion_note = ""
    asyncio.create_task(run_analysis(data))
    return {"success": True, "message": "Анализ запущен"}

async def run_analysis(data: AnalysisRequest):
    start_time = time.time(); output_path = ""
    total_steps = 7 if data.use_llm else 6
    try:
        update_progress(0, total_steps, "Парсинг запросов...")
        parsed = parse_raw_queries(data.queries)
        api_queries = [q["api"] for q in parsed]
        display_queries = [q["display"] for q in parsed]
        if not api_queries:
            app_state.last_error = "Не указаны запросы"
            app_state.analysis_running = False
            return

        llm_expansion_note = ""
        if data.use_llm:
            if not yagpt_configured():
                logger.warning("LLM включён, но YandexGPT не настроен (ключ / Folder ID)")
            else:
                update_progress(1, total_steps, "YandexGPT: расширение запросов...")
                extra, note = expand_queries(display_queries, max_extra=8)
                if extra:
                    api_queries = list(dict.fromkeys(
                        api_queries + [clean_query(p) for p in extra]
                    ))
                if note:
                    llm_expansion_note = note
                    app_state.llm_expansion_note = note

        step_dates = 2 if data.use_llm else 1
        update_progress(step_dates, total_steps, "Парсинг дат...")
        date_from, date_to = parse_date_range(data.date_from, data.date_to)
        step_collect = step_dates + 1
        update_progress(step_collect, total_steps, "Сбор данных Wordstat...")
        def progress_cb(current, total, query):
            if app_state.analysis_cancelled:
                raise InterruptedError("Отменено")
            update_progress(step_collect, total_steps, f"Wordstat [{current}/{total}] {query}")
        try:
            raw_data = collect_all_data(api_queries, date_from, date_to, progress_callback=progress_cb)
        except InterruptedError:
            app_state.last_error = "Анализ отменён"
            app_state.analysis_running = False
            return
        if raw_data.empty:
            app_state.last_error = "Нет данных"
            app_state.analysis_running = False
            return

        step_metrics = step_collect + 1
        update_progress(step_metrics, total_steps, "Расчёт метрик...")
        analyzed = calculate_metrics(raw_data)
        summary = compute_summary(analyzed)

        step_prophet = step_metrics + 1
        update_progress(step_prophet, total_steps, "Прогноз Prophet...")
        forecast_df = forecast_queries(raw_data, periods=6, min_data_points=12)
        forecast_mean_map = compute_forecasted_mean(forecast_df) if forecast_df is not None else {}

        llm_recommendations = None
        llm_market_overview = None
        llm_verdicts: dict[str, str] = {}
        period_str = f"{date_from[:10]} — {date_to[:10]}"

        if data.use_llm and yagpt_configured():
            update_progress(step_prophet, total_steps, "YandexGPT: анализ ниш...")
            all_q = summary["query"].tolist()
            try:
                llm_market_overview = generate_market_overview(summary, all_q, period_str)
            except Exception as e:
                logger.warning("LLM обзор рынка: %s", e)

            top_blue = make_blue_oceans(summary, "growth_momentum")
            if not top_blue.empty:
                top_blue_queries = top_blue["query"].head(5).tolist()
                try:
                    llm_recommendations = generate_recommendations(
                        summary, top_blue_queries, period_str, len(all_q),
                    )
                except Exception as e:
                    logger.warning("LLM рекомендации: %s", e)
                for q in top_blue_queries:
                    row = summary[summary["query"] == q].iloc[0]
                    try:
                        v = generate_verdict(
                            q,
                            float(row["mean_frequency"]),
                            float(row["growth_momentum"]),
                            float(row["competition_score"]),
                            float(row["mean_volatility"]),
                            sum_freq=float(row["sum_frequency"]),
                            growth_year=float(row.get("growth_year", row["growth_momentum"])),
                        )
                    except Exception as e:
                        logger.warning("LLM вердикт для '%s': %s", q, e)
                        v = None
                    if v:
                        llm_verdicts[q] = v

        step_excel = total_steps
        update_progress(step_excel, total_steps, "Создание Excel...")
        output_path = os.path.join(_OUTPUT_DIR, f"yandex_analysis_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
        output_path, chart_payload = create_excel_report(
            analyzed, date_from, date_to, output_path, top_n=config_manager.config.top_n,
            forecast_df=forecast_df, forecast_mean_map=forecast_mean_map,
            llm_recommendations=llm_recommendations,
            llm_market_overview=llm_market_overview,
            llm_expansion_note=llm_expansion_note,
            llm_verdicts=llm_verdicts,
            all_queries=api_queries,
        )
        app_state.last_result_path = output_path
        app_state.chart_data = chart_payload
        send_results = []
        for ch, func in [("telegram", send_telegram), ("vk", send_vk), ("email", send_email)]:
            enabled = getattr(data, f"send_{ch}", False)
            if enabled:
                result = func(output_path)
                send_results.append({"channel": ch.capitalize(), **result})
                if not result.get("success"):
                    logger.error("Отправка %s не удалась: %s", ch, result.get("message"))
        app_state.last_send_results = send_results
        logger.info(f"Анализ завершён за {time.time()-start_time:.1f}с")
    except Exception as e:
        app_state.last_error = f"Ошибка: {e}"
        logger.exception("Ошибка анализа")
    finally: app_state.analysis_running = False; update_progress(0, 0, "Завершено")

@app.get("/api/progress")
async def get_progress():
    return {"running": app_state.analysis_running, "current": app_state.progress_current, "total": app_state.progress_total,
        "message": app_state.progress_message, "last_result": app_state.last_result_path, "last_error": app_state.last_error,
        "last_send_results": app_state.last_send_results, "has_charts": app_state.chart_data is not None}

@app.get("/api/chart-data")
async def get_chart_data():
    if not app_state.chart_data:
        return {"success": False, "message": "Нет данных графиков. Сначала выполните анализ."}
    return {"success": True, "data": app_state.chart_data}


@app.post("/api/yagpt/test")
async def test_yagpt():
    return yagpt_test_connection()

@app.post("/api/email/test")
async def test_email():
    return test_smtp_connection()

@app.post("/api/cancel")
async def cancel_analysis():
    if app_state.analysis_running: app_state.analysis_cancelled = True; return {"success": True, "message": "Отменяется"}
    return {"success": False, "message": "Нет активного анализа"}

@app.get("/api/download")
async def download_result():
    if not app_state.last_result_path or not os.path.exists(app_state.last_result_path):
        return {"success": False, "message": "Файл не найден"}
    return FileResponse(app_state.last_result_path, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=os.path.basename(app_state.last_result_path))

@app.post("/api/send")
async def send_result(data: SendRequest):
    if not app_state.last_result_path or not os.path.exists(app_state.last_result_path):
        return {"success": False, "message": "Файл не найден"}
    funcs = {"telegram": send_telegram, "vk": send_vk, "email": send_email}
    f = funcs.get(data.channel)
    if not f:
        logger.error("Неизвестный канал отправки: %s", data.channel)
        return {"success": False, "message": f"Неизвестный канал: {data.channel}"}
    result = f(app_state.last_result_path)
    if not result.get("success"):
        logger.error("Ручная отправка %s: %s", data.channel, result.get("message"))
    return result

def _stop_previous_instances(port: int) -> None:
    """Завершает другие экземпляры агента (exe) и процесс на порту, кроме текущего PID."""
    import subprocess
    my_pid = os.getpid()
    if sys.platform == "win32":
        for image in ("yandex_wordstat_agent.exe",):
            try:
                out = subprocess.run(
                    ["tasklist", "/FI", f"IMAGENAME eq {image}", "/FO", "CSV", "/NH"],
                    capture_output=True, text=True, timeout=8,
                )
                for line in out.stdout.splitlines():
                    parts = line.strip().strip('"').split('","')
                    if len(parts) >= 2 and parts[1].isdigit():
                        pid = int(parts[1])
                        if pid != my_pid:
                            subprocess.run(["taskkill", "/PID", str(pid), "/F"],
                                           capture_output=True, timeout=5)
            except Exception:
                pass
    try:
        result = subprocess.run(["netstat", "-ano"], capture_output=True, text=True, timeout=5)
        for line in result.stdout.splitlines():
            if f":{port}" in line and "LISTENING" in line:
                pid = line.strip().split()[-1]
                if pid.isdigit() and int(pid) != my_pid:
                    subprocess.run(["taskkill", "/PID", pid, "/F"], capture_output=True, timeout=5)
    except Exception:
        pass

def main():
    port = int(os.environ.get("PORT", 8000)); host = os.environ.get("HOST", "127.0.0.1")
    _stop_previous_instances(port)
    logger.info(f"Сервер: http://{host}:{port}")
    import webbrowser, threading
    threading.Thread(target=lambda: (__import__("time").sleep(2), webbrowser.open(f"http://{host}:{port}")), daemon=True).start()
    uvicorn.run("app.main:app", host=host, port=port, reload=False, log_level="info")

if __name__ == "__main__": main()