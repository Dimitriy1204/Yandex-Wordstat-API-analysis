#!/usr/bin/env python3
"""Модуль управления настройками приложения."""
from __future__ import annotations
import os, sys, json, base64, logging
from typing import Optional
from dataclasses import dataclass, field, asdict
from cryptography.fernet import Fernet
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
logger = logging.getLogger("wordstat.config")
_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
if getattr(sys, 'frozen', False): _PROJECT_DIR = os.path.dirname(os.path.abspath(sys.executable))
else: _PROJECT_DIR = os.path.dirname(_SCRIPT_DIR)
_CONFIG_FILE = os.path.join(_PROJECT_DIR, "settings.enc")

def _derive_key(master_password, salt):
    kdf = PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=salt, iterations=100_000)
    return base64.urlsafe_b64encode(kdf.derive(master_password.encode()))

@dataclass
class AppConfig:
    yandex_api_key: str = ""
    folder_id: str = ""
    api_endpoint: str = "https://searchapi.api.cloud.yandex.net/v2/wordstat/dynamics"
    yagpt_api_key: str = ""
    yagpt_folder_id: str = ""
    yagpt_model: str = "yandexgpt-lite"
    yagpt_endpoint: str = "https://llm.api.cloud.yandex.net/foundationModels/v1/completion"
    tg_bot_token: str = ""
    tg_chat_id: str = ""
    vk_token: str = ""
    vk_group_id: str = ""
    vk_peer_id: str = ""
    smtp_host: str = ""
    smtp_port: int = 587
    smtp_user: str = ""
    smtp_password: str = ""
    email_from: str = ""
    email_to: str = ""
    smtp_use_tls: bool = True
    proxy_type: str = "none"
    proxy_host: str = ""
    proxy_port: int = 1080
    proxy_user: str = ""
    proxy_password: str = ""
    top_n: int = 10
    request_delay: float = 0.4

class ConfigManager:
    def __init__(self, config_path=_CONFIG_FILE):
        self.config_path = config_path
        self.config = AppConfig()
        self._load()
    def _load(self):
        if os.path.exists(self.config_path):
            try:
                with open(self.config_path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                for key, value in data.items():
                    if hasattr(self.config, key):
                        setattr(self.config, key, value)
            except Exception as e:
                logger.warning(f"Не удалось загрузить настройки: {e}")
    def save(self):
        try:
            with open(self.config_path, "w", encoding="utf-8") as f:
                json.dump(asdict(self.config), f, ensure_ascii=False, indent=2)
            logger.info(f"Настройки сохранены в {self.config_path}")
            return True
        except Exception as e:
            logger.error(f"Ошибка сохранения настроек: {e}")
            return False
    def to_dict(self):
        """Возвращает настройки как словарь. Пароли не маскируем,
        иначе фронтенд отправит '****' обратно и затрёт реальные значения."""
        return asdict(self.config)
    def to_dict_masked(self):
        """Возвращает настройки с маскированными паролями для отображения."""
        d = asdict(self.config)
        for secret_field in ["yandex_api_key", "yagpt_api_key", "tg_bot_token",
                             "vk_token", "smtp_password", "yagpt_folder_id", "proxy_password"]:
            if d.get(secret_field):
                d[secret_field] = d[secret_field][:4] + "****"
        return d
    def update(self, data):
        for key, value in data.items():
            if hasattr(self.config, key):
                # Пропускаем маскированные значения "****" — они означают "не менять"
                if isinstance(value, str) and value.endswith("****"):
                    continue
                if key in ("smtp_port", "proxy_port", "top_n"):
                    try:
                        value = int(value)
                    except (TypeError, ValueError):
                        value = 587 if key == "smtp_port" else (1080 if key == "proxy_port" else 10)
                elif key == "request_delay":
                    try:
                        value = float(value)
                    except (TypeError, ValueError):
                        value = 0.4
                setattr(self.config, key, value)
    def is_configured(self):
        return bool(self.config.yandex_api_key) and bool(self.config.folder_id)

config_manager = ConfigManager()
config = config_manager.config