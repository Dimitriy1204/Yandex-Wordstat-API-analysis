#!/usr/bin/env python3
"""
Точка входа для запуска Yandex Wordstat Analysis Agent v2.
Запускает FastAPI веб-сервер и открывает браузер.

Использование:
    python run.py                  # Запуск на 127.0.0.1:8000
    python run.py --port 8080      # Другой порт
    python run.py --host 0.0.0.0   # Доступ с других устройств
"""

import sys
import os

# Добавляем корневую директорию в путь для импорта
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from app.main import main

if __name__ == "__main__":
    main()