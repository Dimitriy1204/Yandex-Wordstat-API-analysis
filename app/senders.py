#!/usr/bin/env python3
"""Модуль отправки отчётов: Telegram, VK, Email."""
from __future__ import annotations
import os, smtplib, logging
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
import requests
from app.config import config
logger = logging.getLogger("wordstat.senders")

def _get_proxies():
    if config.proxy_type == "none" or not config.proxy_host:
        return None
    p = f"{config.proxy_type}://{config.proxy_host}:{config.proxy_port}"
    if config.proxy_user and config.proxy_password:
        p = f"{config.proxy_type}://{config.proxy_user}:{config.proxy_password}@{config.proxy_host}:{config.proxy_port}"
    return {"http": p, "https": p}

def send_telegram(file_path, bot_token=None, chat_id=None):
    if bot_token is None:
        bot_token = config.tg_bot_token
    if chat_id is None:
        chat_id = config.tg_chat_id
    if not bot_token or not chat_id:
        return {"success": False, "message": "Telegram не настроен (token или chat_id)"}
    if not os.path.exists(file_path):
        return {"success": False, "message": f"Файл не найден: {file_path}"}
    try:
        url = f"https://api.telegram.org/bot{bot_token}/sendDocument"
        with open(file_path, "rb") as f:
            resp = requests.post(
                url,
                files={"document": f},
                data={"chat_id": chat_id, "caption": "📊 Анализ поискового спроса - Yandex Wordstat"},
                timeout=90,
                proxies=_get_proxies(),
            )
        if resp.status_code == 200:
            return {"success": True, "message": "Файл отправлен в Telegram"}
        return {"success": False, "message": f"Ошибка Telegram: {resp.text[:200]}"}
    except Exception as e:
        return {"success": False, "message": f"Ошибка отправки в Telegram: {e}"}

def send_vk(file_path, token=None, peer_id=None):
    if token is None:
        token = config.vk_token
    if peer_id is None:
        peer_id = config.vk_peer_id
    if not token or not peer_id:
        return {"success": False, "message": "VK не настроен (token или peer_id)"}
    if not os.path.exists(file_path):
        return {"success": False, "message": f"Файл не найден: {file_path}"}
    proxies = _get_proxies()
    try:
        upload_url_resp = requests.post(
            "https://api.vk.com/method/docs.getUploadServer",
            data={"access_token": token, "v": "5.199", "type": "doc"},
            timeout=30,
            proxies=proxies,
        ).json()
        if "error" in upload_url_resp:
            error_msg = upload_url_resp["error"].get("error_msg", "Unknown error")
            return {"success": False, "message": f"Ошибка VK: {error_msg}"}
        upload_url = upload_url_resp.get("response", {}).get("upload_url")
        if not upload_url:
            return {"success": False, "message": "Не удалось получить URL для загрузки"}
        with open(file_path, "rb") as f:
            upload_resp = requests.post(upload_url, files={"file": f}, timeout=90, proxies=proxies).json()
        if "error" in upload_resp:
            return {"success": False, "message": f"Ошибка загрузки файла: {upload_resp['error']}"}
        file_data = upload_resp.get("file", "")
        save_resp = requests.post(
            "https://api.vk.com/method/docs.save",
            data={"access_token": token, "v": "5.199", "file": file_data},
            timeout=30,
            proxies=proxies,
        ).json()
        if "error" in save_resp:
            error_msg = save_resp["error"].get("error_msg", "Unknown error")
            return {"success": False, "message": f"Ошибка VK save: {error_msg}"}
        doc_info = save_resp.get("response", [{}])[0]
        doc_id = f"doc{doc_info.get('owner_id', '')}_{doc_info.get('id', '')}"
        msg_resp = requests.post(
            "https://api.vk.com/method/messages.send",
            data={
                "access_token": token,
                "v": "5.199",
                "peer_id": peer_id,
                "message": "📊 Анализ поискового спроса - Yandex Wordstat",
                "attachment": doc_id,
                "random_id": 0,
            },
            timeout=30,
            proxies=proxies,
        ).json()
        if "error" in msg_resp:
            error_msg = msg_resp["error"].get("error_msg", "Unknown error")
            return {"success": False, "message": f"Ошибка VK send: {error_msg}"}
        return {"success": True, "message": "Файл отправлен в VK"}
    except Exception as e:
        return {"success": False, "message": f"Ошибка отправки в VK: {e}"}

def send_email(file_path, smtp_host=None, smtp_port=None, smtp_user=None, smtp_password=None, email_from=None, email_to=None, use_tls=None):
    if smtp_host is None:
        smtp_host = config.smtp_host
    if smtp_port is None:
        smtp_port = config.smtp_port
    if smtp_user is None:
        smtp_user = config.smtp_user
    if smtp_password is None:
        smtp_password = config.smtp_password
    if email_from is None:
        email_from = config.email_from
    if email_to is None:
        email_to = config.email_to
    if use_tls is None:
        use_tls = config.smtp_use_tls
    if not all([smtp_host, smtp_port, smtp_user, smtp_password, email_from, email_to]):
        return {"success": False, "message": "Email не настроен (проверьте SMTP параметры)"}
    if not os.path.exists(file_path):
        return {"success": False, "message": f"Файл не найден: {file_path}"}
    try:
        msg = MIMEMultipart()
        msg["From"] = email_from
        msg["To"] = email_to
        msg["Subject"] = "📊 Анализ поискового спроса - Yandex Wordstat"
        msg.attach(MIMEText(
            "Здравствуйте!\n\nВо вложении отчёт по анализу поискового спроса.\n"
            "Сгенерировано с помощью Yandex Wordstat API Agent.\n\nС уважением,\nАналитический модуль"
        ))
        filename = os.path.basename(file_path)
        with open(file_path, "rb") as f:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header("Content-Disposition", f'attachment; filename="{filename}"')
            msg.attach(part)
        if use_tls:
            server = smtplib.SMTP(smtp_host, smtp_port, timeout=30)
            server.starttls()
        else:
            server = smtplib.SMTP_SSL(smtp_host, smtp_port, timeout=30) if smtp_port == 465 else smtplib.SMTP(smtp_host, smtp_port, timeout=30)
            if smtp_port != 465:
                server.starttls()
        server.login(smtp_user, smtp_password)
        server.send_message(msg)
        server.quit()
        return {"success": True, "message": f"Файл отправлен на {email_to}"}
    except smtplib.SMTPAuthenticationError:
        return {"success": False, "message": "Ошибка аутентификации SMTP"}
    except smtplib.SMTPException as e:
        return {"success": False, "message": f"Ошибка SMTP: {e}"}
    except Exception as e:
        return {"success": False, "message": f"Ошибка отправки email: {e}"}
