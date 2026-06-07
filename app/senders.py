#!/usr/bin/env python3
"""Модуль отправки отчётов: Telegram, VK, Email."""
from __future__ import annotations
import os
import re
import ssl
import random
import smtplib
import logging
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders

import requests
from requests.exceptions import ConnectTimeout, ConnectionError as RequestsConnectionError

from app.config import config

logger = logging.getLogger("wordstat.senders")

_TOKEN_IN_URL = re.compile(r"/bot[^/]+/", re.I)


def _get_proxies():
    if config.proxy_type == "none" or not config.proxy_host:
        return None
    p = f"{config.proxy_type}://{config.proxy_host}:{config.proxy_port}"
    if config.proxy_user and config.proxy_password:
        p = (
            f"{config.proxy_type}://{config.proxy_user}:{config.proxy_password}"
            f"@{config.proxy_host}:{config.proxy_port}"
        )
    return {"http": p, "https": p}


def _fail(channel: str, message: str, *, exc: BaseException | None = None) -> dict:
    """Возвращает ошибку и пишет её в app.log."""
    if exc is not None:
        logger.error("[%s] %s", channel, message, exc_info=exc)
    else:
        logger.error("[%s] %s", channel, message)
    return {"success": False, "message": message}


def _ok(channel: str, message: str) -> dict:
    logger.info("[%s] %s", channel, message)
    return {"success": True, "message": message}


def _sanitize_error(text: str) -> str:
    """Убирает токен бота из текста ошибки (не светить в UI/логах URL целиком)."""
    return _TOKEN_IN_URL.sub("/bot***/", str(text))


def _validate_tg_chat_id(chat_id: str) -> str | None:
    """Проверяет chat_id. Возвращает текст ошибки или None."""
    cid = str(chat_id).strip()
    if not cid:
        return "Chat ID не указан"
    if cid.startswith("@"):
        return (
            "Chat ID должен быть числом, не @username. "
            "Напишите боту /start, затем узнайте ID через @userinfobot "
            "или @getmyid_bot"
        )
    if not re.fullmatch(r"-?\d+", cid):
        return "Chat ID должен содержать только цифры (например 123456789)"
    return None


def _normalize_smtp_credentials(
    smtp_user: str,
    smtp_password: str,
    email_from: str,
) -> tuple[str, str, str, str | None]:
    """Нормализует логин/пароль и проверяет обязательные поля."""
    user = (smtp_user or "").strip()
    password = (smtp_password or "").strip()
    sender = (email_from or "").strip()
    if not user and sender:
        user = sender
    if user and "@" not in user and sender and "@" in sender:
        user = sender
    if not user:
        return user, password, sender, "Логин SMTP не указан (полный email, например you@disroot.org)"
    if "@" not in user:
        return user, password, sender, "Логин SMTP должен быть полным email (например you@disroot.org)"
    if sender and user.lower() != sender.lower():
        return (
            user,
            password,
            sender,
            f"Логин ({user}) должен совпадать с полем «От кого» ({sender})",
        )
    if not password:
        return user, password, sender, "Пароль SMTP не указан — введите и сохраните настройки"
    return user, password, sender, None


def _smtp_connect(host: str, port: int, use_tls: bool) -> smtplib.SMTP:
    context = ssl.create_default_context()
    if port == 465:
        return smtplib.SMTP_SSL(host, port, timeout=30, context=context)
    server = smtplib.SMTP(host, port, timeout=30)
    server.ehlo()
    if use_tls or port == 587:
        server.starttls(context=context)
        server.ehlo()
    return server


def _smtp_auth_hint(host: str, port: int, user: str, email_from: str) -> str:
    parts = [
        "Сервер отклонил логин или пароль (код 535).",
        f"Проверьте: хост {host}:{port}, логин {user}.",
    ]
    if email_from and user.lower() != email_from.lower():
        parts.append(f"Логин должен совпадать с «От кого»: {email_from}.")
    if "disroot" in host.lower():
        parts.append(
            "Disroot: пароль как на mail.disroot.org; логин и From = you@disroot.org; "
            "порт 587 + TLS или 465 (SSL)."
        )
    else:
        parts.append("Перезапишите пароль в настройках и нажмите «Сохранить».")
    return " ".join(parts)


def test_smtp_connection(
    smtp_host=None,
    smtp_port=None,
    smtp_user=None,
    smtp_password=None,
    email_from=None,
    use_tls=None,
):
    """Проверяет только подключение и аутентификацию SMTP (без отправки письма)."""
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
    if use_tls is None:
        use_tls = config.smtp_use_tls

    if not smtp_host or not smtp_port:
        return _fail("email", "SMTP не настроен (хост и порт)")

    user, password, sender, err = _normalize_smtp_credentials(
        smtp_user, smtp_password, email_from
    )
    if err:
        return _fail("email", err)

    host = smtp_host.strip()
    port = int(smtp_port)
    logger.info("[email] Тест SMTP: %s:%s, логин=%s, TLS=%s", host, port, user, use_tls)

    try:
        server = _smtp_connect(host, port, use_tls)
        server.login(user, password)
        server.quit()
        return _ok("email", f"SMTP: вход выполнен ({user} @ {host}:{port})")
    except smtplib.SMTPAuthenticationError as e:
        hint = _smtp_auth_hint(host, port, user, sender)
        return _fail("email", hint, exc=e)
    except smtplib.SMTPException as e:
        return _fail("email", f"Ошибка SMTP: {e}", exc=e)
    except Exception as e:
        return _fail("email", f"Ошибка подключения SMTP: {e}", exc=e)


def _vk_hint(error_msg: str) -> str:
    msg = error_msg
    low = error_msg.lower()
    if "group auth" in low:
        msg += (
            ". Для ключа сообщества укажите ID группы и peer_id получателя; "
            "в ключе нужны права «Документы» и «Сообщения сообщества»"
        )
    return msg


def send_telegram(file_path, bot_token=None, chat_id=None):
    if bot_token is None:
        bot_token = config.tg_bot_token
    if chat_id is None:
        chat_id = config.tg_chat_id
    if not bot_token or not chat_id:
        return _fail("telegram", "Telegram не настроен (token или chat_id)")
    chat_err = _validate_tg_chat_id(chat_id)
    if chat_err:
        return _fail("telegram", chat_err)
    if not os.path.exists(file_path):
        return _fail("telegram", f"Файл не найден: {file_path}")
    try:
        url = f"https://api.telegram.org/bot{bot_token}/sendDocument"
        with open(file_path, "rb") as f:
            resp = requests.post(
                url,
                files={"document": f},
                data={
                    "chat_id": str(chat_id).strip(),
                    "caption": "📊 Анализ поискового спроса - Yandex Wordstat",
                },
                timeout=90,
                proxies=_get_proxies(),
            )
        if resp.status_code == 200:
            return _ok("telegram", "Файл отправлен в Telegram")
        body = _sanitize_error(resp.text[:300])
        return _fail("telegram", f"Ошибка Telegram HTTP {resp.status_code}: {body}")
    except (ConnectTimeout, RequestsConnectionError) as e:
        hint = (
            "Не удалось подключиться к api.telegram.org. "
            "В России Telegram часто заблокирован — настройте прокси (SOCKS5/HTTP) "
            "в разделе «Прокси» или используйте VPN."
        )
        return _fail("telegram", f"Ошибка отправки в Telegram: {hint}", exc=e)
    except Exception as e:
        return _fail(
            "telegram",
            f"Ошибка отправки в Telegram: {_sanitize_error(e)}",
            exc=e,
        )


def send_vk(file_path, token=None, peer_id=None, group_id=None):
    if token is None:
        token = config.vk_token
    if group_id is None:
        group_id = config.vk_group_id
    if peer_id is None:
        peer_id = config.vk_peer_id
    if not token or not peer_id:
        return _fail("vk", "VK не настроен (token или peer_id)")
    if not os.path.exists(file_path):
        return _fail("vk", f"Файл не найден: {file_path}")
    gid = str(group_id).strip().lstrip("-") if group_id else ""
    proxies = _get_proxies()
    try:
        # Для вложений в личку — getMessagesUploadServer (поддерживает ключ сообщества).
        # docs.getUploadServer с group auth недоступен → «Group authorization failed».
        upload_data = {
            "access_token": token,
            "v": "5.199",
            "type": "doc",
            "peer_id": str(peer_id).strip(),
        }
        upload_url_resp = requests.post(
            "https://api.vk.com/method/docs.getMessagesUploadServer",
            data=upload_data,
            timeout=30,
            proxies=proxies,
        ).json()
        if "error" in upload_url_resp:
            error_msg = _vk_hint(upload_url_resp["error"].get("error_msg", "Unknown error"))
            code = upload_url_resp["error"].get("error_code")
            if code == 901:
                error_msg += (
                    ". Напишите сообществу в личку или разрешите входящие сообщения "
                    "в настройках группы (Управление → Сообщения)"
                )
            return _fail("vk", f"Ошибка VK upload: {error_msg}")
        upload_url = upload_url_resp.get("response", {}).get("upload_url")
        if not upload_url:
            return _fail("vk", "Не удалось получить URL для загрузки")
        with open(file_path, "rb") as f:
            upload_resp = requests.post(
                upload_url, files={"file": f}, timeout=90, proxies=proxies
            ).json()
        if "error" in upload_resp:
            return _fail("vk", f"Ошибка загрузки файла: {upload_resp['error']}")
        file_data = upload_resp.get("file", "")
        save_data = {"access_token": token, "v": "5.199", "file": file_data}
        if gid:
            save_data["group_id"] = gid
        save_resp = requests.post(
            "https://api.vk.com/method/docs.save",
            data=save_data,
            timeout=30,
            proxies=proxies,
        ).json()
        if "error" in save_resp:
            error_msg = _vk_hint(save_resp["error"].get("error_msg", "Unknown error"))
            return _fail("vk", f"Ошибка VK save: {error_msg}")
        raw = save_resp.get("response", {})
        if isinstance(raw, list):
            doc_info = raw[0] if raw else {}
        elif isinstance(raw, dict) and "doc" in raw:
            doc_info = raw["doc"]
        else:
            doc_info = raw if isinstance(raw, dict) else {}
        doc_id = f"doc{doc_info.get('owner_id', '')}_{doc_info.get('id', '')}"
        if not doc_info.get("id"):
            return _fail("vk", "Не удалось сохранить документ VK")
        send_data = {
            "access_token": token,
            "v": "5.199",
            "peer_id": peer_id,
            "message": "📊 Анализ поискового спроса - Yandex Wordstat",
            "attachment": doc_id,
            "random_id": random.randint(1, 2**31 - 1),
        }
        if gid:
            send_data["group_id"] = gid
        msg_resp = requests.post(
            "https://api.vk.com/method/messages.send",
            data=send_data,
            timeout=30,
            proxies=proxies,
        ).json()
        if "error" in msg_resp:
            error_msg = _vk_hint(msg_resp["error"].get("error_msg", "Unknown error"))
            return _fail("vk", f"Ошибка VK send: {error_msg}")
        return _ok("vk", "Файл отправлен в VK")
    except Exception as e:
        return _fail("vk", f"Ошибка отправки в VK: {e}", exc=e)


def send_email(
    file_path,
    smtp_host=None,
    smtp_port=None,
    smtp_user=None,
    smtp_password=None,
    email_from=None,
    email_to=None,
    use_tls=None,
):
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

    if not all([smtp_host, smtp_port, email_from, email_to]):
        return _fail("email", "Email не настроен (проверьте SMTP параметры)")
    if not os.path.exists(file_path):
        return _fail("email", f"Файл не найден: {file_path}")

    user, password, sender, err = _normalize_smtp_credentials(
        smtp_user, smtp_password, email_from
    )
    if err:
        return _fail("email", err)

    host = smtp_host.strip()
    port = int(smtp_port)
    if port == 587 and not use_tls:
        logger.warning(
            "[email] Порт 587 без TLS: для disroot.org включите «Использовать TLS»"
        )

    try:
        msg = MIMEMultipart()
        msg["From"] = sender
        msg["To"] = email_to.strip()
        msg["Subject"] = "📊 Анализ поискового спроса - Yandex Wordstat"
        msg.attach(MIMEText(
            "Здравствуйте!\n\nВо вложении отчёт по анализу поискового спроса.\n"
            "Сгенерировано с помощью Yandex Wordstat API Agent.\n\n"
            "С уважением,\nАналитический модуль"
        ))
        filename = os.path.basename(file_path)
        with open(file_path, "rb") as f:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header("Content-Disposition", f'attachment; filename="{filename}"')
            msg.attach(part)

        server = _smtp_connect(host, port, use_tls)
        server.login(user, password)
        server.send_message(msg)
        server.quit()
        return _ok("email", f"Файл отправлен на {email_to}")
    except smtplib.SMTPAuthenticationError as e:
        hint = _smtp_auth_hint(host, port, user, sender)
        return _fail("email", hint, exc=e)
    except smtplib.SMTPException as e:
        return _fail("email", f"Ошибка SMTP: {e}", exc=e)
    except Exception as e:
        return _fail("email", f"Ошибка отправки email: {e}", exc=e)
