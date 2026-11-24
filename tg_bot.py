# -*- coding: utf-8 -*-
"""Telegram-бот: распознаёт техпаспорт и возвращает данные для внесения в CRM.

Запуск:
    python tools/tg_bot.py

Токен: переменная окружения TELEGRAM_BOT_TOKEN или файл tools/.tg_token.
Настройки: tools/tg_config.json (белый список chat_id, лимиты).
Журнал: tools/logs/tg_bot.log

Проверка без Telegram:
    python tools/tg_bot.py --test-photo "C:\\путь\\к\\фото.jpg"
"""
from __future__ import annotations

import argparse
import io
import json
import os
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
from datetime import datetime
from pathlib import Path

TOOLS = Path(__file__).resolve().parent
sys.path.insert(0, str(TOOLS))

from car_registry import build_note, to_latin, validate_vehicle  # noqa: E402
from scan_cars import (  # noqa: E402
    ModelTruncated,
    call_model,
    load_api_key,
    prepare_image,
)

API = "https://api.telegram.org/bot{token}/{method}"
CONFIG_PATH = TOOLS / "tg_config.json"
LOG_PATH = TOOLS / "logs" / "tg_bot.log"
LOCK_PATH = TOOLS / ".tg_bot.pid"
DEFAULT_CONFIG = {
    "allowed_chat_ids": [],
    "max_per_minute": 3,
    "admin_note": "Владелец файла tools/tg_config.json",
}

_log_file = None
_lock_handle = None


def _acquire_lock() -> bool:
    """Не даёт запустить два экземпляра: держим эксклюзивную блокировку файла."""
    global _lock_handle
    LOG_PATH.parent.mkdir(parents=True, exist_ok=True)
    handle = open(LOCK_PATH, "a+b")
    try:
        import msvcrt

        msvcrt.locking(handle.fileno(), msvcrt.LK_NBLCK, 1)
    except (ImportError, OSError):
        handle.close()
        return False
    _lock_handle = handle
    handle.seek(0)
    handle.truncate()
    handle.write(str(os.getpid()).encode())
    handle.flush()
    return True


def log(message: str) -> None:
    global _log_file
    stamp = datetime.now().strftime("%d.%m.%Y %H:%M:%S")
    line = f"{stamp}  {message}"
    print(line, flush=True)
    try:
        if _log_file is None:
            LOG_PATH.parent.mkdir(parents=True, exist_ok=True)
            _log_file = open(LOG_PATH, "a", encoding="utf-8")
        _log_file.write(line + "\n")
        _log_file.flush()
    except OSError:
        pass


def load_config() -> dict:
    config = dict(DEFAULT_CONFIG)
    if CONFIG_PATH.exists():
        try:
            config.update(json.loads(CONFIG_PATH.read_text(encoding="utf-8")))
        except json.JSONDecodeError:
            log(f"! {CONFIG_PATH.name} не разобран, использую значения по умолчанию")
    return config


def load_token() -> str | None:
    token = os.environ.get("TELEGRAM_BOT_TOKEN")
    if token:
        return token.strip()
    path = TOOLS / ".tg_token"
    if path.exists():
        return path.read_text(encoding="utf-8").strip() or None
    return None


def api(token: str, method: str, http_timeout: int = 70, **params):
    """Вызов метода Bot API. None-параметры не отправляются."""
    url = API.format(token=token, method=method)
    clean = {k: v for k, v in params.items() if v is not None}
    if clean:
        req = urllib.request.Request(
            url, data=json.dumps(clean).encode("utf-8"),
            headers={"Content-Type": "application/json"}, method="POST",
        )
    else:
        req = urllib.request.Request(url, method="GET")
    with urllib.request.urlopen(req, timeout=http_timeout) as resp:
        payload = json.loads(resp.read().decode("utf-8"))
    if not payload.get("ok"):
        raise RuntimeError(f"Telegram {method}: {payload}")
    return payload["result"]


def send(token: str, chat_id: int, text: str, reply_markup: dict | None = None) -> None:
    try:
        api(token, "sendMessage", chat_id=chat_id, text=text,
            parse_mode="HTML", disable_web_page_preview=True,
            reply_markup=reply_markup)
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode("utf-8", "replace")[:300]
        log(f"! sendMessage {exc.code}: {detail}")


def download_file(token: str, file_id: str) -> bytes:
    info = api(token, "getFile", file_id=file_id)
    path = info["file_path"]
    url = f"https://api.telegram.org/file/bot{token}/{path}"
    with urllib.request.urlopen(url, timeout=120) as resp:
        return resp.read()


def build_message(vehicle: dict, warnings: list[str]) -> str:
    """Карточка в порядке полей формы CRM: марка, модель, VIN, номер, год."""
    plate = to_latin(vehicle.get("plate") or "") or "—"
    vin = (to_latin(vehicle.get("vin")) or "").upper() or "—"
    brand = vehicle.get("brand") or "—"
    model = vehicle.get("model") or "—"
    year = vehicle.get("year") or "—"

    lines = [
        "🚛 <b>Распознано — проверьте перед вставкой</b>",
        "",
        f"Марка: <b>{brand}</b>",
        f"Модель: <b>{model}</b>",
        f"VIN: <code>{vin}</code>",
        f"Гос. номер: <b>{plate}</b>",
        f"Год выпуска: <b>{year}</b>",
    ]
    if vehicle.get("displacement"):
        lines.append(f"Объём двигателя: {vehicle['displacement']}")
    if vehicle.get("code_dvs") or vehicle.get("engine_model"):
        lines.append(f"Код ДВС: {vehicle.get('code_dvs') or vehicle.get('engine_model')}")

    if warnings:
        lines.append("")
        lines.append("⚠️ " + "; ".join(warnings))
    if vehicle.get("unreadable"):
        lines.append("⚠️ не читается: " + ", ".join(map(str, vehicle["unreadable"])))

    lines += ["", "Копировать — кнопками ниже."]
    return "\n".join(lines)


def build_keyboard(vehicle: dict) -> dict | None:
    """Кнопки «скопировать значение» в порядке полей CRM."""
    plate = to_latin(vehicle.get("plate") or "") or ""
    vin = (to_latin(vehicle.get("vin")) or "").upper()
    fields = [
        ("Марка", vehicle.get("brand") or ""),
        ("Модель", vehicle.get("model") or ""),
        ("VIN", vin),
        ("Номер", plate),
        ("Год", str(vehicle.get("year") or "")),
    ]
    if vehicle.get("displacement"):
        fields.append(("Объём", str(vehicle["displacement"])))
    engine = vehicle.get("code_dvs") or vehicle.get("engine_model")
    if engine:
        fields.append(("Код ДВС", str(engine)))
    buttons = [
        {"text": f"📋 {label}", "copy_text": {"text": str(value)}}
        for label, value in fields if value
    ]
    if not buttons:
        return None
    rows = [buttons[i:i + 3] for i in range(0, len(buttons), 3)]
    return {"inline_keyboard": rows}


def process_photo(token: str, chat_id: int, file_id: str, suffix: str, api_key: str,
                  model: str, max_tokens: int) -> None:
    raw = download_file(token, file_id)
    work = TOOLS / "logs" / "_incoming"
    work.mkdir(parents=True, exist_ok=True)
    temp = work / f"in_{int(time.time())}{suffix}"
    temp.write_bytes(raw)
    try:
        image, mime = prepare_image(temp)
    finally:
        temp.unlink(missing_ok=True)

    vehicle = call_model(image, mime, api_key, model, max_tokens)
    errors, warnings = validate_vehicle(vehicle, [])
    notable = list(warnings)

    if errors:
        send(token, chat_id, "⚠️ <b>Не удалось надёжно распознать:</b>\n"
                             + "\n".join(f"• {e}" for e in errors)
                             + "\n\nПопробуйте сфотографировать ровнее и ближе.")
        log(f"chat {chat_id}: отклонено — {errors}")
        return

    send(token, chat_id, build_message(vehicle, notable), build_keyboard(vehicle))
    log(f"chat {chat_id}: распознано {to_latin(vehicle.get('plate'))} "
        f"VIN {(to_latin(vehicle.get('vin')) or '').upper()}")


def main() -> int:
    ap = argparse.ArgumentParser(description="Telegram-бот распознавания техпаспортов")
    ap.add_argument("--test-photo", help="прогнать одно фото без Telegram и напечатать ответ")
    ap.add_argument("--model", default="deepseek-flash")
    ap.add_argument("--max-tokens", type=int, default=6000)
    args = ap.parse_args()

    token = load_token()
    if not token:
        log("Не найден токен: TELEGRAM_BOT_TOKEN или tools/.tg_token")
        return 1

    api_key = load_api_key(None)
    if not api_key:
        log("Не найден ключ DeepSeek: DEEPSEEK_API_KEY или tools/.deepseek_key")
        return 1

    if args.test_photo:
        image, mime = prepare_image(Path(args.test_photo))
        vehicle = call_model(image, mime, api_key, args.model, args.max_tokens)
        errors, warnings = validate_vehicle(vehicle, [])
        print(build_message(vehicle, warnings))
        if errors:
            print("\nОШИБКИ:", errors)
        return 0

    config = load_config()
    allowed = {int(x) for x in config.get("allowed_chat_ids", [])}
    limit = int(config.get("max_per_minute", 3))

    if LOCK_PATH.exists():
        if not _acquire_lock():
            log("Бот уже запущен (замок занят). Второй экземпляр не нужен.")
            return 1
    else:
        _acquire_lock()

    me = api(token, "getMe")
    log(f"Старт: @{me.get('username')} (id {me.get('id')}), "
        f"допущено чатов: {len(allowed) if allowed else 'белый список пуст'}")

    seen: dict[int, list[float]] = {}
    offset = None
    try:
        while True:
            try:
                updates = api(token, "getUpdates", http_timeout=90,
                              offset=offset, timeout=50)
            except urllib.error.HTTPError as exc:
                if exc.code == 409:
                    log("! Конфликт 409: где-то запущен второй экземпляр бота")
                    time.sleep(5)
                    continue
                log(f"! getUpdates {exc.code}")
                time.sleep(5)
                continue
            except (urllib.error.URLError, TimeoutError, OSError) as exc:
                log(f"! сеть: {exc}")
                time.sleep(5)
                continue

            for update in updates:
                offset = update["update_id"] + 1
                message = update.get("message") or update.get("edited_message")
                if not message:
                    continue
                chat_id = message["chat"]["id"]
                text = (message.get("text") or "").strip()

                if not allowed:
                    send(token, chat_id,
                         f"Бот ещё не настроен.\nВаш chat_id: <code>{chat_id}</code>\n"
                         "Передайте его администратору.")
                    log(f"chat {chat_id}: не в списке (список пуст), показан chat_id")
                    continue

                if chat_id not in allowed:
                    send(token, chat_id,
                         f"Доступ закрыт.\nВаш chat_id: <code>{chat_id}</code>\n"
                         "Передайте его администратору — он добавит вас в список.")
                    log(f"chat {chat_id}: доступ закрыт (не в списке)")
                    continue

                if text.startswith("/start") or text.startswith("/help"):
                    send(token, chat_id,
                         "Пришлите фото свидетельства о регистрации ТС "
                         "(лицевая сторона, где VIN и гос. номер).\n"
                         "Я верну данные и готовый блок для вставки в CRM.")
                    continue

                if text.startswith("/id"):
                    send(token, chat_id, f"Ваш chat_id: <code>{chat_id}</code>")
                    continue

                now = time.time()
                recent = [t for t in seen.get(chat_id, []) if now - t < 60]
                if len(recent) >= limit:
                    send(token, chat_id, "Слишком много запросов, подождите минуту.")
                    continue
                recent.append(now)
                seen[chat_id] = recent

                photo = message.get("photo")
                document = message.get("document")
                try:
                    if photo:
                        send(token, chat_id, "Принял, распознаю…")
                        process_photo(token, chat_id, photo[-1]["file_id"], ".jpg",
                                      api_key, args.model, args.max_tokens)
                    elif document and (document.get("mime_type") or "").startswith(("image/", "application/pdf")):
                        suffix = ".pdf" if document["mime_type"] == "application/pdf" else ".jpg"
                        send(token, chat_id, "Принял файл, распознаю…")
                        process_photo(token, chat_id, document["file_id"], suffix,
                                      api_key, args.model, args.max_tokens)
                    elif text:
                        send(token, chat_id, "Нужно фото техпаспорта. "
                                             "Пришлите снимок лицевой стороны.")
                except ModelTruncated as exc:
                    log(f"chat {chat_id}: {exc}")
                    send(token, chat_id, "Модель не успела ответить. Пришлите фото ещё раз.")
                except Exception as exc:  # noqa: BLE001 — бот не должен падать
                    log(f"chat {chat_id}: ошибка — {exc!r}")
                    send(token, chat_id, f"Ошибка обработки: {type(exc).__name__}. "
                                         "Попробуйте ещё раз.")
    except KeyboardInterrupt:
        log("Остановка по Ctrl+C")
    finally:
        LOCK_PATH.unlink(missing_ok=True)
    return 0


if __name__ == "__main__":
    sys.exit(main())
