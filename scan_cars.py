# -*- coding: utf-8 -*-
"""Распознавание свидетельств о регистрации ТС и дозапись в реестр «Автомобили.xlsx».

Примеры:

    # посмотреть, что распознано, но ничего не писать
    python tools/scan_cars.py "D:\\Фото авто" --dry-run

    # записать все найденные карточки в реестр
    python tools/scan_cars.py "D:\\Фото авто"

    # один файл
    python tools/scan_cars.py "D:\\Фото авто\\СТС.jpg"

Ключ API берётся из --api-key, переменной окружения DEEPSEEK_API_KEY
или файла tools/.deepseek_key (одна строка).
"""
from __future__ import annotations

import argparse
import base64
import io
import json
import os
import re
import sys
import urllib.error
import urllib.request
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from car_registry import (  # noqa: E402
    RegistryLocked,
    append_vehicle,
    read_registry,
    validate_vehicle,
)

API_URL = "https://api.deepseek.com/chat/completions"
DEFAULT_MODEL = "deepseek-flash"
DEFAULT_MAX_TOKENS = 6000
IMAGE_EXT = {".jpg", ".jpeg", ".png", ".webp"}
MAX_SIDE = 1600

SYSTEM_PROMPT = (
    "Ты — оператор, который переносит данные из свидетельств о регистрации "
    "транспортного средства (СТС РФ, свидетельство о регистрации РБ) в базу. "
    "Ты никогда не додумываешь значения: если символ не читается — null и запись "
    "в unreadable. Отвечаешь строго одним JSON-объектом без пояснений."
)

USER_PROMPT = """Извлеки данные с фотографии документа. Верни JSON строго такой структуры:

{
  "doc_type": "СТС РФ" | "Свидетельство РБ" | null,
  "plate": "гос. номер как напечатан",
  "vin": "17 символов без пробелов, ЗАГЛАВНЫМИ, латиница",
  "brand": "марка",
  "model": "модель",
  "vehicle_type": "тип ТС как напечатан",
  "category": "категория ТС",
  "year": 2013,
  "color": "цвет",
  "displacement": 11997,
  "engine_model": "модель двигателя, если напечатана",
  "max_mass": 34000,
  "curb_mass": 7700,
  "sts": "серия и номер свидетельства",
  "sts_issue_date": "ДД.ММ.ГГГГ или null",
  "pts": "паспорт ТС / ПТС или null",
  "owner": "собственник/владелец, если напечатан; иначе null",
  "lessor": "лизингодатель, если напечатан; иначе null",
  "unreadable": ["список полей, которые не читаются"],
  "notes": "прочие важные отметки или null"
}

Требования:
- VIN: только символы 0-9 и латинские буквы, ровно 17 знаков; цифру 0 не путать с буквой O,
  единицу с I, восьмёрку с B, пятёрку с S. Если сомневаешься между 0 и D/O, 8 и B — ставь null
  и добавляй поле в unreadable.
- Год выпуска — целое число из поля «Год выпуска ТС».
- Массы — целые числа в кг, без единиц измерения.
- displacement — «Рабочий объём двигателя, см³»; у прицепов его нет, тогда null.
- engine_model — «Модель двигателя»; если поле пустое или стоит «ОТСУТСТВУЕТ» — null.
- Если поля нет в документе (например, владелец не напечатан на лицевой стороне) — null.
- unreadable: пустой список, если всё прочитано уверенно."""


LOG_PATH = Path(__file__).resolve().parent / "logs" / "scan_cars.log"
_log_handle = None


def _log_file():
    global _log_handle
    if _log_handle is None:
        LOG_PATH.parent.mkdir(parents=True, exist_ok=True)
        _log_handle = open(LOG_PATH, "a", encoding="utf-8")
        _log_handle.write("\n" + "=" * 96 + "\n")
        _log_handle.write(
            f"{datetime.now():%d.%m.%Y %H:%M:%S}  argv: {' '.join(sys.argv[1:])}\n"
        )
    return _log_handle


def log(msg: str = "") -> None:
    print(msg, flush=True)
    try:
        handle = _log_file()
        handle.write(msg + "\n")
        handle.flush()
    except Exception:  # лог не должен ломать основной сценарий
        pass


def load_api_key(cli_value: str | None) -> str | None:
    if cli_value:
        return cli_value.strip()
    env = os.environ.get("DEEPSEEK_API_KEY")
    if env:
        return env.strip()
    key_file = Path(__file__).resolve().parent / ".deepseek_key"
    if key_file.exists():
        text = key_file.read_text(encoding="utf-8").strip()
        return text or None
    return None


def collect_files(paths: list[str]) -> list[Path]:
    files: list[Path] = []
    for raw in paths:
        p = Path(raw)
        if p.is_dir():
            files.extend(sorted(q for q in p.iterdir() if q.suffix.lower() in IMAGE_EXT | {".pdf"}))
        elif p.is_file():
            files.append(p)
        else:
            log(f"! пропущено (не найдено): {raw}")
    return [f for f in files if f.suffix.lower() in IMAGE_EXT | {".pdf"}]


def prepare_image(path: Path) -> tuple[bytes, str]:
    """Возвращает (bytes, mime). PDF рендерится в PNG, крупные фото уменьшаются."""
    if path.suffix.lower() == ".pdf":
        import fitz  # PyMuPDF

        doc = fitz.open(path)
        pix = doc[0].get_pixmap(dpi=200)
        raw = pix.tobytes("png")
        doc.close()
        return maybe_shrink(raw, "image/png")
    return maybe_shrink(path.read_bytes(), None)


def maybe_shrink(data: bytes, forced_mime: str | None) -> tuple[bytes, str]:
    try:
        from PIL import Image
    except ImportError:  # Pillow есть в bundled Python, но пусть не падает
        mime = forced_mime or "image/jpeg"
        return data, mime

    img = Image.open(io.BytesIO(data))
    img = img.convert("RGB")
    w, h = img.size
    scale = MAX_SIDE / max(w, h)
    if scale < 1:
        img = img.resize((int(w * scale), int(h * scale)), Image.LANCZOS)
    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=88, optimize=True)
    return buf.getvalue(), "image/jpeg"


class ModelTruncated(RuntimeError):
    """Модель израсходовала лимит токенов на рассуждение и не выдала JSON."""


def _request(image: bytes, mime: str, api_key: str, model: str, max_tokens: int) -> dict:
    b64 = base64.b64encode(image).decode("ascii")
    body = {
        "model": model,
        "messages": [
            {"role": "system", "content": SYSTEM_PROMPT},
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": USER_PROMPT},
                    {"type": "image_url", "image_url": {"url": f"data:{mime};base64,{b64}"}},
                ],
            },
        ],
        "max_tokens": max_tokens,
        "response_format": {"type": "json_object"},
    }
    data = json.dumps(body).encode("utf-8")
    req = urllib.request.Request(
        API_URL,
        data=data,
        headers={
            "Content-Type": "application/json",
            "Authorization": f"Bearer {api_key}",
        },
        method="POST",
    )
    with urllib.request.urlopen(req, timeout=240) as resp:
        return json.loads(resp.read().decode("utf-8"))


def call_model(image: bytes, mime: str, api_key: str, model: str,
               max_tokens: int = DEFAULT_MAX_TOKENS) -> dict:
    """Запрос к модели. Пустой content — повтор с удвоенным лимитом."""
    limits = [max_tokens, max_tokens * 2]
    for attempt, limit in enumerate(limits, start=1):
        payload = _request(image, mime, api_key, model, limit)
        choice = payload["choices"][0]
        message = choice.get("message") or {}
        content = (message.get("content") or "").strip()
        finish = choice.get("finish_reason")
        usage = payload.get("usage") or {}
        spent = usage.get("completion_tokens")

        if content:
            try:
                return json.loads(strip_json(content))
            except json.JSONDecodeError as exc:
                raise RuntimeError(
                    f"модель вернула не-JSON (finish_reason={finish}): {content[:200]!r}"
                ) from exc

        truncated = finish == "length"
        if attempt < len(limits):
            log(f"  ~ модель не успела выдать JSON за {limit} токенов "
                f"(finish_reason={finish}, потрачено {spent}); повтор с лимитом {limits[attempt]}...")
            continue

        if truncated:
            raise ModelTruncated(
                f"модель израсходовала все {limit} токенов на рассуждение и не выдала JSON "
                f"(finish_reason=length). Запустите ещё раз или укажите --max-tokens побольше."
            )
        raise RuntimeError(
            f"пустой ответ модели (finish_reason={finish}, потрачено {spent} токенов)"
        )



def strip_json(text: str) -> str:
    text = text.strip()
    text = re.sub(r"^```(?:json)?\s*|\s*```$", "", text)
    start, end = text.find("{"), text.rfind("}")
    if start != -1 and end != -1:
        return text[start : end + 1]
    return text


def replay_lookup(replay_path: str, path: Path) -> dict | None:
    data = json.loads(Path(replay_path).read_text(encoding="utf-8"))
    return data.get(path.name) or data.get("*")


def main() -> int:
    ap = argparse.ArgumentParser(description="Распознать СТС и дописать в реестр автомобилей")
    ap.add_argument("paths", nargs="*", help="файлы и/или папки с фото документов")
    ap.add_argument("--registry", default=str(Path(__file__).resolve().parent.parent / "Автомобили.xlsx"))
    ap.add_argument("--model", default=os.environ.get("SCAN_CARS_MODEL", DEFAULT_MODEL))
    ap.add_argument("--max-tokens", type=int, default=DEFAULT_MAX_TOKENS, metavar="N",
                    help="потолок выходных токенов (модель тратит их и на рассуждение)")
    ap.add_argument("--api-key", default=None)
    ap.add_argument("--dry-run", action="store_true", help="только показать, что распознано")
    ap.add_argument("--customer", default=None, help="заказчик для всех добавляемых строк")
    ap.add_argument("--owner", default=None, help="владелец для всех добавляемых строк")
    ap.add_argument("--format-plates", action="store_true",
                    help="добавлять пробелы в прицепных номерах: ВР265577 -> ВР 2655 77")
    ap.add_argument("--replay", default=os.environ.get("SCAN_CARS_REPLAY"),
                    help="JSON-файл с готовыми ответами модели (для тестов, без API)")
    ap.add_argument("--dump-json", default=None, help="сохранить сырые ответы модели в файл")
    args = ap.parse_args()

    files = collect_files(args.paths)
    if not files:
        if not args.paths:
            log("Не указаны файлы или папки с документами.")
            log("Перетащите папку (или файл) мышкой на scan_cars.cmd либо укажите путь:")
            log('  scan_cars.cmd "D:\\Фото авто" --dry-run')
        else:
            log("Не найдено ни одного файла (ожидаются .jpg .jpeg .png .webp .pdf).")
        log(f"Журнал запуска: {LOG_PATH}")
        return 1

    api_key = None
    if not args.replay:
        api_key = load_api_key(args.api_key)
        if not api_key:
            log("Не найден ключ API. Укажите --api-key, переменную DEEPSEEK_API_KEY "
                "или файл tools/.deepseek_key")
            return 1

    registry_path = Path(args.registry)
    if not registry_path.exists() and not args.dry_run:
        log(f"Реестр не найден: {registry_path}")
        log('Укажите файл явно: --registry "C:\\путь\\Автомобили.xlsx"')
        log("Либо запустите с --dry-run, чтобы только посмотреть распознанное.")
        log(f"Журнал запуска: {LOG_PATH}")
        return 1

    try:
        existing = read_registry(args.registry)
    except RegistryLocked as exc:
        log(f"! {exc}")
        log("  Ничего не изменено. Закройте файл в Excel и запустите снова.")
        log(f"Журнал запуска: {LOG_PATH}")
        return 1

    dumps, ok, failed, warned = [], 0, 0, 0
    log(f"Файлов к обработке: {len(files)}; реестр: {args.registry}")
    log("-" * 96)

    for path in files:
        log(f"→ {path.name}")
        try:
            if args.replay:
                parsed = replay_lookup(args.replay, path)
                if parsed is None:
                    raise RuntimeError("нет ответа в replay-файле")
            else:
                img, mime = prepare_image(path)
                parsed = call_model(img, mime, api_key, args.model, args.max_tokens)
        except urllib.error.HTTPError as exc:
            detail = exc.read().decode("utf-8", "replace")[:300]
            log(f"  ! API вернул {exc.code}: {detail}")
            failed += 1
            continue
        except Exception as exc:  # noqa: BLE001 — отчёт важнее трейсбека
            log(f"  ! ошибка распознавания: {exc}")
            failed += 1
            continue

        parsed["source_file"] = path.name
        if args.customer and not parsed.get("customer"):
            parsed["customer"] = args.customer
        if args.owner and not parsed.get("owner"):
            parsed["owner"] = args.owner
        dumps.append(parsed)

        errors, warnings = validate_vehicle(parsed, existing)
        summary = " · ".join(
            str(x) for x in [
                parsed.get("plate"), parsed.get("vin"), parsed.get("brand"),
                parsed.get("model"), parsed.get("year"),
            ] if x
        )
        log(f"  распознано: {summary}")

        for w in warnings:
            log(f"  ~ {w}")
        if warnings:
            warned += 1
        if parsed.get("unreadable"):
            log(f"  ~ не читается: {', '.join(map(str, parsed['unreadable']))}")

        if errors:
            for e in errors:
                log(f"  ✗ {e}")
            log("  → строка НЕ добавлена")
            failed += 1
            continue

        if args.dry_run:
            log("  → dry-run: строка не записана")
            ok += 1
            continue

        try:
            result = append_vehicle(args.registry, parsed,
                                    format_plates=args.format_plates)
        except RegistryLocked as exc:
            log(f"  ! {exc}")
            log("  Ничего не изменено. Закройте файл в Excel и запустите снова.")
            failed += 1
            break

        existing.append({"row_no": result["row"] - 1, "plate": parsed.get("plate"),
                         "vin": parsed.get("vin")})
        log(f"  ✓ строка №{result['row'] - 1} добавлена")
        ok += 1

    log("-" * 96)
    if args.dry_run:
        log(f"Итого распознано: {ok}, с ошибками {failed}, с предупреждениями {warned}")
        log("Записи в реестр не было (--dry-run).")
    else:
        log(f"Итого: добавлено {ok}, с ошибками {failed}, с предупреждениями {warned}")

    if args.dump_json:
        Path(args.dump_json).write_text(
            json.dumps(dumps, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        log(f"Сырые ответы сохранены: {args.dump_json}")

    log(f"Журнал запуска: {LOG_PATH}")
    return 0 if failed == 0 else 2


if __name__ == "__main__":
    try:
        sys.exit(main())
    except SystemExit:
        raise
    except BaseException:  # noqa: BLE001 — всё, что не поймали, должно попасть в лог
        import traceback

        log("НЕОБРАБОТАННАЯ ОШИБКА:")
        log(traceback.format_exc())
        sys.exit(3)
