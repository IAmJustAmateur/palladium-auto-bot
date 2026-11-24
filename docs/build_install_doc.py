# -*- coding: utf-8 -*-
"""Сборка инструкции по установке бота на новый ПК (DOCX)."""
import sys
from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

sys.stdout.reconfigure(encoding="utf-8")

OUT = Path(r"C:\Users\ighiz\pycode\palladium-auto-bot\docs\Установка бота на новый компьютер.docx")
OUT.parent.mkdir(parents=True, exist_ok=True)

doc = Document()

# --- базовые стили ---
normal = doc.styles["Normal"]
normal.font.name = "Calibri"
normal.font.size = Pt(10.5)
normal.paragraph_format.space_after = Pt(4)
normal.paragraph_format.line_spacing = 1.08

for section in doc.sections:
    section.top_margin = Cm(2)
    section.bottom_margin = Cm(2)
    section.left_margin = Cm(2.2)
    section.right_margin = Cm(1.8)


def shade(paragraph, color="F2F2F2"):
    pPr = paragraph._p.get_or_add_pPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), color)
    pPr.append(shd)


def code(text, color="F2F2F2"):
    """Блок команды: моноширинный, с заливкой."""
    for line in text.strip("\n").split("\n"):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.left_indent = Cm(0.4)
        run = p.add_run(line if line else " ")
        run.font.name = "Consolas"
        run.font.size = Pt(9)
        run._element.rPr.rFonts.set(qn("w:eastAsia"), "Consolas")
        shade(p, color)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)


def step(title):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(2)
    run = p.add_run(title)
    run.bold = True
    run.font.size = Pt(12)
    return p


def para(text, bold_prefix=None):
    p = doc.add_paragraph()
    if bold_prefix:
        run = p.add_run(bold_prefix)
        run.bold = True
    p.add_run(text)
    return p


def bullet(text, bold_prefix=None):
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.space_after = Pt(2)
    if bold_prefix:
        run = p.add_run(bold_prefix)
        run.bold = True
    p.add_run(text)
    return p


def note(text, label="ВАЖНО"):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(0.3)
    p.paragraph_format.space_before = Pt(4)
    run = p.add_run(f"{label}: ")
    run.bold = True
    run.font.color.rgb = RGBColor(0xB0, 0x2A, 0x2A)
    p.add_run(text)
    shade(p, "FDF3F3")


def table(headers, rows, widths=None):
    t = doc.add_table(rows=1, cols=len(headers))
    t.style = "Light Grid Accent 1"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr = t.rows[0].cells
    for i, h in enumerate(headers):
        hdr[i].text = ""
        run = hdr[i].paragraphs[0].add_run(h)
        run.bold = True
        run.font.size = Pt(9.5)
    for row in rows:
        cells = t.add_row().cells
        for i, value in enumerate(row):
            cells[i].text = ""
            run = cells[i].paragraphs[0].add_run(str(value))
            run.font.size = Pt(9.5)
    if widths:
        for row in t.rows:
            for i, w in enumerate(widths):
                row.cells[i].width = Cm(w)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)


# ============ ТИТУЛ ============
title = doc.add_paragraph()
title.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = title.add_run("Установка Telegram-бота распознавания техпаспортов\nна новый компьютер")
run.bold = True
run.font.size = Pt(17)

sub = doc.add_paragraph()
sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = sub.add_run("ООО «Палладиум – арт» · пошаговая инструкция")
run.font.size = Pt(11)
run.font.color.rgb = RGBColor(0x55, 0x55, 0x55)

doc.add_paragraph()

para("Документ рассчитан на человека без опыта программирования. Все команды нужно "
     "копировать целиком, как написано. Время установки — 20–30 минут.")
para("Секретные данные (токен бота и ключ модели) в этот файл не входят и в переписке не "
     "передаются. Их нужно вставить на месте, в шаге 5.")

note("Если какой-то шаг не совпадает с тем, что видно на экране, — остановитесь и напишите "
     "мне. Не пытайтесь «догадаться», особенно там, где речь о токенах.", "ВНИМАНИЕ")

# ============ ЧТО ПОЛУЧИТСЯ ============
step("Что должно получиться в итоге")
bullet("на компьютере появится папка C:\\PalladiumAutoBot с кодом бота;")
bullet("бот @PalladiumAutoBot запускается сам при входе в систему и работает в фоне;")
bullet("сотрудник присылает боту фото техпаспорта и получает данные для CRM;")
bullet("журнал работы пишется в C:\\PalladiumAutoBot\\logs\\tg_bot.log.")

# ============ ШАГ 1 ============
step("Шаг 1. Установить Python")
para("Открыть https://www.python.org/downloads/ и скачать Python 3.12 (Windows installer, 64-bit). "
     "Запустить установщик.")
note("На первом экране установщика обязательно поставить галочку «Add python.exe to PATH» — "
     "без неё ничего не заработает. Затем нажать «Install Now».", "ВАЖНО")
para("Проверка: нажать Win+R, ввести cmd, Enter. В чёрном окне набрать:")
code("py --version")
para("Должно ответить примерно так: Python 3.12.x. Если написано «не является внутренней или "
     "внешней командой» — Python установлен без галочки, переустановить.")
para("Больше ничего устанавливать не нужно: боту не требуются дополнительные пакеты.")

# ============ ШАГ 2 ============
step("Шаг 2. Установить Git")
para("Открыть https://git-scm.com/download/win, скачать и установить. Во всех окнах жать «Next», "
     "ничего не меняя.")
para("Проверка в том же чёрном окне:")
code("git --version")
para("Должно ответить: git version 2.x.x.")

# ============ ШАГ 3 ============
step("Шаг 3. Получить код бота")
para("Код лежит в репозитории на GitHub. Есть два варианта — какой использовать, зависит от того, "
     "сделан репозиторий публичным или приватным. Уточните у меня, если не знаете.")
doc.add_paragraph()

p = doc.add_paragraph()
run = p.add_run("Вариант А — репозиторий публичный (проще всего)")
run.bold = True
para("Аккаунт на GitHub не нужен. В чёрном окне (cmd) выполнить:")
code("cd C:\\\ngit clone https://github.com/ПОЛЬЗОВАТЕЛЬ/palladium-auto-bot.git C:\\PalladiumAutoBot")
para("Вместо ПОЛЬЗОВАТЕЛЬ подставить логин владельца репозитория (уточнить у меня).")
doc.add_paragraph()

p = doc.add_paragraph()
run = p.add_run("Вариант Б — репозиторий приватный")
run.bold = True
para("Нужен «токен доступа» GitHub (его выдаёт владелец репозитория и присылает отдельно от этого "
     "файла). Команда:")
code("cd C:\\\ngit clone https://ТОКЕН@github.com/ПОЛЬЗОВАТЕЛЬ/palladium-auto-bot.git C:\\PalladiumAutoBot")
para("После клонирования токен останется в файле C:\\PalladiumAutoBot\\.git\\config. Если это "
     "нежелательно — сообщите, покажу, как убрать.")
doc.add_paragraph()

para("Проверка: в папке C:\\PalladiumAutoBot должны появиться файлы tg_bot.py, run_bot.cmd, "
     "README.md, ИНСТРУКЦИЯ.md и папка .git.")

# ============ ШАГ 4 ============
step("Шаг 4. Проверить папку")
para("Открыть Проводник и перейти в C:\\PalladiumAutoBot. Должно быть так:")
table(
    ["Файл", "Что это"],
    [
        ["tg_bot.py", "сам бот"],
        ["run_bot.cmd", "запуск бота (с автоперезапуском)"],
        ["tg_config.json", "список допущенных сотрудников (создаём в шаге 5)"],
        [".tg_token", "токен бота (создаём в шаге 5)"],
        [".deepseek_key", "ключ модели (создаём в шаге 5)"],
        ["README.md, ИНСТРУКЦИЯ.md", "документация"],
    ],
    widths=[6, 11],
)

# ============ ШАГ 5 ============
step("Шаг 5. Вставить секреты")
para("Нужны три файла. Их нет в репозитории и их не присылают по почте — значения вставляет "
     "человек на месте.")

p = doc.add_paragraph()
run = p.add_run("5.1. Токен бота — файл .tg_token")
run.bold = True
para("В папке C:\\PalladiumAutoBot: правый клик → «Создать» → «Текстовый документ». "
     "Назвать его .tg_token (с точкой в начале, без .txt на конце). Открыть Блокнотом, "
     "вставить одну строку с токеном, сохранить.")
code("<ТОКЕН_БОТА>")
note("Токен выглядит как 8123456789:AAH4kLmNoPqRsTuVwXyZ1234567890abcde. Это пароль от бота: "
     "не пересылать в мессенджерах и не вставлять в письма.", "ВАЖНО")

p = doc.add_paragraph()
run = p.add_run("5.2. Ключ модели — файл .deepseek_key")
run.bold = True
para("Таким же образом создать файл .deepseek_key и вставить в него ключ DeepSeek одной строкой.")
code("<КЛЮЧ_DEEPSEEK>")
para("Взять ключ можно на том компьютере, где бот уже работает. В PowerShell выполнить:")
code("[Environment]::GetEnvironmentVariable('DEEPSEEK_API_KEY','User')")
para("Команда покажет ключ — скопировать и перенести на новый компьютер.")

p = doc.add_paragraph()
run = p.add_run("5.3. Список доступа — файл tg_config.json")
run.bold = True
para("Создать файл tg_config.json (обычный текстовый, с расширением .json) со следующим "
     "содержимым:")
code('{\n  "allowed_chat_ids": [5085693770],\n  "max_per_minute": 3\n}')
para("В квадратных скобках — номера сотрудников, которым разрешено пользоваться ботом, через "
     "запятую. 5085693770 — уже известный номер руководителя.")
para("Как узнать номер другого сотрудника: он пишет боту /start, бот отвечает «Бот ещё не "
     "настроен. Ваш chat_id: …». Этот номер добавить в список и перезапустить бота.")

# ============ ШАГ 6 ============
step("Шаг 6. Первый запуск вручную")
para("Двойной клик по run_bot.cmd. В окне появится:")
code("Старт: @PalladiumAutoBot (id 9147971503), допущено чатов: 1")
para("Это значит, что бот подключился к Telegram. Теперь отправить ему фото техпаспорта — "
     "должен прийти ответ с данными и кнопками.")
para("Остановить бота: закрыть чёрное окно.")
note("Если в окне написано «Не найден токен» или «Не найден ключ» — файлы из шага 5 созданы "
     "неправильно: проверьте имя файла (точка в начале, без .txt) и что внутри не пусто.", "ЕСЛИ ОШИБКА")

# ============ ШАГ 7 ============
step("Шаг 7. Автозапуск при включении компьютера")
para("Нужен, чтобы бот работал без ручного запуска.")
bullet("Нажать Win+R, ввести taskschd.msc, Enter.")
bullet("Справа выбрать «Создать задачу…» (именно «Создать задачу», не «Создать простую задачу»).")
bullet("Вкладка «Общие»: имя — Palladium Auto Bot. Галочка «Выполнять только при входе "
       "пользователя».")
bullet("Вкладка «Триггеры» → «Создать» → «Начало задачи: При входе в систему» → ОК.")
bullet("Вкладка «Действия» → «Создать» → «Запуск программы»:")
code("Программа:  C:\\PalladiumAutoBot\\run_bot.cmd\nРабочая папка:  C:\\PalladiumAutoBot")
bullet("Вкладка «Условия»: снять галочку «Запускать задачу только при питании от сети» "
       "(если это ноутбук).")
bullet("ОК. Пароль не спросят, потому что задача запускается при входе пользователя.")
para("Проверка: найти задачу в списке → правый клик → «Выполнить». В файле "
     "C:\\PalladiumAutoBot\\logs\\tg_bot.log появится новая строка «Старт».")
note("Бот работает, пока пользователь в системе. Блокировка экрана не мешает. "
     "«Выход из системы» — останавливает бота.", "ВАЖНО")

# ============ ШАГ 8 ============
step("Шаг 8. Запретить компьютеру засыпать")
para("Открыть PowerShell от имени администратора (правый клик по «Пуск» → «Терминал (Администратор)») "
     "и выполнить три команды:")
code("powercfg /change standby-timeout-ac 0\npowercfg /change hibernate-timeout-ac 0\npowercfg /change monitor-timeout-ac 10")
para("Сон и гибернация отключаются, монитор по-прежнему гаснет через 10 минут.")
para("Если это ноутбук — проверить, что при закрытии крышки он не засыпает: "
     "«Электропитание» → «Действие при закрытии крышки» → «Ничего не делать».")

# ============ ШАГ 9 ============
step("Шаг 9. Проверка после перезагрузки")
bullet("Перезагрузить компьютер и войти в систему.")
bullet("Подождать минуту, открыть C:\\PalladiumAutoBot\\logs\\tg_bot.log — там должна быть "
       "свежая строка «Старт».")
bullet("Отправить боту фото — убедиться, что ответ приходит.")
para("На этом установка закончена.")

# ============ ОБСЛУЖИВАНИЕ ============
step("Обновление кода (если понадобится)")
para("В чёрном окне:")
code("cd C:\\PalladiumAutoBot\ngit pull")
para("Затем перезапустить задачу: Планировщик → Palladium Auto Bot → «Завершить», потом «Выполнить».")

# ============ ПРОБЛЕМЫ ============
step("Если что-то не работает")
table(
    ["Что видно", "Причина", "Что делать"],
    [
        ["py не является внутренней командой", "Python без галочки PATH",
         "переустановить, поставить галочку «Add python.exe to PATH»"],
        ["Не найден токен", "нет файла .tg_token или он пустой",
         "проверить имя (точка в начале, без .txt) и содержимое"],
        ["Не найден ключ DeepSeek", "нет файла .deepseek_key",
         "создать файл по шагу 5.2"],
        ["Бот уже запущен (замок занят)", "работает второй экземпляр",
         "закрыть лишнее окно, удалить C:\\PalladiumAutoBot\\.tg_bot.pid"],
        ["В журнале «Конфликт 409»", "бот запущен на двух компьютерах с одним токеном",
         "оставить один экземпляр"],
        ["Бот молчит", "компьютер спит / нет интернета / задача не запущена",
         "проверить журнал, электропитание, Планировщик"],
        ["Ответ «Доступ закрыт»", "chat_id нет в tg_config.json",
         "вписать номер, перезапустить бота"],
        ["«Не удалось надёжно распознать»", "фото смазано или обрезано",
         "переснять ровнее, ближе, без блика"],
    ],
    widths=[4.5, 5, 7],
)

para("Журналы: C:\\PalladiumAutoBot\\logs\\tg_bot.log — все запуски, ошибки и распознанные "
     "машины.")

# ============ ПРИЛОЖЕНИЕ ============
step("Приложение. Что нельзя делать с секретами")
bullet("не отправлять .tg_token и .deepseek_key по почте, в мессенджерах и не хранить в облаке;")
bullet("не добавлять эти файлы в Git — они уже исключены через .gitignore;")
bullet("если токен попал в чужие руки — в Telegram написать @BotFather команду /revoke и "
       "выпустить новый токен;")
bullet("ключ DeepSeek при утечке перевыпускается в личном кабинете на platform.deepseek.com.")

doc.save(OUT)
print("saved:", OUT)
