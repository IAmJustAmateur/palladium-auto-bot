# -*- coding: utf-8 -*-
"""Краткая инструкция установки бота на новый ПК (DOCX, ~2 страницы)."""
import sys
from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

sys.stdout.reconfigure(encoding="utf-8")

REPO = "https://github.com/IAmJustAmateur/palladium-auto-bot.git"
DOCX = Path(r"C:\Users\ighiz\pycode\palladium-auto-bot\docs\Установка бота на новый компьютер.docx")

doc = Document()
normal = doc.styles["Normal"]
normal.font.name = "Calibri"
normal.font.size = Pt(10.5)
normal.paragraph_format.space_after = Pt(4)
normal.paragraph_format.line_spacing = 1.05

for section in doc.sections:
    section.top_margin = Cm(1.8)
    section.bottom_margin = Cm(1.8)
    section.left_margin = Cm(2.0)
    section.right_margin = Cm(1.6)


def shade(paragraph, color="F2F2F2"):
    pPr = paragraph._p.get_or_add_pPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), color)
    pPr.append(shd)


def code(text):
    for line in text.strip("\n").split("\n"):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.left_indent = Cm(0.3)
        run = p.add_run(line if line else " ")
        run.font.name = "Consolas"
        run.font.size = Pt(9)
        shade(p)
    doc.add_paragraph().paragraph_format.space_after = Pt(0)


def step(text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(9)
    p.paragraph_format.space_after = Pt(2)
    run = p.add_run(text)
    run.bold = True
    run.font.size = Pt(11.5)


def para(text, bold_prefix=None):
    p = doc.add_paragraph()
    if bold_prefix:
        run = p.add_run(bold_prefix)
        run.bold = True
    p.add_run(text)
    return p


def bullet(text, bold_prefix=None):
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.space_after = Pt(1)
    if bold_prefix:
        run = p.add_run(bold_prefix)
        run.bold = True
    p.add_run(text)


def note(text):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(0.3)
    p.paragraph_format.space_before = Pt(3)
    run = p.add_run("ВАЖНО: ")
    run.bold = True
    run.font.color.rgb = RGBColor(0xB0, 0x2A, 0x2A)
    p.add_run(text)
    shade(p, "FDF3F3")


def table(headers, rows, widths):
    t = doc.add_table(rows=1, cols=len(headers))
    t.style = "Light Grid Accent 1"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, h in enumerate(headers):
        cell = t.rows[0].cells[i]
        cell.text = ""
        run = cell.paragraphs[0].add_run(h)
        run.bold = True
        run.font.size = Pt(9)
    for row in rows:
        cells = t.add_row().cells
        for i, value in enumerate(row):
            cells[i].text = ""
            run = cells[i].paragraphs[0].add_run(str(value))
            run.font.size = Pt(9)
            if i == 0:
                run.font.name = "Consolas"
    for row in t.rows:
        for i, w in enumerate(widths):
            row.cells[i].width = Cm(w)
    doc.add_paragraph().paragraph_format.space_after = Pt(0)


# ---------- заголовок ----------
title = doc.add_paragraph()
title.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = title.add_run("Бот распознавания техпаспортов — установка на новый ПК")
run.bold = True
run.font.size = Pt(15)

sub = doc.add_paragraph()
sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = sub.add_run("ООО «Палладиум – арт» · 20 минут · Windows 10/11")
run.font.size = Pt(9.5)
run.font.color.rgb = RGBColor(0x55, 0x55, 0x55)

# ---------- 1 ----------
step("1. Python и Git")
para("Python 3.12: https://www.python.org/downloads/ — при установке обязательно галочка "
     "«Add python.exe to PATH».")
para("Git: https://git-scm.com/download/win — везде «Next».")
para("Проверка в cmd (Win+R → cmd):")
code("py --version\ngit --version")
para("Дополнительные пакеты не нужны.")

# ---------- 2 ----------
step("2. Код бота")
code(f"cd C:\\\ngit clone {REPO} C:\\PalladiumAutoBot")
para("Репозиторий публичный, аккаунт GitHub не требуется. После клонирования в папке "
     "C:\\PalladiumAutoBot появятся tg_bot.py, run_bot.cmd и другие файлы.")

# ---------- 3 ----------
step("3. Секреты — три файла в C:\\PalladiumAutoBot")
table(
    ["Файл", "Что внутри"],
    [
        [".tg_token", "токен бота, одна строка"],
        [".deepseek_key", "ключ DeepSeek, одна строка"],
        ["tg_config.json",
         '{"allowed_chat_ids": [5085693770], "max_per_minute": 3}'],
    ],
    widths=[4.0, 13.0],
)
para("Создавать Блокнотом: «Файл → Сохранить как», имя с точкой в начале, тип «Все файлы», "
     "кодировка UTF-8, без .txt на конце.")
para("Ключ DeepSeek можно посмотреть на текущем ПК командой в PowerShell:")
code("[Environment]::GetEnvironmentVariable('DEEPSEEK_API_KEY','User')")
note("Токен и ключ — это пароли. Их не хранят в репозитории (там они уже исключены) и не "
     "пересылают вместе с этим файлом.")

# ---------- 4 ----------
step("4. Проверка запуска")
para("Двойной клик по run_bot.cmd. Ожидаемая строка в окне и в журнале "
     "C:\\PalladiumAutoBot\\logs\\tg_bot.log:")
code("Старт: @PalladiumAutoBot (id 9147971503), допущено чатов: 1")
para("Отправить боту фото техпаспорта — должен прийти ответ с данными и кнопками. "
     "Остановить: закрыть окно.")

# ---------- 5 ----------
step("5. Автозапуск")
bullet("Win+R → taskschd.msc → «Создать задачу» (не «простую»).")
bullet("«Общие»: имя Palladium Auto Bot, «Выполнять только при входе пользователя».")
bullet("«Триггеры» → Создать → «При входе в систему».")
code("«Действия» → Создать → Запуск программы:\nПрограмма:     C:\\PalladiumAutoBot\\run_bot.cmd\nРабочая папка: C:\\PalladiumAutoBot")
bullet("«Условия»: снять «Запускать только при питании от сети» (для ноутбука).")
para("Проверка: правый клик по задаче → «Выполнить» → в журнале появляется «Старт».")

# ---------- 6 ----------
step("6. Запретить сон (PowerShell от администратора)")
code("powercfg /change standby-timeout-ac 0\npowercfg /change hibernate-timeout-ac 0\npowercfg /change monitor-timeout-ac 10")
para("Ноутбук: «Электропитание» → «Действие при закрытии крышки» → «Ничего не делать».")

# ---------- 7 ----------
step("7. Проверка после перезагрузки")
para("Перезагрузить, войти в систему, через минуту проверить свежую строку «Старт» в журнале "
     "и отправить боту фото.")

# ---------- пользователи ----------
step("8. Как добавить сотрудника")
bullet("Сотрудник пишет боту @PalladiumAutoBot команду /start.")
bullet("Бот отвечает «Доступ закрыт. Ваш chat_id: 1234567890» — это его номер.")
bullet("Администратор вписывает номер в список в файле tg_config.json через запятую.")
bullet("Сохранить файл и перезапустить бота: Планировщик → «Завершить», затем «Выполнить».")
bullet("Сотрудник снова пишет /start — бот начинает отвечать на фото.")
para("Итоговый вид файла tg_config.json с двумя сотрудниками:")
code('{\n  "allowed_chat_ids": [5085693770, 1234567890],\n  "max_per_minute": 3\n}')
para("Убрать доступ: удалить номер из списка и перезапустить бота.")
# ---------- проблемы ----------
step("Если не работает")
table(
    ["Симптом", "Что делать"],
    [
        ["py не является внутренней командой", "переустановить Python с галочкой PATH"],
        ["Не найден токен / ключ", "проверить имя файла (точка в начале, без .txt) и содержимое"],
        ["Бот уже запущен (замок занят)", "закрыть лишнее окно, удалить .tg_bot.pid"],
        ["Ответ «Доступ закрыт»", "добавить chat_id в tg_config.json, перезапустить"],
        ["Бот молчит", "проверить журнал, сон, задачу в Планировщике"],
        ["«Не удалось распознать»", "переснять ровнее, ближе, без блика"],
    ],
    widths=[6.0, 11.0],
)
para("Журнал: C:\\PalladiumAutoBot\\logs\\tg_bot.log")

# ---------- обновление ----------
step("Обновление кода")
code("cd C:\\PalladiumAutoBot\ngit pull")
para("Затем в Планировщике: «Завершить» и «Выполнить».")

doc.save(DOCX)
print("saved:", DOCX)
