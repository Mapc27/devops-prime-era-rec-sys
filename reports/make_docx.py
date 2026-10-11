#!/usr/bin/env python3
"""Сборка отчёта в .docx из измеренных метрик.

Все измеренные значения — таблицы и числа в тексте — читаются из
reports/metrics.csv, reports/data_stats.json и reports/hybrid_weights.json
(их пишет python -m recsys.evaluate), а не вписываются руками: после
перезапуска оценки достаточно пересобрать отчёт.

Скриншоты кладутся в reports/figures/screens/ под именами из SCREENSHOTS;
пока файла нет, на его месте печатается заметная заглушка.

Запуск:
    .venv/bin/python reports/make_docx.py          # только .docx
    .venv/bin/python reports/make_docx.py --pdf    # + .pdf через LibreOffice
"""

from __future__ import annotations

import json
import shutil
import subprocess
import sys
from pathlib import Path

import pandas as pd
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.image.image import Image as DocxImage
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

ROOT = Path(__file__).resolve().parent.parent
REPORTS = ROOT / "reports"
IMG = REPORTS / "figures"
SCREENS = IMG / "screens"
OUT = REPORTS / "RecSys_report.docx"

REPO_URL = "https://github.com/<ваш-аккаунт>/<имя-репозитория>"
AUTHOR = ""  # ФИО, если нужно указать в шапке

# A4 минус поля: все таблицы и изображения нормируются под эту ширину
PAGE_W, PAGE_H = 8.27, 11.69
MARGIN_L, MARGIN_R = 1.0, 0.8
USABLE = PAGE_W - MARGIN_L - MARGIN_R

BODY_FONT = "Times New Roman"
HEAD_FONT = "Arial"
MONO_FONT = "Courier New"
ACCENT = RGBColor(0x1F, 0x3A, 0x5F)
MUTED = RGBColor(0x55, 0x55, 0x55)
WARN = RGBColor(0xB0, 0x30, 0x30)

# Порядок рисунков в документе. Номер рисунка вычисляется из этого списка,
# поэтому ссылки в тексте не сбиваются при добавлении скриншотов.
FIGURES = [
    "data",
    "architecture",
    "scr_evaluate",
    "metrics",
    "scr_swagger",
    "scr_history",
    "scr_recommend",
    "scr_similar",
    "scr_also_bought",
    "scr_cold",
    "scr_422",
]

# пользователь и товар для демонстрации API (есть и в train, и в полном датасете)
DEMO_USER = "AHZUR65AQ7QWJ5BFBZFHBJIYSRPQ"
DEMO_ITEM = "B004HD55VK"  # Tomb Raider (2013)
API = "http://localhost:8000"

# запрос, ответ на который снят на скриншоте: печатается над картинкой,
# потому что адресная строка браузера в кадр не попадает
REQUESTS = {
    "scr_history": f"GET /history/{DEMO_USER}?k=10",
    "scr_recommend": f"GET /recommend/{DEMO_USER}?k=10",
    "scr_cold": "GET /recommend/new_user_123?k=5",
    "scr_similar": f"GET /item/{DEMO_ITEM}/similar?k=5",
    "scr_also_bought": f"GET /item/{DEMO_ITEM}/also_bought?k=5",
    "scr_422": "GET /popular?k=0",
}

# что должно быть на каждом скриншоте — печатается в заглушке
SCREENSHOTS = {
    "scr_evaluate": ("evaluate.png",
                     "терминал: .venv/bin/python -m recsys.evaluate — строки «best hybrid…» "
                     "(подбор весов на валидации) и итоговая таблица метрик"),
    "scr_swagger": ("swagger.png",
                    f"браузер: {API}/docs — список эндпоинтов"),
    "scr_history": ("history.png",
                    f"браузер с видимой адресной строкой и включённым Pretty print: "
                    f"{API}/history/{DEMO_USER}?k=10 — старые игры для Wii и Xbox 360"),
    "scr_recommend": ("recommend_warm.png",
                      f"браузер: {API}/recommend/{DEMO_USER}?k=10 — source=hybrid, "
                      f"в выдаче геймпад PS5, Zelda, аксессуары для Switch"),
    "scr_cold": ("recommend_cold.png",
                 f"браузер: {API}/recommend/new_user_123?k=5 — source=trending_fallback"),
    "scr_similar": ("similar.png",
                    f"браузер: {API}/item/{DEMO_ITEM}/similar?k=5 — другие издания Tomb Raider"),
    "scr_also_bought": ("also_bought.png",
                        f"браузер: {API}/item/{DEMO_ITEM}/also_bought?k=5 — BioShock Infinite, Crysis 3, Far Cry 3"),
    "scr_422": ("validation_422.png",
                f"браузер: {API}/popular?k=0 — ответ 422 с текстом ошибки валидации"),
}


# --------------------------------------------------------------------------- #
# метрики
# --------------------------------------------------------------------------- #
def n(x: float, d: int = 3) -> str:
    """Число с десятичной запятой."""
    return f"{x:.{d}f}".replace(".", ",")


def pc(x: float, d: int = 0) -> str:
    """Доля -> проценты с десятичной запятой."""
    return f"{x * 100:.{d}f}".replace(".", ",") + "%"


def num(x: int) -> str:
    """Целое с пробелом между разрядами: 9 610."""
    return f"{x:,}".replace(",", " ")


def thousands(x: float) -> str:
    return f"{x / 1000:.1f}".replace(".", ",").replace(",0", "") + " тыс."


S = json.loads((REPORTS / "data_stats.json").read_text(encoding="utf-8"))
WEIGHTS = json.loads((REPORTS / "hybrid_weights.json").read_text(encoding="utf-8"))
_m = pd.read_csv(REPORTS / "metrics.csv")
W = _m[_m.segment == "warm"].set_index("model")
C = _m[_m.segment == "cold"].set_index("model")

TREND, HYB, HYB_NT, ALS = W.loc["Trending"], W.loc["Hybrid"], W.loc["Hybrid без Trending"], W.loc["ALS"]
GAIN = HYB["NDCG@10"] / TREND["NDCG@10"] - 1
DROP_NT = 1 - HYB_NT["NDCG@10"] / HYB["NDCG@10"]
TREND_SLOTS = round(HYB["TrendShare"] * 10)
ML_MODELS = ["ContentBased", "ItemKNN", "UserKNN", "ALS"]
BEST_ML = max(ML_MODELS, key=lambda m: W.loc[m, "NDCG@10"])
COLD_SHARE, NEW_ITEMS = pc(S["cold_share"]), pc(S["test_new_items"])
WH, WNT = WEIGHTS["hybrid"], WEIGHTS["hybrid_no_trending"]


def weights_str(w: dict) -> str:
    return ", ".join(f"{k} {n(v, 1)}" for k, v in w.items() if v)


def cov(x: float) -> str:
    return pc(x, 2) if x < 0.01 else pc(x)


# --------------------------------------------------------------------------- #
# инфраструктура оформления
# --------------------------------------------------------------------------- #
def set_font(run, name: str, size: float, bold=False, italic=False, color=None):
    run.font.name = name
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic
    if color is not None:
        run.font.color.rgb = color
    # кириллица берёт шрифт из тех же полей, но их надо проставить явно
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = rpr.makeelement(qn("w:rFonts"), {})
        rpr.append(rfonts)
    for attr in ("w:ascii", "w:hAnsi", "w:cs"):
        rfonts.set(qn(attr), name)


def add_marked(p, text: str, font: str, size: float, bold=False, italic=False, color=None):
    """Текст с разметкой **жирного** внутри строки."""
    for i, chunk in enumerate(text.split("**")):
        if chunk:
            set_font(p.add_run(chunk), font, size, bold or i % 2 == 1, italic, color)


def para(doc, text="", size=12, bold=False, italic=False, align=None,
         space_after=6, space_before=0, font=BODY_FONT, color=None, indent=None):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.line_spacing = 1.15
    if align is not None:
        p.alignment = align
    if indent is not None:
        p.paragraph_format.left_indent = Inches(indent)
    if text:
        add_marked(p, text, font, size, bold, italic, color)
    return p


def text(doc, s: str):
    """Обычный абзац по ширине; **жирное** размечается внутри строки."""
    return para(doc, s, align=WD_ALIGN_PARAGRAPH.JUSTIFY)


def h1(doc, s):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(20)
    p.paragraph_format.space_after = Pt(10)
    p.paragraph_format.keep_with_next = True
    set_font(p.add_run(s), HEAD_FONT, 16, bold=True, color=ACCENT)
    return p


def h2(doc, s):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.keep_with_next = True
    set_font(p.add_run(s), HEAD_FONT, 13, bold=True, color=ACCENT)
    return p


def bullets(doc, items, size=12):
    for item in items:
        p = doc.add_paragraph(style="List Bullet")
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.line_spacing = 1.15
        add_marked(p, item, BODY_FONT, size)


def numbered(doc, items, size=12):
    for i, item in enumerate(items, 1):
        p = para(doc, space_after=3, indent=0.3)
        p.paragraph_format.first_line_indent = Inches(-0.25)
        add_marked(p, f"{i}. {item}", BODY_FONT, size)


def problem(doc, title, prob, sol, sol_prob, discussion):
    """Блок части II-a: проблема → решение → проблема решения → обсуждение."""
    h2(doc, title)
    bullets(doc, [f"**Проблема:** {prob}", f"**Решение:** {sol}",
                  f"**Проблема решения:** {sol_prob}", f"**Обсуждение:** {discussion}"])


def shade(cell, hex_color: str):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.makeelement(qn("w:shd"), {})
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tc_pr.append(shd)


def repeat_header(row):
    """Шапка таблицы повторяется на каждой странице при переносе."""
    tr_pr = row._tr.get_or_add_trPr()
    el = tr_pr.makeelement(qn("w:tblHeader"), {})
    el.set(qn("w:val"), "true")
    tr_pr.append(el)


def table(doc, headers, rows, widths=None, size=10, header_fill="DCE6F1", bold_rows=()):
    t = doc.add_table(rows=1, cols=len(headers))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    repeat_header(t.rows[0])
    for i, s in enumerate(headers):
        cell = t.rows[0].cells[i]
        cell.text = ""
        p = cell.paragraphs[0]
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.space_before = Pt(2)
        set_font(p.add_run(str(s)), HEAD_FONT, size, bold=True)
        shade(cell, header_fill)
    for r, row in enumerate(rows):
        cells = t.add_row().cells
        for i, value in enumerate(row):
            cells[i].text = ""
            p = cells[i].paragraphs[0]
            p.paragraph_format.space_after = Pt(2)
            p.paragraph_format.space_before = Pt(2)
            add_marked(p, str(value), BODY_FONT, size, bold=r in bold_rows)
    if widths:
        # ширины задаются в удобных пропорциях и масштабируются под полосу набора
        scale = (USABLE - 0.07) / sum(widths)
        for row in t.rows:
            for i, w in enumerate(widths):
                row.cells[i].width = Inches(w * scale)
        # LibreOffice берёт ширины из сетки таблицы, а не из ячеек
        for col, w in zip(t._tbl.tblGrid.findall(qn("w:gridCol")), widths):
            col.set(qn("w:w"), str(int(w * scale * 1440)))
    doc.add_paragraph().paragraph_format.space_after = Pt(4)
    return t


def fig(fid: str) -> int:
    return FIGURES.index(fid) + 1


def caption(doc, fid: str, s: str):
    cap = doc.add_paragraph()
    cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cap.paragraph_format.space_after = Pt(12)
    set_font(cap.add_run(f"Рисунок {fig(fid)} — {s}"), BODY_FONT, 10, italic=True, color=MUTED)


MAX_FIG_H = 6.0  # дюймы: высокий скриншот не должен вылезать за страницу


def figure(doc, fid: str, path: Path, width_in: float, s: str):
    if not path.exists():
        para(doc, f"[изображение не найдено: {path.relative_to(ROOT)}]", italic=True, color=WARN)
    else:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(6)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.keep_with_next = True
        img = DocxImage.from_file(str(path))
        width = min(width_in, USABLE - 0.1, MAX_FIG_H * img.px_width / img.px_height)
        p.add_run().add_picture(str(path), width=Inches(width))
    caption(doc, fid, s)


def screenshot(doc, fid: str, s: str):
    """Скриншот запуска; пока файла нет — заметная заглушка с описанием."""
    filename, what = SCREENSHOTS[fid]
    path = SCREENS / filename
    if path.exists():
        if fid in REQUESTS:
            req = para(doc, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=8, space_after=0)
            req.paragraph_format.keep_with_next = True
            set_font(req.add_run(REQUESTS[fid]), MONO_FONT, 9.5, color=MUTED)
        # JSON-ответы сняты с retina-экрана: ~200 px на дюйм дают шрифт, близкий к основному тексту;
        # терминал и Swagger — на всю ширину
        width = 6.4 if fid in ("scr_evaluate", "scr_swagger") else DocxImage.from_file(str(path)).px_width / 200
        figure(doc, fid, path, width, s)
        return
    p = para(doc, f"[СКРИНШОТ НЕ ДОБАВЛЕН — {what}. Положите файл "
                  f"reports/figures/screens/{filename} и пересоберите отчёт]",
             size=10, italic=True, color=WARN, align=WD_ALIGN_PARAGRAPH.CENTER,
             space_before=6, space_after=2)
    p.paragraph_format.keep_with_next = True
    caption(doc, fid, s)


def screenshot_row(doc, items, height_in: float = 4.6):
    """Два скриншота рядом в таблице без рамок: [(fid, подпись), (fid, подпись)].

    Вертикальные JSON-ответы по одному занимают полстраницы, рядом — вдвое меньше.
    Если какого-то файла нет, ряд печатается заглушками по одному.
    """
    if not all((SCREENS / SCREENSHOTS[fid][0]).exists() for fid, _ in items):
        for fid, cap in items:
            screenshot(doc, fid, cap)
        return
    col_w = (USABLE - 0.2) / len(items)
    t = doc.add_table(rows=1, cols=len(items))
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    # строка не рвётся между страницами, иначе запросы отрываются от картинок
    tr_pr = t.rows[0]._tr.get_or_add_trPr()
    tr_pr.append(tr_pr.makeelement(qn("w:cantSplit"), {}))
    for col in t._tbl.tblGrid.findall(qn("w:gridCol")):
        col.set(qn("w:w"), str(int(col_w * 1440)))
    for cell, (fid, cap) in zip(t.rows[0].cells, items):
        cell.width = Inches(col_w)
        path = SCREENS / SCREENSHOTS[fid][0]
        img = DocxImage.from_file(str(path))
        width = min(col_w - 0.1, height_in * img.px_width / img.px_height)
        req = cell.paragraphs[0]
        req.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_font(req.add_run(REQUESTS[fid]), MONO_FONT, 8.5, color=MUTED)
        pic = cell.add_paragraph()
        pic.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pic.add_run().add_picture(str(path), width=Inches(width))
        c = cell.add_paragraph()
        c.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_font(c.add_run(f"Рисунок {fig(fid)} — {cap}"), BODY_FONT, 10, italic=True, color=MUTED)
    doc.add_paragraph().paragraph_format.space_after = Pt(4)


def code(doc, lines):
    for line in lines:
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.left_indent = Inches(0.3)
        p.paragraph_format.line_spacing = 1.0
        set_font(p.add_run(line), MONO_FONT, 9.5)
    doc.add_paragraph().paragraph_format.space_after = Pt(4)


# --------------------------------------------------------------------------- #
# документ
# --------------------------------------------------------------------------- #
doc = Document()
sec = doc.sections[0]
sec.page_width = Inches(PAGE_W)
sec.page_height = Inches(PAGE_H)
sec.left_margin = Inches(MARGIN_L)
sec.right_margin = Inches(MARGIN_R)
sec.top_margin = Inches(0.8)
sec.bottom_margin = Inches(0.8)

style = doc.styles["Normal"]
style.font.name = BODY_FONT
style.font.size = Pt(12)
style.element.rPr.rFonts.set(qn("w:eastAsia"), BODY_FONT)

footer_p = sec.footer.paragraphs[0]
footer_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
_fld = footer_p._p.makeelement(qn("w:fldSimple"), {qn("w:instr"): "PAGE"})
_run = footer_p._p.makeelement(qn("w:r"), {})
_txt = footer_p._p.makeelement(qn("w:t"), {})
_txt.text = "1"
_run.append(_txt)
_fld.append(_run)
footer_p._p.append(_fld)

# ------------------------------- шапка ------------------------------------- #
para(doc, "Рекомендательная система для маркетплейса", size=19, bold=True,
     align=WD_ALIGN_PARAGRAPH.CENTER, font=HEAD_FONT, space_after=4)
para(doc, "Отчёт по заданию PrEng. RecSys", size=13, align=WD_ALIGN_PARAGRAPH.CENTER,
     font=HEAD_FONT, color=ACCENT, space_after=4)
if AUTHOR:
    para(doc, AUTHOR, size=12, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=4)
para(doc, f"Репозиторий: {REPO_URL}", size=11, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=20,
     color=WARN if "<" in REPO_URL else None)

# --------------------------- 1. цель ---------------------------------------- #
h1(doc, "1. Цель и постановка задачи")
text(doc, "**Цель:** разобраться в подходах к построению рекомендательных систем (RecSys) и их ограничениях "
          "на примере маркетплейса: реализовать несколько классических подходов и бизнес-эвристик, выбрать "
          "датасет, сравнить модели, оценить уязвимости и проблемы (в том числе в продакшене), описать "
          "мониторинг и ожидаемый эффект для бизнеса.")
text(doc, "**Предметная область:** маркетплейс, рекомендации товаров. Выбрана потому, что на маркетплейсе "
          "естественно возникают все типовые блоки рекомендаций: «Вам может понравиться», «С этим товаром "
          "покупают», «Похожие товары», «Другие товары бренда», «Популярное», «Недавно покупали».")
text(doc, "Соответствие пунктов задания разделам отчёта:")
table(doc, ["Пункт задания", "Раздел отчёта"], [
    ["I. Классические подходы (content-based, collaborative user/item-based, hybrid)", "5"],
    ["I. Бизнес-эвристики (популярное, «с этим покупают», бренд, недавние)", "5"],
    ["Развернуть существующее решение, выбрать датасет", "2, 3"],
    ["II-a. Проблемы и ограничения подходов, уязвимости, этика", "7"],
    ["II-b. Проблемы в продакшене", "8"],
    ["III. Оценка в продакшене: мониторинг, ML- и бизнес-метрики", "9"],
    ["Ожидания бизнеса против реальных результатов", "9.3"],
    ["Как улучшить модель", "10"],
], widths=[5.4, 1.2])

# --------------------------- 2. стек ---------------------------------------- #
h1(doc, "2. Выбор стека и аргументы")
text(doc, "Вся работа выполнена на открытом (open source) стеке, локально на ноутбуке MacBook Pro "
          "(Apple M4 Pro, 24 ГБ RAM), без GPU и облачных сервисов.")
table(doc, ["Компонент", "Выбор", "Почему"], [
    ["Язык", "Python 3.12", "Стандарт для ML, все нужные библиотеки"],
    ["Данные", "pandas, pyarrow (parquet)", "Быстрая загрузка, parquet-кэш ускоряет повторные запуски"],
    ["Разреженные матрицы", "scipy.sparse", f"Матрица user × item заполнена на {pc(S['density'], 3)}, плотная не нужна"],
    ["Content-based", "scikit-learn (TF-IDF)", "Простой, интерпретируемый, работает на CPU за секунды"],
    ["Collaborative", "implicit (ALS, ItemKNN)", "Готовая промышленная библиотека для неявного фидбека, сборки под "
                                                 "Apple Silicon. Это и есть «развёрнутое существующее решение»"],
    ["Сервис", "FastAPI + uvicorn", "Минимальный REST API с автодокументацией (Swagger) для демонстрации блоков витрины"],
    ["Графики и отчёт", "matplotlib, python-docx", "Графики и этот отчёт собираются скриптами из сохранённых метрик"],
    ["Датасет", "Amazon Reviews 2023 (HuggingFace)", "Открытый доступ без регистрации, реальный маркетплейс, "
                                                     "есть метаданные товаров"],
], widths=[1.5, 1.9, 3.4])
text(doc, "Что рассматривалось и не вошло:")
bullets(doc, [
    "**LightFM** (гибрид с признаками) — у библиотеки известны проблемы сборки на новых версиях Python "
    "и Apple Silicon. Гибрид реализован своими силами.",
    "**RecTools** — удобная обёртка, но для MVP достаточно implicit и scikit-learn, лишний уровень "
    "зависимостей не нужен.",
    "**Нейросетевые модели** (SASRec, two-tower) — требуют GPU и сильно больше времени на настройку. "
    "Для учебного MVP избыточны, упомянуты в разделе 10.",
    "**Prometheus/Grafana** — задание требует описать мониторинг (раздел 9), а не разворачивать его.",
])
text(doc, "**Архитектурный стиль:** модульный пайплайн на Python. Данные → подготовка → общий контекст "
          "(индексы и матрица) → модели с единым интерфейсом (fit, score, similar) → офлайн-оценка или REST API. "
          "Единый интерфейс позволяет одинаково оценивать любую модель и собирать из них гибрид без изменения "
          "остального кода.")

# --------------------------- 3. данные -------------------------------------- #
h1(doc, "3. Данные")
text(doc, "**Датасет:** Amazon Reviews 2023 (McAuley Lab, UCSD), категория Video Games, вариант 5-core "
          "(каждый пользователь и товар имеет не менее 5 взаимодействий). Использованы файлы Video_Games.csv "
          "(47 МБ: пользователь, товар, оценка, время) и meta_Video_Games.jsonl (437 МБ: название, бренд, "
          "категории, описание, цена).")
text(doc, "**Допущение:** отзыв = факт покупки (implicit feedback). Оценка не используется как сила сигнала.")
table(doc, ["Показатель", "Значение", "Что это значит для RecSys"], [
    ["Покупки / пользователи / товары",
     f"{thousands(S['interactions'])} / {thousands(S['users'])} / {thousands(S['items'])}",
     "Небольшой, но реалистичный каталог"],
    ["Период", f"{S['period'][0]}–{S['period'][1]}", "Длинная история, сильная смена ассортимента"],
    ["Плотность матрицы user × item", pc(S["density"], 3), "Сильная разреженность"],
    ["Покупок на пользователя (медиана)", f"{S['median_per_user']:.0f}",
     f"У {pc(S['users_le5'])} пользователей не больше 5 покупок"],
    ["Доля покупок топ-1% / топ-20% товаров", f"{pc(S['top1_share'])} / {pc(S['top20_share'])}",
     "Длинный хвост, риск popularity bias"],
    ["Товаров с ≤ 10 покупками", pc(S["items_le10"]), "Для половины каталога CF почти не работает"],
    ["Покупки в тесте товаров, которых нет в train", NEW_ITEMS, "Холодный старт товаров (новые релизы)"],
    ["Новые пользователи среди тестовых", COLD_SHARE, "Холодный старт пользователей"],
    ["Оценок 1–2 звезды", pc(S["low_rating"]), "«Покупка» не значит «понравилось»"],
    ["Пик покупок", f"декабрь–январь ({pc(S['dec_jan'])} за 2 месяца)", "Сезонность (праздники)"],
    ["Товаров без описания", f"{pc(S['no_description'])} (без описания и характеристик — "
                             f"{pc(S['no_description_features'])})", "Ограничение content-based"],
], widths=[2.4, 1.9, 2.5])
figure(doc, "data", IMG / "data.png", 6.4, "Длинный хвост и сезонность покупок")
text(doc, f"**Схема валидации:** глобальное разбиение по времени. Последние 5% покупок (после {S['test_start']}) — "
          f"тест, остальное — обучение. Модель не видит будущего, и в тест естественно попадают новые "
          f"пользователи. Релевантными считаются товары, купленные в тестовом периоде, кроме уже купленных ранее. "
          f"Метрики считаются для K = 10.")
text(doc, f"Тестовое окно получилось длинным — {S['test_months']} месяцев ({S['test_start']}–{S['test_end']}), "
          f"и модель за это время не переобучается. Это занижает абсолютные метрики и завышает долю новинок "
          f"в тесте (см. 7.2). Веса гибрида подбираются на отдельной валидационной выборке: из train отделяются "
          f"последние 5% покупок (после {pd.Timestamp(WEIGHTS['val_cutoff']):%d.%m.%Y}), модели обучаются на "
          f"оставшейся части, и перебором по сетке выбирается комбинация весов с лучшим NDCG@10. Тестовая "
          f"выборка в подборе не участвует.")

# --------------------------- 4. архитектура --------------------------------- #
h1(doc, "4. Архитектура и ход работы")
figure(doc, "architecture", IMG / "architecture.png", 6.4, "Архитектура решения")
text(doc, "Структура репозитория:")
table(doc, ["Файл", "Назначение"], [
    ["recsys/data.py", "Загрузка, очистка, фильтр метаданных, parquet-кэш, разбиение по времени"],
    ["recsys/models.py", "Все модели с единым интерфейсом, общий контекст, функция top-K без уже купленного"],
    ["recsys/metrics.py", "Precision, Recall, NDCG, MAP, HitRate@K, Coverage, Novelty, TrendShare"],
    ["recsys/evaluate.py", "Подбор весов гибрида на валидации, офлайн-сравнение всех моделей на тесте"],
    ["recsys/api.py", "REST API (FastAPI): по эндпоинту на каждый блок витрины"],
    ["recsys/figures.py", "Графики для отчёта (reports/figures)"],
    ["scripts/download_data.sh", "Скачивание датасета в data/raw"],
    ["reports/metrics.csv, data_stats.json, hybrid_weights.json",
     "Результаты оценки; веса гибрида читает API, все числа отчёта берутся отсюда"],
    ["reports/make_docx.py", "Сборка этого отчёта из сохранённых метрик"],
], widths=[2.6, 4.2])
text(doc, "**Ход работы:**")
numbered(doc, [
    "Разбор задания, выбор предметной области (маркетплейс) и плана работ.",
    "Выбор датасета: сравнение H&M (Kaggle), Retailrocket, Olist и Amazon Reviews 2023. "
    "Выбран Amazon Video Games (раздел 6, проблема 1).",
    "Подготовка окружения: виртуальное окружение Python 3.12, установка библиотек.",
    "Загрузка данных, подготовка, разбиение по времени, расчёт статистики (раздел 3).",
    "Реализация бизнес-эвристик: Popular, Trending, AlsoBought, SameBrand, история покупок.",
    "Реализация content-based (TF-IDF) и коллаборативных моделей (UserKNN, ItemKNN, ALS).",
    "Реализация гибрида: switching для новых пользователей + смешивание рангов (RRF), подбор весов на валидации.",
    "Офлайн-сравнение всех моделей и анализ результатов (раздел 5).",
    "REST API для демонстрации блоков витрины, проверка всех эндпоинтов.",
    "Ревью готовой работы и исправление найденных ошибок (раздел 6, проблемы 10–12).",
    "Теоретический разбор проблем подходов и продакшена, метрики мониторинга (разделы 7–9).",
])
text(doc, "**Запуск:**")
code(doc, [
    "python3.12 -m venv .venv && .venv/bin/pip install -r requirements.txt",
    "./scripts/download_data.sh",
    ".venv/bin/python -m recsys.evaluate",
    ".venv/bin/python -m recsys.figures",
    ".venv/bin/uvicorn recsys.api:app --port 8000",
    ".venv/bin/python reports/make_docx.py --pdf",
])
screenshot(doc, "scr_evaluate", "Вывод офлайн-оценки: подбор весов гибрида на валидации и сравнение моделей на тесте")

# --------------------------- 5. часть I ------------------------------------- #
h1(doc, "5. Часть I: реализованные подходы и результаты")
h2(doc, "5.1 Подходы")
table(doc, ["Группа", "Подход", "Как работает", "Блок на витрине"], [
    ["Эвристика", "Popular", "Самые продаваемые товары за всё время", "«Популярное», fallback"],
    ["Эвристика", "Trending", "Популярность с затуханием по времени (период полураспада 30 дней)",
     "«Популярное сейчас», fallback для новых пользователей"],
    ["Эвристика", "AlsoBought", "Пары товаров, купленные одним человеком в пределах 30 дней. Делим на корень "
                                "популярности, чтобы хиты не попадали к каждому товару", "«С этим товаром покупают»"],
    ["Эвристика", "SameBrand", "Популярные товары брендов, которые пользователь уже покупал", "«Другие товары бренда»"],
    ["Эвристика", "История", "Последние купленные товары", "«Недавно покупали»"],
    ["Content-based", "TF-IDF", "Векторы по названию, бренду, категориям, описанию. Профиль пользователя — "
                                "взвешенная сумма векторов последних покупок", "«Похожие товары»"],
    ["Collaborative", "UserKNN", "50 самых похожих покупателей (косинусная мера), рекомендуем их покупки", "—"],
    ["Collaborative", "ItemKNN", "Похожие товары по векторам покупателей (implicit)", "—"],
    ["Collaborative", "ALS", "Матричная факторизация для неявного фидбека, 64 фактора (implicit)", "—"],
    ["Hybrid", "Switching + RRF", "Нет истории → Trending. Есть история → смешивание рангов ALS, ItemKNN, "
                                  "Content и Trending через Reciprocal Rank Fusion", "«Вам может понравиться»"],
], widths=[1.3, 1.3, 2.7, 1.6])
text(doc, f"Для смешивания выбран Reciprocal Rank Fusion: каждая модель голосует позицией товара в своей выдаче "
          f"(вес / (60 + позиция)). Скоры разных моделей несравнимы между собой (косинус, скалярное произведение "
          f"факторов, счётчики), а позиции сравнимы. Уже купленные товары исключаются до отбора кандидатов каждой "
          f"модели. Веса, подобранные на валидации: {weights_str(WH)} (для варианта без Trending — "
          f"{weights_str(WNT)}).")

h2(doc, "5.2 Результаты офлайн-оценки")
text(doc, f"Пользователи с историей ({num(S['warm_users'])} человек), K = 10:")
order = ["Popular", "Trending", "AlsoBought", "SameBrand", "ContentBased", "ItemKNN", "UserKNN", "ALS",
         "Hybrid без Trending", "Hybrid"]
table(doc, ["Модель", "Recall@10", "NDCG@10", "HitRate@10", "Coverage", "Novelty", "TrendShare", "мс"],
      [[m, n(W.loc[m, "Recall@10"]), n(W.loc[m, "NDCG@10"]), n(W.loc[m, "HitRate@10"]), cov(W.loc[m, "Coverage"]),
        n(W.loc[m, "Novelty"], 1), pc(W.loc[m, "TrendShare"]), n(W.loc[m, "ms/user"], 2)] for m in order],
      widths=[1.28, 0.88, 0.84, 0.96, 0.78, 0.66, 0.98, 0.4], size=9, bold_rows=(len(order) - 1,))
text(doc, "мс — время построения рекомендаций на одного пользователя.")
text(doc, f"Новые пользователи ({num(S['cold_users'])} человека, истории нет, работают только неперсональные модели): "
          f"Popular NDCG@10 = {n(C.loc['Popular', 'NDCG@10'])}, Trending = {n(C.loc['Trending', 'NDCG@10'])}.")
text(doc, "**Coverage** — доля каталога, которая хотя бы раз попала в рекомендации. **Novelty** — средняя "
          "«неожиданность» товара, −log2 доли купивших (чем выше, тем менее заезжен товар). **TrendShare** — доля "
          "выдачи из топ-30 Trending: насколько «персональная» выдача совпадает с общей для всех.")
figure(doc, "metrics", IMG / "metrics.png", 6.4, "Точность, покрытие каталога и доля Trending в выдаче по моделям")
text(doc, "**Выводы:**")
numbered(doc, [
    "Простая эвристика Trending обходит каждую ML-модель по отдельности: в играх спрос сосредоточен вокруг "
    "свежих релизов. Любую модель надо сравнивать с сильным бейзлайном, а не с нулём.",
    f"Среди ML-моделей точнее всех {BEST_ML}. UserKNN точнее ItemKNN, ItemKNN зато покрывает больше каталога.",
    f"Гибрид даёт +{pc(GAIN)} NDCG к Trending, но покрытие падает с {cov(HYB_NT['Coverage'])} (гибрид без Trending) "
    f"до {cov(HYB['Coverage'])}: точность куплена ценой разнообразия.",
    f"Гибрид персонализирован слабо: в среднем {TREND_SLOTS} из 10 его рекомендаций — товары из топ-30 Trending "
    f"(TrendShare {pc(HYB['TrendShare'])}). По сути это Trending, в который ML-модели подмешивают пару "
    f"персональных позиций, и прирост NDCG даёт в основном «свежесть», а не персонализация (см. 7.4).",
    f"Абсолютные значения низкие (HitRate {pc(HYB['HitRate@10'])}). Для каталога в {thousands(S['items'])} товаров "
    f"при плотности {pc(S['density'], 2)} это нормально, важно относительное сравнение.",
])

h2(doc, "5.3 Демонстрация сервиса")
text(doc, "Сервис поднимается командой uvicorn recsys.api:app и при старте обучает модели на всех данных "
          "(около минуты). Каждому блоку витрины соответствует свой эндпоинт.")
screenshot(doc, "scr_swagger", "Эндпоинты REST API (Swagger UI)")
text(doc, f"Для демонстрации персональной выдачи взят пользователь {DEMO_USER}. В его истории — старые игры "
          f"для Wii и Xbox 360 (рисунок {fig('scr_history')}), а гибрид рекомендует ему геймпад PS5, Zelda "
          f"и аксессуары для Switch: 9 из 10 позиций совпадают с Trending (рисунок {fig('scr_recommend')}). "
          f"Это наглядная иллюстрация popularity bias из раздела 7.4.")
screenshot(doc, "scr_history", "Блок «Недавно покупали»: история демонстрационного пользователя")
screenshot(doc, "scr_recommend", "Ответ /recommend для того же пользователя: source = hybrid, "
                                 "но выдача почти совпадает с Trending")
text(doc, f"Блоки на карточке товара показаны на примере Tomb Raider ({DEMO_ITEM}). «Похожие товары» "
          f"(content-based) возвращают другие издания той же игры — это замены и пример переспециализации "
          f"из раздела 7.5. «С этим покупают» (AlsoBought) возвращает другие шутеры того же периода — "
          f"BioShock Infinite, Crysis 3, Far Cry 3, то есть дополнения к покупке.")
screenshot_row(doc, [("scr_similar", "Блок «Похожие товары» (content-based) для Tomb Raider"),
                    ("scr_also_bought", "Блок «С этим товаром покупают» (AlsoBought) для Tomb Raider")])
text(doc, f"Для нового пользователя /recommend возвращает \"source\": \"trending_fallback\" "
          f"(рисунок {fig('scr_cold')}). Так реализован switching-гибрид, решающий холодный старт.")
text(doc, "Параметр k проверяется (от 1 до 100): неверное значение возвращает понятную ошибку 422, а не 500 "
          f"(рисунок {fig('scr_422')}).")
screenshot_row(doc, [("scr_cold", "Новый пользователь получает Trending (switching-гибрид, холодный старт)"),
                    ("scr_422", "Ошибка валидации параметра k")])

# --------------------------- 6. проблемы в ходе работы ---------------------- #
h1(doc, "6. Проблемы в ходе работы и их решение")
table(doc, ["#", "Проблема", "Решение"], [
    ["1", "Изначально выбранный датасет H&M (Kaggle) требует регистрации и принятия правил соревнования, "
          "API-ключа Kaggle не было",
     "Выбран открытый Amazon Reviews 2023: тоже маркетплейс, есть метаданные товаров, скачивается без регистрации"],
    ["2", "Основной сервер датасета (mcauleylab.ucsd.edu) не отвечал", "Использовано зеркало на HuggingFace"],
    ["3", "Системный Python 3.14 слишком новый для ML-библиотек", "Отдельное виртуальное окружение на Python 3.12"],
    ["4", "Файл метаданных 437 МБ, большая часть товаров не нужна",
     "Построчное чтение с фильтром по нужным товарам, результат кэшируется в parquet. Повторный запуск — секунды"],
    ["5", "Риск нечестной оценки при случайном разбиении (утечка будущего)", "Глобальное разбиение по времени"],
    ["6", "Неперсональная эвристика Trending оказалась лучше всех ML-моделей",
     f"Trending добавлен в гибрид компонентом «свежести». Итог: +{pc(GAIN)} NDCG к Trending"],
    ["7", "Скоры моделей несравнимы, простое сложение не работает",
     "Смешивание по позициям (Reciprocal Rank Fusion) вместо скоров"],
    ["8", f"Гибрид в основном повторяет Trending: {pc(HYB['TrendShare'])} выдачи из топ-30 Trending, "
          f"покрытие {cov(HYB['Coverage'])}",
     "Добавлена метрика TrendShare, проблема зафиксирована как popularity bias, предложены решения (разделы 7.4 и 10)"],
    ["9", f"{NEW_ITEMS} тестовых покупок — товары, которых не было в обучении: CF их не видит, а офлайн-оценка "
          f"их отбрасывает", "Зафиксировано как ограничение, предложены решения (разделы 7.2 и 10)"],
    ["10", "В первой версии веса гибрида подбирались на тестовой выборке, результат был завышен",
     "Подбор весов вынесен на валидацию (последние 5% train). Веса сохраняются в reports/hybrid_weights.json, "
     "API берёт их оттуда"],
    ["11", "Уже купленные товары занимали места среди кандидатов компонент гибрида (у ALS ~5 из 200)",
     "Фильтр купленного перенесён до отбора кандидатов"],
    ["12", "API падал с ошибкой 500 при k=100000 и возвращал почти весь каталог при отрицательном k",
     "Проверка параметра k (1–100), неверное значение → 422"],
], widths=[0.35, 3.1, 3.3])
text(doc, "**Ограничения работы:**")
bullets(doc, [
    f"Тестовое окно — {S['test_months']} месяцев без переобучения модели: абсолютные метрики занижены, доля новинок "
    f"завышена. Ближе к проду — скользящая валидация с окном в 1–2 недели и регулярным переобучением.",
    "Каталог ограничен товарами, у которых есть покупки. Для нового товара без продаж «Похожие товары» в MVP "
    "не строятся, хотя content-based это позволяет: достаточно векторизовать его карточку.",
    "Отзыв используется как замена покупки. Просмотров, корзины и возвратов в данных нет.",
    "Эффект на бизнес не измерен: для этого нужен A/B-тест на реальном трафике (раздел 9).",
])

# --------------------------- 7. часть II-a ---------------------------------- #
h1(doc, "7. Часть II-a: проблемы и ограничения подходов")
problem(doc, "7.1 Холодный старт пользователя (все персональные модели)",
        f"{COLD_SHARE} тестовых покупателей новые. CF и content-based им ничего не дают.",
        "switching на Trending (реализовано), онбординг-опрос (любимые платформы и жанры), контекст (страница "
        "входа, рекламная кампания, гео, устройство), быстрая адаптация по первым кликам в сессии.",
        "Trending не персонален. Опрос добавляет шаг и увеличивает отток на старте.",
        "best practice — не заставлять, а предлагать: опрос в 1–2 клика с кнопкой «пропустить», выдача по "
        "популярности с учётом контекста и сессионные модели, которые персонализируют уже после 2–3 действий.")
problem(doc, "7.2 Холодный старт товара (CF: ItemKNN, UserKNN, ALS, AlsoBought)",
        f"{NEW_ITEMS} покупок тестового периода — товары, которых не было в обучении (новые релизы). CF их не видит "
        f"вообще. Офлайн-оценка такие покупки отфильтровывает и тем самым завышает реальное качество.",
        "content-based по атрибутам (реализован: похожесть считается по описанию, история покупок товара не нужна; "
        "в MVP каталог ограничен товарами с покупками, но для нового товара достаточно векторизовать его карточку), "
        "гибридные модели с признаками товаров (LightFM, two-tower), квота показов для новинок (exploration, "
        "многорукие бандиты).",
        f"у {pc(S['no_description'])} товаров нет описания (нет ни описания, ни характеристик — у "
        f"{pc(S['no_description_features'])}), для них остаются название и бренд, и content-based слабее. Квота "
        f"новинок отнимает показы у проверенных хитов, а значит и краткосрочную выручку.",
        "стандарт — двухстадийная схема: кандидаты из нескольких источников, включая content и новинки, затем "
        "ранжирование. Отдельно — контроль качества карточек товара у продавцов (без описания хуже ранжируется).")
problem(doc, "7.3 Разреженность данных (CF)",
        f"плотность {pc(S['density'], 3)}, у {pc(S['users_le5'])} пользователей ≤ 5 покупок, у "
        f"{pc(S['items_le10'])} товаров ≤ 10 покупок. Оценки близости шумные.",
        "добавить более частые сигналы (просмотры, добавления в корзину, избранное) с разными весами; "
        "факторизация (ALS) вместо прямых kNN-соседей.",
        "неявные сигналы шумные: просмотр ≠ интерес, случайный клик ≠ намерение купить.",
        "веса по типу события (покупка > корзина > просмотр), учёт времени просмотра, фильтр ботов.")
problem(doc, "7.4 Popularity bias и длинный хвост (Popular, Trending, ALS, Hybrid)",
        f"{pc(S['top20_share'])} покупок приходится на 20% товаров. Модели воспроизводят и усиливают это: Trending "
        f"покрывает {cov(TREND['Coverage'])} каталога, гибрид — {cov(HYB['Coverage'])}, и "
        f"{pc(HYB['TrendShare'])} его выдачи — товары из топ-30 Trending, то есть большинству пользователей он "
        f"отдаёт почти «то же, что всем». Пример из API (рисунки {fig('scr_history')}–{fig('scr_recommend')}): "
        f"пользователю, который покупал игры для Wii и Xbox 360, гибрид рекомендует геймпад PS5, Zelda и стекло для "
        f"Switch. Для маркетплейса это ещё и проблема продавцов: малые продавцы не получают показов.",
        "штраф за популярность, переранжирование на разнообразие (MMR, квоты по категориям и продавцам), фильтры "
        "по контексту (платформа пользователя).",
        f"офлайн-точность падает: наш гибрид без Trending разнообразнее (coverage {cov(HYB_NT['Coverage'])}, "
        f"TrendShare {pc(HYB_NT['TrendShare'])}), но NDCG на {pc(DROP_NT)} ниже.",
        "баланс задаётся бизнес-целью. Обычная практика — отслеживать coverage, novelty и долю длинного хвоста "
        "рядом с точностью и решать A/B-тестом. Офлайн-метрики систематически переоценивают популярное.")
problem(doc, "7.5 Пузырь фильтров и переспециализация (Content-based, SameBrand)",
        f"купил Tomb Raider — получаешь пять изданий Tomb Raider (рисунок {fig('scr_similar')}). "
        f"Мало неожиданных находок.",
        "подмешивать рекомендации CF или исследовательские слоты, дедуплицировать издания одного тайтла.",
        "часть исследовательских рекомендаций будет нерелевантна.",
        "разделять блоки по назначению («Похожие» и «Вам может понравиться») и ограничивать долю одного бренда "
        "или серии в выдаче.")
problem(doc, "7.6 Взаимозаменяемые и уже купленные товары (все модели)",
        "купил консоль — система предлагает ещё консоли (то же для геймпадов, карт памяти). Повторные покупки "
        "товаров длительного пользования маловероятны, а повторные покупки расходников нормальны.",
        "фильтр уже купленного (реализован), правила «после покупки X предлагать комплементарные товары» "
        "(AlsoBought частично это делает), разметка категорий: длительного пользования или расходник.",
        "нужна ручная или полуавтоматическая разметка, правила устаревают.",
        f"различать «замены» (похожие) и «дополнения» (с этим покупают), использовать каждый тип на своей "
        f"странице. Разница хорошо видна на рисунках {fig('scr_similar')} и {fig('scr_also_bought')}.")
problem(doc, "7.7 Неявный фидбек без негатива (все модели)",
        f"{pc(S['low_rating'])} «покупок» оценены на 1–2 звезды, но модель считает их положительным сигналом. "
        f"Отсутствие покупки тоже не означает «не нравится».",
        "веса по рейтингу или возвратам, исключение возвратов, отдельный учёт «скрыть рекомендацию».",
        "явных оценок мало, они смещены (пишут в основном очень довольные или очень недовольные).",
        "использовать возвраты и кнопку «не интересно» как явный негатив, остальное оставлять implicit.")
problem(doc, "7.8 Сезонность и дрейф (все модели, особенно ALS и kNN)",
        "пик в декабре–январе, спрос смещается к новым релизам. Модель, обученная месяц назад, устаревает.",
        "затухание по времени (Trending), регулярное переобучение, признаки сезона в ранжировщике.",
        "частое переобучение стоит денег и рискует выкатить сломанную модель.",
        "переобучать по расписанию с проверкой качества перед выкаткой, свежие сигналы подмешивать онлайн.")
problem(doc, "7.9 Масштабируемость (UserKNN, Content-based)",
        f"UserKNN на каждый запрос сравнивает пользователя со всеми остальными (~{n(W.loc['UserKNN', 'ms/user'], 1)} мс "
        f"на пользователя на {thousands(S['users'])} пользователей, а на миллионах это секунды). Content-based "
        f"считает близость ко всему каталогу.",
        "ItemKNN или ALS (матрица похожих товаров считается заранее), приближённый поиск соседей (FAISS, HNSW), "
        "предрасчёт.",
        "приближённый поиск теряет часть точности, а предрасчёт устаревает.",
        "user-based в индустрии почти не используют в чистом виде. Стандарт — item-to-item (Amazon, 2003) "
        "и эмбеддинги + ANN.")
problem(doc, "7.10 Этические проблемы (AlsoBought, «люди также покупают»)",
        "без фильтрации «с этим покупают» и персональная выдача показывают неуместные связи. Примеры: Amazon "
        "в 2017 году предлагал вместе покупать компоненты для взрывчатки (расследование Channel 4). Target по "
        "покупкам вычислил беременность подростка раньше её семьи (NYT, 2012). В нашем каталоге — игры с "
        "рейтингом Mature, которые могут попасть детскому аккаунту. Также лутбоксы и внутриигровая валюта "
        "(близко к азартным играм) и продвижение собственных товаров площадки в ущерб продавцам.",
        "список чувствительных категорий и запрещённых связок, фильтр по возрастному рейтингу и возрасту аккаунта, "
        "запрет выводов о чувствительных характеристиках (здоровье, беременность, ориентация), аудит выдачи, "
        "прозрачность («почему мне это показали»).",
        "ручные списки всегда неполны, фильтры режут и легитимные рекомендации, аудит дорогой.",
        "правила + выборочная ручная проверка + жалобы пользователей как сигнал. Для маркетплейса это ещё "
        "и регуляторика: защита персональных данных (152-ФЗ / GDPR), антимонопольные требования "
        "к self-preferencing, DSA в ЕС.")
problem(doc, "7.11 Уязвимости: манипуляции и утечки (CF, AlsoBought)",
        "(1) **Shilling / data poisoning.** Наш датасет — отзывы, а накрутка отзывов — известная проблема "
        "маркетплейсов. Продавец с сотней фейковых аккаунтов, покупающих его товар вместе с хитом, попадает "
        "в «с этим покупают» хита. AlsoBought требует всего 2 совместные покупки (min_support=2), такая атака "
        "дешёвая. (2) **Утечка приватности.** По изменениям публичных списков «с этим покупают» можно "
        "восстановить покупки конкретного человека (Calandrino et al., «You Might Also Like», 2011).",
        "учитывать только верифицированные покупки, антифрод (новые аккаунты, аномальные всплески), ограничить "
        "вклад одного пользователя, повысить порог поддержки, добавить шум или задержку в публичные блоки.",
        "антифрод даёт ложные срабатывания, высокий порог ухудшает холодный старт, шум снижает точность.",
        "вклад каждого аккаунта взвешивать по «доверию», мониторить резкие изменения в топах, обновлять "
        "публичные блоки батчами.")

# --------------------------- 8. часть II-b ---------------------------------- #
h1(doc, "8. Часть II-b: проблемы в продакшене")
table(doc, ["Проблема", "Решение", "Проблема решения", "Обсуждение / best practice"], [
    [f"**Медленный ответ.** Гибрид считает 4 модели по всему каталогу (~{n(HYB['ms/user'], 1)} мс/пользователь "
     f"здесь, на миллионах товаров — сотни мс)",
     "Предрасчёт рекомендаций офлайн → кэш (Redis), ANN-индекс для эмбеддингов",
     "Выдача не учитывает действия в текущей сессии",
     "Гибрид: офлайн-кандидаты + лёгкое онлайн-ранжирование. Бюджет латентности p99 < 100 мс"],
    ["**Падение сервиса рекомендаций**",
     "Fallback на Trending (реализовано для новых пользователей), таймауты, circuit breaker",
     "Fallback неперсонален, просадка конверсии незаметна без мониторинга",
     "Пустой блок хуже популярного. Отслеживать долю fallback-ответов"],
    ["**Устаревание модели**",
     f"Переобучение по расписанию (ALS у нас {ALS['fit_s']:.0f} с, на реальном каталоге — часы)",
     "Стоимость, риск выкатить сломанную модель",
     "Проверка метрик перед выкаткой, shadow/canary-релиз, быстрый откат"],
    ["**Рекомендация отсутствующих товаров**", "Постфильтр по остаткам и статусу карточки в реальном времени",
     "Если фильтр вычеркнул много, выдача короткая", "Брать с запасом (top-50 → фильтр → top-10)"],
    ["**Расхождение признаков при обучении и в проде** (training/serving skew)",
     "Общий код признаков, feature store", "Сложность инфраструктуры",
     "Логировать признаки, с которыми был сделан прогноз, и обучаться на них"],
    ["**Петля обратной связи и position bias.** Модель учится на кликах по собственным рекомендациям",
     "Логировать позицию показа, 1–5% трафика со случайной/exploration-выдачей",
     "Exploration-трафик зарабатывает меньше",
     "Без него модель всё сильнее сужает каталог (см. popularity bias)"],
    ["**Расхождение офлайн- и онлайн-метрик**", "A/B-тест каждого изменения", "Нужен трафик и время (2–4 недели)",
     "Офлайн — фильтр для кандидатов в A/B, решение — только по онлайну"],
    ["**Качество данных** (дубли, боты, сломанный ETL)", "Валидация данных на входе пайплайна, фильтр ботов",
     "Ложные срабатывания проверок", "Алерт при резком изменении объёма или распределения событий"],
    ["**Масштабирование** (миллионы пользователей и товаров)",
     "Item-based / эмбеддинги вместо user-based, шардирование, батч-инференс", "Стоимость инфраструктуры",
     "Считать стоимость рекомендации против прироста выручки"],
], widths=[1.8, 1.8, 1.5, 1.8], size=9.5)

# --------------------------- 9. часть III ----------------------------------- #
h1(doc, "9. Часть III: оценка работы в продакшене")
h2(doc, "9.1 Уровни метрик")
table(doc, ["Уровень", "Метрики", "Когда и как"], [
    ["**Офлайн ML**", "Precision/Recall/NDCG/MAP/HitRate@K, Coverage, Novelty, TrendShare, Diversity",
     "Перед выкаткой, на сплите по времени. Порог: не хуже текущей модели"],
    ["**Онлайн-поведение**", "CTR блока, add-to-cart rate, конверсия в покупку из блока",
     "Ежедневно, в разрезе блоков и сегментов (новые/старые)"],
    ["**Бизнес**", "Выручка (GMV) на пользователя, средний чек, число товаров в заказе, retention, доля продаж "
                   "длинного хвоста и малых продавцов", "A/B-тест, сравнение с контрольной группой (Trending)"],
    ["**Guardrail**", "Возвраты, жалобы, «скрыть рекомендацию», скорость загрузки страницы, нарушения политики "
                      "(18+ детям)", "Тест останавливают, если ухудшились, даже при росте выручки"],
    ["**Системные**", "Латентность p50/p95/p99, RPS, ошибки, **доля fallback-ответов**, попадание в кэш",
     "Онлайн, алерты"],
    ["**Здоровье модели и данных**", "Возраст модели, свежесть данных, объём событий, дрейф распределений, "
                                     "coverage и доля топ-100 в выдаче", "Ежедневно, алерт на скачки"],
], widths=[1.6, 3.0, 2.2])
h2(doc, "9.2 Как измерять эффект")
bullets(doc, [
    "**Только A/B-тест** против бейзлайна. Контроль — Trending, а не «без рекомендаций»: иначе мы измеряем "
    "ценность блока, а не модели.",
    "**Главная метрика — инкрементальная выручка на пользователя**, а не «доля выручки через рекомендации». "
    "Многие клики по рекомендациям — покупки, которые произошли бы и так (каннибализация поиска и каталога).",
    "Длительность от 2 недель (недельная сезонность), учёт эффекта новизны.",
    "Отдельный разрез по новым пользователям: холодный старт улучшается или нет.",
])
h2(doc, "9.3 Что бизнес ожидает и что получит на самом деле")
table(doc, ["Ожидание", "Реальность"], [
    ["«+35% выручки, как у Amazon»", "Цифра 35% — давняя оценка McKinsey, не результат A/B. В исследованиях "
                                     "(Jannach & Jugovac, 2019) реальный инкрементальный эффект обычно в единицах "
                                     "процентов"],
    ["«ML сразу лучше текущей выдачи»", f"В нашем эксперименте каждая ML-модель по отдельности хуже Trending. "
                                        f"Выигрыш дал только гибрид с эвристикой (+{pc(GAIN)} NDCG офлайн), и "
                                        f"{pc(HYB['TrendShare'])} его выдачи — те же хиты"],
    ["«Персонализация для всех»", f"{COLD_SHARE} тестовых пользователей без истории, {NEW_ITEMS} покупок — новинки, "
                                  f"которых модель не видела. Для них работает content-based и популярное"],
    [f"«Офлайн +{pc(GAIN)} = онлайн +{pc(GAIN)}»", "Офлайн-прирост обычно не переносится один к одному, бывает "
                                                   "и обратный знак. Нужен A/B"],
    ["«Внедрили и забыли»", "Нужны переобучение, мониторинг, фильтры, разбор жалоб: постоянная стоимость команды "
                            "и инфраструктуры"],
    ["«Рост продаж хитов»", "Рекомендации усиливают хиты. Если цель — продажи длинного хвоста и малых продавцов, "
                            "это надо явно закладывать в выдачу"],
], widths=[2.2, 4.6])
text(doc, "Реалистичный результат для MVP: заметный прирост CTR и конверсии блоков по сравнению со случайным или "
          "просто популярным, единицы процентов к выручке на пользователя. Основная ценность — в блоках «с этим "
          "покупают» и «похожие» на карточке товара, где они увеличивают чек.")

# --------------------------- 10. улучшения ---------------------------------- #
h1(doc, "10. Как улучшить модель")
numbered(doc, [
    f"**Валидация ближе к проду:** скользящее окно (обучение → неделя теста → дообучение) вместо одного теста "
    f"длиной {S['test_months']} месяцев; та же процедура — для проверки модели перед выкаткой.",
    "**Новинки:** включить в оценку покупки новых товаров. Добавить в гибрид content-based кандидатов среди "
    "свежих товаров.",
    "**Двухстадийная схема:** кандидаты (ALS, ItemKNN, content, AlsoBought, Trending) → ранжировщик на "
    "градиентном бустинге (CatBoost) с признаками пользователя, товара, цены, платформы и сезона.",
    "**Контекстная популярность:** Trending по платформе пользователя (PS / Xbox / Switch / PC) вместо "
    "глобального. Ожидаемо снижает проблему «Zelda владельцу Xbox 360».",
    "**Разнообразие:** MMR-переранжирование, лимит на бренд или серию в выдаче, дедупликация изданий одного тайтла.",
    "**Больше сигналов:** просмотры, корзина, возвраты, «не интересно» в качестве негатива.",
    f"**Эмбеддинги текста** (sentence-transformers) вместо TF-IDF для похожих товаров, а для "
    f"{pc(S['no_description'])} товаров без описания — изображения.",
])

# --------------------------- 11. выводы ------------------------------------- #
h1(doc, "11. Выводы")
numbered(doc, [
    "Реализованы 4 бизнес-эвристики, content-based, три коллаборативных подхода (UserKNN, ItemKNN, ALS из готовой "
    "библиотеки implicit) и гибрид. Всё работает на ноутбуке без GPU: полный цикл обучения, подбора весов гибрида "
    "и оценки всех моделей занимает около двух минут.",
    f"Главный практический урок: сильный бейзлайн. Неперсональная эвристика Trending обошла каждую ML-модель по "
    f"отдельности, а лучший результат дал гибрид ML-моделей с этой эвристикой (+{pc(GAIN)} NDCG@10, веса подобраны "
    f"на валидации).",
    f"Точность и разнообразие конфликтуют: лучший по точности гибрид покрывает {cov(HYB['Coverage'])} каталога, "
    f"и {pc(HYB['TrendShare'])} его выдачи — общие для всех хиты; менее точный покрывает {cov(HYB_NT['Coverage'])}. "
    f"Выбор между ними — бизнес-решение, принимается A/B-тестом.",
    f"Холодный старт в этих данных в первую очередь касается товаров ({NEW_ITEMS} тестовых покупок — новинки), а не "
    f"пользователей ({COLD_SHARE}). Коллаборативные модели здесь бессильны, нужны content-based и квоты на новинки.",
    "Офлайн-метрики — только фильтр перед A/B-тестом. Реальный эффект для бизнеса меряется инкрементальной "
    "выручкой против бейзлайна, и обычно он заметно скромнее ожиданий.",
])

# --------------------------- 12. ссылки ------------------------------------- #
h1(doc, "12. Ссылки")
para(doc, f"Репозиторий проекта: {REPO_URL}", bold=True)
if "<" in REPO_URL:
    para(doc, "(ссылку подставить в REPO_URL в reports/make_docx.py после публикации)",
         size=10, italic=True, color=WARN, space_after=10)
text(doc, "Данные и библиотеки (открытый доступ):")
bullets(doc, [
    "Amazon Reviews 2023: https://huggingface.co/datasets/McAuley-Lab/Amazon-Reviews-2023",
    "Hou Y. et al. Bridging Language and Items for Retrieval and Recommendation, 2024 (статья о датасете): "
    "https://arxiv.org/abs/2403.03952",
    "implicit: https://github.com/benfred/implicit",
    "scikit-learn: https://scikit-learn.org",
    "FastAPI: https://fastapi.tiangolo.com",
])
text(doc, "Литература и кейсы:")
bullets(doc, [
    "Hu Y., Koren Y., Volinsky C. Collaborative Filtering for Implicit Feedback Datasets. ICDM, 2008 "
    "(ALS для неявного фидбека).",
    "Linden G., Smith B., York J. Amazon.com Recommendations: Item-to-Item Collaborative Filtering. "
    "IEEE Internet Computing, 2003.",
    "Cormack G., Clarke C., Buettcher S. Reciprocal Rank Fusion outperforms Condorcet and individual Rank "
    "Learning Methods. SIGIR, 2009.",
    "Calandrino J. et al. «You Might Also Like»: Privacy Risks of Collaborative Filtering. IEEE S&P, 2011.",
    "Jannach D., Jugovac M. Measuring the Business Value of Recommender Systems. ACM TMIS, 2019.",
    "Duhigg C. How Companies Learn Your Secrets. The New York Times Magazine, 2012 (кейс Target).",
    "Channel 4 News, 2017: расследование о рекомендациях Amazon «Frequently bought together» для компонентов "
    "взрывчатки.",
])

doc.save(str(OUT))
print("документ →", OUT)
missing = [SCREENSHOTS[f][0] for f in SCREENSHOTS if not (SCREENS / SCREENSHOTS[f][0]).exists()]
if missing:
    print(f"скриншотов не хватает ({len(missing)}): {', '.join(missing)}")

if "--pdf" in sys.argv:
    soffice = shutil.which("soffice") or "/Applications/LibreOffice.app/Contents/MacOS/soffice"
    subprocess.run([soffice, "--headless", "--convert-to", "pdf", "--outdir", str(REPORTS), str(OUT)],
                   check=True, capture_output=True)
    print("pdf →", OUT.with_suffix(".pdf"))
