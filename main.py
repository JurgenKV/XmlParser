import os
import html
import threading
from datetime import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from collections import Counter, defaultdict

from lxml import etree as ET

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.lib import colors
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
)
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont


# ============================================================
# 0. КОНСТАНТЫ
# ============================================================

QUIK_NS = "urn:quik:trans-info:v1.0"

# Точные ключи (совпадают с тем, что реально кладётся в rec)
QUIK_FIELD_LABELS = {
    # Атрибуты Trans
    "@TransNum":   "№ транзакции",
    "@UID":        "UID",
    "@TransID":    "ID транзакции",
    "@SessionID":  "ID сессии",
    "@TradeDate":  "Дата сделки",
    "@Status":     "Статус",
    "@QuikDate":   "Дата QUIK",
    "@QuikTime":   "Время QUIK",
    "@ReplyTime":  "Время ответа",
    "@OrderNum":   "№ заявки",
    "@ClientCode": "Код клиента",

    # Текстовые потомки Trans
    "Data":     "Данные",
    "Reply":    "Ответ",
    "PureData": "PureData",

    # Атрибуты UserInfo
    "UserInfo.@Name1":   "Имя 1",
    "UserInfo.@Name2":   "Имя 2",
    "UserInfo.@Name3":   "Имя 3",
    "UserInfo.@OrgCode": "Код организации",
    "UserInfo.@OrgName": "Организация",
    "UserInfo.@Login":   "Логин",
}

# Числовые поля QUIK (для отображения чисел, даже если PreparedValue вдруг пустое)
QUIK_NUMERIC_FIELDS = {
    "PRICE", "QUANTITY", "ORDERVALUE", "VALUE", "VOLUME",
    "ACCruedInterest", "AccruedInterest",
}


# ============================================================
# 1. ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
# ============================================================

def strip_ns(tag):
    if not isinstance(tag, str):
        return str(tag)
    if tag.startswith("{"):
        return tag.split("}", 1)[1]
    return tag


def qname(local_name, namespace):
    if namespace:
        return f"{{{namespace}}}{local_name}"
    return local_name


def base_field_name(col):
    """'@TransID' -> 'TransID'; 'UserInfo.@Login' -> 'Login'."""
    if col.startswith("@"):
        return col[1:]
    if "." in col:
        last = col.split(".")[-1]
        if last.startswith("@"):
            return last[1:]
        return last
    return col


def label_for(field_name, descriptions=None):
    # 1) точный ключ
    if field_name in QUIK_FIELD_LABELS:
        return QUIK_FIELD_LABELS[field_name]

    # 2) базовое имя
    base = base_field_name(field_name)
    if base in QUIK_FIELD_LABELS:
        return QUIK_FIELD_LABELS[base]

    # 3) Description из TransData/Field
    if descriptions:
        if field_name in descriptions:
            return descriptions[field_name]
        if base in descriptions:
            return descriptions[base]

    # 4) fallback — вернуть как есть
    return field_name


def format_date_string(s):
    """YYYY-MM-DD или YYYYMMDD -> DD.MM.YYYY"""
    if not s:
        return s
    if len(s) == 10 and s[4] == "-" and s[7] == "-":
        y, m, d = s.split("-")
        if y.isdigit() and m.isdigit() and d.isdigit():
            return f"{d}.{m}.{y}"
    if len(s) == 8 and s.isdigit():
        return f"{s[6:8]}.{s[4:6]}.{s[0:4]}"
    return s


def format_number(value, scale=None):
    """Форматирует число с учётом Scale (кол-во знаков после запятой)."""
    if value is None or value == "":
        return ""
    s = str(value).strip().replace(",", ".")
    # если scale > 0 — QUIK хранит как целое, делим на 10^scale
    if scale and scale > 0:
        try:
            num = float(s) / (10 ** scale)
            return f"{num:.{scale}f}".rstrip("0").rstrip(".")
        except ValueError:
            pass
    # обычное число
    try:
        num = float(s)
        if num == int(num):
            return str(int(num))
        return f"{num:g}"
    except ValueError:
        return s


def format_quik_value(col, value, numeric_types=None):
    """
    Форматирует значение.
    col — ключ колонки,
    value — уже подготовленное к отображению (обычно PreparedValue),
    numeric_types — не используется (оставлено для совместимости).
    """
    base = base_field_name(col)
    if value is None:
        return ""
    s = str(value).strip()

    # даты
    if len(s) == 10 and s[4] == "-" and s[7] == "-":
        y, m, d = s.split("-")
        if y.isdigit() and m.isdigit() and d.isdigit():
            return f"{d}.{m}.{y}"

    if base in ("TradeDate", "QuikDate", "Date", "SettleDate"):
        return format_date_string(s)

    # время
    if base in ("QuikTime", "ReplyTime", "Time"):
        return s  # обычно уже HH:MM:SS.xxxxxx

    # числа
    if base in QUIK_NUMERIC_FIELDS:
        return format_number(s)

    return s


def detect_namespace(path):
    with open(path, "rb") as f:
        for _, elem in ET.iterparse(f, events=("start",)):
            tag = elem.tag
            if isinstance(tag, str) and tag.startswith("{"):
                return tag.split("}", 1)[0][1:]
            return None
    return None


def read_report_header(path):
    """
    Читает атрибуты корневого узла TransactionsReport:
    ProgramVersion, StartDate, EndDate.
    """
    info = {"ProgramVersion": "", "StartDate": "", "EndDate": ""}
    with open(path, "rb") as f:
        for _, elem in ET.iterparse(f, events=("start",)):
            for k in ("ProgramVersion", "StartDate", "EndDate"):
                v = elem.get(k)
                if v:
                    info[k] = v
            break
    return info


# ============================================================
# 2. АНАЛИЗ СТРУКТУРЫ
# ============================================================

def analyze_structure(path, sample_limit=200000):
    stats = defaultdict(lambda: {"count": 0, "children": Counter()})

    with open(path, "rb") as f:
        context = ET.iterparse(f, events=("start", "end"))
        _, root = next(context)

        stack = []
        processed = 0

        for event, elem in context:
            tag = elem.tag if isinstance(elem.tag, str) else None

            if event == "start":
                stack.append(elem)
                if tag is not None:
                    stats[tag]["count"] += 1
            else:
                if stack:
                    stack.pop()
                if tag is not None and stack:
                    parent = stack[-1]
                    ptag = parent.tag if isinstance(parent.tag, str) else None
                    if ptag is not None:
                        stats[ptag]["children"][tag] += 1

                elem.clear()
                while elem.getprevious() is not None:
                    del elem.getparent()[0]

                processed += 1
                if processed >= sample_limit:
                    break

        del root
    return stats


def detect_record_tag(stats, preferred="Trans"):
    for tag, info in stats.items():
        local = strip_ns(tag)
        if local == preferred and info["count"] >= 1:
            ns = tag[1:tag.index("}")] if tag.startswith("{") else None
            return ns, local

    best = None
    best_score = 0
    for tag, info in stats.items():
        cnt = info["count"]
        n_children = len(info["children"])
        if cnt < 2 or n_children < 1:
            continue
        total_child_uses = sum(info["children"].values())
        score = cnt * min(n_children, 5) + total_child_uses
        if len(info["children"]) == 1:
            score *= 0.6
        if score > best_score:
            best_score = score
            best = tag

    if not best:
        return None, None
    ns = best[1:best.index("}")] if best.startswith("{") else None
    return ns, strip_ns(best)


# ============================================================
# 3. ИЗВЛЕЧЕНИЕ ЗАПИСЕЙ
# ============================================================

class ExtractConfig:
    def __init__(self,
                 container_tag="TransData",
                 field_tag="Field",
                 name_attr="Name",
                 description_attr="Description",
                 value_attr="Value",
                 prepared_attr="PreparedValue"):
        self.container_tag = container_tag
        self.field_tag = field_tag
        self.name_attr = name_attr
        self.description_attr = description_attr
        self.value_attr = value_attr
        self.prepared_attr = prepared_attr


def extract_records(path, record_tag, namespace=None,
                    config=None, descriptions=None,
                    progress_cb=None, cancel_flag=None, max_records=None):
    """
    Порядок ключей в rec — строго как в XML:
      1. Атрибуты Trans (в порядке следования).
      2. Дочерние узлы Trans в порядке следования: текст узла → затем его атрибуты.
      3. Только когда доходим до TransData — её Field (отсортированные по Number).
    """
    if config is None:
        config = ExtractConfig()
    if descriptions is None:
        descriptions = {}

    search_tag = qname(record_tag, namespace)

    with open(path, "rb") as f:
        context = ET.iterparse(f, events=("end",), tag=search_tag)

        count = 0
        for _, elem in context:
            if cancel_flag and cancel_flag.is_set():
                return

            rec = {}

            # 1. Атрибуты Trans
            for k, v in elem.attrib.items():
                rec[f"@{strip_ns(k)}"] = v

            # 2. Обход детей Trans в порядке XML
            for child in elem:
                if not isinstance(child.tag, str):
                    continue
                tag_local = strip_ns(child.tag)

                if tag_local == config.container_tag:
                    # 3. Field-ы внутри TransData — сортируем по Number
                    fields = list(child.findall(config.field_tag))
                    if namespace and not fields:
                        fields = list(child.findall(
                            f"{{{namespace}}}{config.field_tag}"))

                    def _num_key(fe):
                        try:
                            return int(fe.get("Number") or 0)
                        except ValueError:
                            return 0
                    fields.sort(key=_num_key)

                    for field in fields:
                        name = (field.get(config.name_attr) or "").strip()
                        desc = (field.get(config.description_attr) or "").strip()
                        value = (field.get(config.value_attr) or "").strip()
                        prepared = (field.get(config.prepared_attr) or "").strip()
                        scale_attr = (field.get("Scale") or "").strip()

                        if not name:
                            continue

                        scale = int(scale_attr) if scale_attr.isdigit() else 0

                        if desc and name not in descriptions:
                            descriptions[name] = desc

                        # приоритет: PreparedValue → Value
                        if prepared:
                            display = prepared
                        elif scale > 0:
                            display = format_number(value, scale)
                        else:
                            display = value

                        if name in rec:
                            i = 2
                            while f"{name}_{i}" in rec:
                                i += 1
                            rec[f"{name}_{i}"] = display
                        else:
                            rec[name] = display
                    continue  # TransData обработана, дальше не идём

                # обычный дочерний узел — сначала текст, потом атрибуты
                text = (child.text or "").strip()
                if text:
                    key = tag_local
                    if key in rec:
                        i = 2
                        while f"{key}_{i}" in rec:
                            i += 1
                        rec[f"{key}_{i}"] = text
                    else:
                        rec[key] = text

                for ak, av in child.attrib.items():
                    key = f"{tag_local}.@{strip_ns(ak)}"
                    if key in rec:
                        i = 2
                        while f"{key}_{i}" in rec:
                            i += 1
                        rec[f"{key}_{i}"] = av
                    else:
                        rec[key] = av

            yield rec

            elem.clear()
            while elem.getprevious() is not None:
                del elem.getparent()[0]

            count += 1
            if progress_cb and count % 1000 == 0:
                progress_cb(count)
            if max_records and count >= max_records:
                return


def dump_first_record(xml_path, record_tag="Trans", namespace=None,
                      max_depth=10):
    search_tag = qname(record_tag, namespace)

    with open(xml_path, "rb") as f:
        for _, elem in ET.iterparse(f, events=("end",), tag=search_tag):
            lines = []

            def walk(node, depth=0):
                if depth > max_depth:
                    lines.append("  " * depth + "...")
                    return
                pad = "  " * depth
                t = strip_ns(node.tag)
                attrs = " ".join(f'{strip_ns(k)}="{v}"'
                                 for k, v in node.attrib.items())
                text = (node.text or "").strip()
                head = f"{pad}<{t}"
                if attrs:
                    head += " " + attrs
                head += ">"
                if text:
                    head += f" {text!r}"
                lines.append(head)

                for c in node:
                    if not isinstance(c.tag, str):
                        continue
                    walk(c, depth + 1)

            walk(elem)
            return "\n".join(lines)
    return "(записей не найдено)"


# ============================================================
# 4. КОЛОНКИ — строго по XML
# ============================================================

def build_columns(records, descriptions=None):
    """
    Порядок ключей — как в первой записи (то есть как в XML).
    Если в других записях встречаются новые ключи — они идут в конец
    в порядке появления.
    """
    seen = {}
    for rec in records:
        for k in rec:
            if k not in seen:
                seen[k] = True

    columns = list(seen.keys())
    header_map = {c: label_for(c, descriptions) for c in columns}
    return columns, header_map


# ============================================================
# 5. DOCX
# ============================================================

def _add_table_docx(doc, records, columns, header_map, title):
    doc.add_heading(title, level=1)

    table = doc.add_table(rows=1, cols=len(columns))
    try:
        table.style = "Light Grid Accent 1"
    except KeyError:
        pass
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    hdr = table.rows[0].cells
    for i, col in enumerate(columns):
        hdr[i].text = header_map[col]
        for p in hdr[i].paragraphs:
            for r in p.runs:
                r.bold = True

    for rec in records:
        row = table.add_row().cells
        for i, col in enumerate(columns):
            row[i].text = format_quik_value(col, rec.get(col, ""))
    return table


def _add_cards_docx(doc, records, columns, header_map, title):
    doc.add_heading(title, level=1)
    for i, rec in enumerate(records, 1):
        p = doc.add_paragraph()
        run = p.add_run(f"Запись {i}")
        run.bold = True
        run.font.size = Pt(13)
        run.font.color.rgb = RGBColor(0x00, 0x4A, 0x99)

        for col in columns:
            if col not in rec:
                continue
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Cm(0.5)
            r1 = p.add_run(f"{header_map[col]}: ")
            r1.bold = True
            p.add_run(format_quik_value(col, rec[col]))
        doc.add_paragraph()


def export_docx(xml_path, out_path, record_tag, namespace=None,
                config=None, limit=100000, progress_cb=None,
                cancel_flag=None, as_cards=False):
    doc = Document()

    header = read_report_header(xml_path)

    title = doc.add_heading("Отчёт по транзакциям QUIK", level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    meta = doc.add_paragraph()
    meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
    meta.add_run("Дата формирования отчёта: "
                 f"{datetime.now().strftime('%d.%m.%Y %H:%M')}\n").italic = True
    if header.get("StartDate") or header.get("EndDate"):
        s = format_date_string(header.get("StartDate", ""))
        e = format_date_string(header.get("EndDate", ""))
        meta.add_run(f"Период: {s} — {e}\n").italic = True
    if header.get("ProgramVersion"):
        meta.add_run(f"Версия QUIK: {header['ProgramVersion']}\n").italic = True
    meta.add_run(f"Источник: {os.path.basename(xml_path)}").italic = True

    descriptions = {}
    records = list(extract_records(
        xml_path, record_tag, namespace,
        config=config, descriptions=descriptions,
        progress_cb=progress_cb, cancel_flag=cancel_flag,
        max_records=limit))

    if not records:
        doc.add_paragraph("Записи не найдены.")
        doc.save(out_path)
        return

    columns, header_map = build_columns(records, descriptions)

    if as_cards:
        _add_cards_docx(doc, records, columns, header_map, "Транзакции")
    else:
        _add_table_docx(doc, records, columns, header_map,
                        "Таблица транзакций")

    doc.save(out_path)


# ============================================================
# 6. HTML
# ============================================================

HTML_HEAD = """<!DOCTYPE html>
<html lang="ru"><head><meta charset="utf-8">
<title>{title}</title>
<style>
  body {{ font-family: -apple-system, Segoe UI, Arial, sans-serif;
         margin: 2em; color: #222; }}
  h1 {{ border-bottom: 2px solid #444; padding-bottom: .3em; }}
  .meta {{ color: #555; font-size: 13px; margin-bottom: 1em; }}
  table {{ border-collapse: collapse; margin: 1em 0; width: 100%;
           font-size: 13px; }}
  th, td {{ border: 1px solid #ccc; padding: 6px 10px; text-align: left;
            vertical-align: top; }}
  th {{ background: #f0f4f8; position: sticky; top: 0; }}
  tr:nth-child(even) {{ background: #fafbfc; }}
  td.num {{ text-align: right; font-variant-numeric: tabular-nums; }}
  .card {{ border: 1px solid #ddd; border-radius: 6px; padding: 12px 16px;
           margin: 8px 0; background: #fafbfc; }}
  .card h3 {{ margin-top: 0; color: #004a99; }}
  .field {{ margin: 2px 0; }}
  .field b {{ color: #333; }}
</style></head><body>
<h1>{title}</h1>
<div class="meta">{meta}</div>
"""


def _esc(s):
    return html.escape(str(s))


def export_html(xml_path, out_path, record_tag, namespace=None,
                config=None, limit=100000, progress_cb=None,
                cancel_flag=None, as_cards=False):
    header = read_report_header(xml_path)

    descriptions = {}
    records = list(extract_records(
        xml_path, record_tag, namespace,
        config=config, descriptions=descriptions,
        progress_cb=progress_cb, cancel_flag=cancel_flag,
        max_records=limit))

    columns, header_map = build_columns(records, descriptions)

    meta_lines = [
        f"Дата формирования: {datetime.now().strftime('%d.%m.%Y %H:%M')}",
    ]
    if header.get("StartDate") or header.get("EndDate"):
        s = format_date_string(header.get("StartDate", ""))
        e = format_date_string(header.get("EndDate", ""))
        meta_lines.append(f"Период: {s} — {e}")
    if header.get("ProgramVersion"):
        meta_lines.append(f"Версия QUIK: {header['ProgramVersion']}")
    meta_lines.append(f"Источник: {os.path.basename(xml_path)}")
    meta_lines.append(f"Записей: {len(records)}")

    with open(out_path, "w", encoding="utf-8") as f:
        f.write(HTML_HEAD.format(
            title="Отчёт по транзакциям QUIK",
            meta="<br>".join(_esc(m) for m in meta_lines),
        ))

        if as_cards:
            for i, rec in enumerate(records, 1):
                f.write(f'<div class="card"><h3>Запись {i}</h3>\n')
                for col in columns:
                    if col in rec:
                        f.write(
                            f'<div class="field"><b>{_esc(header_map[col])}:</b> '
                            f'{_esc(format_quik_value(col, rec[col]))}</div>\n')
                f.write("</div>\n")
        else:
            f.write("<table><thead><tr>")
            for col in columns:
                f.write(f"<th>{_esc(header_map[col])}</th>")
            f.write("</tr></thead><tbody>\n")
            for rec in records:
                f.write("<tr>")
                for col in columns:
                    cls = ("num"
                           if base_field_name(col) in QUIK_NUMERIC_FIELDS
                           else "")
                    f.write(
                        f'<td class="{cls}">'
                        f'{_esc(format_quik_value(col, rec.get(col, "")))}</td>')
                f.write("</tr>\n")
            f.write("</tbody></table>\n")

        f.write("</body></html>\n")


# ============================================================
# 7. PDF
# ============================================================

def register_cyrillic_font():
    for path in [
        "C:/Windows/Fonts/DejaVuSans.ttf",
        "C:/Windows/Fonts/arial.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/Library/Fonts/Arial Unicode.ttf",
    ]:
        if os.path.exists(path):
            try:
                pdfmetrics.registerFont(TTFont("UnicodeFont", path))
                return "UnicodeFont"
            except Exception:
                continue
    return "Helvetica"


def export_pdf(xml_path, out_path, record_tag, namespace=None,
               config=None, limit=100000, progress_cb=None,
               cancel_flag=None, as_cards=False):
    font = register_cyrillic_font()

    styles = getSampleStyleSheet()
    h1 = ParagraphStyle("H1", parent=styles["Heading1"], fontName=font,
                        fontSize=16, leading=20)
    normal = ParagraphStyle("N", parent=styles["Normal"], fontName=font,
                            fontSize=9, leading=12)
    small = ParagraphStyle("S", parent=styles["Normal"], fontName=font,
                           fontSize=8, leading=10, textColor=colors.grey)

    header = read_report_header(xml_path)

    descriptions = {}
    records = list(extract_records(
        xml_path, record_tag, namespace,
        config=config, descriptions=descriptions,
        progress_cb=progress_cb, cancel_flag=cancel_flag,
        max_records=limit))

    columns, header_map = build_columns(records, descriptions)

    use_landscape = len(columns) > 5
    pagesize = landscape(A4) if use_landscape else A4

    doc = SimpleDocTemplate(
        out_path, pagesize=pagesize,
        leftMargin=1.2 * cm, rightMargin=1.2 * cm,
        topMargin=1.2 * cm, bottomMargin=1.2 * cm,
    )

    story = []
    story.append(Paragraph("Отчёт по транзакциям QUIK", h1))

    meta_lines = [
        f"Дата формирования: {datetime.now().strftime('%d.%m.%Y %H:%M')}",
    ]
    if header.get("StartDate") or header.get("EndDate"):
        s = format_date_string(header.get("StartDate", ""))
        e = format_date_string(header.get("EndDate", ""))
        meta_lines.append(f"Период: {s} — {e}")
    if header.get("ProgramVersion"):
        meta_lines.append(f"Версия QUIK: {header['ProgramVersion']}")
    meta_lines.append(f"Источник: {os.path.basename(xml_path)}")
    meta_lines.append(f"Записей: {len(records)}")

    for line in meta_lines:
        story.append(Paragraph(_esc(line), small))
    story.append(Spacer(1, 8))

    if as_cards:
        for i, rec in enumerate(records, 1):
            story.append(Paragraph(f"<b>Запись {i}</b>", normal))
            for col in columns:
                if col in rec:
                    story.append(Paragraph(
                        f"<b>{_esc(header_map[col])}:</b> "
                        f"{_esc(format_quik_value(col, rec[col]))}", normal))
            story.append(Spacer(1, 6))
    else:
        data = [[Paragraph(f"<b>{_esc(header_map[c])}</b>", normal)
                 for c in columns]]
        for rec in records:
            data.append([
                Paragraph(_esc(format_quik_value(c, rec.get(c, ""))), normal)
                for c in columns
            ])

        avail = (pagesize[0] - 2.4 * cm) / max(len(columns), 1)
        col_widths = [avail] * len(columns)

        tbl = Table(data, colWidths=col_widths, repeatRows=1)
        tbl.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#f0f4f8")),
            ("GRID", (0, 0), (-1, -1), 0.4, colors.grey),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING", (0, 0), (-1, -1), 4),
            ("RIGHTPADDING", (0, 0), (-1, -1), 4),
            ("TOPPADDING", (0, 0), (-1, -1), 3),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ]))
        story.append(tbl)

    doc.build(story)


# ============================================================
# 8. GUI
# ============================================================

class DumpWindow(tk.Toplevel):
    def __init__(self, master, text):
        super().__init__(master)
        self.title("Сырой дамп первой записи")
        self.geometry("900x600")

        frame = ttk.Frame(self)
        frame.pack(fill="both", expand=True, padx=8, pady=8)

        txt = tk.Text(frame, wrap="none", font=("Consolas", 10))
        txt.pack(side="left", fill="both", expand=True)

        sb_y = ttk.Scrollbar(frame, orient="vertical", command=txt.yview)
        sb_y.pack(side="right", fill="y")
        sb_x = ttk.Scrollbar(self, orient="horizontal", command=txt.xview)
        sb_x.pack(side="bottom", fill="x")

        txt.configure(yscrollcommand=sb_y.set, xscrollcommand=sb_x.set)
        txt.insert("1.0", text)
        txt.configure(state="disabled")


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("QUIK XML → отчёт (DOCX / PDF / HTML)")
        self.geometry("820x760")
        self.resizable(False, False)

        self.xml_path = tk.StringVar()
        self.format_var = tk.StringVar(value="html")
        self.record_tag = tk.StringVar(value="Trans")
        self.limit = tk.IntVar(value=100000)
        self.view_mode = tk.StringVar(value="table")

        self.container_tag = tk.StringVar(value="TransData")
        self.field_tag = tk.StringVar(value="Field")
        self.name_attr = tk.StringVar(value="Name")
        self.desc_attr = tk.StringVar(value="Description")
        self.value_attr = tk.StringVar(value="Value")
        self.prepared_attr = tk.StringVar(value="PreparedValue")

        self.cancel_flag = threading.Event()
        self.detected_tags = []

        self._build_ui()

    def _build_ui(self):
        pad = {"padx": 10, "pady": 6}

        frame_file = ttk.LabelFrame(self, text="1. XML-файл QUIK")
        frame_file.pack(fill="x", **pad)
        ttk.Entry(frame_file, textvariable=self.xml_path, width=70).pack(
            side="left", padx=6, pady=6, fill="x", expand=True)
        ttk.Button(frame_file, text="Обзор…", command=self.choose_file).pack(
            side="right", padx=6, pady=6)

        frame_an = ttk.LabelFrame(self, text="2. Диагностика")
        frame_an.pack(fill="x", **pad)
        ttk.Button(frame_an, text="Анализ структуры",
                   command=self.run_analysis).pack(
            side="left", padx=6, pady=6)
        ttk.Button(frame_an, text="Сырой дамп первой записи",
                   command=self.show_dump).pack(side="left", padx=6, pady=6)

        frame_rec = ttk.LabelFrame(
            self, text="3. Тег записи (по умолчанию Trans)")
        frame_rec.pack(fill="x", **pad)
        self.cmb_tag = ttk.Combobox(frame_rec, textvariable=self.record_tag,
                                    values=[], width=40)
        self.cmb_tag.pack(side="left", padx=6, pady=6, fill="x", expand=True)
        ttk.Label(frame_rec, text="Макс. записей:").pack(side="left", padx=6)
        ttk.Entry(frame_rec, textvariable=self.limit, width=10).pack(
            side="left", padx=6)

        frame_fields = ttk.LabelFrame(
            self, text="4. Параметры TransData / Field (атрибуты)")
        frame_fields.pack(fill="x", **pad)

        r1 = ttk.Frame(frame_fields)
        r1.pack(fill="x", padx=6, pady=4)
        ttk.Label(r1, text="Контейнер:").pack(side="left")
        ttk.Entry(r1, textvariable=self.container_tag, width=14).pack(
            side="left", padx=4)
        ttk.Label(r1, text="Тег поля:").pack(side="left", padx=(12, 0))
        ttk.Entry(r1, textvariable=self.field_tag, width=10).pack(
            side="left", padx=4)

        r2 = ttk.Frame(frame_fields)
        r2.pack(fill="x", padx=6, pady=4)
        ttk.Label(r2, text="Атрибут имени:").pack(side="left")
        ttk.Entry(r2, textvariable=self.name_attr, width=12).pack(
            side="left", padx=4)
        ttk.Label(r2, text="Атрибут описания:").pack(side="left", padx=(12, 0))
        ttk.Entry(r2, textvariable=self.desc_attr, width=14).pack(
            side="left", padx=4)

        r3 = ttk.Frame(frame_fields)
        r3.pack(fill="x", padx=6, pady=4)
        ttk.Label(r3, text="Атрибут значения:").pack(side="left")
        ttk.Entry(r3, textvariable=self.value_attr, width=12).pack(
            side="left", padx=4)
        ttk.Label(r3, text="Атрибут PreparedValue:").pack(
            side="left", padx=(12, 0))
        ttk.Entry(r3, textvariable=self.prepared_attr, width=16).pack(
            side="left", padx=4)

        frame_view = ttk.LabelFrame(self, text="5. Представление")
        frame_view.pack(fill="x", **pad)
        ttk.Radiobutton(frame_view, text="Таблица",
                        value="table", variable=self.view_mode).pack(
            side="left", padx=12, pady=6)
        ttk.Radiobutton(frame_view, text="Карточки",
                        value="cards", variable=self.view_mode).pack(
            side="left", padx=12, pady=6)

        frame_fmt = ttk.LabelFrame(self, text="6. Формат вывода")
        frame_fmt.pack(fill="x", **pad)
        for fmt, label in [("html", "HTML"), ("docx", "Word (.docx)"),
                           ("pdf", "PDF")]:
            ttk.Radiobutton(frame_fmt, text=label, value=fmt,
                            variable=self.format_var).pack(
                side="left", padx=12, pady=6)

        frame_prog = ttk.LabelFrame(self, text="Прогресс")
        frame_prog.pack(fill="x", **pad)
        self.progress = ttk.Progressbar(frame_prog, mode="indeterminate")
        self.progress.pack(fill="x", padx=6, pady=8)

        frame_btn = ttk.Frame(self)
        frame_btn.pack(fill="x", **pad)
        self.btn_convert = ttk.Button(frame_btn, text="Сформировать отчёт",
                                      command=self.convert)
        self.btn_convert.pack(side="left", padx=6)
        self.btn_cancel = ttk.Button(frame_btn, text="Отмена",
                                     command=self.cancel, state="disabled")
        self.btn_cancel.pack(side="left", padx=6)

        self.status = ttk.Label(self, text="Готов к работе", foreground="gray")
        self.status.pack(pady=4)

    def choose_file(self):
        path = filedialog.askopenfilename(
            title="Выберите XML-файл QUIK",
            filetypes=[("XML files", "*.xml"), ("All files", "*.*")])
        if path:
            self.xml_path.set(path)

    def cancel(self):
        self.cancel_flag.set()

    def _set_status(self, text, color="gray"):
        self.after(0, lambda: self.status.config(text=text, foreground=color))

    def _progress_cb(self, count):
        if count % 5000 == 0:
            self._set_status(f"Обработано записей: {count:,}", "blue")

    def _make_config(self):
        return ExtractConfig(
            container_tag=self.container_tag.get().strip() or "TransData",
            field_tag=self.field_tag.get().strip() or "Field",
            name_attr=self.name_attr.get().strip() or "Name",
            description_attr=self.desc_attr.get().strip() or "Description",
            value_attr=self.value_attr.get().strip() or "Value",
            prepared_attr=self.prepared_attr.get().strip() or "PreparedValue",
        )

    def run_analysis(self):
        xml_file = self.xml_path.get().strip()
        if not xml_file or not os.path.isfile(xml_file):
            messagebox.showwarning("Внимание", "Сначала выберите XML-файл")
            return

        self.progress.start(10)
        self._set_status("Анализ структуры…", "blue")
        self.btn_convert.config(state="disabled")

        def worker():
            try:
                stats = analyze_structure(xml_file, sample_limit=200000)
                ns, guessed = detect_record_tag(stats, preferred="Trans")

                candidates = sorted(
                    ((strip_ns(tag), info["count"], len(info["children"]))
                     for tag, info in stats.items()
                     if info["count"] >= 2 and len(info["children"]) >= 1),
                    key=lambda x: -x[1]
                )
                top = [c[0] for c in candidates[:20]]
                top = list(dict.fromkeys(top))

                def apply():
                    self.detected_tags = top
                    self.cmb_tag["values"] = top
                    if guessed and not self.record_tag.get():
                        self.record_tag.set(guessed)
                    self._set_status(
                        f"Кандидатов: {len(top)}. Предполагаемая запись: "
                        f"{guessed or '—'} (ns: {ns or '—'})", "green")
                self.after(0, apply)

            except Exception as exc:
                err = f"{type(exc).__name__}: {exc}"
                self.after(0, self._show_error, "Ошибка анализа", err)
            finally:
                self.after(0, self._reset_ui)

        threading.Thread(target=worker, daemon=True).start()

    def show_dump(self):
        xml_file = self.xml_path.get().strip()
        if not xml_file or not os.path.isfile(xml_file):
            messagebox.showwarning("Внимание", "Сначала выберите XML-файл")
            return

        tag = self.record_tag.get().strip() or "Trans"

        def worker():
            try:
                ns = detect_namespace(xml_file)
                dump = dump_first_record(xml_file, tag, ns)
                self.after(0, lambda: DumpWindow(self, dump))
            except Exception as exc:
                err = f"{type(exc).__name__}: {exc}"
                self.after(0, self._show_error, "Ошибка дампа", err)

        threading.Thread(target=worker, daemon=True).start()

    def convert(self):
        xml_file = self.xml_path.get().strip()
        if not xml_file or not os.path.isfile(xml_file):
            messagebox.showwarning("Внимание", "Выберите существующий XML-файл")
            return

        fmt = self.format_var.get()
        ext = fmt
        record_tag = self.record_tag.get().strip() or "Trans"
        as_cards = self.view_mode.get() == "cards"
        limit = max(1, int(self.limit.get() or 100000))
        config = self._make_config()

        out_path = filedialog.asksaveasfilename(
            title="Сохранить как",
            defaultextension=f".{ext}",
            filetypes=[(f"{ext.upper()} files", f"*.{ext}")],
            initialfile=os.path.splitext(os.path.basename(xml_file))[0]
                        + f"_report.{ext}")
        if not out_path:
            return

        self.cancel_flag.clear()
        self.btn_convert.config(state="disabled")
        self.btn_cancel.config(state="normal")
        self.progress.start(10)
        self._set_status("Формирование отчёта…", "blue")

        def worker():
            try:
                namespace = detect_namespace(xml_file)
                self._set_status(
                    f"Namespace: {namespace or '—'}. Разбор транзакций…",
                    "blue")

                if self.cancel_flag.is_set():
                    raise RuntimeError("Отменено")

                if fmt == "html":
                    export_html(xml_file, out_path, record_tag,
                                namespace=namespace, config=config,
                                limit=limit,
                                progress_cb=self._progress_cb,
                                cancel_flag=self.cancel_flag,
                                as_cards=as_cards)
                elif fmt == "docx":
                    export_docx(xml_file, out_path, record_tag,
                                namespace=namespace, config=config,
                                limit=limit,
                                progress_cb=self._progress_cb,
                                cancel_flag=self.cancel_flag,
                                as_cards=as_cards)
                elif fmt == "pdf":
                    export_pdf(xml_file, out_path, record_tag,
                               namespace=namespace, config=config,
                               limit=limit,
                               progress_cb=self._progress_cb,
                               cancel_flag=self.cancel_flag,
                               as_cards=as_cards)
                else:
                    raise ValueError(f"Неизвестный формат: {fmt}")

                self.after(0, self._show_success,
                           f"Файл сохранён:\n{out_path}")

            except Exception as exc:
                err = f"{type(exc).__name__}: {exc}"
                self.after(0, self._show_error, "Ошибка", err)
            finally:
                self.after(0, self._reset_ui)

        threading.Thread(target=worker, daemon=True).start()

    def _show_error(self, title, message):
        self.status.config(text=title, foreground="red")
        messagebox.showerror(title, message)

    def _show_success(self, message):
        self.status.config(text="Готово", foreground="green")
        messagebox.showinfo("Успех", message)

    def _reset_ui(self):
        self.progress.stop()
        self.btn_convert.config(state="normal")
        self.btn_cancel.config(state="disabled")


if __name__ == "__main__":
    App().mainloop()