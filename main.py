import os
import html
import threading
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

# Дополнительные русские подписи для тех полей,
# которые QUIK может не описать в Description.
QUIK_FIELD_LABELS = {
    "TransId":          "ID транзакции",
    "TradeNum":         "Номер сделки",
    "OrderNum":         "Номер заявки",
    "FirmCode":         "Код фирмы",
    "FirmName":         "Наименование фирмы",
    "ClientCode":       "Код клиента",
    "Account":          "Счёт",
    "Login":            "Логин",
    "Date":             "Дата",
    "Time":             "Время",
    "TradeDate":        "Дата сделки",
    "SettleDate":       "Дата расчётов",
    "SecCode":          "Код бумаги",
    "SecName":          "Наименование бумаги",
    "ClassCode":        "Код класса",
    "Exchange":         "Биржа",
    "Operation":        "Операция",
    "ExecType":         "Тип исполнения",
    "Status":           "Статус",
    "SettleCode":       "Код расчётов",
    "Currency":         "Валюта",
    "Quantity":         "Количество",
    "Price":            "Цена",
    "Volume":           "Объём",
    "Value":            "Стоимость",
    "AccruedInterest":  "НКД",
    "BrokerRef":        "Примечание",
    "Comment":          "Комментарий",
    "Partner":          "Контрагент",
    "Inout":            "Ввод/вывод",
    "Reason":           "Основание",
    "RejectReason":     "Причина отказа",
}

QUIK_PREFERRED_ORDER = [
    "TransId", "TradeNum", "OrderNum",
    "Date", "Time", "TradeDate", "SettleDate",
    "SecCode", "SecName", "ClassCode", "Exchange",
    "Operation", "ExecType", "Status",
    "Quantity", "Price", "Volume", "Value",
    "Currency", "AccruedInterest",
    "FirmCode", "FirmName", "ClientCode", "Account",
    "BrokerRef", "Comment",
]

QUIK_NUMERIC_FIELDS = {
    "Quantity", "Price", "Volume", "Value",
    "AccruedInterest", "PosValue",
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
    if col.startswith("@"):
        return col[1:]
    if "." in col:
        last = col.split(".")[-1]
        if last.startswith("@"):
            return last[1:]
        return last
    return col


def format_quik_value(col, value):
    base = base_field_name(col)
    if value is None:
        return ""
    s = str(value).strip()

    if base in ("Date", "TradeDate", "SettleDate") and len(s) == 8 and s.isdigit():
        return f"{s[6:8]}.{s[4:6]}.{s[0:4]}"

    if base == "Time" and len(s) == 6 and s.isdigit():
        return f"{s[0:2]}:{s[2:4]}:{s[4:6]}"

    if base in QUIK_NUMERIC_FIELDS:
        try:
            num = float(s.replace(",", "."))
            if num == int(num):
                return str(int(num))
            return f"{num:g}"
        except ValueError:
            return s

    return s


def detect_namespace(path):
    with open(path, "rb") as f:
        for _, elem in ET.iterparse(f, events=("start",)):
            tag = elem.tag
            if isinstance(tag, str) and tag.startswith("{"):
                return tag.split("}", 1)[0][1:]
            return None
    return None


def _path_under(root, node):
    parts = []
    cur = node
    guard = 0
    while cur is not None and cur is not root:
        parts.append(strip_ns(cur.tag))
        cur = cur.getparent()
        guard += 1
        if guard > 1000:
            break
    return ".".join(reversed(parts)) or strip_ns(node.tag)


def _add_field(rec, key, value):
    if key not in rec:
        rec[key] = value
        return
    i = 2
    while f"{key}_{i}" in rec:
        i += 1
    rec[f"{key}_{i}"] = value


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
# 3. ИЗВЛЕЧЕНИЕ ЗАПИСЕЙ С РАЗВОРОТОМ TransData/Field
# ============================================================

class ExtractConfig:
    """Параметры распаковки QUIK-структуры."""
    def __init__(self,
                 container_tag="TransData",
                 field_tag="Field",
                 name_tag="Name",
                 description_tag="Description",
                 value_tag="Value"):
        self.container_tag = container_tag
        self.field_tag = field_tag
        self.name_tag = name_tag
        self.description_tag = description_tag
        self.value_tag = value_tag


def extract_records(path, record_tag, namespace=None,
                    config=None, descriptions=None,
                    progress_cb=None, cancel_flag=None, max_records=None):
    """
    Потоково извлекает записи.

    Логика:
      1) Атрибуты самой записи -> "@attr".
      2) Прямой текст -> "_text".
      3) Обычные вложенные поля -> "path.to.field".
      4) Внутри контейнера config.container_tag (по умолчанию TransData):
         находим узлы config.field_tag (по умолчанию Field),
         у каждого берём Name / Description / Value,
         в запись кладём {Name: Value},
         а описания из Description запоминаем в словаре descriptions,
         чтобы использовать их как заголовки колонок.

    Параметр descriptions — общий dict (Name -> Description).
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

            # 1. атрибуты самой записи
            for k, v in elem.attrib.items():
                rec[f"@{strip_ns(k)}"] = v

            # 2. прямой текст
            direct_text = (elem.text or "").strip()
            if direct_text:
                rec["_text"] = direct_text

            # 3. найдём все контейнеры (TransData) и соберём их Field-ы
            containers = elem.findall(f".//{config.container_tag}")
            # если namespace задан, попробуем и с ним
            if namespace and not containers:
                containers = elem.findall(
                    f".//{{{namespace}}}{config.container_tag}")

            handled_field_elements = set()

            for container in containers:
                # Field-ы внутри контейнера
                fields = container.findall(config.field_tag)
                if namespace and not fields:
                    fields = container.findall(
                        f"{{{namespace}}}{config.field_tag}")

                for field in fields:
                    name_el = field.find(config.name_tag)
                    if name_el is None and namespace:
                        name_el = field.find(f"{{{namespace}}}{config.name_tag}")

                    value_el = field.find(config.value_tag)
                    if value_el is None and namespace:
                        value_el = field.find(f"{{{namespace}}}{config.value_tag}")

                    desc_el = field.find(config.description_tag)
                    if desc_el is None and namespace:
                        desc_el = field.find(
                            f"{{{namespace}}}{config.description_tag}")

                    name = (name_el.text or "").strip() if name_el is not None else ""
                    value = (value_el.text or "").strip() if value_el is not None else ""
                    desc = (desc_el.text or "").strip() if desc_el is not None else ""

                    # запоминаем описание (один раз)
                    if name and desc and name not in descriptions:
                        descriptions[name] = desc

                    if not name:
                        continue

                    # если поля с таким именем ещё нет — пишем значение
                    # если уже есть — суффикс _2, _3, ...
                    _add_field(rec, name, value)

                    # помечаем, что этот Field уже обработан,
                    # чтобы не задваивать его в общем обходе
                    handled_field_elements.add(id(field))
                    # и сам контейнер тоже помечаем
                    handled_field_elements.add(id(container))

            # 4. обход остальных потомков (кроме Field в TransData)
            for child in elem.iter():
                if child is elem:
                    continue
                if not isinstance(child.tag, str):
                    continue
                if id(child) in handled_field_elements:
                    continue

                # пропускаем всё, что внутри TransData — уже разобрано
                in_container = False
                cur = child
                while cur is not None and cur is not elem:
                    if strip_ns(cur.tag) == config.container_tag:
                        in_container = True
                        break
                    cur = cur.getparent()
                if in_container:
                    continue

                path_to_child = _path_under(elem, child)

                text = (child.text or "").strip()
                if text:
                    _add_field(rec, path_to_child, text)

                for ak, av in child.attrib.items():
                    _add_field(rec, f"{path_to_child}.@{strip_ns(ak)}", av)

                tail = (child.tail or "").strip()
                if tail:
                    _add_field(rec, f"{path_to_child}._tail", tail)

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
    """Текстовый дамп первой записи — для диагностики структуры."""
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
                tail = (node.tail or "").strip()

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

                if tail:
                    lines.append(f"{pad}  # tail: {tail!r}")

            walk(elem)
            return "\n".join(lines)
    return "(записей не найдено)"


# ============================================================
# 4. КОЛОНКИ
# ============================================================

def label_for(field_name, descriptions=None):
    if field_name in QUIK_FIELD_LABELS:
        return QUIK_FIELD_LABELS[field_name]
    base = base_field_name(field_name)
    if base in QUIK_FIELD_LABELS:
        return QUIK_FIELD_LABELS[base]
    if descriptions and base in descriptions:
        return descriptions[base]
    if descriptions and field_name in descriptions:
        return descriptions[field_name]
    return field_name


def build_columns(records, descriptions=None):
    seen = {}
    for rec in records:
        for k in rec:
            seen[k] = True

    all_keys = list(seen.keys())

    preferred = []
    others = []
    used = set()

    for pref in QUIK_PREFERRED_ORDER:
        for k in all_keys:
            if k in used:
                continue
            if base_field_name(k) == pref:
                preferred.append(k)
                used.add(k)

    for k in all_keys:
        if k not in used:
            others.append(k)

    columns = preferred + others
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
    title = doc.add_heading("Отчёт по транзакциям QUIK", level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    meta = doc.add_paragraph()
    meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = meta.add_run(f"Источник: {os.path.basename(xml_path)}")
    r.italic = True

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
<p><i>Источник: {source}<br>Записей: {count}</i></p>
"""


def _esc(s):
    return html.escape(str(s))


def export_html(xml_path, out_path, record_tag, namespace=None,
                config=None, limit=100000, progress_cb=None,
                cancel_flag=None, as_cards=False):
    descriptions = {}
    records = list(extract_records(
        xml_path, record_tag, namespace,
        config=config, descriptions=descriptions,
        progress_cb=progress_cb, cancel_flag=cancel_flag,
        max_records=limit))

    columns, header_map = build_columns(records, descriptions)

    with open(out_path, "w", encoding="utf-8") as f:
        f.write(HTML_HEAD.format(
            title="Отчёт по транзакциям QUIK",
            source=_esc(os.path.basename(xml_path)),
            count=len(records),
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
    story.append(Paragraph(
        f"Источник: {_esc(os.path.basename(xml_path))} &nbsp;|&nbsp; "
        f"Записей: {len(records)}", small))
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
        self.geometry("820x600")

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
        self.geometry("780x720")
        self.resizable(False, False)

        self.xml_path = tk.StringVar()
        self.format_var = tk.StringVar(value="html")
        self.record_tag = tk.StringVar(value="Trans")
        self.limit = tk.IntVar(value=100000)
        self.view_mode = tk.StringVar(value="table")

        # параметры распаковки TransData/Field
        self.container_tag = tk.StringVar(value="TransData")
        self.field_tag = tk.StringVar(value="Field")
        self.name_tag = tk.StringVar(value="Name")
        self.desc_tag = tk.StringVar(value="Description")
        self.value_tag = tk.StringVar(value="Value")

        self.cancel_flag = threading.Event()
        self.detected_tags = []

        self._build_ui()

    def _build_ui(self):
        pad = {"padx": 10, "pady": 6}

        frame_file = ttk.LabelFrame(self, text="1. XML-файл QUIK")
        frame_file.pack(fill="x", **pad)
        ttk.Entry(frame_file, textvariable=self.xml_path, width=65).pack(
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

        # --- параметры распаковки Field ---
        frame_fields = ttk.LabelFrame(
            self, text="4. Параметры TransData / Field")
        frame_fields.pack(fill="x", **pad)

        row1 = ttk.Frame(frame_fields)
        row1.pack(fill="x", padx=6, pady=4)
        ttk.Label(row1, text="Контейнер:").pack(side="left")
        ttk.Entry(row1, textvariable=self.container_tag, width=14).pack(
            side="left", padx=4)
        ttk.Label(row1, text="Тег поля:").pack(side="left", padx=(12, 0))
        ttk.Entry(row1, textvariable=self.field_tag, width=10).pack(
            side="left", padx=4)

        row2 = ttk.Frame(frame_fields)
        row2.pack(fill="x", padx=6, pady=4)
        ttk.Label(row2, text="Тег имени:").pack(side="left")
        ttk.Entry(row2, textvariable=self.name_tag, width=12).pack(
            side="left", padx=4)
        ttk.Label(row2, text="Тег описания:").pack(side="left", padx=(12, 0))
        ttk.Entry(row2, textvariable=self.desc_tag, width=14).pack(
            side="left", padx=4)
        ttk.Label(row2, text="Тег значения:").pack(side="left", padx=(12, 0))
        ttk.Entry(row2, textvariable=self.value_tag, width=10).pack(
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

    # ---------- действия ----------

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
            name_tag=self.name_tag.get().strip() or "Name",
            description_tag=self.desc_tag.get().strip() or "Description",
            value_tag=self.value_tag.get().strip() or "Value",
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