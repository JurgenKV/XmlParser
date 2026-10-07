import os
import re
import html
import queue
import tempfile
import threading
from datetime import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from lxml import etree as ET

from reportlab.lib.pagesizes import A4, landscape, A3, A2, A1, A0
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

RECORD_TAG = "Trans"
CONTAINER_TAG = "TransData"
FIELD_TAG = "Field"
NAME_ATTR = "Name"
DESC_ATTR = "Description"
VALUE_ATTR = "Value"
PREPARED_ATTR = "PreparedValue"
DEFAULT_LIMIT = 1000000

EMPTY_MARK = "-"
RECORD_NUM_COL = "__record_num__"

# Размер чанка при чтении/записи
CLEAN_CHUNK = 4 * 1024 * 1024   # 4 МБ

# Глубина очередей (сколько чанков может висеть между потоками)
QUEUE_DEPTH = 4

QUIK_FIELD_LABELS = {
    RECORD_NUM_COL: "Запись",

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

    "Data":     "Данные",
    "Reply":    "Ответ",
    "PureData": "PureData",

    "UserInfo.@Name1":   "Имя 1",
    "UserInfo.@Name2":   "Имя 2",
    "UserInfo.@Name3":   "Имя 3",
    "UserInfo.@OrgCode": "Код организации",
    "UserInfo.@OrgName": "Организация",
    "UserInfo.@Login":   "Логин",
}

QUIK_NUMERIC_FIELDS = {
    "PRICE", "QUANTITY", "ORDERVALUE", "VALUE", "VOLUME",
    "AccruedInterest",
}


# ============================================================
# 0.1. САНИТИЗАЦИЯ XML — ускоренная + конвейерная
# ============================================================

# Неэкранированные &, кроме уже валидных сущностей
_AMP_FIX = re.compile(rb'&(?!amp;|lt;|gt;|quot;|apos;|#\d+;)')

# Сломанный QUIK-шаблон: =""X""
_BROKEN_QUOTE = re.compile(rb'=""([^"\s/<>][^"]*)""(?=[\s/>])')


def _sanitize_tag(tag_bytes):
    """Чистит один тег: незаэкранированный & и QUIK-битые ""...""."""
    if b'&' in tag_bytes:
        tag_bytes = _AMP_FIX.sub(b'&amp;', tag_bytes)
    if b'=""' in tag_bytes:
        tag_bytes = _BROKEN_QUOTE.sub(rb'="&quot;\1&quot;"', tag_bytes)
    return tag_bytes


def _process_clean_data(raw, final):
    """
    Ускоренная очистка куска XML.
    Разбиение на теги через bytes.split — цикл в C, не в Python.
    Возвращает (cleaned, pending).
    """
    out = bytearray()
    pos = 0
    n = len(raw)

    while pos < n:
        lt = raw.find(b'<', pos)
        if lt == -1:
            out.extend(raw[pos:])
            return bytes(out), b""

        # текст до тега — не трогаем
        out.extend(raw[pos:lt])

        # комментарий
        if raw[lt:lt + 4] == b'<!--':
            end = raw.find(b'-->', lt)
            if end == -1:
                return bytes(out), raw[lt:]
            out.extend(raw[lt:end + 3])
            pos = end + 3
            continue

        # CDATA
        if raw[lt:lt + 9] == b'<![CDATA[':
            end = raw.find(b']]>', lt)
            if end == -1:
                return bytes(out), raw[lt:]
            out.extend(raw[lt:end + 3])
            pos = end + 3
            continue

        gt = raw.find(b'>', lt)
        if gt == -1:
            if final:
                out.extend(raw[lt:])
                return bytes(out), b""
            return bytes(out), raw[lt:]

        # ---- быстрая санитизация одного тега ----
        tag = raw[lt:gt + 1]

        # быстрая проверка, нужно ли чистить вообще
        if b'&' in tag or b'=""' in tag:
            tag = _sanitize_tag(tag)

        out.extend(tag)
        pos = gt + 1

    return bytes(out), b""


def _reader_thread(fi, raw_q, cancel_flag, clean_progress_cb, total_size):
    """Поток 1: читает исходный файл и кладёт чанки в raw_q."""
    try:
        total_read = 0
        while True:
            if cancel_flag and cancel_flag.is_set():
                break
            chunk = fi.read(CLEAN_CHUNK)
            if not chunk:
                break
            raw_q.put(chunk)
            total_read += len(chunk)
            if clean_progress_cb:
                clean_progress_cb(total_read, total_size)
    finally:
        raw_q.put(None)  # сигнал конца


def _cleaner_thread(raw_q, clean_q, cancel_flag):
    """Поток 2: чистит чанки из raw_q и кладёт в clean_q."""
    pending = b""
    try:
        while True:
            if cancel_flag and cancel_flag.is_set():
                break
            chunk = raw_q.get()
            if chunk is None:
                # конец — обработаем остаток
                if pending:
                    cleaned, _ = _process_clean_data(pending, final=True)
                    if cleaned:
                        clean_q.put(cleaned)
                break

            data = pending + chunk
            cleaned, pending = _process_clean_data(data, final=False)
            if cleaned:
                clean_q.put(cleaned)
    finally:
        clean_q.put(None)  # сигнал конца


def _writer_thread(fo, clean_q, cancel_flag):
    """Поток 3: пишет очищенные чанки в файл."""
    try:
        while True:
            chunk = clean_q.get()
            if chunk is None:
                break
            if cancel_flag and cancel_flag.is_set():
                break
            fo.write(chunk)
    finally:
        fo.flush()


def clean_xml_file(src_path, dst_path, progress_cb=None, cancel_flag=None):
    """
    Очистка XML в конвейере из трёх потоков.
    progress_cb получает (bytes_read, total_size).
    """
    total_size = os.path.getsize(src_path)

    raw_q = queue.Queue(maxsize=QUEUE_DEPTH)
    clean_q = queue.Queue(maxsize=QUEUE_DEPTH)

    def progress_wrap(read, total):
        if progress_cb:
            progress_cb(read, total)

    with open(src_path, "rb") as fi, open(dst_path, "wb") as fo:
        t_reader = threading.Thread(
            target=_reader_thread,
            args=(fi, raw_q, cancel_flag, progress_wrap, total_size),
            daemon=True)
        t_cleaner = threading.Thread(
            target=_cleaner_thread,
            args=(raw_q, clean_q, cancel_flag),
            daemon=True)
        t_writer = threading.Thread(
            target=_writer_thread,
            args=(fo, clean_q, cancel_flag),
            daemon=True)

        t_reader.start()
        t_cleaner.start()
        t_writer.start()

        t_writer.join()
        t_cleaner.join()
        t_reader.join()


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


def label_for(field_name, descriptions=None):
    if field_name in QUIK_FIELD_LABELS:
        return QUIK_FIELD_LABELS[field_name]
    base = base_field_name(field_name)
    if base in QUIK_FIELD_LABELS:
        return QUIK_FIELD_LABELS[base]
    if descriptions:
        if field_name in descriptions:
            return descriptions[field_name]
        if base in descriptions:
            return descriptions[base]
    return field_name


def format_quik_value(col, value):
    """Форматирование PreparedValue. Пустое -> «-»."""
    base = base_field_name(col)
    if value is None:
        return EMPTY_MARK
    s = str(value).strip()
    if not s:
        return EMPTY_MARK

    if len(s) == 10 and s[4] == "-" and s[7] == "-":
        y, m, d = s.split("-")
        if y.isdigit() and m.isdigit() and d.isdigit():
            return f"{d}.{m}.{y}"

    if base in ("TradeDate", "QuikDate", "Date", "SettleDate") \
            and len(s) == 8 and s.isdigit():
        return f"{s[6:8]}.{s[4:6]}.{s[0:4]}"

    if base in ("QuikTime", "ReplyTime", "Time"):
        return s

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
    try:
        with open(path, "rb") as f:
            head = f.read(8192)
        m = re.search(rb'xmlns\s*=\s*"([^"]+)"', head)
        if m:
            return m.group(1).decode("ascii", errors="ignore")
    except Exception:
        pass
    return None


def read_report_header(path):
    info = {"ProgramVersion": "", "StartDate": "", "EndDate": ""}
    try:
        with open(path, "rb") as f:
            head = f.read(8192)
        for key in ("ProgramVersion", "StartDate", "EndDate"):
            m = re.search(
                rb'%s\s*=\s*"([^"]*)"' % key.encode("ascii"),
                head
            )
            if m:
                info[key] = m.group(1).decode("utf-8", errors="replace")
    except Exception:
        pass
    return info


def format_date_string(s):
    if not s:
        return s
    if len(s) == 10 and s[4] == "-" and s[7] == "-":
        y, m, d = s.split("-")
        if y.isdigit() and m.isdigit() and d.isdigit():
            return f"{d}.{m}.{y}"
    if len(s) == 8 and s.isdigit():
        return f"{s[6:8]}.{s[4:6]}.{s[0:4]}"
    return s


# ============================================================
# 2. КОНФИГ ИЗВЛЕЧЕНИЯ
# ============================================================

class ExtractConfig:
    def __init__(self,
                 container_tag=CONTAINER_TAG,
                 field_tag=FIELD_TAG,
                 name_attr=NAME_ATTR,
                 description_attr=DESC_ATTR,
                 value_attr=VALUE_ATTR,
                 prepared_attr=PREPARED_ATTR):
        self.container_tag = container_tag
        self.field_tag = field_tag
        self.name_attr = name_attr
        self.description_attr = description_attr
        self.value_attr = value_attr
        self.prepared_attr = prepared_attr


# ============================================================
# 3. ИЗВЛЕЧЕНИЕ ЗАПИСЕЙ
# ============================================================

def extract_records(path, record_tag=RECORD_TAG, namespace=None,
                    config=None, descriptions=None,
                    progress_cb=None, cancel_flag=None, max_records=None):
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

            for k, v in elem.attrib.items():
                rec[f"@{strip_ns(k)}"] = v

            for child in elem:
                if not isinstance(child.tag, str):
                    continue
                tag_local = strip_ns(child.tag)

                if tag_local == config.container_tag:
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
                        prepared = (field.get(config.prepared_attr) or "").strip()

                        if not name:
                            continue

                        if desc and name not in descriptions:
                            descriptions[name] = desc

                        display = prepared

                        if name in rec:
                            i = 2
                            while f"{name}_{i}" in rec:
                                i += 1
                            rec[f"{name}_{i}"] = display
                        else:
                            rec[name] = display
                    continue

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


# ============================================================
# 4. КОЛОНКИ
# ============================================================

def inject_record_numbers(records):
    for i, rec in enumerate(records, 1):
        rec[RECORD_NUM_COL] = str(i)


def build_card_columns(records, descriptions=None):
    seen = {}
    for rec in records:
        for k in rec:
            if k == RECORD_NUM_COL:
                continue
            if k not in seen:
                seen[k] = True
    columns = list(seen.keys())
    header_map = {c: label_for(c, descriptions) for c in columns}
    return columns, header_map


def build_table_columns(card_columns):
    return [RECORD_NUM_COL] + list(card_columns)


# ============================================================
# 5. HTML
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


def export_html(clean_xml_path, out_path, record_tag=RECORD_TAG,
                namespace=None, config=None, limit=DEFAULT_LIMIT,
                progress_cb=None, cancel_flag=None, as_cards=False):
    header = read_report_header(clean_xml_path)

    descriptions = {}
    records = list(extract_records(
        clean_xml_path, record_tag, namespace,
        config=config, descriptions=descriptions,
        progress_cb=progress_cb, cancel_flag=cancel_flag,
        max_records=limit))

    inject_record_numbers(records)

    card_columns, card_header_map = build_card_columns(records, descriptions)
    table_columns = build_table_columns(card_columns)
    table_header_map = {RECORD_NUM_COL: "Запись"}
    table_header_map.update(card_header_map)

    meta_lines = [
        f"Дата формирования: {datetime.now().strftime('%d.%m.%Y %H:%M')}",
    ]
    if header.get("StartDate") or header.get("EndDate"):
        s = format_date_string(header.get("StartDate", ""))
        e = format_date_string(header.get("EndDate", ""))
        meta_lines.append(f"Период: {s} — {e}")
    if header.get("ProgramVersion"):
        meta_lines.append(f"Версия QUIK: {header['ProgramVersion']}")
    meta_lines.append(f"Источник: {os.path.basename(clean_xml_path)}")
    meta_lines.append(f"Записей: {len(records)}")

    # Буферизованная запись — крупный буфер в 1 МБ
    with open(out_path, "w", encoding="utf-8",
              buffering=1024 * 1024) as f:
        f.write(HTML_HEAD.format(
            title="Отчёт по транзакциям QUIK",
            meta="<br>".join(_esc(m) for m in meta_lines),
        ))

        if as_cards:
            for i, rec in enumerate(records, 1):
                f.write(f'<div class="card"><h3>Запись {i}</h3>\n')
                for col in card_columns:
                    if col not in rec:
                        continue
                    f.write(
                        f'<div class="field"><b>{_esc(card_header_map[col])}:</b> '
                        f'{_esc(format_quik_value(col, rec[col]))}</div>\n')
                f.write("</div>\n")
        else:
            f.write("<table><thead><tr>")
            for col in table_columns:
                f.write(f"<th>{_esc(table_header_map[col])}</th>")
            f.write("</tr></thead><tbody>\n")
            for rec in records:
                f.write("<tr>")
                for col in table_columns:
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
# 6. PDF
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


def export_pdf(clean_xml_path, out_path, record_tag=RECORD_TAG,
               namespace=None, config=None, limit=DEFAULT_LIMIT,
               progress_cb=None, cancel_flag=None, as_cards=False):
    font = register_cyrillic_font()

    styles = getSampleStyleSheet()

    header = read_report_header(clean_xml_path)

    descriptions = {}
    records = list(extract_records(
        clean_xml_path, record_tag, namespace,
        config=config, descriptions=descriptions,
        progress_cb=progress_cb, cancel_flag=cancel_flag,
        max_records=limit))

    inject_record_numbers(records)

    card_columns, card_header_map = build_card_columns(records, descriptions)
    table_columns = build_table_columns(card_columns)
    table_header_map = {RECORD_NUM_COL: "Запись"}
    table_header_map.update(card_header_map)

    ncols_table = len(table_columns)
    ncols_cards = len(card_columns)
    ncols = ncols_table if not as_cards else ncols_cards

    if as_cards or ncols <= 6:
        pagesize = A4
        body_font_size = 9
    elif ncols <= 10:
        pagesize = landscape(A3)
        body_font_size = 8
    elif ncols <= 16:
        pagesize = landscape(A2)
        body_font_size = 8
    elif ncols <= 24:
        pagesize = landscape(A1)
        body_font_size = 7
    else:
        pagesize = landscape(A0)
        body_font_size = 7

    h1 = ParagraphStyle("H1", parent=styles["Heading1"], fontName=font,
                        fontSize=16, leading=20)
    normal = ParagraphStyle("N", parent=styles["Normal"], fontName=font,
                            fontSize=body_font_size,
                            leading=body_font_size + 2)
    small = ParagraphStyle("S", parent=styles["Normal"], fontName=font,
                           fontSize=8, leading=10, textColor=colors.grey)

    doc = SimpleDocTemplate(
        out_path, pagesize=pagesize,
        leftMargin=1.0 * cm, rightMargin=1.0 * cm,
        topMargin=1.0 * cm, bottomMargin=1.0 * cm,
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
    meta_lines.append(f"Источник: {os.path.basename(clean_xml_path)}")
    meta_lines.append(f"Записей: {len(records)}")

    for line in meta_lines:
        story.append(Paragraph(_esc(line), small))
    story.append(Spacer(1, 8))

    if as_cards:
        for i, rec in enumerate(records, 1):
            story.append(Paragraph(f"<b>Запись {i}</b>", normal))
            for col in card_columns:
                if col not in rec:
                    continue
                story.append(Paragraph(
                    f"<b>{_esc(card_header_map[col])}:</b> "
                    f"{_esc(format_quik_value(col, rec[col]))}", normal))
            story.append(Spacer(1, 6))
        doc.build(story)
        return

    data = [[Paragraph(f"<b>{_esc(table_header_map[c])}</b>", normal)
             for c in table_columns]]
    for rec in records:
        data.append([
            Paragraph(_esc(format_quik_value(c, rec.get(c, ""))), normal)
            for c in table_columns
        ])

    avail = (pagesize[0] - 2.0 * cm) / max(len(table_columns), 1)
    col_widths = [avail] * len(table_columns)

    tbl = Table(data, colWidths=col_widths, repeatRows=1)
    tbl.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#f0f4f8")),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.grey),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 3),
        ("RIGHTPADDING", (0, 0), (-1, -1), 3),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
    ]))
    story.append(tbl)

    doc.build(story)


# ============================================================
# 7. GUI
# ============================================================

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("QUIK XML → отчёт (HTML / PDF)")
        self.geometry("720x420")
        self.resizable(False, False)

        self.xml_path = tk.StringVar()
        self.format_var = tk.StringVar(value="html")
        self.view_mode = tk.StringVar(value="cards")

        self.cancel_flag = threading.Event()

        self._build_ui()

    def _build_ui(self):
        pad = {"padx": 10, "pady": 8}

        frame_file = ttk.LabelFrame(self, text="1. XML-файл QUIK")
        frame_file.pack(fill="x", **pad)
        ttk.Entry(frame_file, textvariable=self.xml_path, width=70).pack(
            side="left", padx=6, pady=6, fill="x", expand=True)
        ttk.Button(frame_file, text="Обзор…", command=self.choose_file).pack(
            side="right", padx=6, pady=6)

        frame_view = ttk.LabelFrame(self, text="2. Представление")
        frame_view.pack(fill="x", **pad)
        ttk.Radiobutton(frame_view, text="Карточки",
                        value="cards", variable=self.view_mode).pack(
            side="left", padx=16, pady=8)
        ttk.Radiobutton(frame_view, text="Таблица",
                        value="table", variable=self.view_mode).pack(
            side="left", padx=16, pady=8)

        frame_fmt = ttk.LabelFrame(self, text="3. Формат вывода")
        frame_fmt.pack(fill="x", **pad)
        for fmt, label in [("html", "HTML"), ("pdf", "PDF")]:
            ttk.Radiobutton(frame_fmt, text=label, value=fmt,
                            variable=self.format_var).pack(
                side="left", padx=16, pady=8)

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
        self.status.pack(pady=6)

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

    def _clean_progress_cb(self, bytes_read, total_size):
        if total_size > 0:
            pct = bytes_read * 100 // total_size
            mb = bytes_read // (1024 * 1024)
            total_mb = total_size // (1024 * 1024)
            self._set_status(
                f"Очистка XML: {pct}% ({mb} / {total_mb} МБ)", "blue")

    def convert(self):
        xml_file = self.xml_path.get().strip()
        if not xml_file or not os.path.isfile(xml_file):
            messagebox.showwarning("Внимание", "Выберите существующий XML-файл")
            return

        fmt = self.format_var.get()
        ext = fmt
        as_cards = self.view_mode.get() == "cards"
        config = ExtractConfig()

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
            clean_path = None
            try:
                src_dir = os.path.dirname(os.path.abspath(xml_file))
                tmp_fd, clean_path = tempfile.mkstemp(
                    prefix="quik_clean_", suffix=".xml", dir=src_dir)
                os.close(tmp_fd)

                self._set_status("Очистка XML…", "blue")
                clean_xml_file(
                    xml_file, clean_path,
                    progress_cb=self._clean_progress_cb,
                    cancel_flag=self.cancel_flag)

                if self.cancel_flag.is_set():
                    raise RuntimeError("Отменено")

                namespace = detect_namespace(clean_path)
                self._set_status(
                    f"Namespace: {namespace or '—'}. Разбор транзакций…",
                    "blue")

                if fmt == "html":
                    export_html(clean_path, out_path, RECORD_TAG,
                                namespace=namespace, config=config,
                                limit=DEFAULT_LIMIT,
                                progress_cb=self._progress_cb,
                                cancel_flag=self.cancel_flag,
                                as_cards=as_cards)
                elif fmt == "pdf":
                    export_pdf(clean_path, out_path, RECORD_TAG,
                               namespace=namespace, config=config,
                               limit=DEFAULT_LIMIT,
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
                if clean_path and os.path.exists(clean_path):
                    try:
                        os.remove(clean_path)
                    except Exception:
                        pass

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