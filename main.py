import os
import re
import html
import queue
import shutil
import tempfile
import threading
import time
from datetime import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from lxml import etree as ET

from reportlab.lib.pagesizes import A4, landscape, A3, A2, A1, A0
from reportlab.lib.units import cm
from reportlab.pdfgen import canvas as rl_canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.lib.utils import simpleSplit


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
TRADE_DATE_ATTR = "@TradeDate"

CLEAN_CHUNK = 4 * 1024 * 1024
QUEUE_DEPTH = 4

FAST_LEN_THRESHOLD = 20

# Оставлять ли временные файлы после работы (для отладки)
KEEP_TEMP_FILES = False

RU_MONTHS = [
    "Январь", "Февраль", "Март", "Апрель", "Май", "Июнь",
    "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь",
]

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
# 0.1. ФОРМАТИРОВАНИЕ ВРЕМЕНИ И МЕСЯЦЕВ
# ============================================================

def fmt_duration(seconds):
    seconds = int(seconds)
    h = seconds // 3600
    m = (seconds % 3600) // 60
    s = seconds % 60
    parts = []
    if h:
        parts.append(f"{h} ч")
    if m or h:
        parts.append(f"{m} мин")
    parts.append(f"{s} с")
    return " ".join(parts)


def date_str_to_month(d):
    """'2019-06-04' или '20190604' → ('2019-06', '2019_Июнь')."""
    d = (d or "").strip()

    if len(d) == 10 and d[4] == "-" and d[7] == "-":
        try:
            y = int(d[0:4])
            m = int(d[5:7])
            if 1 <= m <= 12:
                return (f"{y:04d}-{m:02d}",
                        f"{y:04d}_{RU_MONTHS[m - 1]}")
        except ValueError:
            pass

    if len(d) == 8 and d.isdigit():
        try:
            y = int(d[0:4])
            m = int(d[4:6])
            if 1 <= m <= 12:
                return (f"{y:04d}-{m:02d}",
                        f"{y:04d}_{RU_MONTHS[m - 1]}")
        except ValueError:
            pass

    return ("0000-00", "0000_БезДаты")


def extract_month_key(rec):
    return date_str_to_month(rec.get(TRADE_DATE_ATTR, ""))


# ============================================================
# 0.2. САНИТИЗАЦИЯ XML
# ============================================================

_AMP_FIX = re.compile(rb'&(?!amp;|lt;|gt;|quot;|apos;|#\d+;)')
_BROKEN_QUOTE = re.compile(rb'=""([^"]+)""(?=[\s/>])')



def _sanitize_tag(tag_bytes):
    if b'&' in tag_bytes:
        tag_bytes = _AMP_FIX.sub(b'&amp;', tag_bytes)
    if b'=""' in tag_bytes:
        tag_bytes = _BROKEN_QUOTE.sub(rb'="&quot;\1&quot;"', tag_bytes)
    return tag_bytes


def _process_clean_data(raw, final):
    out = bytearray()
    pos = 0
    n = len(raw)

    while pos < n:
        lt = raw.find(b'<', pos)
        if lt == -1:
            out.extend(raw[pos:])
            return bytes(out), b""

        out.extend(raw[pos:lt])

        if raw[lt:lt + 4] == b'<!--':
            end = raw.find(b'-->', lt)
            if end == -1:
                return bytes(out), raw[lt:]
            out.extend(raw[lt:end + 3])
            pos = end + 3
            continue

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

        tag = raw[lt:gt + 1]
        if b'&' in tag or b'=""' in tag:
            tag = _sanitize_tag(tag)

        out.extend(tag)
        pos = gt + 1

    return bytes(out), b""


def _reader_thread(fi, raw_q, cancel_flag, clean_progress_cb, total_size):
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
        raw_q.put(None)


def _cleaner_thread(raw_q, clean_q, cancel_flag):
    pending = b""
    try:
        while True:
            if cancel_flag and cancel_flag.is_set():
                break
            chunk = raw_q.get()
            if chunk is None:
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
        clean_q.put(None)


def _writer_thread(fo, clean_q, cancel_flag):
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
# 0.3. РАЗБИВКА XML ПО МЕСЯЦАМ — ОДИН ПРОХОД
# ============================================================

_TRANS_START_RE = re.compile(rb'<Trans(?=[\s>])')
_TRADE_DATE_RE = re.compile(rb'\bTradeDate\s*=\s*"([^"]*)"')


def split_xml_by_month(clean_src_path, split_dir,
                       progress_cb=None, cancel_flag=None):
    """
    Читает clean_src_path (уже очищенный) и режет его на файлы
    по месяцам — по границам </Trans>.

    Возвращает dict: {ym: {'label': ..., 'path': ..., 'count': N}}
    """
    os.makedirs(split_dir, exist_ok=True)

    with open(clean_src_path, "rb") as f:
        head = f.read(16384)

    m = re.search(rb'<TransactionsReport\b[^>]*>', head, re.DOTALL)
    if not m:
        raise RuntimeError(
            "Не найден корневой тег <TransactionsReport> в первых "
            "16 КБ. Возможно, файл не QUIK-XML.")
    root_open_end = m.end()
    preamble = head[:root_open_end]
    closing_tag = b'\n</TransactionsReport>\n'

    open_files = {}

    def get_writer(ym, label):
        entry = open_files.get(ym)
        if entry is None:
            path = os.path.join(split_dir, f"{ym}.xml")
            fd = open(path, "wb", buffering=1024 * 1024)
            fd.write(preamble)
            fd.write(b"\n")
            entry = {"fd": fd, "path": path, "label": label,
                     "count": 0}
            open_files[ym] = entry
        return entry

    total_read = 0
    file_size = os.path.getsize(clean_src_path)

    buf = bytearray()
    in_trans = False
    current_ym = None
    current_label = None
    date_parsed = False

    with open(clean_src_path, "rb") as f:
        f.seek(root_open_end)

        while True:
            if cancel_flag and cancel_flag.is_set():
                break

            chunk = f.read(CLEAN_CHUNK)
            if not chunk:
                break

            total_read += len(chunk)
            if progress_cb:
                progress_cb(total_read, file_size)

            data = chunk
            pos = 0

            while pos < len(data):
                if not in_trans:
                    m = _TRANS_START_RE.search(data, pos)
                    if not m:
                        pos = len(data)
                        break

                    in_trans = True
                    date_parsed = False
                    current_ym = None
                    current_label = None
                    buf = bytearray()
                    pos = m.start()

                if not date_parsed:
                    gt = data.find(b'>', pos)
                    if gt != -1:
                        tag_bytes = bytes(buf) + data[pos:gt + 1]
                        m2 = _TRADE_DATE_RE.search(tag_bytes)
                        if m2:
                            d = m2.group(1).decode("ascii",
                                                   errors="ignore")
                            current_ym, current_label = \
                                date_str_to_month(d)
                        else:
                            current_ym, current_label = (
                                "0000-00", "0000_БезДаты")
                        date_parsed = True

                end = data.find(b'</Trans>', pos)
                if end == -1:
                    buf.extend(data[pos:])
                    pos = len(data)
                    break

                end_pos = end + len(b'</Trans>')
                buf.extend(data[pos:end_pos])

                if current_ym is None:
                    current_ym, current_label = (
                        "0000-00", "0000_БезДаты")

                entry = get_writer(current_ym, current_label)
                entry["fd"].write(bytes(buf))
                entry["fd"].write(b"\n")
                entry["count"] += 1

                buf = bytearray()
                in_trans = False
                date_parsed = False
                current_ym = None
                current_label = None
                pos = end_pos

    result = {}
    for ym, entry in open_files.items():
        entry["fd"].write(closing_tag)
        entry["fd"].close()
        result[ym] = {
            "label": entry["label"],
            "path": entry["path"],
            "count": entry["count"],
        }

    return result


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

def build_columns_from_seen(seen, descriptions=None):
    columns = list(seen.keys())
    header_map = {c: label_for(c, descriptions) for c in columns}
    return columns, header_map


def build_table_columns(card_columns):
    return [RECORD_NUM_COL] + list(card_columns)


# ============================================================
# 5. HTML (один месяц)
# ============================================================

def _esc(s):
    return html.escape(str(s))


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


def export_html_month(out_path, month_label, records_iter,
                      card_columns, card_header_map,
                      header, started_at, finished_at,
                      as_cards, progress_cb=None, cancel_flag=None):
    table_columns = build_table_columns(card_columns)
    table_header_map = {RECORD_NUM_COL: "Запись"}
    table_header_map.update(card_header_map)

    now = datetime.now()
    started_at = started_at or now
    finished_at = finished_at or now
    duration_sec = (finished_at - started_at).total_seconds()

    title = f"Отчёт по транзакциям QUIK — {month_label}"
    count = 0

    meta_lines = [
        f"Дата формирования: {now.strftime('%d.%m.%Y %H:%M:%S')}",
        f"Начало обработки: {started_at.strftime('%d.%m.%Y %H:%M:%S')}",
        f"Конец обработки: {finished_at.strftime('%d.%m.%Y %H:%M:%S')}",
        f"Затрачено времени: {fmt_duration(duration_sec)}",
        f"Месяц: {month_label}",
    ]
    if header.get("StartDate") or header.get("EndDate"):
        s = format_date_string(header.get("StartDate", ""))
        e = format_date_string(header.get("EndDate", ""))
        meta_lines.append(f"Период (из XML): {s} — {e}")
    if header.get("ProgramVersion"):
        meta_lines.append(f"Версия QUIK: {header['ProgramVersion']}")

    with open(out_path, "w", encoding="utf-8",
              buffering=1024 * 1024) as f:
        f.write(HTML_HEAD.format(
            title=_esc(title),
            meta="<br>".join(_esc(m) for m in meta_lines),
        ))

        if as_cards:
            for rec in records_iter:
                if cancel_flag and cancel_flag.is_set():
                    break
                count += 1
                f.write(f'<div class="card"><h3>Запись {count}</h3>\n')
                for col in card_columns:
                    if col not in rec:
                        continue
                    f.write(
                        f'<div class="field"><b>{_esc(card_header_map[col])}:</b> '
                        f'{_esc(format_quik_value(col, rec[col]))}</div>\n')
                f.write("</div>\n")
                if progress_cb and count % 5000 == 0:
                    progress_cb(count)
        else:
            f.write("<table><thead><tr>")
            for col in table_columns:
                f.write(f"<th>{_esc(table_header_map[col])}</th>")
            f.write("</tr></thead><tbody>\n")
            for rec in records_iter:
                if cancel_flag and cancel_flag.is_set():
                    break
                count += 1
                f.write("<tr>")
                f.write(f'<td>{count}</td>')
                for col in card_columns:
                    cls = ("num"
                           if base_field_name(col) in QUIK_NUMERIC_FIELDS
                           else "")
                    f.write(
                        f'<td class="{cls}">'
                        f'{_esc(format_quik_value(col, rec.get(col, "")))}</td>')
                f.write("</tr>\n")
                if progress_cb and count % 5000 == 0:
                    progress_cb(count)
            f.write("</tbody></table>\n")

        f.write("</body></html>\n")

    return count


# ============================================================
# 6. PDF (один месяц) через Canvas
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


def _wrap_text(text, font_name, font_size, max_width):
    if not text:
        return [""]
    text = str(text)
    if not text:
        return [""]

    if len(text) <= FAST_LEN_THRESHOLD and text.isascii():
        return [text]

    if stringWidth(text, font_name, font_size) <= max_width:
        return [text]

    lines = simpleSplit(text, font_name, font_size, max_width)

    result = []
    for line in lines:
        if stringWidth(line, font_name, font_size) <= max_width:
            result.append(line)
            continue

        chunk = ""
        for ch in line:
            test = chunk + ch
            if stringWidth(test, font_name, font_size) <= max_width:
                chunk = test
            else:
                if chunk:
                    result.append(chunk)
                chunk = ch
        if chunk:
            result.append(chunk)

    return result or [""]


def _calc_row_height(row_values, table_columns, col_widths,
                     font_name, font_size, line_h):
    wrapped = {}
    max_lines = 1

    for col, w in zip(table_columns, col_widths):
        val = row_values.get(col, "")
        lines = _wrap_text(val, font_name, font_size, w - 4)
        wrapped[col] = lines
        if len(lines) > max_lines:
            max_lines = len(lines)

    return max_lines * line_h, wrapped


def export_pdf_month(out_path, month_label, records_iter,
                     card_columns, card_header_map,
                     header, started_at, finished_at,
                     as_cards, progress_cb=None, cancel_flag=None):
    font = register_cyrillic_font()

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

    page_w, page_h = pagesize
    margin = 1.0 * cm
    usable_w = page_w - 2 * margin

    if ncols_table > 0:
        record_col_w = min(1.4 * cm, usable_w * 0.06)
        other_cols = max(ncols_table - 1, 1)
        other_col_w = (usable_w - record_col_w) / other_cols
        col_widths = [record_col_w] + [other_col_w] * other_cols
    else:
        col_widths = []

    header_font_size = body_font_size
    meta_line_h = 10
    line_h = body_font_size * 1.25
    row_pad = 2

    c = rl_canvas.Canvas(out_path, pagesize=pagesize)
    title_text = f"Отчёт по транзакциям QUIK — {month_label}"
    c.setTitle(title_text)

    now = datetime.now()
    started_at = started_at or now
    finished_at = finished_at or now
    duration_sec = (finished_at - started_at).total_seconds()

    def draw_report_header(is_first_page):
        y = page_h - margin

        c.setFont(font, 16)
        c.drawString(margin, y - 16, title_text)
        y -= 26

        if is_first_page:
            c.setFont(font, 8)
            meta_lines = [
                f"Дата формирования: "
                f"{now.strftime('%d.%m.%Y %H:%M:%S')}",
                f"Начало обработки: "
                f"{started_at.strftime('%d.%m.%Y %H:%M:%S')}",
                f"Конец обработки: "
                f"{finished_at.strftime('%d.%m.%Y %H:%M:%S')}",
                f"Затрачено времени: {fmt_duration(duration_sec)}",
                f"Месяц: {month_label}",
            ]
            if header.get("StartDate") or header.get("EndDate"):
                s = format_date_string(header.get("StartDate", ""))
                e = format_date_string(header.get("EndDate", ""))
                meta_lines.append(f"Период (из XML): {s} — {e}")
            if header.get("ProgramVersion"):
                meta_lines.append(
                    f"Версия QUIK: {header['ProgramVersion']}")

            for line in meta_lines:
                c.drawString(margin, y, line)
                y -= meta_line_h
            y -= 4
        else:
            y -= 4

        return y

    def draw_table_header(y):
        header_line_h = header_font_size * 1.25

        wrapped_headers = {}
        max_lines = 1
        for col, w in zip(table_columns, col_widths):
            text = table_header_map.get(col, col)
            lines = _wrap_text(text, font, header_font_size, w - 4)
            wrapped_headers[col] = lines
            if len(lines) > max_lines:
                max_lines = len(lines)

        header_h = max_lines * header_line_h + 4

        c.setFont(font, header_font_size)
        x = margin
        for col, w in zip(table_columns, col_widths):
            for j, ln in enumerate(wrapped_headers[col]):
                y_text = y - header_font_size - j * header_line_h
                c.drawString(x + 2, y_text, ln)
            x += w

        y_line = y - header_h + 2
        c.setLineWidth(0.6)
        c.setStrokeColorRGB(0.3, 0.3, 0.3)
        c.line(margin, y_line, margin + usable_w, y_line)
        c.setStrokeColorRGB(0, 0, 0)
        c.setLineWidth(0.25)

        return y - header_h

    count = 0

    if as_cards:
        page_y = draw_report_header(True)
        card_line_h = body_font_size * 1.3

        for rec in records_iter:
            if cancel_flag and cancel_flag.is_set():
                break
            count += 1

            card_lines = []
            card_lines.append((True, f"Запись {count}"))
            for col in card_columns:
                if col not in rec:
                    continue
                label = card_header_map.get(col, col)
                val = format_quik_value(col, rec[col])
                line = f"{label}: {val}"
                for ln in _wrap_text(line, font, body_font_size,
                                     usable_w - 6):
                    card_lines.append((False, ln))

            card_h = len(card_lines) * card_line_h + 8

            if page_y - card_h < margin:
                c.showPage()
                page_y = draw_report_header(False)

            for is_title, ln in card_lines:
                if is_title:
                    c.setFont(font, body_font_size + 1)
                    c.drawString(margin,
                                 page_y - body_font_size - 1, ln)
                    page_y -= card_line_h
                else:
                    c.setFont(font, body_font_size)
                    c.drawString(margin + 6,
                                 page_y - body_font_size, ln)
                    page_y -= card_line_h

            page_y -= 4

            if progress_cb and count % 5000 == 0:
                progress_cb(count)

        c.showPage()
        c.save()
        return count

    page_y = draw_report_header(True)
    page_y = draw_table_header(page_y)

    for rec in records_iter:
        if cancel_flag and cancel_flag.is_set():
            break
        count += 1

        row_vals = {RECORD_NUM_COL: str(count)}
        for col in card_columns:
            row_vals[col] = format_quik_value(col, rec.get(col, ""))

        row_h, wrapped = _calc_row_height(
            row_vals, table_columns, col_widths,
            font, body_font_size, line_h)
        row_h += row_pad

        if page_y - row_h < margin:
            c.showPage()
            page_y = draw_report_header(False)
            page_y = draw_table_header(page_y)

        if count % 2 == 0:
            c.setFillColorRGB(0.97, 0.97, 0.97)
            c.rect(margin, page_y - row_h, usable_w, row_h,
                   stroke=0, fill=1)
            c.setFillColorRGB(0, 0, 0)

        c.setFont(font, body_font_size)
        x = margin
        for col, w in zip(table_columns, col_widths):
            lines = wrapped.get(col, [""])
            for j, ln in enumerate(lines):
                y_text = page_y - body_font_size - j * line_h
                c.drawString(x + 2, y_text, ln)
            x += w

        c.setStrokeColorRGB(0.85, 0.85, 0.85)
        c.setLineWidth(0.2)
        c.line(margin, page_y - row_h,
               margin + usable_w, page_y - row_h)
        c.setStrokeColorRGB(0, 0, 0)
        c.setLineWidth(0.25)

        page_y -= row_h

        if progress_cb and count % 5000 == 0:
            progress_cb(count)

    c.showPage()
    c.save()
    return count


# ============================================================
# 7. GUI
# ============================================================

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("QUIK XML → отчёты по месяцам (HTML / PDF)")
        self.geometry("860x640")
        self.resizable(False, False)

        self.xml_path = tk.StringVar()
        self.format_var = tk.StringVar(value="html")
        self.view_mode = tk.StringVar(value="cards")

        self.cancel_flag = threading.Event()
        self._start_time = None
        self._end_time = None

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

        # ─── Блок времени и целостности ───
        frame_time = ttk.LabelFrame(self, text="Время и целостность")
        frame_time.pack(fill="x", **pad)

        # Левая колонка — время
        time_left = ttk.Frame(frame_time)
        time_left.grid(row=0, column=0, sticky="nw", padx=6, pady=3)

        ttk.Label(time_left, text="Начало:").grid(
            row=0, column=0, sticky="w", padx=4, pady=2)
        self.lbl_start = ttk.Label(time_left, text="—")
        self.lbl_start.grid(row=0, column=1, sticky="w", padx=4, pady=2)

        ttk.Label(time_left, text="Конец:").grid(
            row=1, column=0, sticky="w", padx=4, pady=2)
        self.lbl_end = ttk.Label(time_left, text="—")
        self.lbl_end.grid(row=1, column=1, sticky="w", padx=4, pady=2)

        ttk.Label(time_left, text="Прошло:").grid(
            row=2, column=0, sticky="w", padx=4, pady=2)
        self.lbl_elapsed = ttk.Label(time_left, text="—")
        self.lbl_elapsed.grid(row=2, column=1, sticky="w", padx=4, pady=2)

        # Разделитель
        ttk.Separator(frame_time, orient="vertical").grid(
            row=0, column=1, sticky="ns", padx=12, pady=4)

        # Правая колонка — целостность
        time_right = ttk.Frame(frame_time)
        time_right.grid(row=0, column=2, sticky="nw", padx=6, pady=3)

        ttk.Label(time_right, text="Всего записей в XML:").grid(
            row=0, column=0, sticky="w", padx=4, pady=2)
        self.lbl_total_xml = ttk.Label(
            time_right, text="—",
            font=("Segoe UI", 10, "bold"))
        self.lbl_total_xml.grid(row=0, column=1, sticky="w",
                                padx=4, pady=2)

        ttk.Label(time_right, text="Записей в отчётах:").grid(
            row=1, column=0, sticky="w", padx=4, pady=2)
        self.lbl_total_reports = ttk.Label(
            time_right, text="—",
            font=("Segoe UI", 10, "bold"))
        self.lbl_total_reports.grid(row=1, column=1, sticky="w",
                                    padx=4, pady=2)

        ttk.Label(time_right, text="Проверка:").grid(
            row=2, column=0, sticky="w", padx=4, pady=2)
        self.lbl_integrity = ttk.Label(
            time_right, text="—",
            font=("Segoe UI", 10, "bold"))
        self.lbl_integrity.grid(row=2, column=1, sticky="w",
                                padx=4, pady=2)

        frame_prog = ttk.LabelFrame(self, text="Прогресс")
        frame_prog.pack(fill="x", **pad)
        self.progress = ttk.Progressbar(frame_prog, mode="indeterminate")
        self.progress.pack(fill="x", padx=6, pady=8)

        frame_btn = ttk.Frame(self)
        frame_btn.pack(fill="x", **pad)
        self.btn_convert = ttk.Button(frame_btn, text="Сформировать отчёты",
                                      command=self.convert)
        self.btn_convert.pack(side="left", padx=6)
        self.btn_cancel = ttk.Button(frame_btn, text="Отмена",
                                     command=self.cancel, state="disabled")
        self.btn_cancel.pack(side="left", padx=6)

        self.status = ttk.Label(self, text="Готов к работе", foreground="gray")
        self.status.pack(pady=6)

        self._tick()

    def _tick(self):
        if self._start_time is not None and self._end_time is None:
            elapsed = time.time() - self._start_time
            self.lbl_elapsed.config(text=fmt_duration(elapsed))
        self.after(500, self._tick)

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

    def _split_progress_cb(self, bytes_read, total_size):
        if total_size > 0:
            pct = bytes_read * 100 // total_size
            mb = bytes_read // (1024 * 1024)
            total_mb = total_size // (1024 * 1024)
            self._set_status(
                f"Резка по месяцам: {pct}% ({mb} / {total_mb} МБ)",
                "blue")

    def _reset_integrity_labels(self):
        self.lbl_total_xml.config(text="—", foreground="black")
        self.lbl_total_reports.config(text="—", foreground="black")
        self.lbl_integrity.config(text="—", foreground="gray")

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
            title="Куда сохранить отчёты (будет создана папка)",
            defaultextension=f".{ext}",
            filetypes=[(f"{ext.upper()} files", f"*.{ext}")],
            initialfile=os.path.splitext(os.path.basename(xml_file))[0]
                        + f"_report.{ext}")
        if not out_path:
            return

        self._start_time = time.time()
        self._end_time = None

        started_at = datetime.now()
        self.lbl_start.config(text=started_at.strftime('%d.%m.%Y %H:%M:%S'))
        self.lbl_end.config(text="—")
        self.lbl_elapsed.config(text="—")
        self._reset_integrity_labels()

        self.cancel_flag.clear()
        self.btn_convert.config(state="disabled")
        self.btn_cancel.config(state="normal")
        self.progress.start(10)
        self._set_status("Формирование отчётов…", "blue")

        def worker():
            work_dir = None
            try:
                src_dir = os.path.dirname(os.path.abspath(xml_file))
                work_dir = tempfile.mkdtemp(
                    prefix="quik_work_", dir=src_dir)

                clean_path = os.path.join(work_dir, "clean.xml")
                split_dir = os.path.join(work_dir, "months")
                os.makedirs(split_dir, exist_ok=True)

                # ─── ЭТАП 1: очистка XML ───
                self._set_status("Этап 1/3: очистка XML…", "blue")
                clean_xml_file(
                    xml_file, clean_path,
                    progress_cb=self._clean_progress_cb,
                    cancel_flag=self.cancel_flag)

                if self.cancel_flag.is_set():
                    raise RuntimeError("Отменено")

                # ─── ЭТАП 2: резка по месяцам ───
                self._set_status(
                    "Этап 2/3: резка XML по месяцам…", "blue")

                months_info = split_xml_by_month(
                    clean_path, split_dir,
                    progress_cb=self._split_progress_cb,
                    cancel_flag=self.cancel_flag)

                if self.cancel_flag.is_set():
                    raise RuntimeError("Отменено")

                if not months_info:
                    raise RuntimeError(
                        "В файле не найдено ни одной транзакции.")

                # Сумма записей по месяцам = сколько записей в исходном XML
                total_in_xml = sum(info["count"]
                                   for info in months_info.values())

                self.after(0, lambda: self.lbl_total_xml.config(
                    text=f"{total_in_xml:,}".replace(",", " "),
                    foreground="black"))

                if not KEEP_TEMP_FILES:
                    try:
                        os.remove(clean_path)
                    except Exception:
                        pass

                first_month_path = list(months_info.values())[0]["path"]
                namespace = detect_namespace(first_month_path)
                header = read_report_header(first_month_path)

                sample_seen = {}
                sample_desc = {}
                for i, rec in enumerate(extract_records(
                        first_month_path, RECORD_TAG, namespace,
                        config=config, descriptions=sample_desc,
                        progress_cb=None, cancel_flag=self.cancel_flag,
                        max_records=200)):
                    for k in rec:
                        if k == RECORD_NUM_COL:
                            continue
                        if k not in sample_seen:
                            sample_seen[k] = True

                card_columns, card_header_map = build_columns_from_seen(
                    sample_seen, sample_desc)

                # ─── ЭТАП 3: генерация отчётов ───
                out_dir_base = os.path.dirname(os.path.abspath(out_path))
                base_name = os.path.splitext(
                    os.path.basename(out_path))[0]
                reports_dir = os.path.join(out_dir_base,
                                           f"{base_name}_reports")
                os.makedirs(reports_dir, exist_ok=True)

                months_sorted = sorted(months_info.keys())
                K = len(months_sorted)
                written = []

                for i, ym in enumerate(months_sorted, 1):
                    if self.cancel_flag.is_set():
                        break

                    info = months_info[ym]
                    month_label = info["label"]
                    month_path = info["path"]
                    file_name = f"{base_name}_{month_label}.{ext}"
                    file_path = os.path.join(reports_dir, file_name)

                    self._set_status(
                        f"Этап 3/3: {i}/{K} — {month_label} "
                        f"({info['count']} записей)…", "blue")

                    records_iter = extract_records(
                        month_path, RECORD_TAG, namespace,
                        config=config, descriptions=sample_desc,
                        progress_cb=self._progress_cb,
                        cancel_flag=self.cancel_flag,
                        max_records=DEFAULT_LIMIT)

                    finished_at = datetime.now()

                    if fmt == "html":
                        count = export_html_month(
                            file_path, month_label, records_iter,
                            card_columns, card_header_map,
                            header, started_at, finished_at,
                            as_cards=as_cards,
                            progress_cb=self._progress_cb,
                            cancel_flag=self.cancel_flag)
                    else:
                        count = export_pdf_month(
                            file_path, month_label, records_iter,
                            card_columns, card_header_map,
                            header, started_at, finished_at,
                            as_cards=as_cards,
                            progress_cb=self._progress_cb,
                            cancel_flag=self.cancel_flag)

                    written.append((file_path, count))

                # Сумма записей в отчётах
                total_in_reports = sum(c for _, c in written)

                def update_integrity():
                    self.lbl_total_reports.config(
                        text=f"{total_in_reports:,}".replace(",", " "),
                        foreground="black")
                    if total_in_reports == total_in_xml:
                        self.lbl_integrity.config(
                            text=f"OK — {total_in_xml:,}".replace(",", " "),
                            foreground="green")
                    else:
                        diff = total_in_xml - total_in_reports
                        self.lbl_integrity.config(
                            text=f"РАСХОЖДЕНИЕ: {diff:+,}"
                                 .replace(",", " "),
                            foreground="red")

                self.after(0, update_integrity)

                self._end_time = time.time()
                end_dt = datetime.now()
                total_time = self._end_time - self._start_time

                def apply_end():
                    self.lbl_end.config(
                        text=end_dt.strftime('%d.%m.%Y %H:%M:%S'))
                    self.lbl_elapsed.config(text=fmt_duration(total_time))

                self.after(0, apply_end)

                if not written:
                    raise RuntimeError("Отчёты не сформированы (отменено)")

                msg_lines = [
                    f"Всего записей в XML: {total_in_xml}",
                    f"Записей в отчётах: {total_in_reports}",
                ]
                if total_in_reports == total_in_xml:
                    msg_lines.append("Целостность: OK")
                else:
                    msg_lines.append(
                        f"Целостность: РАСХОЖДЕНИЕ "
                        f"({total_in_xml - total_in_reports:+d})")
                msg_lines.append("")
                msg_lines.append(f"Месяцев: {len(written)}")
                msg_lines.append(f"Папка: {reports_dir}")
                msg_lines.append("")
                for p, c in written[:15]:
                    msg_lines.append(
                        f"  • {os.path.basename(p)} — {c}")
                if len(written) > 15:
                    msg_lines.append(
                        f"  … и ещё {len(written) - 15}")
                msg_lines.append("")
                msg_lines.append(f"Всего: {fmt_duration(total_time)}")

                self.after(0, self._show_success, "\n".join(msg_lines))

            except Exception as exc:
                err = f"{type(exc).__name__}: {exc}"
                self._end_time = time.time()
                end_dt = datetime.now()
                total_time = self._end_time - self._start_time

                def apply_end_err():
                    self.lbl_end.config(
                        text=end_dt.strftime('%d.%m.%Y %H:%M:%S'))
                    self.lbl_elapsed.config(text=fmt_duration(total_time))

                self.after(0, apply_end_err)
                self.after(0, self._show_error, "Ошибка", err)

            finally:
                if work_dir and os.path.exists(work_dir):
                    if KEEP_TEMP_FILES:
                        print(f"[DEBUG] Временные файлы в: {work_dir}")
                    else:
                        try:
                            shutil.rmtree(work_dir, ignore_errors=True)
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