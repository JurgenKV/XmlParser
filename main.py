import os
import gc
import re
import html
import queue
import shutil
import tempfile
import threading
import time
import traceback
from datetime import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from lxml import etree as ET

from reportlab import rl_config
rl_config.pageCompression = 1   # сжимать страницы в PDF — меньше RAM при save()

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

# 1 МБ вместо 4 МБ: при 3-поточной чистке пик памяти
# (raw_q + clean_q + текущий chunk + pending) падает с ~48 МБ до ~12 МБ.
# На скорость почти не влияет: 1 МБ всё равно >> буфера ОС.
CLEAN_CHUNK = 1 * 1024 * 1024
QUEUE_DEPTH = 4

FAST_LEN_THRESHOLD = 20

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
# 0.0. ИЗМЕРЕНИЕ ПАМЯТИ
# ============================================================

def rss_mb():
    """Возвращает текущий рабочий set процесса в МБ.

    Windows — через psapi.GetProcessMemoryInfo.
    Linux — ru_maxrss в KB, macOS — в байтах.
    При любой ошибке возвращает -1.0.
    """
    try:
        import sys as _sys
        if _sys.platform.startswith("win"):
            import ctypes
            from ctypes import wintypes

            class _PMC(ctypes.Structure):
                _fields_ = [
                    ("cb", wintypes.DWORD),
                    ("PageFaultCount", wintypes.DWORD),
                    ("PeakWorkingSetSize", ctypes.c_size_t),
                    ("WorkingSetSize", ctypes.c_size_t),
                    ("QuotaPeakPagedPoolUsage", ctypes.c_size_t),
                    ("QuotaPagedPoolUsage", ctypes.c_size_t),
                    ("QuotaPeakNonPagedPoolUsage", ctypes.c_size_t),
                    ("QuotaNonPagedPoolUsage", ctypes.c_size_t),
                    ("PagefileUsage", ctypes.c_size_t),
                    ("PeakPagefileUsage", ctypes.c_size_t),
                ]

            pmc = _PMC()
            pmc.cb = ctypes.sizeof(pmc)
            ctypes.windll.psapi.GetProcessMemoryInfo(
                ctypes.windll.kernel32.GetCurrentProcess(),
                ctypes.byref(pmc), pmc.cb)
            return pmc.WorkingSetSize / (1024.0 * 1024.0)

        import resource
        rss = resource.getrusage(resource.RUSAGE_SELF).ru_maxrss
        if _sys.platform == "darwin":
            return rss / (1024.0 * 1024.0)
        return rss / 1024.0
    except Exception:
        return -1.0


# ============================================================
# 0.1. ЛОГЕР
# ============================================================

class Logger:
    def __init__(self, path):
        self.path = path
        self._lock = threading.Lock()
        # 64 КБ буфер вместо построчного (buffering=1):
        # меньше syscall'ов, при kрэше всё равно сбрасывается
        # через Logger.close() в finally.
        self._f = open(path, "w", encoding="utf-8",
                       buffering=64 * 1024)

    def log(self, msg):
        ts = datetime.now().strftime("%H:%M:%S.%f")[:-3]
        line = f"[{ts}] {msg}"
        with self._lock:
            try:
                self._f.write(line + "\n")
            except Exception:
                pass
            try:
                print(line)
            except Exception:
                pass

    def close(self):
        try:
            self._f.close()
        except Exception:
            pass


# ============================================================
# 0.2. ФОРМАТИРОВАНИЕ ВРЕМЕНИ И МЕСЯЦЕВ
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
# 0.3. САНИТИЗАЦИЯ XML
# ============================================================

_AMP_FIX = re.compile(rb'&(?!amp;|lt;|gt;|quot;|apos;|#\d+;)')

# =""X""  — лишняя кавычка в НАЧАЛЕ значения
_BROKEN_QUOTE = re.compile(rb'=""([^"]+)""(?=[\s/>])')

# ="...=""  — лишняя кавычка в КОНЦЕ значения
# Пример: ClientCode="N=""   ->  ClientCode="N=&quot;"
_EXTRA_TRAILING_QUOTE = re.compile(rb'(="[^"]*?)""(?=[\s/>])')

_INVALID_XML_BYTES = bytes(
    b for b in range(0x20)
    if b not in (0x09, 0x0A, 0x0D)
)
_INVALID_BYTE_RE = re.compile(
    rb'[' + re.escape(_INVALID_XML_BYTES) + rb']'
)

# ─── Умные кавычки и тире cp1251 — ПРОСТО УДАЛЯЕМ ───
# 0x91 ‘  0x92 ’  0x93 “  0x94 ”  0x96 –  0x97 —
# Эти символы ломают границы атрибутов в QUIK-отчётах
# (парсер видит внутри значения обычную кавычку, хотя
# это другой байт), поэтому безопаснее от них избавиться.
_SMART_CHARS_MAP = {
    0x91: b"",
    0x92: b"",
    0x93: b"",
    0x94: b"",
    0x96: b"",
    0x97: b"",
}
_SMART_CHARS_RE = re.compile(
    b'[' + bytes(_SMART_CHARS_MAP.keys()) + b']'
)


def _smart_replace(match):
    b = match.group(0)[0]
    return _SMART_CHARS_MAP.get(b, b'')


def _fix_smart_quotes(data: bytes) -> bytes:
    if not data:
        return data
    if _SMART_CHARS_RE.search(data) is None:
        return data
    return _SMART_CHARS_RE.sub(_smart_replace, data)


def _sanitize_tag(tag_bytes):
    # 0) умные кавычки/тире — просто удаляем
    tag_bytes = _fix_smart_quotes(tag_bytes)

    # 1) невалидные управляющие байты — удаляем
    if _INVALID_BYTE_RE.search(tag_bytes):
        tag_bytes = _INVALID_BYTE_RE.sub(b'', tag_bytes)

    # 2) & -> &amp;
    if b'&' in tag_bytes:
        tag_bytes = _AMP_FIX.sub(b'&amp;', tag_bytes)

    # 3) =""X""  ->  ="&quot;X&quot;"
    if b'=""' in tag_bytes:
        tag_bytes = _BROKEN_QUOTE.sub(
            rb'="&quot;\1&quot;"', tag_bytes)

    # 3б) ="...=""  ->  ="...=&quot;"
    # Лишняя хвостовая кавычка внутри значения атрибута:
    # ClientCode="N=""   ->   ClientCode="N=&quot;"
    if b'""' in tag_bytes:
        tag_bytes = _EXTRA_TRAILING_QUOTE.sub(
            rb'\1&quot;"', tag_bytes)

    return tag_bytes


def _process_clean_data(raw, final):
    # Возвращаем bytearray напрямую: fo.write() принимает
    # bytes-like объекты. Раньше был bytes(out) — лишняя копия
    # всего 4-МБ чанка на каждый вызов.
    out = bytearray()
    pos = 0
    n = len(raw)

    while pos < n:
        lt = raw.find(b'<', pos)
        if lt == -1:
            tail = raw[pos:]
            tail = _fix_smart_quotes(tail)
            if _INVALID_BYTE_RE.search(tail):
                tail = _INVALID_BYTE_RE.sub(b'', tail)
            out.extend(tail)
            return out, b""

        text_chunk = raw[pos:lt]
        text_chunk = _fix_smart_quotes(text_chunk)
        if _INVALID_BYTE_RE.search(text_chunk):
            text_chunk = _INVALID_BYTE_RE.sub(b'', text_chunk)
        out.extend(text_chunk)

        if raw[lt:lt + 4] == b'<!--':
            end = raw.find(b'-->', lt)
            if end == -1:
                return out, raw[lt:]
            out.extend(raw[lt:end + 3])
            pos = end + 3
            continue

        if raw[lt:lt + 9] == b'<![CDATA[':
            end = raw.find(b']]>', lt)
            if end == -1:
                return out, raw[lt:]
            out.extend(raw[lt:end + 3])
            pos = end + 3
            continue

        gt = raw.find(b'>', lt)
        if gt == -1:
            if final:
                tag = raw[lt:]
                tag = _sanitize_tag(tag)
                out.extend(tag)
                return out, b""
            return out, raw[lt:]

        tag = raw[lt:gt + 1]

        # Проверяем любые двойные кавычки подряд — этого достаточно,
        # чтобы отловить и `=""X""`, и `="...=""` одновременно.
        if (b'&' in tag
                or b'""' in tag
                or _INVALID_BYTE_RE.search(tag)
                or _SMART_CHARS_RE.search(tag)):
            tag = _sanitize_tag(tag)

        out.extend(tag)
        pos = gt + 1

    return out, b""


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
            # chunk может быть bytearray — это нормально
            fo.write(chunk)
    finally:
        fo.flush()


def clean_xml_file(src_path, dst_path, progress_cb=None,
                   cancel_flag=None, logger=None):
    total_size = os.path.getsize(src_path)

    if logger:
        logger.log(f"clean_xml_file: {src_path} -> {dst_path} "
                   f"({total_size} bytes)")

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

    if logger:
        logger.log("clean_xml_file: done")


# ============================================================
# 0.4. РАЗБИВКА XML ПО МЕСЯЦАМ — ОДИН ПРОХОД
# ============================================================

_TRANS_START_RE = re.compile(rb'<Trans(?=[\s>])')
_TRADE_DATE_RE = re.compile(rb'\bTradeDate\s*=\s*"([^"]*)"')


def split_xml_by_month(clean_src_path, split_dir,
                       progress_cb=None, cancel_flag=None,
                       logger=None):
    os.makedirs(split_dir, exist_ok=True)

    if logger:
        logger.log(f"split_xml_by_month: {clean_src_path} -> {split_dir}")

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

    if logger:
        logger.log(f"preamble: {preamble[:200]!r}...")

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
            if logger:
                logger.log(f"  new writer for {ym} ({label}): {path}")
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
                        # buf и data[..] — bytes-like; regex работает и с bytearray
                        tag_bytes = buf + data[pos:gt + 1]
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
                # write() принимает bytearray — без bytes(buf)
                entry["fd"].write(buf)
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
        if logger:
            logger.log(f"  closed {ym}: {entry['count']} records")

    if logger:
        logger.log(f"split_xml_by_month: {len(result)} months done")
        logger.log(f"  total records: "
                   f"{sum(info['count'] for info in result.values())}")

    return result


# ============================================================
# 0.5. ТОЧНЫЙ ПОДСЧЁТ <Trans> ЧЕРЕЗ LXML
# ============================================================
#
# Байтовый регэксп в split_xml_by_month (rb'<Trans(?=[\s>])')
# может находить ложные срабатывания: например, подстроку
# "<Trans " внутри комментария <!-- ... -->, внутри CDATA
# или внутри текстового узла (Data/Reply/PureData).
#
# В результате счётчик split_xml_by_month даёт число больше,
# чем реальное число элементов <Trans>, и проверка
# целостности показывает ложное расхождение (обычно -1).
#
# lxml видит только настоящие XML-элементы, поэтому пересчёт
# через XMLPullParser даёт "источник истины".

def count_records_lxml(path, namespace=None, cancel_flag=None,
                       progress_cb=None):
    """Считает реальные <Trans> через lxml XMLPullParser.

    Не извлекает поля, только считает элементы — быстро даже
    на больших файлах, потому что не строит словарей на запись.
    """
    search_tag = qname(RECORD_TAG, namespace) if namespace else RECORD_TAG

    parser = ET.XMLPullParser(
        events=("end",),
        tag=search_tag,
        huge_tree=True,
        recover=True,
        resolve_entities=False,
        no_network=True,
    )

    n = 0
    CHUNK = 4 * 1024 * 1024

    try:
        with open(path, "rb") as f:
            while True:
                if cancel_flag and cancel_flag.is_set():
                    break
                chunk = f.read(CHUNK)
                if not chunk:
                    break
                try:
                    parser.feed(chunk)
                except Exception:
                    pass
                for _ in parser.read_events():
                    n += 1
                    if progress_cb and n % 50000 == 0:
                        progress_cb(n)
                if cancel_flag and cancel_flag.is_set():
                    break

        try:
            parser.close()
        except Exception:
            pass
    finally:
        # явно отпускаем ссылку на парсер — не ждём GC
        parser = None

    return n


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
# 1.1. ДИАГНОСТИКА XML-ФАЙЛА
# ============================================================

def _read_selected_lines(path, target_lines, before=3, after=3):
    """Читает из файла только нужные строки (и их контекст).

    target_lines: iterable номеров строк (1-based).
    Возвращает dict {lineno: bytes} (без завершающего \\n,
    как при data.split(b"\\n")).

    Память — O(число нужных строк), а не O(размер файла).
    На 16-ГБ файле это критично: раньше diagnose_xml_error
    делал data = f.read() и мгновенно съедал всю RAM.
    """
    needed = set()
    for t in target_lines:
        if not t:
            continue
        for i in range(max(1, t - before), t + after + 1):
            needed.add(i)
    if not needed:
        return {}

    lo = min(needed)
    hi = max(needed)
    result = {}

    with open(path, "rb") as f:
        for lineno, raw in enumerate(f, 1):
            if lineno < lo:
                continue
            if lineno > hi:
                break
            if lineno in needed:
                if raw.endswith(b"\n"):
                    raw = raw[:-1]
                result[lineno] = raw

    return result


def diagnose_xml_error(xml_path, error, logger):
    lines_report = []
    lines_report.append(f"Файл: {xml_path}")
    lines_report.append(f"Ошибка: {error}")

    entries = list(getattr(error, "error_log", []) or [])
    if not entries:
        m = re.search(r'line (\d+), column (\d+)', str(error))
        if m:
            entries = [type("E", (), {
                "line": int(m.group(1)),
                "column": int(m.group(2)),
                "message": str(error),
            })()]

    # Читаем только окрестности строк с ошибками, не весь файл
    target_lines = [e.line for e in entries[:5] if e.line]
    try:
        line_data = _read_selected_lines(xml_path, target_lines,
                                         before=3, after=3)
    except Exception as exc:
        lines_report.append(f"Не удалось прочитать файл: {exc}")
        return "\n".join(lines_report)

    for e in entries[:5]:
        ln = e.line
        col = e.column
        msg = e.message
        lines_report.append("")
        lines_report.append(f"--- Ошибка на строке {ln}, "
                            f"колонке {col}: {msg} ---")

        raw = line_data.get(ln)
        if raw is None:
            lines_report.append("  (не удалось найти эту строку)")
            continue

        try:
            txt = raw.decode("cp1251", errors="replace")
        except Exception:
            txt = raw.decode("utf-8", errors="replace")

        col0 = max(0, (col or 1) - 1)
        left = max(0, col0 - 40)
        right = min(len(txt), col0 + 40)

        lines_report.append(f"  строка (срез {left}..{right}):")
        lines_report.append(f"    ...{txt[left:right]}...")
        lines_report.append(f"  байты вокруг (hex):")
        try:
            hex_slice = raw[max(0, col0 - 20): col0 + 20]
            lines_report.append(f"    {hex_slice.hex(' ')}")
        except Exception:
            pass

        caret_pos = col0 - left
        lines_report.append(f"    {' ' * caret_pos}^")

        if len(txt) > 500:
            lines_report.append(f"  полная строка (первые 500):")
            lines_report.append(f"    {txt[:500]}")
        else:
            lines_report.append(f"  полная строка:")
            lines_report.append(f"    {txt}")

        lines_report.append(f"  контекст:")
        for i in range(max(1, ln - 3), ln + 4):
            raw_i = line_data.get(i)
            if raw_i is None:
                continue
            try:
                t_i = raw_i.decode("cp1251", errors="replace")
            except Exception:
                t_i = raw_i.decode("utf-8", errors="replace")
            if len(t_i) > 200:
                t_i = t_i[:200] + "..."
            marker = ">>>" if i == ln else "   "
            lines_report.append(f"    {marker} {i:>6}: {t_i}")

    text = "\n".join(lines_report)

    if logger:
        logger.log("=" * 60)
        logger.log("XML SYNTAX ERROR")
        logger.log("=" * 60)
        for line in lines_report:
            logger.log(line)
        logger.log("=" * 60)

    return text


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
# 3. ИЗВЛЕЧЕНИЕ ЗАПИСЕЙ — через XMLPullParser (huge_tree=True)
# ============================================================

def extract_records(path, record_tag=RECORD_TAG, namespace=None,
                    config=None, descriptions=None,
                    progress_cb=None, cancel_flag=None, max_records=None):
    if config is None:
        config = ExtractConfig()
    if descriptions is None:
        descriptions = {}

    search_tag = qname(record_tag, namespace)

    parser = ET.XMLPullParser(
        events=("end",),
        tag=search_tag,
        huge_tree=True,
        recover=True,
        resolve_entities=False,
        no_network=True,
    )

    count = 0
    CHUNK = 4 * 1024 * 1024

    try:
        with open(path, "rb") as f:
            while True:
                if cancel_flag and cancel_flag.is_set():
                    return

                data = f.read(CHUNK)
                if not data:
                    break

                try:
                    parser.feed(data)
                except Exception:
                    pass

                for _, elem in parser.read_events():
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

        try:
            parser.close()
        except Exception:
            pass
    finally:
        # Явно отпускаем ссылку на парсер. Если генератор "висит"
        # (не вызван .close() и не исчерпан), парсер жил бы до
        # следующей сборки мусора и держал lxml-дерево.
        parser = None


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
                      as_cards, progress_cb=None, cancel_flag=None,
                      logger=None):
    table_columns = build_table_columns(card_columns)
    table_header_map = {RECORD_NUM_COL: "Запись"}
    table_header_map.update(card_header_map)

    now = datetime.now()
    started_at = started_at or now
    finished_at = finished_at or now
    duration_sec = (finished_at - started_at).total_seconds()

    title = f"Отчёт по транзакциям QUIK — {month_label}"
    count = 0

    if logger:
        logger.log(f"  export_html_month: {out_path} ({month_label})")

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
                    if logger:
                        logger.log(f"    {month_label}: {count} records")
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
                    if logger:
                        logger.log(f"    {month_label}: {count} records")
            f.write("</tbody></table>\n")

        f.write("</body></html>\n")

    if logger:
        logger.log(f"  export_html_month done: {count} records")

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
                     as_cards, progress_cb=None, cancel_flag=None,
                     logger=None):
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

    if logger:
        logger.log(f"  export_pdf_month: {out_path} ({month_label}), "
                   f"cols={ncols}, page={pagesize}")

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
                if logger:
                    logger.log(f"    {month_label}: {count} records")

        c.showPage()
        c.save()
        if logger:
            logger.log(f"  export_pdf_month done: {count} records")
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
            if logger:
                logger.log(f"    {month_label}: {count} records")

    c.showPage()
    c.save()
    if logger:
        logger.log(f"  export_pdf_month done: {count} records")
    return count


# ============================================================
# 7. GUI
# ============================================================

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("QUIK XML → отчёты по месяцам (HTML / PDF)")
        self.geometry("880x700")
        self.resizable(False, False)

        self.xml_path = tk.StringVar()
        self.format_var = tk.StringVar(value="html")
        self.view_mode = tk.StringVar(value="cards")

        self.delete_temp_files = tk.BooleanVar(value=True)
        self.save_bad_files = tk.BooleanVar(value=True)

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

        frame_opts = ttk.LabelFrame(self, text="4. Опции обработки")
        frame_opts.pack(fill="x", **pad)

        ttk.Checkbutton(
            frame_opts,
            text="Удалять временные файлы после работы",
            variable=self.delete_temp_files,
        ).pack(anchor="w", padx=10, pady=4)

        ttk.Checkbutton(
            frame_opts,
            text="Сохранять битый XML-файл при ошибке "
                 "(__ERROR__YYYY-MM.xml)",
            variable=self.save_bad_files,
        ).pack(anchor="w", padx=10, pady=4)

        frame_time = ttk.LabelFrame(self, text="5. Время и целостность")
        frame_time.pack(fill="x", **pad)

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

        ttk.Separator(frame_time, orient="vertical").grid(
            row=0, column=1, sticky="ns", padx=12, pady=4)

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

    def _recount_progress_cb(self, ym, label, done_months, total_months):
        self._set_status(
            f"Пересчёт записей (lxml): {done_months}/{total_months} — "
            f"{label}", "blue")

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

        keep_temp = not self.delete_temp_files.get()
        save_bad = self.save_bad_files.get()

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
            logger = None
            try:
                src_dir = os.path.dirname(os.path.abspath(xml_file))
                work_dir = tempfile.mkdtemp(
                    prefix="quik_work_", dir=src_dir)

                log_path = os.path.join(work_dir, "debug.log")
                logger = Logger(log_path)

                logger.log("=" * 60)
                logger.log("START")
                logger.log(f"  src: {xml_file}")
                logger.log(f"  out: {out_path}")
                logger.log(f"  fmt: {fmt}, view: {'cards' if as_cards else 'table'}")
                logger.log(f"  work_dir: {work_dir}")
                logger.log(f"  keep_temp_files: {keep_temp}")
                logger.log(f"  save_bad_files: {save_bad}")
                logger.log(f"  RSS на старте: {rss_mb():.0f} МБ")
                logger.log("=" * 60)

                clean_path = os.path.join(work_dir, "clean.xml")
                split_dir = os.path.join(work_dir, "months")
                os.makedirs(split_dir, exist_ok=True)

                # ─── ЭТАП 1: очистка XML ───
                self._set_status("Этап 1/4: очистка XML…", "blue")
                logger.log("ЭТАП 1: очистка XML")
                clean_xml_file(
                    xml_file, clean_path,
                    progress_cb=self._clean_progress_cb,
                    cancel_flag=self.cancel_flag,
                    logger=logger)
                logger.log(f"  RSS после очистки: {rss_mb():.0f} МБ")

                if self.cancel_flag.is_set():
                    raise RuntimeError("Отменено")

                # ─── ЭТАП 2: резка по месяцам ───
                self._set_status(
                    "Этап 2/4: резка XML по месяцам…", "blue")
                logger.log("ЭТАП 2: резка по месяцам")

                months_info = split_xml_by_month(
                    clean_path, split_dir,
                    progress_cb=self._split_progress_cb,
                    cancel_flag=self.cancel_flag,
                    logger=logger)
                logger.log(f"  RSS после резки: {rss_mb():.0f} МБ")

                if self.cancel_flag.is_set():
                    raise RuntimeError("Отменено")

                if not months_info:
                    raise RuntimeError(
                        "В файле не найдено ни одной транзакции.")

                # Промежуточный итог от байтового регэкспа (может
                # содержать ложные срабатывания — см. 0.5)
                byte_total = sum(info["count"]
                                 for info in months_info.values())
                logger.log(f"Быстрый счёт (по байтам): {byte_total}")

                if not keep_temp:
                    try:
                        os.remove(clean_path)
                    except Exception:
                        pass

                first_month_path = list(months_info.values())[0]["path"]
                namespace = detect_namespace(first_month_path)
                header = read_report_header(first_month_path)
                logger.log(f"namespace: {namespace}")
                logger.log(f"header: {header}")

                # ─── ЭТАП 2.5: точный пересчёт записей через lxml ───
                # Байтовый регэксп rb'<Trans(?=[\s>])' может давать
                # ложные срабатывания (например, "<Trans " внутри
                # комментария или внутри текстового узла Data).
                # lxml видит только настоящие XML-элементы — это
                # источник истины.
                self._set_status(
                    "Этап 2.5/4: пересчёт записей (lxml)…", "blue")
                logger.log("ЭТАП 2.5: пересчёт записей через lxml")

                months_sorted_tmp = sorted(months_info.keys())
                total_in_xml = 0
                total_m_diff = 0

                for idx, ym in enumerate(months_sorted_tmp, 1):
                    if self.cancel_flag.is_set():
                        raise RuntimeError("Отменено")

                    info = months_info[ym]
                    real_n = count_records_lxml(
                        info["path"], namespace,
                        cancel_flag=self.cancel_flag)

                    if real_n != info["count"]:
                        diff = info["count"] - real_n
                        total_m_diff += abs(diff)
                        logger.log(
                            f"  WARN {ym} ({info['label']}): "
                            f"byte_count={info['count']} "
                            f"lxml_count={real_n} "
                            f"diff={diff:+d}  "
                            f"(ложные срабатывания байтового "
                            f"регэкспа)")
                    else:
                        logger.log(f"  {ym} ({info['label']}): "
                                   f"{real_n} records")

                    info["count"] = real_n
                    total_in_xml += real_n

                    self._recount_progress_cb(
                        ym, info["label"], idx, len(months_sorted_tmp))

                logger.log(f"Точный счёт (lxml): {total_in_xml}")
                if total_m_diff:
                    logger.log(
                        f"Суммарное расхождение байт/lxml: "
                        f"{total_m_diff} (ложные срабатывания)")

                for ym in sorted(months_info.keys()):
                    info = months_info[ym]
                    logger.log(f"  {ym} ({info['label']}): "
                               f"{info['count']} records, "
                               f"file={info['path']}")

                self.after(0, lambda: self.lbl_total_xml.config(
                    text=f"{total_in_xml:,}".replace(",", " "),
                    foreground="black"))

                # ─── Сэмпл колонок ───
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

                logger.log(f"sample cols: {len(sample_seen)}")
                logger.log(f"  {list(sample_seen.keys())}")

                card_columns, card_header_map = build_columns_from_seen(
                    sample_seen, sample_desc)

                # sample_seen больше не нужен — освобождаем явно.
                # sample_desc ОСТАВЛЯЕМ: он идёт дальше в extract_records
                # как кэш описаний полей.
                del sample_seen
                gc.collect()
                logger.log(f"  RSS после сэмпла колонок: {rss_mb():.0f} МБ")

                # ─── ЭТАП 3: генерация отчётов ───
                logger.log("ЭТАП 3: генерация отчётов")
                out_dir_base = os.path.dirname(os.path.abspath(out_path))
                base_name = os.path.splitext(
                    os.path.basename(out_path))[0]
                reports_dir = os.path.join(out_dir_base,
                                           f"{base_name}_reports")
                os.makedirs(reports_dir, exist_ok=True)
                logger.log(f"reports_dir: {reports_dir}")

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

                    logger.log(f"[{i}/{K}] {ym} — {month_label} "
                               f"({info['count']} records)")

                    self._set_status(
                        f"Этап 3/4: {i}/{K} — {month_label} "
                        f"({info['count']} записей)…", "blue")

                    records_iter = extract_records(
                        month_path, RECORD_TAG, namespace,
                        config=config, descriptions=sample_desc,
                        progress_cb=self._progress_cb,
                        cancel_flag=self.cancel_flag,
                        max_records=DEFAULT_LIMIT)

                    finished_at = datetime.now()

                    try:
                        if fmt == "html":
                            count = export_html_month(
                                file_path, month_label, records_iter,
                                card_columns, card_header_map,
                                header, started_at, finished_at,
                                as_cards=as_cards,
                                progress_cb=self._progress_cb,
                                cancel_flag=self.cancel_flag,
                                logger=logger)
                        else:
                            count = export_pdf_month(
                                file_path, month_label, records_iter,
                                card_columns, card_header_map,
                                header, started_at, finished_at,
                                as_cards=as_cards,
                                progress_cb=self._progress_cb,
                                cancel_flag=self.cancel_flag,
                                logger=logger)

                        written.append((file_path, count))
                        logger.log(f"  [{i}/{K}] done: {count} records")

                    except ET.XMLSyntaxError as xml_exc:
                        diag = diagnose_xml_error(month_path, xml_exc,
                                                  logger)

                        if save_bad:
                            saved_path = os.path.join(
                                reports_dir,
                                f"__ERROR___{ym}.xml")
                            try:
                                shutil.copy2(month_path, saved_path)
                            except Exception:
                                saved_path = month_path
                        else:
                            saved_path = month_path

                        msg = (f"Ошибка XML в месяце {month_label} "
                               f"({ym}).\n\n"
                               f"Файл: {saved_path}\n\n"
                               f"Диагностика:\n{diag}\n\n"
                               f"Лог: {log_path}")
                        logger.log("!!! XMLSyntaxError !!!")
                        logger.log(msg)

                        def show_err(m=msg):
                            self._show_error("Ошибка XML", m)
                        self.after(0, show_err)
                        written.append((file_path, -1))

                    finally:
                        # ─── Явное освобождение памяти после месяца ───
                        # 1) .close() бросает GeneratorExit в текущую
                        #    точку yield внутри extract_records, срабатывают
                        #    все finally внутри — закрывается lxml-парсер,
                        #    файл, последний <Trans>.
                        try:
                            records_iter.close()
                        except Exception:
                            pass
                        del records_iter

                        # 2) GC собирает циклы lxml/reportlab.
                        #    Без этого RSS может не падать минутами.
                        gc.collect()

                        rss = rss_mb()
                        logger.log(
                            f"  [{i}/{K}] RSS после {month_label}: "
                            f"{rss:.0f} МБ")

                total_in_reports = sum(c for _, c in written if c >= 0)

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

                logger.log(f"total_in_xml (lxml): {total_in_xml}")
                logger.log(f"total_in_reports: {total_in_reports}")
                logger.log(f"total_time: {fmt_duration(total_time)}")
                logger.log(f"RSS на финише: {rss_mb():.0f} МБ")

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
                if total_m_diff:
                    msg_lines.append(
                        f"(байтовый счётчик дал лишних "
                        f"{total_m_diff} — ложные срабатывания, "
                        f"см. лог)")
                msg_lines.append("")
                msg_lines.append(f"Месяцев: {len(written)}")
                msg_lines.append(f"Папка: {reports_dir}")
                msg_lines.append(f"Лог: {log_path}")
                if keep_temp:
                    msg_lines.append(f"Временные файлы: {work_dir}")
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
                if logger:
                    logger.log("!!!" * 20)
                    logger.log("ОБЩАЯ ОШИБКА")
                    logger.log(traceback.format_exc())
                    logger.log("!!!" * 20)

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
                if logger:
                    logger.log("FINISHED")
                    logger.close()

                if work_dir and os.path.exists(work_dir):
                    if keep_temp:
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