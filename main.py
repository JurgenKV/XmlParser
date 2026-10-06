import os
import html
import threading
from datetime import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from collections import Counter, defaultdict

from lxml import etree as ET

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

RECORD_TAG = "Trans"
CONTAINER_TAG = "TransData"
FIELD_TAG = "Field"
NAME_ATTR = "Name"
DESC_ATTR = "Description"
PREPARED_ATTR = "PreparedValue"

DEFAULT_LIMIT = 1000000
MIN_FILL_RATIO = 1.0
EMPTY_MARK = "—"

QUIK_FIELD_LABELS = {
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
    """Показывает PreparedValue как есть; пустое -> «—»."""
    base = base_field_name(col)
    if value is None:
        return EMPTY_MARK
    s = str(value).strip()
    if not s:
        return EMPTY_MARK

    # YYYY-MM-DD -> DD.MM.YYYY
    if len(s) == 10 and s[4] == "-" and s[7] == "-":
        y, m, d = s.split("-")
        if y.isdigit() and m.isdigit() and d.isdigit():
            return f"{d}.{m}.{y}"

    # YYYYMMDD -> DD.MM.YYYY
    if base in ("TradeDate", "QuikDate", "Date", "SettleDate") \
            and len(s) == 8 and s.isdigit():
        return f"{s[6:8]}.{s[4:6]}.{s[0:4]}"

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
    info = {"ProgramVersion": "", "StartDate": "", "EndDate": ""}
    with open(path, "rb") as f:
        for _, elem in ET.iterparse(f, events=("start",)):
            for k in ("ProgramVersion", "StartDate", "EndDate"):
                v = elem.get(k)
                if v:
                    info[k] = v
            break
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


def _is_empty(value):
    return value is None or str(value).strip() == ""


# ============================================================
# 2. ИЗВЛЕЧЕНИЕ ЗАПИСЕЙ
# ============================================================

def extract_records(path, record_tag=RECORD_TAG, namespace=None,
                    descriptions=None,
                    progress_cb=None, cancel_flag=None, max_records=None):
    """
    Порядок ключей — строго как в XML.
    Значения из TransData/Field берутся ТОЛЬКО из PreparedValue.
    Пустые дочерние узлы Trans (например, PureData) тоже попадают в rec,
    чтобы в отчёте отобразиться как «—».
    """
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

            # 2. Дочерние узлы Trans в порядке XML
            for child in elem:
                if not isinstance(child.tag, str):
                    continue
                tag_local = strip_ns(child.tag)

                if tag_local == CONTAINER_TAG:
                    fields = list(child.findall(FIELD_TAG))
                    if namespace and not fields:
                        fields = list(child.findall(
                            f"{{{namespace}}}{FIELD_TAG}"))

                    def _num_key(fe):
                        try:
                            return int(fe.get("Number") or 0)
                        except ValueError:
                            return 0
                    fields.sort(key=_num_key)

                    for field in fields:
                        name = (field.get(NAME_ATTR) or "").strip()
                        desc = (field.get(DESC_ATTR) or "").strip()
                        prepared = (field.get(PREPARED_ATTR) or "").strip()

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

                # 3. Обычный дочерний узел — сначала текст, потом атрибуты.
                #    Даже если текста нет — добавляем ключ (будет «—»).
                text = (child.text or "").strip()

                if tag_local not in rec:
                    rec[tag_local] = text
                else:
                    i = 2
                    while f"{tag_local}_{i}" in rec:
                        i += 1
                    rec[f"{tag_local}_{i}"] = text

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
# 3. КОЛОНКИ
# ============================================================

def build_columns(records, descriptions=None, min_fill_ratio=MIN_FILL_RATIO):
    """
    Оставляем колонки, которые присутствуют в КАЖДОЙ записи
    (наличие ключа, а не непустое значение).
    """
    if not records:
        return [], {}

    total = len(records)
    counts = defaultdict(int)
    order = {}

    for rec in records:
        for k in rec:
            if k not in order:
                order[k] = len(order)
            counts[k] += 1

    ordered_keys = sorted(order, key=order.get)
    columns = [
        k for k in ordered_keys
        if counts.get(k, 0) / total >= min_fill_ratio
    ]

    header_map = {c: label_for(c, descriptions) for c in columns}
    return columns, header_map


# ============================================================
# 4. HTML
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
  td.empty {{ color: #999; }}
  .card {{ border: 1px solid #ddd; border-radius: 6px; padding: 12px 16px;
           margin: 8px 0; background: #fafbfc; }}
  .card h3 {{ margin-top: 0; color: #004a99; }}
  .field {{ margin: 2px 0; }}
  .field b {{ color: #333; }}
  .field .empty {{ color: #999; }}
</style></head><body>
<h1>{title}</h1>
<div class="meta">{meta}</div>
"""


def _esc(s):
    return html.escape(str(s))


def export_html(xml_path, out_path, namespace=None,
                limit=DEFAULT_LIMIT, progress_cb=None,
                cancel_flag=None, as_cards=False):
    header = read_report_header(xml_path)

    descriptions = {}
    records = list(extract_records(
        xml_path, RECORD_TAG, namespace,
        descriptions=descriptions,
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
                    raw = rec.get(col, "")
                    val = format_quik_value(col, raw)
                    css = ' class="empty"' if _is_empty(raw) else ""
                    f.write(
                        f'<div class="field"><b>{_esc(header_map[col])}:</b> '
                        f'<span{css}>{_esc(val)}</span></div>\n')
                f.write("</div>\n")
        else:
            f.write("<table><thead><tr>")
            for col in columns:
                f.write(f"<th>{_esc(header_map[col])}</th>")
            f.write("</tr></thead><tbody>\n")
            for rec in records:
                f.write("<tr>")
                for col in columns:
                    raw = rec.get(col, "")
                    val = format_quik_value(col, raw)
                    classes = []
                    if base_field_name(col) in QUIK_NUMERIC_FIELDS:
                        classes.append("num")
                    if _is_empty(raw):
                        classes.append("empty")
                    cls_attr = f' class="{" ".join(classes)}"' if classes else ""
                    f.write(f'<td{cls_attr}>{_esc(val)}</td>')
                f.write("</tr>\n")
            f.write("</tbody></table>\n")

        f.write("</body></html>\n")


# ============================================================
# 5. PDF
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


def export_pdf(xml_path, out_path, namespace=None,
               limit=DEFAULT_LIMIT, progress_cb=None,
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
        xml_path, RECORD_TAG, namespace,
        descriptions=descriptions,
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
                raw = rec.get(col, "")
                val = format_quik_value(col, raw)
                if _is_empty(raw):
                    val_html = f'<font color="#999">{_esc(val)}</font>'
                else:
                    val_html = _esc(val)
                story.append(Paragraph(
                    f"<b>{_esc(header_map[col])}:</b> {val_html}", normal))
            story.append(Spacer(1, 6))
    else:
        data = [[Paragraph(f"<b>{_esc(header_map[c])}</b>", normal)
                 for c in columns]]
        for rec in records:
            row = []
            for c in columns:
                raw = rec.get(c, "")
                val = format_quik_value(c, raw)
                if _is_empty(raw):
                    row.append(Paragraph(
                        f'<font color="#999">{_esc(val)}</font>', normal))
                else:
                    row.append(Paragraph(_esc(val), normal))
            data.append(row)

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
# 6. GUI
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

    def convert(self):
        xml_file = self.xml_path.get().strip()
        if not xml_file or not os.path.isfile(xml_file):
            messagebox.showwarning("Внимание", "Выберите существующий XML-файл")
            return

        fmt = self.format_var.get()
        ext = fmt
        as_cards = self.view_mode.get() == "cards"

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
                    export_html(xml_file, out_path,
                                namespace=namespace,
                                limit=DEFAULT_LIMIT,
                                progress_cb=self._progress_cb,
                                cancel_flag=self.cancel_flag,
                                as_cards=as_cards)
                elif fmt == "pdf":
                    export_pdf(xml_file, out_path,
                               namespace=namespace,
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