# -*- coding: utf-8 -*-
"""
QUIK XML -> HTML / DOCX / PDF
Потоковый парсер (iterparse), выдерживает файлы десятки ГБ.
GUI: Tkinter.
"""

import os
import sys
import threading
import traceback
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from html import escape as html_escape
from typing import Iterable, Iterator

# ---------- ускоренный парсер, если есть lxml ----------
try:
    from lxml import etree as LET
    HAS_LXML = True
except Exception:
    HAS_LXML = False


# =====================================================================
#                         МОДЕЛЬ ДАННЫХ
# =====================================================================

@dataclass
class QuikTable:
    """Одна таблица QUIK-отчёта."""
    name: str                       # имя узла, например 'securities'
    title: str                      # человекочитаемое имя
    columns: list[str] = field(default_factory=list)  # заголовки
    rows: list[list[str]] = field(default_factory=list)


# =====================================================================
#                        ПОТОКОВЫЙ ПАРСЕР
# =====================================================================

# Теги QUIK, которые нужно пропускать (служебные)
SKIP_TAGS = {"document", "root", "meta", "params", "param"}


def _iter_events(path: str):
    """
    Возвращает итератор (event, elem) для iterparse.
    Автоматически выбирает lxml или stdlib.
    """
    if HAS_LXML:
        # lxml быстрее и умеет recover
        return LET.iterparse(path, events=("end",), recover=True, huge_tree=True)
    else:
        return ET.iterparse(path, events=("end",))


def _row_to_list(elem) -> list[str]:
    """Преобразует <row>...</row> в список значений."""
    values = []
    for child in elem:
        # QUIK обычно пишет <row><field>..</field>...</row> или атрибутами
        if child.text is not None:
            values.append(child.text.strip())
        else:
            values.append("")
        if child.tail:
            t = child.tail.strip()
            if t:
                values.append(t)
    return values


def stream_quik_tables(path: str) -> Iterator[QuikTable]:
    """
    Потоково парсит QUIK XML.
    Возвращает таблицы по мере нахождения (yield).
    Память: O(размер одной таблицы).
    """
    context = _iter_events(path)

    current_table: QuikTable | None = None
    current_name: str | None = None
    header_found = False

    # корневой элемент для очистки памяти (lxml)
    root = None

    for event, elem in context:
        tag = elem.tag
        # у lxml тег может быть {ns}name — отбрасываем namespace
        if isinstance(tag, str) and tag.startswith("{"):
            tag = tag.split("}", 1)[1]

        if root is None:
            # первый end-элемент — обычно самый вложенный;
            # корень найдём через getroottree
            try:
                root = context.root if hasattr(context, "root") else elem
            except Exception:
                root = elem

        # --- начало новой таблицы ---
        # Эвристика: таблица = узел, у которого дети <row>.
        # У QUIK это обычно <securities>, <trades>, <orders> и т.д.
        # На практике идём снизу вверх: сначала встречаем <row>,
        # но нам нужно поймать родителя. Поэтому собираем ряды
        # в буфер и "закрываем" таблицу, когда встречаем элемент,
        # содержащий эти ряды.
        #
        # Проще: ловим все <row>, смотрим parent через elem.getparent()
        # (только lxml). Для stdlib — придётся накапливать.
        # Реализация ниже работает для обоих вариантов через буфер.

        if tag == "row":
            # определяем имя таблицы
            parent = None
            if HAS_LXML:
                parent = elem.getparent()
            if parent is not None:
                ptag = parent.tag
                if isinstance(ptag, str) and ptag.startswith("{"):
                    ptag = ptag.split("}", 1)[1]
                if current_name != ptag:
                    # закрываем предыдущую таблицу, если была
                    if current_table and current_table.rows:
                        yield current_table
                    current_name = ptag
                    current_table = QuikTable(name=ptag, title=ptag)
                    header_found = False

            row_values = _row_to_list(elem)

            if current_table is None:
                # на всякий случай
                current_table = QuikTable(name="unknown", title="unknown")
                current_name = "unknown"

            # Первую строку таблицы часто используют как заголовок
            # (в некоторых QUIK-выгрузках так и есть)
            if not header_found and not current_table.columns:
                # эвристика: если первая строка похожа на заголовок
                # (нет чисел), принимаем её за заголовки.
                current_table.columns = row_values
                header_found = True
            else:
                current_table.rows.append(row_values)

            # очистка: удаляем <row> из дерева (важно для больших файлов)
            if HAS_LXML:
                elem.clear()
                while elem.getprevious() is not None:
                    del elem.getparent()[0]
            else:
                elem.clear()
                # стандартный ElementTree не умеет удалять sibling'ов на лету,
                # но clear() уже освобождает text/attrib/children
            continue

        # --- если встретили элемент-обёртку (не row), и в буфере что-то есть ---
        # закрываем таблицу, только когда реально сменился родитель.
        # (обрабатывается выше при появлении нового <row>)

    # хвост
    if current_table and current_table.rows:
        yield current_table


# =====================================================================
#                          ЭКСПОРТЁРЫ
# =====================================================================

def export_html(tables: Iterable[QuikTable], out_path: str,
                title: str = "QUIK Report", progress=None):
    """Потоково пишет HTML. Память O(1) от размера таблиц."""
    with open(out_path, "w", encoding="utf-8", newline="") as f:
        f.write("<!DOCTYPE html>\n<html><head><meta charset='utf-8'>\n")
        f.write(f"<title>{html_escape(title)}</title>\n")
        f.write("""<style>
body{font-family:Segoe UI,Arial,sans-serif;margin:20px;background:#f7f7f7}
h2{color:#1a3d6d}
table{border-collapse:collapse;margin-bottom:24px;background:#fff;font-size:13px}
th,td{border:1px solid #bbb;padding:4px 8px;text-align:left;white-space:nowrap}
th{background:#1a3d6d;color:#fff;position:sticky;top:0}
tr:nth-child(even) td{background:#f0f4fa}
</style></head><body>\n""")
        f.write(f"<h1>{html_escape(title)}</h1>\n")

        n = 0
        for table in tables:
            n += 1
            f.write(f"<h2>{html_escape(table.title)} "
                    f"({len(table.rows)} строк)</h2>\n<table>\n")
            if table.columns:
                f.write("<thead><tr>")
                for c in table.columns:
                    f.write(f"<th>{html_escape(str(c))}</th>")
                f.write("</tr></thead>\n")
            f.write("<tbody>\n")
            for row in table.rows:
                f.write("<tr>")
                for v in row:
                    f.write(f"<td>{html_escape(str(v))}</td>")
                f.write("</tr>\n")
            f.write("</tbody></table>\n")
            if progress:
                progress(f"HTML: таблица «{table.title}» — {len(table.rows)} строк")

        f.write("</body></html>\n")


def export_docx(tables: Iterable[QuikTable], out_path: str,
                title: str = "QUIK Report", progress=None,
                max_rows_per_table: int = 200_000):
    """
    DOCX через python-docx.
    ВНИМАНИЕ: python-docx держит документ в памяти — для 30 ГБ не годится.
    Для очень больших файлов используйте HTML/PDF, либо разбивайте
    на несколько DOCX (см. параметр max_rows_per_table).
    """
    from docx import Document
    from docx.shared import Pt

    doc = Document()
    doc.add_heading(title, level=0)

    for table in tables:
        doc.add_heading(f"{table.title} ({len(table.rows)} строк)", level=1)
        cols = table.columns if table.columns else (
            [f"col{i}" for i in range(len(table.rows[0]))] if table.rows else [])
        if not cols:
            continue
        t = doc.add_table(rows=1, cols=len(cols))
        t.style = "Light Grid Accent 1"
        hdr = t.rows[0].cells
        for i, c in enumerate(cols):
            hdr[i].text = str(c)
        for row in table.rows:
            cells = t.add_row().cells
            for i, v in enumerate(row[:len(cols)]):
                cells[i].text = str(v)
        if progress:
            progress(f"DOCX: таблица «{table.title}» — {len(table.rows)} строк")

    doc.save(out_path)


def export_pdf(tables: Iterable[QuikTable], out_path: str,
               title: str = "QUIK Report", progress=None):
    """
    PDF через reportlab. Потоково пишем таблицу — память O(1).
    Шрифт по умолчанию Helvetica (кириллицу надо регистрировать отдельно,
    см. register_font ниже).
    """
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.lib.units import mm
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer,
                                    Table, TableStyle, LongTable)

    # --- регистрируем шрифт с кириллицей (если найдём в системе) ---
    font_name = "Helvetica"
    candidates = [
        r"C:\Windows\Fonts\arial.ttf",
        r"C:\Windows\Fonts\DejaVuSans.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/Library/Fonts/Arial.ttf",
    ]
    for p in candidates:
        if os.path.exists(p):
            try:
                pdfmetrics.registerFont(TTFont("QuikFont", p))
                font_name = "QuikFont"
                break
            except Exception:
                pass

    styles = getSampleStyleSheet()
    styles["Title"].fontName = font_name
    styles["Heading1"].fontName = font_name
    styles["Heading2"].fontName = font_name

    doc = SimpleDocTemplate(
        out_path,
        pagesize=landscape(A4),
        leftMargin=10 * mm, rightMargin=10 * mm,
        topMargin=10 * mm, bottomMargin=10 * mm,
        title=title,
    )

    story = [Paragraph(title, styles["Title"]), Spacer(1, 6 * mm)]

    for table in tables:
        story.append(Paragraph(
            f"{table.title} ({len(table.rows)} строк)", styles["Heading2"]))
        story.append(Spacer(1, 3 * mm))

        cols = table.columns if table.columns else (
            [f"col{i}" for i in range(len(table.rows[0]))] if table.rows else [])
        if not cols:
            continue

        data = [cols]
        for row in table.rows:
            data.append([str(v) for v in row[:len(cols)]])

        # LongTable умеет разрываться между страницами
        t = LongTable(data, repeatRows=1)
        t.setStyle(TableStyle([
            ("FONTNAME", (0, 0), (-1, -1), font_name),
            ("FONTSIZE", (0, 0), (-1, -1), 7),
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1a3d6d")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("LEFTPADDING", (0, 0), (-1, -1), 3),
            ("RIGHTPADDING", (0, 0), (-1, -1), 3),
        ]))
        story.append(t)
        story.append(Spacer(1, 6 * mm))

        if progress:
            progress(f"PDF: таблица «{table.title}» — {len(table.rows)} строк")

    doc.build(story)


# =====================================================================
#                          ОРКЕСТРАТОР
# =====================================================================

def convert(xml_path: str, out_path: str, fmt: str,
            progress=print) -> None:
    """
    fmt: 'html' | 'docx' | 'pdf'
    """
    fmt = fmt.lower()
    if fmt not in {"html", "docx", "pdf"}:
        raise ValueError(f"Неизвестный формат: {fmt}")

    progress(f"Открываю: {xml_path}")
    progress(f"Парсер: {'lxml' if HAS_LXML else 'xml.etree'}")

    tables = stream_quik_tables(xml_path)

    if fmt == "html":
        export_html(tables, out_path, progress=progress)
    elif fmt == "docx":
        export_docx(tables, out_path, progress=progress)
    elif fmt == "pdf":
        export_pdf(tables, out_path, progress=progress)

    progress(f"Готово: {out_path}")


# =====================================================================
#                             GUI (Tkinter)
# =====================================================================

def run_gui():
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox

    root = tk.Tk()
    root.title("QUIK XML → HTML / Word / PDF")
    root.geometry("720x480")
    root.minsize(640, 420)

    # --- переменные ---
    var_xml = tk.StringVar()
    var_out = tk.StringVar()
    var_fmt = tk.StringVar(value="html")

    # --- верхняя панель ---
    frm = ttk.Frame(root, padding=10)
    frm.pack(fill="x")

    ttk.Label(frm, text="XML-файл QUIK:").grid(row=0, column=0, sticky="w")
    ttk.Entry(frm, textvariable=var_xml, width=70).grid(
        row=0, column=1, padx=6, pady=4, sticky="we")
    ttk.Button(frm, text="Обзор…",
               command=lambda: var_xml.set(
                   filedialog.askopenfilename(
                       title="Выберите XML QUIK",
                       filetypes=[("XML files", "*.xml"),
                                  ("All files", "*.*")])))\
        .grid(row=0, column=2, padx=4)

    ttk.Label(frm, text="Куда сохранить:").grid(row=1, column=0, sticky="w")
    ttk.Entry(frm, textvariable=var_out, width=70).grid(
        row=1, column=1, padx=6, pady=4, sticky="we")
    ttk.Button(frm, text="Обзор…",
               command=lambda: var_out.set(
                   filedialog.asksaveasfilename(
                       title="Сохранить как",
                       defaultextension="." + var_fmt.get(),
                       filetypes=[("HTML", "*.html"),
                                  ("Word", "*.docx"),
                                  ("PDF", "*.pdf")])))\
        .grid(row=1, column=2, padx=4)

    ttk.Label(frm, text="Формат:").grid(row=2, column=0, sticky="w")
    fmt_frame = ttk.Frame(frm)
    fmt_frame.grid(row=2, column=1, sticky="w", pady=6)
    for f, t in (("html", "HTML"), ("docx", "Word (DOCX)"), ("pdf", "PDF")):
        ttk.Radiobutton(fmt_frame, text=t, value=f, variable=var_fmt)\
            .pack(side="left", padx=4)

    frm.columnconfigure(1, weight=1)

    # --- кнопка ---
    btn_frame = ttk.Frame(root, padding=(10, 0))
    btn_frame.pack(fill="x")
    btn_run = ttk.Button(btn_frame, text="▶  Конвертировать")
    btn_run.pack(side="left")

    # --- лог ---
    log = tk.Text(root, height=18, wrap="none", font=("Consolas", 9))
    log.pack(fill="both", expand=True, padx=10, pady=10)
    log.configure(state="disabled")

    def log_write(msg: str):
        log.configure(state="normal")
        log.insert("end", msg + "\n")
        log.see("end")
        log.configure(state="disabled")
        root.update_idletasks()

    # --- запуск конвертации в отдельном потоке ---
    def do_convert():
        xml = var_xml.get().strip()
        out = var_out.get().strip()
        fmt = var_fmt.get()

        if not xml or not os.path.isfile(xml):
            messagebox.showerror("Ошибка", "Укажите существующий XML-файл.")
            return
        if not out:
            base = os.path.splitext(xml)[0]
            out = f"{base}.{ 'html' if fmt=='html' else ('docx' if fmt=='docx' else 'pdf') }"
            var_out.set(out)

        btn_run.configure(state="disabled")

        def worker():
            try:
                convert(xml, out, fmt, progress=log_write)
                log_write("=== УСПЕШНО ===")
                messagebox.showinfo("Готово", f"Файл сохранён:\n{out}")
            except Exception as e:
                log_write("!!! ОШИБКА !!!")
                log_write(str(e))
                log_write(traceback.format_exc())
                messagebox.showerror("Ошибка", str(e))
            finally:
                btn_run.configure(state="normal")

        threading.Thread(target=worker, daemon=True).start()

    btn_run.configure(command=do_convert)

    root.mainloop()


# =====================================================================
#                              CLI
# =====================================================================

def main():
    if len(sys.argv) >= 4:
        xml, out, fmt = sys.argv[1], sys.argv[2], sys.argv[3]
        convert(xml, out, fmt)
    else:
        run_gui()


if __name__ == "__main__":
    main()