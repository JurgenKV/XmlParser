import os
import html
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import xml.etree.ElementTree as ET

# --- DOCX ---
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

# --- PDF ---
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.enums import TA_LEFT


# ============================================================
# 1. ПОТОКОВЫЙ ПАРСИНГ XML
# ============================================================

def iter_xml_nodes(path, progress_cb=None):
    """
    Потоковый обход XML. Возвращает (level, tag, attrib, text) для каждого узла.
    Память: O(глубина дерева), а не O(размер файла).
    """
    context = ET.iterparse(path, events=("start", "end"))
    _, root = next(context)  # корень

    level = 0
    count = 0
    for event, elem in context:
        if event == "start":
            level += 1
        else:  # end
            level -= 1

            text = (elem.text or "").strip()
            yield (level, elem.tag, dict(elem.attrib), text)

            # Освобождаем память
            elem.clear()
            while elem.getprevious() is not None:
                del elem.getparent()[0]

            count += 1
            if progress_cb and count % 1000 == 0:
                progress_cb(count)

    del root


def xml_root_tag(path):
    """Быстро достаёт имя корневого тега."""
    for event, elem in ET.iterparse(path, events=("start",)):
        return elem.tag
    return "root"


# ============================================================
# 2. HTML ЭКСПОРТ (потоковый)
# ============================================================

def export_html(xml_path, out_path, progress_cb=None):
    root_tag = xml_root_tag(xml_path)

    with open(out_path, "w", encoding="utf-8") as f:
        # Шапка HTML
        f.write("<!DOCTYPE html>\n<html lang='ru'>\n<head>\n")
        f.write("<meta charset='utf-8'>\n")
        f.write(f"<title>XML: {html.escape(root_tag)}</title>\n")
        f.write("""
<style>
  body { font-family: -apple-system, Segoe UI, Arial, sans-serif;
         margin: 2em; line-height: 1.5; color: #222; }
  h1 { border-bottom: 2px solid #444; padding-bottom: .3em; }
  .node { margin: 2px 0; padding: 2px 4px; border-left: 3px solid #ddd; }
  .tag { font-weight: 600; color: #004a99; }
  .attr { color: #a31515; }
  .attr-name { color: #7a3e00; }
  .text { color: #333; }
  .depth-0 { border-left-color: #004a99; }
  .depth-1 { border-left-color: #2e7d32; }
  .depth-2 { border-left-color: #ef6c00; }
  .depth-3 { border-left-color: #6a1b9a; }
</style>
</head>
<body>
""")
        f.write(f"<h1>XML: {html.escape(root_tag)}</h1>\n")

        for level, tag, attrib, text in iter_xml_nodes(xml_path, progress_cb):
            indent = "&nbsp;" * (4 * level)
            depth_class = f"depth-{level % 4}"

            attrs_html = ""
            if attrib:
                parts = [
                    f'<span class="attr-name">{html.escape(k)}</span>'
                    f'<span class="attr">="{html.escape(v)}"</span>'
                    for k, v in attrib.items()
                ]
                attrs_html = " " + " ".join(parts)

            text_html = ""
            if text:
                text_html = f'<span class="text">: {html.escape(text)}</span>'

            f.write(
                f'<div class="node {depth_class}">{indent}'
                f'<span class="tag">&lt;{html.escape(tag)}&gt;</span>'
                f'{attrs_html}{text_html}</div>\n'
            )

        f.write("</body>\n</html>\n")


# ============================================================
# 3. PDF ЭКСПОРТ (потоковый через генератор)
# ============================================================

def register_cyrillic_font():
    candidates = [
        "C:/Windows/Fonts/DejaVuSans.ttf",
        "C:/Windows/Fonts/arial.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/Library/Fonts/Arial Unicode.ttf",
    ]
    for path in candidates:
        if os.path.exists(path):
            try:
                pdfmetrics.registerFont(TTFont("UnicodeFont", path))
                return "UnicodeFont"
            except Exception:
                continue
    return "Helvetica"


def _esc(s):
    return (s.replace("&", "&amp;")
             .replace("<", "&lt;")
             .replace(">", "&gt;"))


def export_pdf(xml_path, out_path, progress_cb=None):
    font_name = register_cyrillic_font()
    root_tag = xml_root_tag(xml_path)

    styles = getSampleStyleSheet()
    tag_style = ParagraphStyle(
        "TagStyle", parent=styles["Normal"],
        fontName=font_name, fontSize=10, leading=13, spaceAfter=1,
        alignment=TA_LEFT,
    )
    title_style = ParagraphStyle(
        "TitleStyle", parent=styles["Title"],
        fontName=font_name, fontSize=18, leading=22,
    )

    doc = SimpleDocTemplate(
        out_path, pagesize=A4,
        leftMargin=2 * cm, rightMargin=2 * cm,
        topMargin=2 * cm, bottomMargin=2 * cm,
    )

    def story():
        yield Paragraph(f"XML: {_esc(root_tag)}", title_style)
        yield Spacer(1, 12)

        for level, tag, attrib, text in iter_xml_nodes(xml_path, progress_cb):
            indent = "&nbsp;" * (4 * level)
            attrs = ""
            if attrib:
                attrs = " " + " ".join(
                    f'{_esc(k)}="{_esc(v)}"' for k, v in attrib.items()
                )
            line = f"{indent}<b>&lt;{_esc(tag)}&gt;</b>{_esc(attrs)}"
            if text:
                line += f": {_esc(text)}"
            yield Paragraph(line, tag_style)

    # build() принимает итератор — не держит всё в памяти
    doc.build(story())


# ============================================================
# 4. DOCX ЭКСПОРТ — чанками (python-docx не умеет потоково)
# ============================================================

def export_docx_chunked(xml_path, out_base_path, nodes_per_chunk=50000,
                        progress_cb=None):
    """
    Разбивает вывод на несколько .docx файлов.
    Если всё умещается в один чанк — создаёт один файл.
    Иначе — <base>_part1.docx, _part2.docx, ...
    """
    root_tag = xml_root_tag(xml_path)

    base, ext = os.path.splitext(out_base_path)
    ext = ext or ".docx"

    part = 1
    count_in_part = 0
    saved_files = []

    def new_doc():
        d = Document()
        title = d.add_heading(f"XML: {root_tag} (часть {part})", level=0)
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        return d

    def save_doc(d, p):
        if p == 1 and not saved_files:
            path = out_base_path
        else:
            path = f"{base}_part{p}{ext}"
        d.save(path)
        saved_files.append(path)
        return path

    doc = new_doc()

    for level, tag, attrib, text in iter_xml_nodes(xml_path, progress_cb):
        if count_in_part >= nodes_per_chunk:
            save_doc(doc, part)
            part += 1
            count_in_part = 0
            doc = new_doc()

        attrs = ""
        if attrib:
            attrs = " " + " ".join(f'{k}="{v}"' for k, v in attrib.items())

        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Inches(0.2 * level)
        run = p.add_run(f"<{tag}>{attrs}")
        run.bold = True
        run.font.size = Pt(11)
        if text:
            p.add_run(f": {text}")

        count_in_part += 1

    save_doc(doc, part)
    return saved_files


# ============================================================
# 5. GUI
# ============================================================

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("XML → DOCX / PDF / HTML")
        self.geometry("620x420")
        self.resizable(False, False)

        self.xml_path = tk.StringVar()
        self.format_var = tk.StringVar(value="html")
        self.chunk_size = tk.IntVar(value=50000)
        self.cancel_flag = threading.Event()

        self._build_ui()

    def _build_ui(self):
        pad = {"padx": 10, "pady": 6}

        # Файл
        frame_file = ttk.LabelFrame(self, text="1. XML-файл")
        frame_file.pack(fill="x", **pad)
        ttk.Entry(frame_file, textvariable=self.xml_path, width=60).pack(
            side="left", padx=6, pady=6, fill="x", expand=True
        )
        ttk.Button(frame_file, text="Обзор…", command=self.choose_file).pack(
            side="right", padx=6, pady=6
        )

        # Формат
        frame_fmt = ttk.LabelFrame(self, text="2. Формат вывода")
        frame_fmt.pack(fill="x", **pad)
        for fmt, label in [
            ("html", "HTML (.html)"),
            ("docx", "Word (.docx)"),
            ("pdf",  "PDF (.pdf)"),
        ]:
            ttk.Radiobutton(
                frame_fmt, text=label, value=fmt, variable=self.format_var,
                command=self._on_format_change,
            ).pack(side="left", padx=12, pady=8)

        # Опции DOCX
        self.frame_docx = ttk.LabelFrame(self, text="3. Опции DOCX")
        self.frame_docx.pack(fill="x", **pad)
        ttk.Label(self.frame_docx, text="Узлов на файл:").pack(
            side="left", padx=6, pady=8
        )
        ttk.Entry(self.frame_docx, textvariable=self.chunk_size, width=10).pack(
            side="left", padx=6, pady=8
        )
        ttk.Label(
            self.frame_docx,
            text="(при больших XML создастся несколько файлов _partN.docx)",
            foreground="gray",
        ).pack(side="left", padx=6)

        # Прогресс
        frame_prog = ttk.LabelFrame(self, text="Прогресс")
        frame_prog.pack(fill="x", **pad)
        self.progress = ttk.Progressbar(frame_prog, mode="indeterminate")
        self.progress.pack(fill="x", padx=6, pady=8)

        # Кнопки
        frame_btn = ttk.Frame(self)
        frame_btn.pack(fill="x", **pad)
        self.btn_convert = ttk.Button(
            frame_btn, text="Конвертировать", command=self.convert
        )
        self.btn_convert.pack(side="left", padx=6)
        self.btn_cancel = ttk.Button(
            frame_btn, text="Отмена", command=self.cancel, state="disabled"
        )
        self.btn_cancel.pack(side="left", padx=6)

        # Статус
        self.status = ttk.Label(self, text="Готов к работе", foreground="gray")
        self.status.pack(pady=4)

        self._on_format_change()

    def _on_format_change(self):
        state = "normal" if self.format_var.get() == "docx" else "disabled"
        for child in self.frame_docx.winfo_children():
            try:
                child.configure(state=state)
            except tk.TclError:
                pass

    def choose_file(self):
        path = filedialog.askopenfilename(
            title="Выберите XML-файл",
            filetypes=[("XML files", "*.xml"), ("All files", "*.*")],
        )
        if path:
            self.xml_path.set(path)

    def cancel(self):
        self.cancel_flag.set()

    def _set_status(self, text, color="gray"):
        self.after(0, lambda: self.status.config(text=text, foreground=color))

    def _progress_cb(self, count):
        if count % 10000 == 0:
            self._set_status(f"Обработано узлов: {count:,}", "blue")

    def convert(self):
        xml_file = self.xml_path.get().strip()
        if not xml_file or not os.path.isfile(xml_file):
            messagebox.showwarning("Внимание", "Выберите существующий XML-файл")
            return

        fmt = self.format_var.get()
        ext = fmt  # html / docx / pdf

        out_path = filedialog.asksaveasfilename(
            title="Сохранить как",
            defaultextension=f".{ext}",
            filetypes=[(f"{ext.upper()} files", f"*.{ext}")],
            initialfile=os.path.splitext(os.path.basename(xml_file))[0] + f".{ext}",
        )
        if not out_path:
            return

        self.cancel_flag.clear()
        self.btn_convert.config(state="disabled")
        self.btn_cancel.config(state="normal")
        self.progress.start(10)
        self._set_status("Обработка…", "blue")

        chunk_size = max(1000, int(self.chunk_size.get() or 50000))

        def worker():
            try:
                if self.cancel_flag.is_set():
                    raise RuntimeError("Отменено пользователем")

                if fmt == "html":
                    export_html(xml_file, out_path, self._progress_cb)
                    result = [out_path]
                elif fmt == "pdf":
                    export_pdf(xml_file, out_path, self._progress_cb)
                    result = [out_path]
                elif fmt == "docx":
                    result = export_docx_chunked(
                        xml_file, out_path,
                        nodes_per_chunk=chunk_size,
                        progress_cb=self._progress_cb,
                    )
                else:
                    raise ValueError(f"Неизвестный формат: {fmt}")

                msg = "Готово:\n" + "\n".join(result)
                self._set_status(f"Готово ({len(result)} файл(ов))", "green")
                self.after(0, lambda: messagebox.showinfo("Успех", msg))

            except ET.ParseError as e:
                self._set_status("Ошибка XML", "red")
                self.after(0, lambda: messagebox.showerror(
                    "Ошибка XML", f"Не удалось разобрать XML:\n{e}"
                ))
            except Exception as e:
                self._set_status("Ошибка", "red")
                self.after(0, lambda: messagebox.showerror(
                    "Ошибка", f"{type(e).__name__}: {e}"
                ))
            finally:
                self.after(0, self._reset_ui)

        threading.Thread(target=worker, daemon=True).start()

    def _reset_ui(self):
        self.progress.stop()
        self.btn_convert.config(state="normal")
        self.btn_cancel.config(state="disabled")


if __name__ == "__main__":
    App().mainloop()