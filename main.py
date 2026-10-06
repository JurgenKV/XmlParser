import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import xml.etree.ElementTree as ET
from xml.dom import minidom

from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.enums import TA_LEFT


# ---------- Парсинг XML ----------

def element_to_dict(element, level=0):
    """Рекурсивно превращает XML-элемент в структуру dict."""
    node = {
        "tag": element.tag,
        "attrib": dict(element.attrib),
        "text": (element.text or "").strip(),
        "tail": (element.tail or "").strip(),
        "level": level,
        "children": [element_to_dict(child, level + 1) for child in element],
    }
    return node


def flatten_nodes(node, result=None):
    """Разворачивает дерево в плоский список (в порядке обхода)."""
    if result is None:
        result = []
    result.append(node)
    for child in node["children"]:
        flatten_nodes(child, result)
    return result


def parse_xml(path):
    tree = ET.parse(path)
    root = tree.getroot()
    return element_to_dict(root)


# ---------- Экспорт в DOCX ----------

def export_docx(root_node, out_path):
    doc = Document()

    # Заголовок документа
    title = doc.add_heading(f"XML: {root_node['tag']}", level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    nodes = flatten_nodes(root_node)

    for node in nodes:
        indent = "    " * node["level"]

        # Формируем строку заголовка узла
        attrs = ""
        if node["attrib"]:
            attrs = " " + " ".join(f'{k}="{v}"' for k, v in node["attrib"].items())

        # Тег + атрибуты — жирным
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Inches(0.2 * node["level"])

        run = p.add_run(f"{node['tag']}{attrs}")
        run.bold = True
        run.font.size = Pt(12)

        # Текстовое содержимое
        if node["text"]:
            p.add_run(f": {node['text']}")

    doc.save(out_path)


# ---------- Экспорт в PDF ----------

def register_cyrillic_font():
    """Регистрирует шрифт с поддержкой кириллицы."""
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
    return "Helvetica"  # fallback (без кириллицы)


def escape_xml_text(s):
    return (s.replace("&", "&amp;")
             .replace("<", "&lt;")
             .replace(">", "&gt;"))


def export_pdf(root_node, out_path):
    font_name = register_cyrillic_font()

    styles = getSampleStyleSheet()
    tag_style = ParagraphStyle(
        "TagStyle",
        parent=styles["Normal"],
        fontName=font_name,
        fontSize=11,
        leading=14,
        spaceAfter=2,
        alignment=TA_LEFT,
    )
    title_style = ParagraphStyle(
        "TitleStyle",
        parent=styles["Title"],
        fontName=font_name,
        fontSize=18,
        leading=22,
    )

    doc = SimpleDocTemplate(
        out_path,
        pagesize=A4,
        leftMargin=2 * cm, rightMargin=2 * cm,
        topMargin=2 * cm, bottomMargin=2 * cm,
    )

    story = [Paragraph(f"XML: {escape_xml_text(root_node['tag'])}", title_style),
             Spacer(1, 12)]

    nodes = flatten_nodes(root_node)
    for node in nodes:
        indent = "&nbsp;" * (4 * node["level"])

        attrs = ""
        if node["attrib"]:
            attrs = " " + " ".join(
                f'{escape_xml_text(k)}="{escape_xml_text(v)}"'
                for k, v in node["attrib"].items()
            )

        text = f"{indent}<b>{escape_xml_text(node['tag'])}{escape_xml_text(attrs)}</b>"
        if node["text"]:
            text += f": {escape_xml_text(node['text'])}"

        story.append(Paragraph(text, tag_style))

    doc.build(story)


# ---------- GUI ----------

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("XML → DOCX / PDF Конвертер")
        self.geometry("560x300")
        self.resizable(False, False)

        self.xml_path = tk.StringVar()
        self.format_var = tk.StringVar(value="docx")

        self._build_ui()

    def _build_ui(self):
        pad = {"padx": 10, "pady": 6}

        # Выбор файла
        frame_file = ttk.LabelFrame(self, text="1. Выберите XML-файл")
        frame_file.pack(fill="x", **pad)

        ttk.Entry(frame_file, textvariable=self.xml_path, width=55).pack(
            side="left", padx=6, pady=6, fill="x", expand=True
        )
        ttk.Button(frame_file, text="Обзор…", command=self.choose_file).pack(
            side="right", padx=6, pady=6
        )

        # Формат
        frame_fmt = ttk.LabelFrame(self, text="2. Формат вывода")
        frame_fmt.pack(fill="x", **pad)

        for fmt, label in [("docx", "Word (.docx)"), ("pdf", "PDF (.pdf)")]:
            ttk.Radiobutton(
                frame_fmt, text=label, value=fmt, variable=self.format_var
            ).pack(side="left", padx=15, pady=8)

        # Кнопка
        ttk.Button(self, text="Конвертировать", command=self.convert).pack(
            pady=10, ipadx=10
        )

        # Статус
        self.status = ttk.Label(self, text="Готов к работе", foreground="gray")
        self.status.pack(pady=4)

    def choose_file(self):
        path = filedialog.askopenfilename(
            title="Выберите XML-файл",
            filetypes=[("XML files", "*.xml"), ("All files", "*.*")],
        )
        if path:
            self.xml_path.set(path)

    def convert(self):
        xml_file = self.xml_path.get().strip()
        if not xml_file:
            messagebox.showwarning("Внимание", "Сначала выберите XML-файл")
            return
        if not os.path.isfile(xml_file):
            messagebox.showerror("Ошибка", "Файл не найден")
            return

        fmt = self.format_var.get()
        ext = "docx" if fmt == "docx" else "pdf"

        out_path = filedialog.asksaveasfilename(
            title="Сохранить как",
            defaultextension=f".{ext}",
            filetypes=[(f"{ext.upper()} files", f"*.{ext}")],
            initialfile=os.path.splitext(os.path.basename(xml_file))[0] + f".{ext}",
        )
        if not out_path:
            return

        try:
            self.status.config(text="Парсинг XML…", foreground="blue")
            self.update_idletasks()

            root_node = parse_xml(xml_file)

            self.status.config(text="Создание документа…", foreground="blue")
            self.update_idletasks()

            if fmt == "docx":
                export_docx(root_node, out_path)
            else:
                export_pdf(root_node, out_path)

            self.status.config(text=f"Готово: {out_path}", foreground="green")
            messagebox.showinfo("Успех", f"Файл сохранён:\n{out_path}")

        except ET.ParseError as e:
            self.status.config(text="Ошибка парсинга XML", foreground="red")
            messagebox.showerror("Ошибка XML", f"Не удалось разобрать XML:\n{e}")
        except Exception as e:
            self.status.config(text="Ошибка", foreground="red")
            messagebox.showerror("Ошибка", f"Произошла ошибка:\n{e}")


if __name__ == "__main__":
    App().mainloop()