import sys
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox


def extract_lines(input_file, output_file, start_line, end_line=None,
                  encoding='utf-8', progress_cb=None, stop_flag=None):
    """Копирует строки из input_file в output_file (построчно)."""
    start_line = max(1, int(start_line))
    if end_line is not None:
        end_line = int(end_line)
        if end_line < start_line:
            raise ValueError("Строка 'конец' меньше строки 'начало'")

    total_written = 0
    with open(input_file, 'r', encoding=encoding, errors='replace') as fin, \
         open(output_file, 'w', encoding=encoding) as fout:

        for i, line in enumerate(fin, start=1):
            if stop_flag and stop_flag():
                break
            if i < start_line:
                continue
            if end_line is not None and i > end_line:
                break
            fout.write(line)
            total_written += 1
            if progress_cb and total_written % 10_000 == 0:
                progress_cb(i, total_written)

    return total_written


class ExtractorApp:
    def __init__(self, root):
        self.root = root
        root.title("XML Extractor")
        root.geometry("620x340")
        root.resizable(False, False)

        self.stop_requested = False
        self.worker = None

        pad = {'padx': 8, 'pady': 4}

        # --- Input file ---
        ttk.Label(root, text="Входной файл:").grid(row=0, column=0, sticky='w', **pad)
        self.in_var = tk.StringVar()
        ttk.Entry(root, textvariable=self.in_var, width=58).grid(row=0, column=1, columnspan=2, sticky='we', **pad)
        ttk.Button(root, text="…", width=3, command=self.pick_input).grid(row=0, column=3, **pad)

        # --- Output file ---
        ttk.Label(root, text="Выходной файл:").grid(row=1, column=0, sticky='w', **pad)
        self.out_var = tk.StringVar()
        ttk.Entry(root, textvariable=self.out_var, width=58).grid(row=1, column=1, columnspan=2, sticky='we', **pad)
        ttk.Button(root, text="…", width=3, command=self.pick_output).grid(row=1, column=3, **pad)

        # --- Start line ---
        ttk.Label(root, text="Строка начало:").grid(row=2, column=0, sticky='w', **pad)
        self.start_var = tk.StringVar(value="1")
        ttk.Entry(root, textvariable=self.start_var, width=20).grid(row=2, column=1, sticky='w', **pad)

        # --- End line ---
        ttk.Label(root, text="Строка конец:").grid(row=3, column=0, sticky='w', **pad)
        self.end_var = tk.StringVar()
        ttk.Entry(root, textvariable=self.end_var, width=20).grid(row=3, column=1, sticky='w', **pad)
        ttk.Label(root, text="(пусто = до конца файла)").grid(row=3, column=2, sticky='w', **pad)

        # --- Encoding ---
        ttk.Label(root, text="Кодировка:").grid(row=4, column=0, sticky='w', **pad)
        self.enc_var = tk.StringVar(value="utf-8")
        enc_combo = ttk.Combobox(root, textvariable=self.enc_var, width=17,
                                 values=["utf-8", "cp1251", "utf-16", "latin-1"])
        enc_combo.grid(row=4, column=1, sticky='w', **pad)

        # --- Progress ---
        self.progress = ttk.Progressbar(root, mode='indeterminate', length=580)
        self.progress.grid(row=5, column=0, columnspan=4, sticky='we', padx=8, pady=(10, 4))

        self.status_var = tk.StringVar(value="Готов")
        ttk.Label(root, textvariable=self.status_var).grid(row=6, column=0, columnspan=4, sticky='w', padx=8)

        # --- Buttons ---
        btn_frame = ttk.Frame(root)
        btn_frame.grid(row=7, column=0, columnspan=4, pady=10)

        self.run_btn = ttk.Button(btn_frame, text="▶ Запустить", command=self.run)
        self.run_btn.pack(side='left', padx=5)

        self.stop_btn = ttk.Button(btn_frame, text="■ Стоп", command=self.stop, state='disabled')
        self.stop_btn.pack(side='left', padx=5)

        ttk.Button(btn_frame, text="Выход", command=root.destroy).pack(side='left', padx=5)

    # --- file pickers ---
    def pick_input(self):
        p = filedialog.askopenfilename(
            title="Выберите входной файл",
            filetypes=[("XML/Text", "*.xml *.txt *.log"), ("Все файлы", "*.*")]
        )
        if p:
            self.in_var.set(p)
            # подставим имя выходного файла
            if not self.out_var.get():
                self.out_var.set(p + ".part.txt")

    def pick_output(self):
        p = filedialog.asksaveasfilename(
            title="Куда сохранить",
            defaultextension=".txt",
            filetypes=[("Text", "*.txt"), ("Все файлы", "*.*")]
        )
        if p:
            self.out_var.set(p)

    # --- controls ---
    def run(self):
        inp = self.in_var.get().strip()
        out = self.out_var.get().strip()

        if not inp or not out:
            messagebox.showerror("Ошибка", "Укажите входной и выходной файлы")
            return

        try:
            start = int(self.start_var.get().strip() or "1")
        except ValueError:
            messagebox.showerror("Ошибка", "Строка 'начало' должна быть числом")
            return

        end_str = self.end_var.get().strip()
        end = None
        if end_str:
            try:
                end = int(end_str)
            except ValueError:
                messagebox.showerror("Ошибка", "Строка 'конец' должна быть числом")
                return

        enc = self.enc_var.get().strip() or 'utf-8'

        # сброс флага и блокировка UI
        self.stop_requested = False
        self.run_btn.config(state='disabled')
        self.stop_btn.config(state='normal')
        self.progress.start(10)
        self.status_var.set("Работаю…")

        def progress_cb(line_no, written):
            self.root.after(0, lambda: self.status_var.set(
                f"Обработано строк: {line_no:,} | записано: {written:,}"
            ))

        def stop_flag():
            return self.stop_requested

        def job():
            try:
                n = extract_lines(inp, out, start, end, enc,
                                  progress_cb=progress_cb, stop_flag=stop_flag)
                self.root.after(0, lambda: self.finish_ok(n))
            except Exception as e:
                self.root.after(0, lambda: self.finish_err(e))

        self.worker = threading.Thread(target=job, daemon=True)
        self.worker.start()

    def stop(self):
        self.stop_requested = True
        self.status_var.set("Останавливаю…")

    def finish_ok(self, n):
        self.progress.stop()
        self.run_btn.config(state='normal')
        self.stop_btn.config(state='disabled')
        if self.stop_requested:
            self.status_var.set(f"Остановлено. Записано строк: {n:,}")
        else:
            self.status_var.set(f"Готово. Записано строк: {n:,}")
            messagebox.showinfo("Готово", f"Записано строк: {n:,}\n\nФайл: {self.out_var.get()}")

    def finish_err(self, e):
        self.progress.stop()
        self.run_btn.config(state='normal')
        self.stop_btn.config(state='disabled')
        self.status_var.set("Ошибка")
        messagebox.showerror("Ошибка", str(e))


if __name__ == "__main__":
    root = tk.Tk()
    try:
        # аккуратный вид на Windows
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    ExtractorApp(root)
    root.mainloop()