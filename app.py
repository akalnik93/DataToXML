"""Простая Windows-обёртка для тестового генератора XML межевого плана."""

from __future__ import annotations

import queue
import re
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from prototype_generate import build_application_only_xml, docx_text


APP_DIR = Path(__file__).resolve().parent
TEMPLATE = APP_DIR / "shablon_mezhevogo_plana.xml"


class XmlDraftApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("МП на уточнение границ")
        self.root.geometry("650x470")
        self.root.minsize(520, 390)
        self.files: list[Path] = []
        self.results: queue.Queue = queue.Queue()

        outer = ttk.Frame(root, padding=16)
        outer.pack(fill="both", expand=True)
        ttk.Label(
            outer,
            text="МП на уточнение границ",
            font=("Segoe UI", 15, "bold"),
        ).pack(anchor="w")
        ttk.Label(
            outer,
            text="Выберите документы DOCX. Индикаторы показывают, какие из трёх ключевых файлов распознаны; они не блокируют создание XML.",
            wraplength=600,
        ).pack(anchor="w", pady=(6, 12))

        checks = ttk.Frame(outer)
        checks.pack(fill="x", pady=(0, 8))
        self.document_checks: dict[str, tk.Label] = {}
        for key, title in (("application", "Заявление"), ("contract", "Договор"), ("consent", "ОПД")):
            row = ttk.Frame(checks)
            row.pack(side="left", padx=(0, 24))
            ttk.Label(row, text=f"{title}:").pack(side="left")
            indicator = tk.Label(row, text="✗", fg="#C62828", font=("Segoe UI Symbol", 12, "bold"))
            indicator.pack(side="left", padx=(5, 0))
            self.document_checks[key] = indicator

        self.listbox = tk.Listbox(outer, height=10, activestyle="dotbox")
        self.listbox.pack(fill="both", expand=True)

        controls = ttk.Frame(outer)
        controls.pack(fill="x", pady=(12, 8))
        self.choose_button = ttk.Button(controls, text="Список документов", command=self.choose_files)
        self.choose_button.pack(side="left")
        self.create_button = ttk.Button(controls, text="Создать XML из документов", command=self.prepare)
        self.create_button.pack(side="left", padx=(8, 0))

        self.status = tk.StringVar(value="Готово к работе")
        ttk.Label(outer, textvariable=self.status, wraplength=600).pack(anchor="w")

    def choose_files(self) -> None:
        selected = filedialog.askopenfilenames(
            title="Выберите документы",
            filetypes=[("Документы", "*.docx *.pdf *.xml"), ("Все файлы", "*.*")],
        )
        if not selected:
            return
        self.files = [Path(item) for item in selected]
        self._update_document_checks()
        self.listbox.delete(0, tk.END)
        for path in self.files:
            self.listbox.insert(tk.END, str(path))
        self.status.set(f"Выбрано файлов: {len(self.files)}")

    def _update_document_checks(self) -> None:
        detected = {"application": False, "contract": False, "consent": False}
        patterns = {
            "application": r"ЗАЯВЛЕНИЕ\s*№",
            "contract": r"ДОГОВОР\s+ПОДРЯДА",
            "consent": r"СОГЛАСИЕ\s+на\s+обработку\s+персональных\s+данных",
        }
        for path in self.files:
            if path.suffix.lower() != ".docx":
                continue
            try:
                content = docx_text(path)
            except Exception:
                continue
            for key, pattern in patterns.items():
                if not detected[key] and re.search(pattern, content, re.I):
                    detected[key] = True
        for key, indicator in self.document_checks.items():
            if detected[key]:
                indicator.configure(text="✓", fg="#16823B")
            else:
                indicator.configure(text="✗", fg="#C62828")

    def prepare(self) -> None:
        if not self.files:
            messagebox.showinfo("Документы не выбраны", "Сначала нажмите «Список документов».")
            return
        if not TEMPLATE.is_file():
            messagebox.showerror(
                "Не найден шаблон",
                f"Положите файл шаблона рядом с программой:\n{TEMPLATE.name}",
            )
            return

        self.choose_button.configure(state="disabled")
        self.create_button.configure(state="disabled")
        self.status.set("Создаю черновик: анализирую заявление, договор и ОПД…")
        threading.Thread(target=self._worker, daemon=True).start()
        self.root.after(100, self._poll_result)

    def _worker(self) -> None:
        try:
            data, summary = build_application_only_xml(self.files, TEMPLATE)
            names = ", ".join(path.name for path in self.files if path.suffix.lower() == ".docx")
            self.results.put(("success", data, summary, names))
        except Exception as exc:
            self.results.put(("error", str(exc)))

    def _poll_result(self) -> None:
        try:
            result = self.results.get_nowait()
        except queue.Empty:
            self.root.after(100, self._poll_result)
            return

        self.choose_button.configure(state="normal")
        self.create_button.configure(state="normal")
        if result[0] == "error":
            self.status.set("Не удалось подготовить XML")
            messagebox.showerror("Ошибка подготовки", result[1])
            return

        _, data, summary, source_name = result
        self.status.set("Подготовка завершена. Выберите имя и место сохранения XML.")
        output = filedialog.asksaveasfilename(
            title="Сохранить черновик XML",
            initialfile="draft_mpw_0.0.12.xml",
            defaultextension=".xml",
            filetypes=[("XML-файлы", "*.xml")],
        )
        if not output:
            self.status.set("Сохранение отменено")
            return
        try:
            Path(output).write_bytes(data)
        except OSError as exc:
            self.status.set("Ошибка сохранения")
            messagebox.showerror("Ошибка сохранения", str(exc))
            return
        self.status.set(f"Создан черновик из {source_name}: {summary}")
        messagebox.showinfo(
            "XML создан",
            f"Файл сохранён:\n{output}\n\nВзято из {source_name}: {summary}\n\n"
            "Числовые значения и координаты участка сохранены из шаблона и требуют проверки инженером. "
            "Это черновик для проверки и дозаполнения в АРГО.",
        )


def main() -> None:
    root = tk.Tk()
    XmlDraftApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
