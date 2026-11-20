from __future__ import annotations

import re
import threading
import queue
import os
import sys
import json
import io
import ctypes
from ctypes import wintypes
from collections import Counter
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

try:
    import pymupdf as fitz
    import pytesseract
    from PIL import Image
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
except ImportError as exc:
    raise SystemExit(
        "کتابخانه‌های لازم نصب نیستند. در PowerShell این دستور را اجرا کنید:\n"
        "py -m pip install -r requirements.txt"
    ) from exc


def app_directory() -> Path:
    return Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent


def bundled_directory() -> Path:
    return Path(getattr(sys, "_MEIPASS", app_directory()))


def find_ocr_runtime() -> tuple[Path | None, Path | None]:
    roots = [app_directory() / "ocr" / "tesseract", bundled_directory() / "ocr" / "tesseract",
             app_directory() / "resources" / "ocr" / "tesseract"]
    for root in roots:
        executable = root / "tesseract.exe"
        tessdata = root / "tessdata"
        if executable.is_file() and tessdata.is_dir():
            return executable, tessdata
    return None, None


def settings_file() -> Path:
    base = Path(os.environ.get("APPDATA", Path.home())) / "PDFKeywordFinder"
    return base / "settings.json"


def normalize(text: str) -> str:
    """Normalize common Persian/Arabic letter, diacritic, digit, and spacing variants."""
    text = text.translate(str.maketrans({
        "ي": "ی", "ى": "ی", "ے": "ی", "ې": "ی", "ك": "ک", "ڪ": "ک",
        "ۀ": "ه", "ة": "ه", "ؤ": "و", "أ": "ا", "إ": "ا", "ٱ": "ا", "آ": "ا",
        "ـ": "",
    })).casefold()
    # Remove Arabic/Persian vowel marks and Quranic annotation marks.
    text = re.sub(r"[\u064b-\u065f\u0670\u06d6-\u06ed]", "", text)
    # Normalize Arabic-Indic and Eastern Arabic-Indic numerals to ASCII digits.
    text = text.translate(str.maketrans("٠١٢٣٤٥٦٧٨٩۰۱۲۳۴۵۶۷۸۹", "01234567890123456789"))
    # Keep one canonical space; matching later allows it to be absent or a ZWNJ.
    text = re.sub(r"[\s\u200c\u200d\ufeff]+", " ", text)
    return text.strip()


def term_pattern(normalized_term: str) -> re.Pattern[str]:
    """Allow a term's spaces to appear as spaces, ZWNJs, or no separator."""
    if not normalized_term.strip():
        return re.compile(r"(?!)")
    pieces = re.split(r"\s+", normalized_term.strip())
    return re.compile(r"[\s\u200c\u200d\ufeff]*".join(re.escape(piece) for piece in pieces))


def write_workbook(destination: Path, terms: list[str], matches: list[tuple[str, str, str, int, int]],
                   errors: list[tuple[str, str]]) -> None:
    book = Workbook()
    results = book.active
    results.title = "نتایج"
    results.append(["واژه یا عبارت", "نام فایل PDF", "مسیر نسبی", "شماره صفحه", "تعداد تکرار در صفحه"])
    for term, filename, relative_path, page, count in matches:
        results.append([term, filename, relative_path, page, count])

    file_counts: Counter[tuple[str, str, str]] = Counter()
    for term, _filename, relative_path, _page, count in matches:
        file_counts[(term, relative_path, _filename)] += count
    file_summary = book.create_sheet("شمارش در هر فایل")
    file_summary.append(["واژه یا عبارت", "نام فایل PDF", "مسیر نسبی", "تعداد تکرار در فایل"])
    for (term, relative_path, filename), count in sorted(file_counts.items()):
        file_summary.append([term, filename, relative_path, count])

    summary = book.create_sheet("خلاصه")
    summary.append(["واژه یا عبارت", "تعداد فایل‌های دارای نتیجه", "تعداد صفحه‌های دارای نتیجه", "مجموع تکرارها"])
    for term in terms:
        rows = [row for row in matches if row[0] == term]
        summary.append([term, len({row[2] for row in rows}), len(rows), sum(row[4] for row in rows)])

    if errors:
        problems = book.create_sheet("فایل‌های مشکل‌دار")
        problems.append(["نام فایل PDF", "خطا"])
        for filename, error in errors:
            problems.append([filename, error])

    for sheet in book.worksheets:
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = sheet.dimensions
        for cell in sheet[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="245A75")
            cell.alignment = Alignment(horizontal="center")
        for column in sheet.columns:
            letter = get_column_letter(column[0].column)
            width = min(max(max(len(str(c.value or "")) for c in column) + 2, 14), 60)
            sheet.column_dimensions[letter].width = width
        for row in sheet.iter_rows(min_row=2):
            for cell in row:
                cell.alignment = Alignment(vertical="top", wrap_text=True)
    book.save(destination)


class WindowsRtlTextBox:
    """A native Windows Rich Edit input with real bidirectional text handling."""
    def __init__(self, parent: tk.Widget) -> None:
        self.parent = parent
        self.library = ctypes.WinDLL("Msftedit.dll")
        self.user32 = ctypes.WinDLL("user32", use_last_error=True)
        self.user32.CreateWindowExW.restype = wintypes.HWND
        self.user32.CreateWindowExW.argtypes = [
            wintypes.DWORD, wintypes.LPCWSTR, wintypes.LPCWSTR, wintypes.DWORD,
            ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int,
            wintypes.HWND, wintypes.HMENU, wintypes.HINSTANCE, wintypes.LPVOID,
        ]
        self.user32.SendMessageW.restype = ctypes.c_ssize_t
        self.user32.SendMessageW.argtypes = [wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM]
        self.user32.SetWindowPos.restype = wintypes.BOOL
        self.user32.SetWindowPos.argtypes = [wintypes.HWND, wintypes.HWND, ctypes.c_int, ctypes.c_int,
                                              ctypes.c_int, ctypes.c_int, wintypes.UINT]
        self.user32.GetWindowTextLengthW.argtypes = [wintypes.HWND]
        self.user32.GetWindowTextLengthW.restype = ctypes.c_int
        self.user32.GetWindowTextW.argtypes = [wintypes.HWND, wintypes.LPWSTR, ctypes.c_int]
        self.user32.GetWindowTextW.restype = ctypes.c_int
        self.user32.DestroyWindow.argtypes = [wintypes.HWND]
        self.user32.DestroyWindow.restype = wintypes.BOOL

        self.parent.update_idletasks()
        self.hwnd = self.user32.CreateWindowExW(
            0x00000200 | 0x00001000 | 0x00002000,  # client edge, right aligned, RTL reading
            "RICHEDIT50W", "",
            0x40000000 | 0x10000000 | 0x00200000 | 0x0004 | 0x0040 | 0x1000 | 0x0100,
            0, 0, max(parent.winfo_width(), 100), max(parent.winfo_height(), 100),
            parent.winfo_id(), None, self.library._handle, None,
        )
        if not self.hwnd:
            raise ctypes.WinError(ctypes.get_last_error())

        # Plain-text mode enables contextual bidi reading and paragraph alignment.
        self.user32.SendMessageW(self.hwnd, 0x400 + 89, 0x0001, 0)  # EM_SETTEXTMODE, TM_PLAINTEXT

        class BidiOptions(ctypes.Structure):
            _fields_ = [("cbSize", wintypes.UINT), ("wMask", wintypes.WORD), ("wEffects", wintypes.WORD)]

        options = BidiOptions(ctypes.sizeof(BidiOptions), 0x0008 | 0x0010, 0x0008 | 0x0010)
        self.user32.SendMessageW(self.hwnd, 0x400 + 200, 0, ctypes.addressof(options))  # EM_SETBIDIOPTIONS
        self.parent.bind("<Configure>", self._resize)
        self.parent.bind("<Destroy>", self._destroy)

    def _resize(self, event: tk.Event) -> None:
        if event.widget is self.parent and self.hwnd:
            self.user32.SetWindowPos(self.hwnd, None, 0, 0, max(event.width, 1), max(event.height, 1), 0x0004)

    def _destroy(self, event: tk.Event) -> None:
        if event.widget is self.parent and self.hwnd:
            self.user32.DestroyWindow(self.hwnd)
            self.hwnd = None

    def get(self, _start: str = "1.0", _end: str = "end") -> str:
        if not self.hwnd:
            return ""
        length = self.user32.GetWindowTextLengthW(self.hwnd)
        buffer = ctypes.create_unicode_buffer(length + 1)
        self.user32.GetWindowTextW(self.hwnd, buffer, length + 1)
        return buffer.value


class App(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.settings_path = settings_file()
        self.settings = self._load_settings()
        self.title("پیدا — جست‌وجوی هوشمند PDF")
        self.geometry("820x740")
        self.minsize(700, 650)
        self.configure(padx=20, pady=16)
        self.style = ttk.Style(self)
        self.style.theme_use("clam")
        self.dark_mode = tk.BooleanVar(value=bool(self.settings.get("dark_mode", False)))
        self.ocr_enabled = tk.BooleanVar(value=bool(self.settings.get("ocr_enabled", True)))
        self.ocr_language = tk.StringVar(value=self.settings.get("ocr_language", "فارسی + انگلیسی"))
        self.folder = tk.StringVar()
        self.recursive = tk.BooleanVar(value=True)
        self.events: queue.Queue = queue.Queue()
        self.cancel_event = threading.Event()
        self.running = False
        self._build()
        self._apply_theme()
        self._apply_app_icon()
        self.after(120, self._poll_events)

    def _load_settings(self) -> dict:
        try:
            return json.loads(settings_file().read_text(encoding="utf-8"))
        except (OSError, ValueError):
            return {}

    def _save_settings(self) -> None:
        self.settings.update({
            "dark_mode": self.dark_mode.get(),
            "ocr_enabled": self.ocr_enabled.get(),
            "ocr_language": self.ocr_language.get(),
        })
        try:
            self.settings_path.parent.mkdir(parents=True, exist_ok=True)
            self.settings_path.write_text(json.dumps(self.settings, ensure_ascii=False, indent=2), encoding="utf-8")
        except OSError:
            pass

    def _apply_theme(self) -> None:
        dark = self.dark_mode.get()
        bg, panel = ("#20252b", "#2b323a") if dark else ("#f3f6f8", "#ffffff")
        fg = "#edf2f7" if dark else "#1f2933"
        accent = "#55b5a6" if dark else "#167d75"
        self.configure(bg=bg)
        self.style.configure("TFrame", background=bg)
        self.style.configure("TLabel", background=bg, foreground=fg, font=("Segoe UI", 10))
        self.style.configure("Title.TLabel", background=bg, foreground=accent, font=("Segoe UI", 19, "bold"))
        self.style.configure("TButton", background=panel, foreground=fg, padding=(10, 6), font=("Segoe UI", 10))
        self.style.map("TButton", background=[("active", accent)], foreground=[("active", "white")])
        self.style.configure("TCheckbutton", background=bg, foreground=fg, font=("Segoe UI", 10))
        self.style.map("TCheckbutton", background=[("active", bg)], foreground=[("active", accent)])
        self.style.configure("TEntry", fieldbackground=panel, foreground=fg)
        self.style.configure("TCombobox", fieldbackground=panel, foreground=fg)
        self.style.configure("Horizontal.TProgressbar", background=accent, troughcolor=panel)
        if hasattr(self, "theme_button"):
            self.theme_button.configure(text="حالت روشن" if dark else "حالت تیره")
        if hasattr(self, "terms_box") and isinstance(self.terms_box, tk.Text):
            self.terms_box.configure(background=panel, foreground=fg, insertbackground=fg)

    def _toggle_theme(self) -> None:
        self.dark_mode.set(not self.dark_mode.get())
        self._apply_theme()
        self._save_settings()

    def _apply_app_icon(self) -> None:
        icon_paths = [bundled_directory() / "assets" / "pdf_finder.ico", app_directory() / "assets" / "pdf_finder.ico"]
        for icon_path in icon_paths:
            if icon_path.is_file():
                try:
                    self.iconbitmap(str(icon_path))
                    return
                except tk.TclError:
                    continue

    def _build(self) -> None:
        header = ttk.Frame(self)
        header.pack(fill="x", pady=(0, 12))
        self.theme_button = ttk.Button(header, text="حالت تیره", command=self._toggle_theme)
        self.theme_button.pack(side="left")
        ttk.Label(header, text="پیدا | جست‌وجوی واژه در PDF", style="Title.TLabel").pack(side="right")
        ttk.Label(self, text="پوشهٔ PDFها:").pack(anchor="e")
        folder_row = ttk.Frame(self)
        folder_row.pack(fill="x", pady=(4, 10))
        ttk.Button(folder_row, text="انتخاب پوشه", command=self._choose_folder).pack(side="right")
        ttk.Entry(folder_row, textvariable=self.folder, justify="right").pack(side="right", fill="x", expand=True, padx=(0, 8))

        ttk.Label(self, text="واژه‌ها یا عبارت‌ها (هر مورد در یک خط):").pack(anchor="e")
        self.terms_frame = ttk.Frame(self, height=160)
        self.terms_frame.pack(fill="x", pady=(4, 8))
        self.terms_frame.pack_propagate(False)
        if os.name == "nt":
            try:
                self.terms_box = WindowsRtlTextBox(self.terms_frame)
            except (OSError, AttributeError):
                self.terms_box = self._fallback_terms_box()
        else:
            self.terms_box = self._fallback_terms_box()

        ocr_row = ttk.Frame(self)
        ocr_row.pack(fill="x", pady=(0, 7))
        ttk.Checkbutton(ocr_row, text="OCR آفلاین برای صفحه‌های اسکن‌شده", variable=self.ocr_enabled,
                        command=self._save_settings).pack(side="right")
        ttk.Label(ocr_row, text="زبان:").pack(side="right", padx=(14, 4))
        self.language_box = ttk.Combobox(ocr_row, textvariable=self.ocr_language, state="readonly", width=19,
                                         values=("فارسی + انگلیسی", "فقط فارسی", "فقط انگلیسی"))
        self.language_box.pack(side="right")
        self.language_box.bind("<<ComboboxSelected>>", lambda _event: self._save_settings())

        ttk.Checkbutton(self, text="پوشه‌های داخلی را هم جست‌وجو کن", variable=self.recursive).pack(anchor="e", pady=(0, 10))
        ttk.Label(self, text="گزارش Excel تعداد تکرار را برای هر صفحه و فایل نشان می‌دهد. OCR هنگام نیاز، کاملاً آفلاین اجرا می‌شود.", wraplength=760, justify="right").pack(anchor="e", pady=(0, 12))

        controls = ttk.Frame(self)
        controls.pack(fill="x")
        self.start_button = ttk.Button(controls, text="شروع جست‌وجو و ذخیرهٔ Excel", command=self._start)
        self.start_button.pack(side="right")
        self.cancel_button = ttk.Button(controls, text="لغو", command=self.cancel_event.set, state="disabled")
        self.cancel_button.pack(side="right", padx=8)
        self.progress = ttk.Progressbar(self, mode="determinate")
        self.progress.pack(fill="x", pady=(12, 4))
        self.status = ttk.Label(self, text="آماده", anchor="e")
        self.status.pack(fill="x")

    def _fallback_terms_box(self) -> tk.Text:
        box = tk.Text(self.terms_frame, height=8, wrap="word", font=("Segoe UI", 11), undo=True)
        box.pack(fill="both", expand=True)
        box.tag_configure("right", justify="right")
        box.tag_add("right", "1.0", "end")
        box.bind("<KeyRelease>", lambda _event: box.tag_add("right", "1.0", "end"))
        return box

    def _choose_folder(self) -> None:
        selected = filedialog.askdirectory(title="پوشهٔ PDFها را انتخاب کنید")
        if selected:
            self.folder.set(selected)

    def _start(self) -> None:
        root = Path(self.folder.get().strip())
        terms = list(dict.fromkeys(line.strip() for line in self.terms_box.get("1.0", "end").splitlines() if line.strip()))
        if not root.is_dir():
            messagebox.showerror("پوشه انتخاب نشده", "لطفاً یک پوشهٔ معتبر انتخاب کنید.")
            return
        if not terms:
            messagebox.showerror("واژه‌ای وارد نشده", "حداقل یک واژه یا عبارت وارد کنید.")
            return
        ocr_enabled = self.ocr_enabled.get()
        ocr_engine, tessdata = find_ocr_runtime()
        language_codes = {"فارسی + انگلیسی": "fas+eng", "فقط فارسی": "fas", "فقط انگلیسی": "eng"}
        ocr_lang = language_codes.get(self.ocr_language.get(), "fas+eng")
        if ocr_enabled:
            if not ocr_engine or not tessdata:
                messagebox.showerror("موتور OCR پیدا نشد", "فایل‌های OCR همراه برنامه نیستند. پوشهٔ کامل برنامه را استخراج کنید و دوباره اجرا کنید.")
                return
            missing_languages = [code for code in ocr_lang.split("+") if not (tessdata / f"{code}.traineddata").is_file()]
            if missing_languages:
                messagebox.showerror("مدل زبان پیدا نشد", "مدل OCR این زبان‌ها در بسته موجود نیست: " + ", ".join(missing_languages))
                return
        destination = filedialog.asksaveasfilename(
            title="ذخیرهٔ گزارش Excel", defaultextension=".xlsx",
            initialfile="گزارش جست‌وجوی PDF.xlsx",
            filetypes=[("Excel Workbook", "*.xlsx")],
        )
        if not destination:
            return
        candidates = root.rglob("*") if self.recursive.get() else root.glob("*")
        pdfs = sorted(path for path in candidates if path.is_file() and path.suffix.casefold() == ".pdf")
        if not pdfs:
            messagebox.showinfo("PDF پیدا نشد", "در پوشهٔ انتخاب‌شده فایل PDF پیدا نشد.")
            return
        self.cancel_event.clear()
        self.running = True
        self._save_settings()
        self.start_button.configure(state="disabled")
        self.cancel_button.configure(state="normal")
        self.progress.configure(maximum=len(pdfs), value=0)
        self.status.configure(text=f"۰ از {len(pdfs)} فایل")
        threading.Thread(target=self._search, args=(root, pdfs, terms, Path(destination), ocr_enabled, ocr_lang,
                                                    ocr_engine, tessdata), daemon=True).start()

    def _search(self, root: Path, pdfs: list[Path], terms: list[str], destination: Path,
                ocr_enabled: bool, ocr_lang: str, ocr_engine: Path | None, tessdata: Path | None) -> None:
        if ocr_enabled and ocr_engine and tessdata:
            pytesseract.pytesseract.tesseract_cmd = str(ocr_engine)
        normalized_terms = [(term, term_pattern(normalize(term))) for term in terms]
        matches: list[tuple[str, str, str, int, int]] = []
        errors: list[tuple[str, str]] = []
        ocr_pages = 0
        for index, pdf in enumerate(pdfs, 1):
            if self.cancel_event.is_set():
                self.events.put(("cancelled", index - 1, len(pdfs)))
                return
            try:
                with fitz.open(pdf) as document:
                    for page_number, page in enumerate(document, 1):
                        raw_text = page.get_text("text")
                        if ocr_enabled and len(re.sub(r"\s+", "", raw_text)) < 25:
                            self.events.put(("status", f"OCR: {pdf.name} — صفحهٔ {page_number}"))
                            try:
                                pixmap = page.get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
                                image = Image.open(io.BytesIO(pixmap.tobytes("png")))
                                config = f'--tessdata-dir "{tessdata}" --psm 3'
                                raw_text = pytesseract.image_to_string(image, lang=ocr_lang, config=config)
                                ocr_pages += 1
                            except Exception as exc:
                                errors.append((pdf.name, f"صفحهٔ {page_number}: خطای OCR: {str(exc)[:400]}"))
                        page_text = normalize(raw_text)
                        for original, pattern in normalized_terms:
                            count = sum(1 for _match in pattern.finditer(page_text))
                            if count:
                                matches.append((original, pdf.name, str(pdf.relative_to(root)), page_number, count))
            except Exception as exc:
                errors.append((pdf.name, str(exc)[:500]))
            self.events.put(("progress", index, len(pdfs), pdf.name))
        try:
            write_workbook(destination, terms, matches, errors)
        except Exception as exc:
            self.events.put(("failed", str(exc)))
            return
        self.events.put(("done", str(destination), len(matches), len(pdfs), len(errors), ocr_pages))

    def _poll_events(self) -> None:
        try:
            while True:
                event = self.events.get_nowait()
                kind = event[0]
                if kind == "progress":
                    self.progress.configure(value=event[1])
                    self.status.configure(text=f"{event[1]} از {event[2]} فایل بررسی شد — {event[3]}")
                elif kind == "status":
                    self.status.configure(text=event[1])
                elif kind == "done":
                    self.running = False
                    self._set_idle()
                    messagebox.showinfo("جست‌وجو تمام شد", f"{event[2]} نتیجه در {event[1]} ذخیره شد.\nفایل‌های بررسی‌شده: {event[3]}\nصفحه‌های OCRشده: {event[5]}\nخطاها: {event[4]}")
                elif kind == "cancelled":
                    self.running = False
                    self._set_idle()
                    self.status.configure(text="جست‌وجو لغو شد؛ گزارشی ذخیره نشد.")
                elif kind == "failed":
                    self.running = False
                    self._set_idle()
                    messagebox.showerror("خطا در ذخیرهٔ گزارش", event[1])
        except queue.Empty:
            pass
        self.after(120, self._poll_events)

    def _set_idle(self) -> None:
        self.start_button.configure(state="normal")
        self.cancel_button.configure(state="disabled")


if __name__ == "__main__":
    App().mainloop()
