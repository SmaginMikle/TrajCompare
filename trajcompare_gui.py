# -*- coding: utf-8 -*-
"""
TrajCompare — оконная программа сравнения траекторий скважин РН-ПД и РНТ.
Точка входа для сборки в .exe (PyInstaller).

Без аргументов — открывается окно.
С аргументами — работает как консольная утилита (те же команды, что у trajcompare.py):
    TrajCompare.exe batch --root D:\\wells -o D:\\out --workers 8
"""
import multiprocessing
import os
import sys

# В оконной сборке нет консоли: перенаправляем вывод, чтобы print/logging не падали
if sys.stdout is None:
    sys.stdout = open(os.devnull, "w", encoding="utf-8")
if sys.stderr is None:
    sys.stderr = open(os.devnull, "w", encoding="utf-8")

import json
import logging
import queue
import subprocess
import threading
import traceback
from pathlib import Path
from types import SimpleNamespace

APP_TITLE = "TrajCompare — сравнение траекторий скважин"
APP_DIR = Path(sys.executable).parent if getattr(sys, "frozen", False) else Path(__file__).parent
SETTINGS_FILE = APP_DIR / "TrajCompare_settings.json"


def open_path(p):
    p = str(p)
    try:
        if sys.platform.startswith("win"):
            os.startfile(p)  # noqa
        elif sys.platform == "darwin":
            subprocess.Popen(["open", p])
        else:
            subprocess.Popen(["xdg-open", p])
    except Exception:  # noqa: BLE001
        pass


class QueueHandler(logging.Handler):
    def __init__(self, q):
        super().__init__()
        self.q = q

    def emit(self, record):
        self.q.put(("log", self.format(record)))


def run_gui(selftest=False):
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox
    import trajcompare as tc

    root = tk.Tk()
    root.title(APP_TITLE)
    root.geometry("900x660")
    root.minsize(780, 560)
    try:
        ttk.Style().theme_use("vista" if sys.platform.startswith("win") else "clam")
    except tk.TclError:
        pass

    q: "queue.Queue" = queue.Queue()
    h = QueueHandler(q)
    h.setFormatter(logging.Formatter("%(asctime)s  %(message)s", "%H:%M:%S"))
    logging.getLogger("trajcompare").addHandler(h)
    logging.getLogger("trajcompare").setLevel(logging.INFO)

    # ---------- настройки ----------
    saved = {}
    if SETTINGS_FILE.exists():
        try:
            saved = json.loads(SETTINGS_FILE.read_text(encoding="utf-8"))
        except Exception:  # noqa: BLE001
            saved = {}
    S = {k: tk.StringVar(value=str(saved.get(k, v))) for k, v in tc.DEFAULTS.items()
         if k in ("step_md", "wellhead_tol", "target_tol", "sep_tol", "sep_warn", "landing_inc", "ddi_tol")}
    plots_var = tk.BooleanVar(value=saved.get("plots", True))
    traj_var = tk.BooleanVar(value=saved.get("traj_sheets", True))
    METHOD_BY_NAME = {v: k for k, v in tc.DDI_METHODS.items()}
    ddi_method_var = tk.StringVar(value=tc.DDI_METHODS.get(saved.get("ddi_method", "oag"), tc.DDI_METHODS["oag"]))
    paths = {k: tk.StringVar(value=saved.get("_" + k, "")) for k in
             ("ref", "alt", "out_pair", "root", "ref_dir", "alt_dir", "manifest", "out_batch",
              "wb_ref", "wb_alt", "wb_out")}
    mode = tk.StringVar(value=saved.get("_mode", "root"))
    ref_pat = tk.StringVar(value=saved.get("_ref_pattern", r"рн-?пд|точн|exact|ref|факт|план"))
    alt_pat = tk.StringVar(value=saved.get("_alt_pattern", r"рнт|альтерн|alt"))
    workers = tk.IntVar(value=saved.get("_workers", max(1, (os.cpu_count() or 2) - 1)))
    skip_var = tk.BooleanVar(value=saved.get("_skip", False))

    def cfg():
        c = dict(tc.DEFAULTS)
        for k, v in S.items():
            try:
                c[k] = float(v.get().replace(",", "."))
            except ValueError:
                raise ValueError(f"Неверное значение настройки «{LABELS[k]}»: {v.get()}")
        c["plots"] = bool(plots_var.get())
        c["traj_sheets"] = bool(traj_var.get())
        c["ddi_method"] = METHOD_BY_NAME.get(ddi_method_var.get(), "oag")
        return c

    def save_settings():
        d = {k: v.get() for k, v in S.items()}
        d["plots"] = bool(plots_var.get())
        d["traj_sheets"] = bool(traj_var.get())
        d["ddi_method"] = METHOD_BY_NAME.get(ddi_method_var.get(), "oag")
        d.update({"_" + k: v.get() for k, v in paths.items()})
        d.update(_mode=mode.get(), _ref_pattern=ref_pat.get(), _alt_pattern=alt_pat.get(),
                 _workers=int(workers.get()), _skip=bool(skip_var.get()), _csv_sep=SEP_NAMES.get(csv_sep.get(), "auto"))
        try:
            SETTINGS_FILE.write_text(json.dumps(d, ensure_ascii=False, indent=1), encoding="utf-8")
        except Exception:  # noqa: BLE001
            pass

    # ---------- общие элементы ----------
    TRAJ_FT = [("Траектории (Excel, CSV, TXT)", "*.xlsx *.xlsm *.xls *.csv *.txt"), ("Excel", "*.xlsx *.xlsm *.xls"),
               ("CSV / TXT", "*.csv *.txt"), ("Все файлы", "*.*")]

    def path_row2(parent, r, label, var):
        """Строка выбора: файл или папка."""
        ttk.Label(parent, text=label).grid(row=r, column=0, sticky="w", padx=6, pady=4)
        ttk.Entry(parent, textvariable=var).grid(row=r, column=1, sticky="ew", padx=6, pady=4)
        fr = ttk.Frame(parent)
        fr.grid(row=r, column=2, padx=6)

        def pick_dir():
            p = filedialog.askdirectory()
            if p:
                var.set(p)

        def pick_file():
            p = filedialog.askopenfilename(filetypes=TRAJ_FT)
            if p:
                var.set(p)
        ttk.Button(fr, text="Папка…", command=pick_dir).pack(side="left")
        ttk.Button(fr, text="Файл…", command=pick_file).pack(side="left", padx=(4, 0))

    def path_row(parent, r, label, var, kind, ftypes=None):
        ttk.Label(parent, text=label).grid(row=r, column=0, sticky="w", padx=6, pady=4)
        ttk.Entry(parent, textvariable=var).grid(row=r, column=1, sticky="ew", padx=6, pady=4)

        def browse():
            if kind == "save":
                p = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")],
                                                 initialfile="Сравнение_траекторий.xlsx")
            elif kind == "csvdir":
                p = filedialog.askopenfilename(filetypes=[("CSV", "*.csv"), ("Все", "*.*")])
            elif kind == "file":
                p = filedialog.askopenfilename(filetypes=ftypes or TRAJ_FT)
            else:
                p = filedialog.askdirectory()
            if p:
                var.set(p)
        ttk.Button(parent, text="Обзор…", command=browse).grid(row=r, column=2, padx=6, pady=4)

    nb = ttk.Notebook(root)
    nb.pack(fill="both", expand=True, padx=8, pady=(8, 0))

    # ================= Вкладка «Книга + CSV» =================
    t0 = ttk.Frame(nb, padding=8)
    nb.add(t0, text="  Сравнение по скважинам  ")
    t0.columnconfigure(1, weight=1)
    path_row2(t0, 0, "РН-ПД (книга, файл или папка):", paths["wb_ref"])
    path_row2(t0, 1, "РНТ (папка или файл):", paths["wb_alt"])
    hint0 = ttk.Frame(t0)
    hint0.grid(row=2, column=1, columnspan=2, sticky="w", padx=6)
    ttk.Label(hint0, foreground="#666", text="Формат определяется автоматически: «Структура файла траекторий» (CSV), "
              "«Профиль скважины» (Excel), книга «лист = скважина», CSV с колонкой «Скважина».").pack(side="left")
    ttk.Button(hint0, text="? Форматы и примеры", command=lambda: open_formats_help()).pack(side="left", padx=8)
    path_row(t0, 3, "Итоговый файл:", paths["wb_out"], "save")
    ttk.Label(t0, foreground="#666", text="Один Excel-файл: лист «Сводная», затем лист на каждую скважину, "
              "в конце «Ошибки».").grid(row=4, column=1, sticky="w", padx=6)
    o0 = ttk.Frame(t0)
    o0.grid(row=5, column=0, columnspan=3, sticky="w", padx=6, pady=4)
    ttk.Label(o0, text="Процессов:").pack(side="left")
    ttk.Spinbox(o0, from_=1, to=64, textvariable=workers, width=4).pack(side="left", padx=4)
    ttk.Label(o0, text="   Разделитель CSV:").pack(side="left")
    SEP_NAMES = {"авто": "auto", "запятая": ",", "точка с запятой": ";", "табуляция": "\t"}
    _sv = saved.get("_csv_sep", "auto")
    csv_sep = tk.StringVar(value={v: k for k, v in SEP_NAMES.items()}.get("auto" if _sv == "," else _sv, "авто"))
    ttk.Combobox(o0, textvariable=csv_sep, values=list(SEP_NAMES), width=16, state="readonly").pack(side="left", padx=4)
    b0 = ttk.Frame(t0)
    b0.grid(row=6, column=0, columnspan=3, sticky="w", padx=6, pady=6)
    run0 = ttk.Button(b0, text="▶  Сравнить все скважины")
    stop0 = ttk.Button(b0, text="■  Остановить", state="disabled")
    open0 = ttk.Button(b0, text="Открыть итоговый файл", state="disabled")
    for b in (run0, stop0, open0):
        b.pack(side="left", padx=(0, 8))
    pb0 = ttk.Progressbar(t0, mode="determinate")
    pb0.grid(row=7, column=0, columnspan=3, sticky="ew", padx=6)
    stat0 = ttk.Label(t0, text="")
    stat0.grid(row=8, column=0, columnspan=3, sticky="w", padx=6)
    log0 = tk.Text(t0, height=12, font=("Consolas", 9), relief="flat", bg="#f7f7f7")
    log0.grid(row=9, column=0, columnspan=3, sticky="nsew", padx=6, pady=6)
    t0.rowconfigure(9, weight=1)
    cancel0 = threading.Event()

    # ================= Вкладка «Одна пара» =================
    t1 = ttk.Frame(nb, padding=8)
    nb.add(t1, text="  Одна пара  ")
    t1.columnconfigure(1, weight=1)
    path_row(t1, 0, "РН-ПД:", paths["ref"], "file")
    path_row(t1, 1, "РНТ:", paths["alt"], "file")
    path_row(t1, 2, "Папка результатов:", paths["out_pair"], "dir")
    ttk.Button(t1, text="? Форматы и примеры", command=lambda: open_formats_help()).grid(
        row=0, column=3, rowspan=2, padx=(0, 6))
    ttk.Label(t1, text="Если папка не указана — результаты кладутся рядом с файлом РН-ПД, в папку «Сравнение».",
              foreground="#666").grid(row=3, column=1, sticky="w", padx=6)

    btns1 = ttk.Frame(t1)
    btns1.grid(row=4, column=0, columnspan=3, sticky="w", pady=10, padx=6)
    run1 = ttk.Button(btns1, text="▶  Сравнить")
    run1.pack(side="left")
    b_xlsx = ttk.Button(btns1, text="Открыть отчёт Excel", state="disabled")
    b_png = ttk.Button(btns1, text="Открыть графики", state="disabled")
    b_dir1 = ttk.Button(btns1, text="Открыть папку", state="disabled")
    for b in (b_xlsx, b_png, b_dir1):
        b.pack(side="left", padx=(8, 0))

    verdict_lbl = tk.Label(t1, text="", font=("Segoe UI", 13, "bold"), anchor="w")
    verdict_lbl.grid(row=5, column=0, columnspan=3, sticky="ew", padx=6)
    res_box = tk.Text(t1, height=16, font=("Consolas", 10), relief="flat", bg="#f7f7f7")
    res_box.grid(row=6, column=0, columnspan=3, sticky="nsew", padx=6, pady=6)
    t1.rowconfigure(6, weight=1)
    last = {}

    # ================= Вкладка «Пакет» =================
    t2 = ttk.Frame(nb, padding=8)
    nb.add(t2, text="  Пакетная обработка  ")
    t2.columnconfigure(1, weight=1)
    mf = ttk.LabelFrame(t2, text="Как найти пары файлов", padding=6)
    mf.grid(row=0, column=0, columnspan=3, sticky="ew")
    mf.columnconfigure(1, weight=1)
    ttk.Radiobutton(mf, text="Одна папка: в каждой подпапке РН-ПД + РНТ", variable=mode,
                    value="root").grid(row=0, column=0, columnspan=3, sticky="w")
    path_row(mf, 1, "   Папка со скважинами:", paths["root"], "dir")
    ttk.Radiobutton(mf, text="Две папки: РН-ПД и РНТ (пары по номеру скважины = имя листа/файла)",
                    variable=mode, value="dirs").grid(row=2, column=0, columnspan=3, sticky="w")
    path_row(mf, 3, "   Папка РН-ПД:", paths["ref_dir"], "dir")
    path_row(mf, 4, "   Папка РНТ:", paths["alt_dir"], "dir")
    ttk.Radiobutton(mf, text="Список пар CSV (колонки well;ref;alt)", variable=mode,
                    value="manifest").grid(row=5, column=0, columnspan=3, sticky="w")
    path_row(mf, 6, "   Файл CSV:", paths["manifest"], "file", [("CSV", "*.csv"), ("Все", "*.*")])

    of = ttk.Frame(t2)
    of.grid(row=1, column=0, columnspan=3, sticky="ew", pady=(6, 0))
    of.columnconfigure(1, weight=1)
    path_row(of, 0, "Папка результатов:", paths["out_batch"], "dir")
    opt = ttk.Frame(t2)
    opt.grid(row=2, column=0, columnspan=3, sticky="w", padx=6, pady=4)
    ttk.Label(opt, text="Признак РН-ПД в имени:").pack(side="left")
    ttk.Entry(opt, textvariable=ref_pat, width=22).pack(side="left", padx=4)
    ttk.Label(opt, text="РНТ:").pack(side="left")
    ttk.Entry(opt, textvariable=alt_pat, width=14).pack(side="left", padx=4)
    ttk.Label(opt, text="  Процессов:").pack(side="left")
    ttk.Spinbox(opt, from_=1, to=64, textvariable=workers, width=4).pack(side="left", padx=4)
    ttk.Checkbutton(opt, text="Пропускать уже посчитанные", variable=skip_var).pack(side="left", padx=8)

    btns2 = ttk.Frame(t2)
    btns2.grid(row=3, column=0, columnspan=3, sticky="w", padx=6, pady=6)
    run2 = ttk.Button(btns2, text="▶  Запустить")
    stop2 = ttk.Button(btns2, text="■  Остановить", state="disabled")
    b_sum = ttk.Button(btns2, text="Открыть сводную таблицу", state="disabled")
    b_dir2 = ttk.Button(btns2, text="Открыть папку", state="disabled")
    for b in (run2, stop2, b_sum, b_dir2):
        b.pack(side="left", padx=(0, 8))
    pb = ttk.Progressbar(t2, mode="determinate")
    pb.grid(row=4, column=0, columnspan=3, sticky="ew", padx=6)
    stat = ttk.Label(t2, text="")
    stat.grid(row=5, column=0, columnspan=3, sticky="w", padx=6)
    logbox = tk.Text(t2, height=10, font=("Consolas", 9), relief="flat", bg="#f7f7f7")
    logbox.grid(row=6, column=0, columnspan=3, sticky="nsew", padx=6, pady=6)
    t2.rowconfigure(6, weight=1)
    cancel_flag = threading.Event()

    # ================= Вкладка «Настройки» =================
    t3 = ttk.Frame(nb, padding=12)
    nb.add(t3, text="  Настройки  ")
    LABELS = {
        "step_md": "Шаг расчёта по стволу, м",
        "wellhead_tol": "Допуск совпадения устьев, м",
        "target_tol": "Допуск попадания в цель, м",
        "sep_tol": "Расхождение стволов для «Сходятся», м",
        "sep_warn": "Расхождение стволов для «Условно сходятся», м",
        "landing_inc": "Зенит начала горизонтального участка, град",
        "ddi_tol": "Разница DDI, при которой траектории «сопоставимы»",
    }
    for i, (k, lab) in enumerate(LABELS.items()):
        ttk.Label(t3, text=lab).grid(row=i, column=0, sticky="w", pady=5)
        ttk.Entry(t3, textvariable=S[k], width=10).grid(row=i, column=1, sticky="w", padx=10)
    ttk.Checkbutton(t3, text="Строить графики (без них расчёт в 2–3 раза быстрее)",
                    variable=plots_var).grid(row=len(LABELS), column=0, columnspan=2, sticky="w", pady=(8, 0))
    ttk.Checkbutton(t3, text="Добавлять листы с исходными траекториями и формулами DDI "
                             "(≈ +170 КБ на скважину)",
                    variable=traj_var).grid(row=len(LABELS) + 1, column=0, columnspan=2, sticky="w", pady=(2, 8))
    mfr = ttk.Frame(t3)
    mfr.grid(row=len(LABELS) + 2, column=0, columnspan=3, sticky="w", pady=(0, 8))
    ttk.Label(mfr, text="Методика расчёта DDI:").pack(side="left")
    ttk.Combobox(mfr, textvariable=ddi_method_var, values=list(tc.DDI_METHODS.values()),
                 state="readonly", width=72).pack(side="left", padx=8)

    def reset():
        for k, v in S.items():
            v.set(str(tc.DEFAULTS[k]))
        plots_var.set(True)
        traj_var.set(True)
        ddi_method_var.set(tc.DDI_METHODS["oag"])
    ttk.Button(t3, text="Вернуть по умолчанию", command=reset).grid(row=len(LABELS) + 3, column=0, sticky="w")
    ttk.Label(t3, foreground="#666", justify="left", text=(
        "\nЦели берутся из колонки комментариев (метки Т1, Т2, T3…). Если меток нет — целью считается забой.\n"
        "Вердикт «Сходятся»: устья и цели в допуске и макс. расхождение стволов ≤ первого порога.\n"
        "«Условно сходятся»: устья и цели в допуске, расхождение ≤ второго порога. Иначе — «Не сходятся».\n"
        "\nDDI — индекс сложности проводки. Методика 1 (Oag & Williams): DDI = log10(AHD · MD · Tortuosity / TVD),\n"
        "AHD — путь в плане вдоль ствола. Методика 2: DDI = log10(M · D · T / V), D = √(ΔX² + ΔY²) — смещение\n"
        "от устья на забое, T(i) = T(i−1) + α/10 · (M(i) − M(i−1)) — извилистость по интенсивности α.\n"
        "Шкала: < 6,0 низкая; 6,0–6,4 средняя; 6,4–6,8 высокая; ≥ 6,8 экстремальная.\n"
        f"\nНастройки сохраняются в {SETTINGS_FILE.name} рядом с программой."
    )).grid(row=len(LABELS) + 4, column=0, columnspan=3, sticky="w")

    status = ttk.Label(root, text="Готово к работе", anchor="w", padding=(10, 4))
    status.pack(fill="x")

    # ================= Справка: форматы, примеры, проверка файла =================
    import shutil
    import tempfile

    def example_file(fmt):
        p = tc.examples_dir() / fmt["example"] if fmt.get("example") else None
        return p if p and p.exists() else None

    def export_example(fmt):
        """Копия примера в папку пользователя (файлы внутри exe открывать напрямую нельзя)."""
        src = example_file(fmt)
        if not src:
            messagebox.showwarning(APP_TITLE, "Пример для этого формата не найден в составе программы.")
            return None
        dst_dir = Path(tempfile.gettempdir()) / "TrajCompare_примеры"
        dst_dir.mkdir(parents=True, exist_ok=True)
        dst = dst_dir / src.name
        try:
            shutil.copy(src, dst)
        except PermissionError:
            pass            # файл уже открыт — покажем существующую копию
        return dst

    def preview_file(path, title=None):
        """Окно предпросмотра: как программа распознала файл + первые строки таблицы."""
        path = Path(path)
        try:
            info = tc.describe_file(path)
            tables = tc.read_tables(path)
        except Exception as ex:  # noqa: BLE001
            messagebox.showerror(APP_TITLE, f"Не удалось прочитать файл:\n{path}\n\n{ex}")
            return
        win = tk.Toplevel(root)
        win.title(title or f"Просмотр — {path.name}")
        win.geometry("1100x620")
        top = ttk.Frame(win, padding=8)
        top.pack(fill="x")
        lines = [f"Файл: {path}", f"Тип по содержимому: {info['Тип по содержимому']}"]
        if "Кодировка" in info:
            lines.append(f"Кодировка: {info['Кодировка']};  разделитель: {info['Разделитель']}")
        for tb in info["Таблицы"]:
            sh = f"Лист «{tb['Лист']}»: " if tb.get("Лист") else ""
            if "Причина" in tb:
                lines.append(f"{sh}формат не распознан — {tb['Причина']}")
            else:
                lines.append(f"{sh}формат «{tb['Формат']}», строка заголовков {tb['Строка заголовков']}, точек {tb['Точек']}")
        ttk.Label(top, text="\n".join(lines[:12]), justify="left").pack(side="left", anchor="w")
        bar = ttk.Frame(win, padding=(8, 0))
        bar.pack(fill="x")
        sheet_var = tk.StringVar(value=str(tables[0][0]) if tables[0][0] is not None else "")
        if len(tables) > 1:
            ttk.Label(bar, text="Лист:").pack(side="left")
            ttk.Combobox(bar, textvariable=sheet_var, values=[str(t[0]) for t in tables], state="readonly",
                         width=20).pack(side="left", padx=4)
        ttk.Button(bar, text="Открыть в программе по умолчанию", command=lambda: open_path(path)).pack(side="right")
        ttk.Label(bar, foreground="#666", text="Показаны первые 300 строк. Подсвечены распознанные столбцы.").pack(
            side="left", padx=10)
        frm = ttk.Frame(win, padding=8)
        frm.pack(fill="both", expand=True)
        tree = ttk.Treeview(frm, show="headings")
        vs = ttk.Scrollbar(frm, orient="vertical", command=tree.yview)
        hs = ttk.Scrollbar(frm, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vs.set, xscrollcommand=hs.set)
        tree.grid(row=0, column=0, sticky="nsew"); vs.grid(row=0, column=1, sticky="ns"); hs.grid(row=1, column=0, sticky="ew")
        frm.rowconfigure(0, weight=1); frm.columnconfigure(0, weight=1)

        def show(*_):
            raw = dict((str(k), v) for k, v in tables).get(sheet_var.get(), tables[0][1])
            try:
                hr = tc._find_header_row(raw)
            except ValueError:
                hr = 0
            hdr = [str(c).replace("\n", " ").strip() if str(c) != "nan" else f"[{j + 1}]"
                   for j, c in enumerate(raw.iloc[hr].tolist())]
            known = set()
            try:
                hn = [tc._norm(c) for c in raw.iloc[hr].tolist()]
                d = raw.iloc[hr + 1:]
                known = {tc._pick(hn, d, k) for k in tc.COLS} - {None}
            except Exception:  # noqa: BLE001
                pass
            cols = [f"c{j}" for j in range(len(hdr))]
            tree.delete(*tree.get_children())
            tree["columns"] = cols
            for j, (c, h) in enumerate(zip(cols, hdr)):
                tree.heading(c, text=("✔ " if j in known else "") + h)
                tree.column(c, width=max(70, min(220, 9 * len(h))), stretch=False)
            for _, r in raw.iloc[hr + 1:hr + 301].iterrows():
                tree.insert("", "end", values=["" if str(v) in ("nan", "None") else v for v in r.tolist()])
        sheet_var.trace_add("write", show)
        show()

    def check_file():
        p = filedialog.askopenfilename(title="Какой формат у файла?", filetypes=TRAJ_FT)
        if p:
            preview_file(p, f"Проверка формата — {Path(p).name}")

    def open_formats_help(select=None):
        win = tk.Toplevel(root)
        win.title("Справка — форматы файлов траекторий")
        win.geometry("1000x600")
        pan = ttk.Frame(win, padding=8)
        pan.pack(fill="both", expand=True)
        lb = tk.Listbox(pan, width=34, exportselection=False, font=("Segoe UI", 10))
        lb.grid(row=0, column=0, rowspan=3, sticky="ns", padx=(0, 10))
        for f in tc.FORMATS:
            lb.insert("end", f["name"])
        head = ttk.Label(pan, font=("Segoe UI", 12, "bold"))
        head.grid(row=0, column=1, sticky="w")
        txt = tk.Text(pan, height=6, wrap="word", relief="flat", bg="#f7f7f7", font=("Segoe UI", 10))
        txt.grid(row=1, column=1, sticky="ew", pady=4)
        tv = ttk.Treeview(pan, columns=("c", "d", "r"), show="headings", height=10)
        for c, h, w in (("c", "Столбец (заголовок в файле)", 260), ("d", "Что содержит", 420), ("r", "Обязательный", 110)):
            tv.heading(c, text=h); tv.column(c, width=w, stretch=(c == "d"))
        tv.grid(row=2, column=1, sticky="nsew")
        pan.columnconfigure(1, weight=1); pan.rowconfigure(2, weight=1)
        btn = ttk.Frame(win, padding=8)
        btn.pack(fill="x")
        b_show = ttk.Button(btn, text="Показать пример здесь")
        b_open = ttk.Button(btn, text="Открыть пример в Excel / программе по умолчанию")
        b_save = ttk.Button(btn, text="Сохранить пример как…")
        for b in (b_show, b_open, b_save):
            b.pack(side="left", padx=(0, 8))
        ttk.Button(btn, text="Проверить свой файл…", command=check_file).pack(side="right")
        ttk.Label(win, foreground="#666", padding=(8, 0, 8, 8), justify="left", text=(
            "Тип файла (Excel или текст), кодировка, разделитель и формат таблицы определяются автоматически — "
            "расширение файла не важно.\nCSV с запятой в Excel с русскими настройками может открыться в один столбец: "
            "используйте «Показать пример здесь» или в Excel «Данные → Из текста/CSV».")).pack(fill="x")

        def sel(*_):
            i = lb.curselection()
            f = tc.FORMATS[i[0] if i else 0]
            head.config(text=f["name"])
            txt.config(state="normal"); txt.delete("1.0", "end")
            txt.insert("end", f"Файл: {f['file']}\n\n{f['note']}")
            txt.config(state="disabled")
            tv.delete(*tv.get_children())
            for c, d, req in f["columns"]:
                tv.insert("", "end", values=(c, d, "да" if req else ""))
            ex = example_file(f)
            st = "normal" if ex else "disabled"
            b_show.config(state=st, command=lambda: preview_file(ex, f"Пример — {f['name']}"))
            b_open.config(state=st, command=lambda: (lambda p: p and open_path(p))(export_example(f)))

            def save_as():
                p = filedialog.asksaveasfilename(initialfile=ex.name, defaultextension=ex.suffix,
                                                 filetypes=[("Пример", "*" + ex.suffix)])
                if p:
                    shutil.copy(ex, p)
                    status.config(text=f"Пример сохранён: {p}")
            b_save.config(state=st, command=save_as)
        lb.bind("<<ListboxSelect>>", sel)
        idx = next((i for i, f in enumerate(tc.FORMATS) if f["code"] == select), 0)
        lb.selection_set(idx)
        sel()
        return win

    menubar = tk.Menu(root)
    m_help = tk.Menu(menubar, tearoff=0)
    m_help.add_command(label="Форматы файлов и примеры…", command=lambda: open_formats_help())
    m_help.add_command(label="Пример: «Структура файла траекторий» (CSV)",
                       command=lambda: preview_file(example_file(tc.FORMATS[0]), "Пример — Структура файла траекторий")
                       if example_file(tc.FORMATS[0]) else None)
    m_help.add_command(label="Проверить формат файла…", command=check_file)
    m_help.add_separator()
    m_help.add_command(label="Открыть папку с примерами",
                       command=lambda: [export_example(f) for f in tc.FORMATS if f.get("example")] and
                       open_path(Path(tempfile.gettempdir()) / "TrajCompare_примеры"))
    m_help.add_command(label="О программе", command=lambda: messagebox.showinfo(
        APP_TITLE, "TrajCompare — сравнение траекторий скважин РН-ПД и РНТ.\n"
                   "Поддерживаемые форматы: «Структура файла траекторий» (CSV), «Профиль скважины» (Excel), "
                   "книга «лист = скважина», CSV с колонкой «Скважина», произвольные таблицы."))
    menubar.add_cascade(label="Справка", menu=m_help)
    root.config(menu=menubar)
    root.bind("<F1>", lambda e: open_formats_help())

    VCOL = {"Сходятся": "#2e7d32", "Условно сходятся": "#b26a00", "Не сходятся": "#c62828"}

    # ---------- действия ----------
    def do_pair():
        ref, alt = paths["ref"].get().strip(), paths["alt"].get().strip()
        if not (ref and alt):
            messagebox.showwarning(APP_TITLE, "Выберите файл РН-ПД и файл РНТ.")
            return
        try:
            c = cfg()
        except ValueError as ex:
            messagebox.showerror(APP_TITLE, str(ex))
            return
        out = paths["out_pair"].get().strip() or str(Path(ref).parent / "Сравнение")
        save_settings()
        run1.config(state="disabled")
        status.config(text="Идёт расчёт…")
        verdict_lbl.config(text="")
        res_box.delete("1.0", "end")

        def work():
            try:
                res = tc.run_pair(ref, alt, out, c)
                q.put(("pair_done", (res.row, out, res.well)))
            except Exception as ex:  # noqa: BLE001
                q.put(("pair_err", f"{type(ex).__name__}: {ex}\n\n{traceback.format_exc(limit=4)}"))
        threading.Thread(target=work, daemon=True).start()

    def do_batch():
        try:
            c = cfg()
        except ValueError as ex:
            messagebox.showerror(APP_TITLE, str(ex))
            return
        m = mode.get()
        a = SimpleNamespace(root=None, ref_dir=None, alt_dir=None, manifest=None,
                            ref_pattern=ref_pat.get(), alt_pattern=alt_pat.get(),
                            workers=max(1, int(workers.get())), skip_existing=bool(skip_var.get()))
        if m == "root":
            a.root = paths["root"].get().strip()
            base = a.root
        elif m == "dirs":
            a.ref_dir, a.alt_dir = paths["ref_dir"].get().strip(), paths["alt_dir"].get().strip()
            base = a.ref_dir and str(Path(a.ref_dir).parent)
        else:
            a.manifest = paths["manifest"].get().strip()
            base = a.manifest and str(Path(a.manifest).parent)
        if not base or (m == "dirs" and not a.alt_dir):
            messagebox.showwarning(APP_TITLE, "Укажите исходные папки или файл списка пар.")
            return
        a.out = paths["out_batch"].get().strip() or str(Path(base) / "Сравнение_траекторий")
        save_settings()
        cancel_flag.clear()
        run2.config(state="disabled"); stop2.config(state="normal")
        b_sum.config(state="disabled"); b_dir2.config(state="disabled")
        logbox.delete("1.0", "end")
        pb.config(value=0, maximum=1)
        status.config(text="Идёт пакетная обработка…")

        def prog(i, n, r):
            q.put(("prog", (i, n, r)))

        def work():
            try:
                rows, errs = tc.run_batch(a, c, progress=prog, cancel=cancel_flag.is_set)
                q.put(("batch_done", (rows, errs, a.out)))
            except BaseException as ex:  # noqa: BLE001
                q.put(("batch_err", f"{type(ex).__name__}: {ex}\n{traceback.format_exc(limit=4)}"))
        threading.Thread(target=work, daemon=True).start()


    def do_workbook():
        ref, alt = paths["wb_ref"].get().strip(), paths["wb_alt"].get().strip()
        if not (ref and alt):
            messagebox.showwarning(APP_TITLE, "Выберите траектории РН-ПД и РНТ (файл или папку).")
            return
        try:
            c = cfg()
        except ValueError as ex:
            messagebox.showerror(APP_TITLE, str(ex))
            return
        out = paths["wb_out"].get().strip() or str((Path(ref) if Path(ref).is_dir() else Path(ref).parent)
                                                   / "Сравнение_траекторий.xlsx")
        if not out.lower().endswith(".xlsx"):
            out += ".xlsx"
        paths["wb_out"].set(out)
        save_settings()
        tc.CSV_SEP = SEP_NAMES.get(csv_sep.get(), "auto")
        cancel0.clear()
        run0.config(state="disabled"); stop0.config(state="normal"); open0.config(state="disabled")
        log0.delete("1.0", "end"); pb0.config(value=0, maximum=1)
        status.config(text="Идёт сравнение…")
        import workbook_mode as wm

        def work():
            try:
                rows, errs, path = wm.run_workbook(ref, alt, out, c, max(1, int(workers.get())),
                                                   progress=lambda i, n, r: q.put(("wb_prog", (i, n, r))),
                                                   cancel=cancel0.is_set)
                q.put(("wb_done", (rows, errs, path)))
            except PermissionError:
                q.put(("wb_err", f"Не удалось записать {out}.\nВозможно, файл открыт в Excel — закройте его и повторите."))
            except BaseException as ex:  # noqa: BLE001
                q.put(("wb_err", f"{type(ex).__name__}: {ex}\n{traceback.format_exc(limit=4)}"))
        threading.Thread(target=work, daemon=True).start()

    run0.config(command=do_workbook)
    stop0.config(command=lambda: (cancel0.set(), status.config(text="Остановка… итоговый файл будет записан по готовым скважинам")))
    run1.config(command=do_pair)
    run2.config(command=do_batch)
    stop2.config(command=lambda: (cancel_flag.set(), status.config(text="Остановка после текущих пар…")))
    counts = {"ok": 0, "err": 0}

    def poll():
        try:
            while True:
                kind, data = q.get_nowait()
                if kind == "wb_prog":
                    i, n, r = data
                    pb0.config(maximum=max(n, 1), value=i)
                    stat0.config(text=f"Посчитано {i} из {n}")
                    msg = r.get("Вердикт") or ("ОШИБКА: " + str(r.get("Ошибка", "")))
                    log0.insert("end", f"{r.get('Скважина')}: {msg}\n"); log0.see("end")
                elif kind in ("wb_done", "wb_err"):
                    run0.config(state="normal"); stop0.config(state="disabled")
                    if kind == "wb_err":
                        status.config(text="Ошибка")
                        messagebox.showerror(APP_TITLE, data)
                    else:
                        rows, errs, path = data
                        open0.config(state="normal", command=lambda p=path: open_path(p))
                        vc = {}
                        for r in rows:
                            vc[r["row"]["Вердикт"]] = vc.get(r["row"]["Вердикт"], 0) + 1
                        txt = ", ".join(f"{k}: {v}" for k, v in vc.items()) or "нет сравнённых скважин"
                        status.config(text=f"Готово. {txt}; проблем: {len(errs)}.  Файл: {path}")
                        log0.insert("end", f"\nГотово: {path}\n{txt}; проблем: {len(errs)}\n"); log0.see("end")
                elif kind == "log":
                    box = log0 if nb.index(nb.select()) == 0 else logbox
                    box.insert("end", data + "\n"); box.see("end")
                elif kind == "pair_done":
                    row, out, well = data
                    v = row["Вердикт"]
                    verdict_lbl.config(text=f"Скважина {well}:  {v}", fg=VCOL.get(v, "black"))
                    lines = []
                    for k, val in row.items():
                        if k in ("Скважина", "Вердикт"):
                            continue
                        lines.append(f"{k:<46} {val:,.2f}".replace(",", " ") if isinstance(val, float)
                                     else f"{k:<46} {val}")
                    res_box.insert("end", "\n".join(lines))
                    stem = tc._safe(well)
                    last.update(xlsx=Path(out) / f"{stem}_сравнение.xlsx", png=Path(out) / f"{stem}_графики.png",
                                dir=Path(out))
                    b_xlsx.config(state="normal", command=lambda: open_path(last["xlsx"]))
                    b_png.config(state="normal" if last["png"].exists() else "disabled",
                                 command=lambda: open_path(last["png"]))
                    b_dir1.config(state="normal", command=lambda: open_path(last["dir"]))
                    run1.config(state="normal")
                    status.config(text=f"Готово: {out}")
                elif kind == "pair_err":
                    run1.config(state="normal")
                    status.config(text="Ошибка")
                    messagebox.showerror(APP_TITLE, "Не удалось сравнить файлы:\n\n" + data)
                elif kind == "prog":
                    i, n, r = data
                    pb.config(maximum=max(n, 1), value=i)
                    counts["ok" if "Вердикт" in r else "err"] += 1
                    stat.config(text=f"Обработано {i} из {n}   •   ошибок: {counts['err']}")
                    msg = r.get("Вердикт") or ("ОШИБКА: " + r.get("Ошибка", ""))
                    logbox.insert("end", f"{r.get('Скважина')}: {msg}\n"); logbox.see("end")
                elif kind in ("batch_done", "batch_err"):
                    run2.config(state="normal"); stop2.config(state="disabled")
                    counts.update(ok=0, err=0)
                    if kind == "batch_err":
                        status.config(text="Ошибка")
                        messagebox.showerror(APP_TITLE, data)
                    else:
                        rows, errs, out = data
                        summ = Path(out) / "СВОДНАЯ_сравнение_траекторий.xlsx"
                        b_sum.config(state="normal", command=lambda s=summ: open_path(s))
                        b_dir2.config(state="normal", command=lambda o=out: open_path(o))
                        vc = {}
                        for r in rows:
                            vc[r["Вердикт"]] = vc.get(r["Вердикт"], 0) + 1
                        txt = ", ".join(f"{k}: {v}" for k, v in vc.items()) or "нет пар"
                        status.config(text=f"Готово. {txt}; ошибок: {len(errs)}.  Результаты: {out}")
                        if not rows and not errs:
                            messagebox.showinfo(APP_TITLE, "Пары файлов не найдены. Проверьте папку и признаки "
                                                           "РН-ПД/РНТ в именах файлов.")
        except queue.Empty:
            pass
        root.after(150, poll)

    def on_close():
        save_settings()
        root.destroy()

    root.protocol("WM_DELETE_WINDOW", on_close)
    root.after(150, poll)
    if selftest:
        errors = []
        root.report_callback_exception = lambda *a: errors.append("".join(traceback.format_exception(*a)))

        def step1():
            w = open_formats_help("traj_struct")
            for f in tc.FORMATS:
                if example_file(f):
                    preview_file(example_file(f))
            nb.select(1); nb.select(2); nb.select(3); nb.select(0)
            root.after(1500, finish)

        def finish():
            (APP_DIR / "selftest_gui.log").write_text("OK" if not errors else "\n".join(errors), encoding="utf-8")
            root.destroy()
            if errors:
                sys.exit(1)
        root.after(500, step1)
    root.mainloop()


def main():
    multiprocessing.freeze_support()
    if sys.argv[1:] == ["--selftest-gui"]:
        try:
            run_gui(selftest=True)
        except Exception:  # noqa: BLE001
            (APP_DIR / "selftest_gui.log").write_text(traceback.format_exc(), encoding="utf-8")
            sys.exit(1)
    elif len(sys.argv) > 1:
        import trajcompare
        trajcompare.main()
    else:
        run_gui()


if __name__ == "__main__":
    main()
