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


def run_gui():
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
         if k in ("step_md", "wellhead_tol", "target_tol", "sep_tol", "sep_warn", "landing_inc")}
    plots_var = tk.BooleanVar(value=saved.get("plots", True))
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
        return c

    def save_settings():
        d = {k: v.get() for k, v in S.items()}
        d["plots"] = bool(plots_var.get())
        d.update({"_" + k: v.get() for k, v in paths.items()})
        d.update(_mode=mode.get(), _ref_pattern=ref_pat.get(), _alt_pattern=alt_pat.get(),
                 _workers=int(workers.get()), _skip=bool(skip_var.get()), _csv_sep=csv_sep.get())
        try:
            SETTINGS_FILE.write_text(json.dumps(d, ensure_ascii=False, indent=1), encoding="utf-8")
        except Exception:  # noqa: BLE001
            pass

    # ---------- общие элементы ----------
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
                p = filedialog.askopenfilename(filetypes=ftypes or [("Excel", "*.xlsx *.xlsm *.xls"), ("Все", "*.*")])
            else:
                p = filedialog.askdirectory()
            if p:
                var.set(p)
        ttk.Button(parent, text="Обзор…", command=browse).grid(row=r, column=2, padx=6, pady=4)

    nb = ttk.Notebook(root)
    nb.pack(fill="both", expand=True, padx=8, pady=(8, 0))

    # ================= Вкладка «Книга + CSV» =================
    t0 = ttk.Frame(nb, padding=8)
    nb.add(t0, text="  Книга РН-ПД + CSV  ")
    t0.columnconfigure(1, weight=1)
    path_row(t0, 0, "Книга РН-ПД (лист = скважина):", paths["wb_ref"], "file")
    ttk.Label(t0, text="РНТ (CSV):").grid(row=1, column=0, sticky="w", padx=6, pady=4)
    ttk.Entry(t0, textvariable=paths["wb_alt"]).grid(row=1, column=1, sticky="ew", padx=6, pady=4)
    altb = ttk.Frame(t0)
    altb.grid(row=1, column=2, padx=6)

    def pick_alt_dir():
        p = filedialog.askdirectory()
        if p:
            paths["wb_alt"].set(p)

    def pick_alt_file():
        p = filedialog.askopenfilename(filetypes=[("CSV", "*.csv"), ("Все", "*.*")])
        if p:
            paths["wb_alt"].set(p)
    ttk.Button(altb, text="Папка…", command=pick_alt_dir).pack(side="left")
    ttk.Button(altb, text="Файл…", command=pick_alt_file).pack(side="left", padx=(4, 0))
    ttk.Label(t0, foreground="#666", text="Папка с CSV (файл на скважину, номер скважины в имени файла) "
              "или один CSV с колонкой «Скважина». Разделитель — запятая.").grid(row=2, column=1, sticky="w", padx=6)
    path_row(t0, 3, "Итоговый файл:", paths["wb_out"], "save")
    ttk.Label(t0, foreground="#666", text="Один Excel-файл: лист «Сводная», затем лист на каждую скважину, "
              "в конце «Ошибки».").grid(row=4, column=1, sticky="w", padx=6)
    o0 = ttk.Frame(t0)
    o0.grid(row=5, column=0, columnspan=3, sticky="w", padx=6, pady=4)
    ttk.Label(o0, text="Процессов:").pack(side="left")
    ttk.Spinbox(o0, from_=1, to=64, textvariable=workers, width=4).pack(side="left", padx=4)
    ttk.Label(o0, text="   Разделитель CSV:").pack(side="left")
    csv_sep = tk.StringVar(value=saved.get("_csv_sep", ","))
    ttk.Combobox(o0, textvariable=csv_sep, values=[",", ";", "auto"], width=6, state="readonly").pack(side="left", padx=4)
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
    }
    for i, (k, lab) in enumerate(LABELS.items()):
        ttk.Label(t3, text=lab).grid(row=i, column=0, sticky="w", pady=5)
        ttk.Entry(t3, textvariable=S[k], width=10).grid(row=i, column=1, sticky="w", padx=10)
    ttk.Checkbutton(t3, text="Строить графики (без них расчёт в 2–3 раза быстрее)",
                    variable=plots_var).grid(row=len(LABELS), column=0, columnspan=2, sticky="w", pady=8)

    def reset():
        for k, v in S.items():
            v.set(str(tc.DEFAULTS[k]))
        plots_var.set(True)
    ttk.Button(t3, text="Вернуть по умолчанию", command=reset).grid(row=len(LABELS) + 1, column=0, sticky="w")
    ttk.Label(t3, foreground="#666", justify="left", text=(
        "\nЦели берутся из колонки комментариев (метки Т1, Т2, T3…). Если меток нет — целью считается забой.\n"
        "Вердикт «Сходятся»: устья и цели в допуске и макс. расхождение стволов ≤ первого порога.\n"
        "«Условно сходятся»: устья и цели в допуске, расхождение ≤ второго порога. Иначе — «Не сходятся».\n"
        f"\nНастройки сохраняются в {SETTINGS_FILE.name} рядом с программой."
    )).grid(row=len(LABELS) + 2, column=0, columnspan=3, sticky="w")

    status = ttk.Label(root, text="Готово к работе", anchor="w", padding=(10, 4))
    status.pack(fill="x")

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
            messagebox.showwarning(APP_TITLE, "Выберите книгу РН-ПД и CSV РНТ (папку или файл).")
            return
        try:
            c = cfg()
        except ValueError as ex:
            messagebox.showerror(APP_TITLE, str(ex))
            return
        out = paths["wb_out"].get().strip() or str(Path(ref).with_name("Сравнение_траекторий.xlsx"))
        if not out.lower().endswith(".xlsx"):
            out += ".xlsx"
        paths["wb_out"].set(out)
        save_settings()
        tc.CSV_SEP = csv_sep.get()
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
    root.mainloop()


def main():
    multiprocessing.freeze_support()
    if len(sys.argv) > 1:
        import trajcompare
        trajcompare.main()
    else:
        run_gui()


if __name__ == "__main__":
    main()
