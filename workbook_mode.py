# -*- coding: utf-8 -*-
"""
Режим «одна книга»:
  РН-ПД  — один Excel-файл, каждая скважина на своём листе (имя листа = номер скважины);
  РНТ — CSV (разделитель запятая): папка с файлом на скважину ИЛИ один CSV с колонкой «Скважина»;
  результат — один Excel-файл: лист «Сводная», затем по листу на скважину, в конце «Ошибки».
"""
from __future__ import annotations

import io
import logging
import re
import time
import traceback
from concurrent.futures import ProcessPoolExecutor, as_completed
from pathlib import Path

import numpy as np
import pandas as pd

import trajcompare as tc

log = logging.getLogger("trajcompare")

# Кириллица/латиница одинакового начертания — чтобы «377Г-А» и «377Г-A» считались одной скважиной
_HOMO = str.maketrans("АВЕКМНОРСТХаеорсух", "ABEKMHOPCTXaeopcyx")


def well_key(s) -> str:
    return re.sub(r"[\s_\-\.]+", "", str(s)).translate(_HOMO).upper()


def _tokens(s) -> set[str]:
    parts = [p for p in re.split(r"[^0-9A-Za-zА-Яа-яЁё]+", str(s)) if p]
    toks = {well_key(p) for p in parts}
    for i in range(len(parts) - 1):          # «377Г_1» → тоже «377Г1»
        toks.add(well_key(parts[i] + parts[i + 1]))
    return toks


# ----------------------------------------------------------------------------
# Загрузка исходных данных
# ----------------------------------------------------------------------------
def load_reference_book(path) -> dict[str, pd.DataFrame]:
    log.info("Читаю книгу РН-ПД %s …", Path(path).name)
    t0 = time.time()
    book = pd.read_excel(path, sheet_name=None, header=None)
    book = {str(k).strip(): v for k, v in book.items() if v.dropna(how="all").shape[0] > 3}
    log.info("  листов со скважинами: %d (%.0f с)", len(book), time.time() - t0)
    return book


def load_alternatives(path) -> dict[str, tuple[str, pd.DataFrame]]:
    """Возвращает {имя скважины: (источник, сырая таблица)}."""
    path = Path(path)
    out: dict[str, tuple[str, pd.DataFrame]] = {}
    files = sorted(p for p in path.rglob("*.csv")) if path.is_dir() else [path]
    for f in files:
        try:
            raw = tc.read_csv_raw(f)
        except Exception as ex:  # noqa: BLE001
            log.warning("Не читается %s: %s", f.name, ex)
            out[f.stem] = (str(f), ex)          # ошибку покажем на листе «Ошибки»
            continue
        # Один CSV на много скважин — ищем колонку «Скважина»
        try:
            hr = tc._find_header_row(raw)
        except ValueError as ex:
            out[f.stem] = (str(f), ex)
            continue
        headers = [tc._norm(c) for c in raw.iloc[hr].tolist()]
        data = raw.iloc[hr + 1:]
        wc = tc._pick(headers, data, "well")
        if wc is not None and data.iloc[:, wc].nunique(dropna=True) > 1:
            for w, part in data.groupby(data.iloc[:, wc].astype(str).str.strip(), sort=False):
                if w and w.lower() != "nan":
                    out[w] = (f"{f.name} [{w}]", pd.concat([raw.iloc[[hr]], part]))
        else:
            w = str(data.iloc[:, wc].dropna().iloc[0]).strip() if wc is not None and data.iloc[:, wc].notna().any() else f.stem
            out[w if wc is not None else f.stem] = (f.name, raw)
    log.info("РНТ найдено: %d", len(out))
    return out


def match_wells(ref_names, alt_names):
    """Сопоставление листов РН-ПД и РНТ. Возвращает (пары, РН-ПД без пары, РНТ без пары)."""
    alt_by_key = {}
    for a in alt_names:
        alt_by_key.setdefault(well_key(a), []).append(a)
    used, pairs, no_alt = set(), [], []
    for r in ref_names:
        k = well_key(r)
        cand = [a for a in alt_by_key.get(k, []) if a not in used]
        if not cand:   # имя файла содержит номер скважины: «377Г_alt.csv», «РНТ 377Г.csv»
            cand = [a for a in alt_names if a not in used and k in _tokens(a)]
        if len(cand) >= 1:
            a = sorted(cand, key=len)[0]
            used.add(a)
            pairs.append((r, a))
        else:
            no_alt.append(r)
    no_ref = [a for a in alt_names if a not in used]
    return pairs, no_alt, no_ref


# ----------------------------------------------------------------------------
# Расчёт одной скважины (выполняется в отдельном процессе)
# ----------------------------------------------------------------------------
def _traj_arrays(t) -> dict:
    marks = {round(float(m), 3): c for m, c in t.marks}
    return {"md": t.md, "inc": t.inc, "azi": t.azi, "tvd": t.tvd, "z": t.z, "n": t.n, "e": t.e,
            "comment": [marks.get(round(float(m), 3), "") for m in t.md]}


def _job(args):
    well, ref_raw, ref_src, alt_raw, alt_src, cfg = args
    t0 = time.time()
    try:
        ref = tc.parse_raw(ref_raw, "РН-ПД", ref_src, well)
        alt = tc.parse_raw(alt_raw, "РНТ", alt_src, well)
        res = tc.compare(ref, alt, cfg, well)
        png = None
        if cfg.get("plots", True):
            buf = io.BytesIO()
            tc.plot_pair(res, buf, cfg, dpi=cfg.get("dpi_embed", 80))
            try:   # 128 цветов: в 3 раза меньше, на графиках разницы не видно
                from PIL import Image
                im = Image.open(io.BytesIO(buf.getvalue())).convert("RGB").quantize(128)
                buf = io.BytesIO()
                im.save(buf, "PNG", optimize=True)
            except Exception:  # noqa: BLE001
                pass
            png = buf.getvalue()
        light = dict(row=res.row, summary=res.summary, targets=res.targets,
                     profile=res.profile, png=png, ref_src=ref_src, alt_src=alt_src,
                     traj={"РН-ПД": _traj_arrays(ref), "РНТ": _traj_arrays(alt)})
        return True, well, light
    except Exception as ex:  # noqa: BLE001
        return False, well, {"Скважина": well, "Проблема": "Ошибка расчёта",
                             "Подробности": f"{type(ex).__name__}: {ex}",
                             "РН-ПД": ref_src, "РНТ": alt_src,
                             "trace": traceback.format_exc(limit=3)}


# ----------------------------------------------------------------------------
# Итоговая книга
# ----------------------------------------------------------------------------
def _sheet_name(name, used):
    base = re.sub(r"[\[\]:*?/\\']", "_", str(name)).strip() or "Скважина"
    base = base[:31]
    n, i = base, 2
    while n.lower() in used:
        suf = f"_{i}"
        n = base[:31 - len(suf)] + suf
        i += 1
    used.add(n.lower())
    return n


XL_MAX_ROWS = 1_048_576

TRAJ_COLS = [   # методика 1 (Oag & Williams)
    ("Скважина", 12), ("MD, м", 10), ("Зенит, град", 9), ("Азимут, град", 9), ("TVD, м", 10),
    ("TVDSS, м", 10), ("Север (X), м", 13), ("Восток (Y), м", 13), ("Комментарий", 18),
    ("ΔMD, м", 8), ("Угол искривления β, град", 11), ("ΔAHD (путь в плане), м", 11),
    ("AHD нарастающим итогом, м", 12), ("Tortuosity нарастающим итогом, град", 13),
    ("DDI нарастающим итогом", 11),
]
TRAJ_COLS_CLOSURE = [   # методика 2 (по смещению от устья) — столбцы по заданию
    ("Скважина", 12), ("Глубина по стволу M, м", 11), ("Глубина по вертикали V, м", 11),
    ("Пространственная интенсивность α, град/10 м", 14), ("Смещение от устья на Север ΔX, м", 13),
    ("Смещение от устья на Восток ΔY, м", 13), ("Смещение от устья D, м", 12), ("Извилистость T, град", 12),
    ("DDI нарастающим итогом", 11),
    ("Зенит, град", 9), ("Азимут, град", 9), ("TVDSS, м", 10), ("Север (X), м", 13), ("Восток (Y), м", 13),
    ("Комментарий", 18),
]
DDI_KEYS = ("DDI", "ΔDDI", "Сравнение по DDI", "Tortuosity", "Сложность")
SUMMARY_DROP = ("DDI (индекс сложности)", "Категория сложности по DDI", "Методика DDI",
                "AHD — горизонтальный путь вдоль ствола, м", "D — смещение от устья на забое, м",
                "Tortuosity — сумма искривлений, град", "TVD забоя от устья, м")


def _cell(row, col, row_abs=False, col_abs=False):
    from xlsxwriter.utility import xl_rowcol_to_cell
    return xl_rowcol_to_cell(row, col, row_abs, col_abs)


def _q(sheet):
    return "'" + sheet.replace("'", "''") + "'"


def _traj_calc(t: dict):
    """Значения промежуточных колонок DDI — те же, что дают формулы Excel (кэш для просмотра без пересчёта)."""
    md, inc, azi, n, e, tvd = (np.asarray(t[k], float) for k in ("md", "inc", "azi", "n", "e", "tvd"))
    i, a = np.radians(inc), np.radians(azi)
    cosb = np.cos(i[1:] - i[:-1]) - np.sin(i[:-1]) * np.sin(i[1:]) * (1 - np.cos(a[1:] - a[:-1]))
    beta = np.r_[0, np.degrees(np.arccos(np.clip(cosb, -1, 1)))]
    dmd = np.r_[0, np.diff(md)]
    dahd = np.r_[0, np.hypot(np.diff(n), np.diff(e))]
    ahd, tort = np.cumsum(dahd), np.cumsum(beta)
    mdr, tvr = md - md[0], tvd - tvd[0]
    with np.errstate(divide="ignore", invalid="ignore"):
        ddi = np.where((ahd > 0) & (tort > 0) & (tvr > 0) & (mdr > 0), np.log10(ahd * mdr * tort / tvr), np.nan)
    return dmd, beta, dahd, ahd, tort, ddi


def _traj_calc_closure(t: dict):
    """Методика 2: α, ΔX, ΔY, D, T и DDI нарастающим итогом — как в формулах Excel."""
    md, inc, azi, n, e, tvd = (np.asarray(t[k], float) for k in ("md", "inc", "azi", "n", "e", "tvd"))
    i, a = np.radians(inc), np.radians(azi)
    cosb = np.cos(i[1:] - i[:-1]) - np.sin(i[:-1]) * np.sin(i[1:]) * (1 - np.cos(a[1:] - a[:-1]))
    beta = np.degrees(np.arccos(np.clip(cosb, -1, 1)))
    dmd = np.diff(md)
    alpha = np.r_[0.0, np.where(dmd > 0, beta / np.where(dmd > 0, dmd, 1) * 10, 0.0)]
    dx, dy = n - n[0], e - e[0]
    d = np.hypot(dx, dy)
    tort = np.r_[0.0, np.cumsum(alpha[1:] / 10 * dmd)]
    with np.errstate(divide="ignore", invalid="ignore"):
        ddi = np.where((md > 0) & (d > 0) & (tort > 0) & (tvd > 0), np.log10(md * d * tort / tvd), np.nan)
    return alpha, dx, dy, d, tort, ddi


def _layout_traj(results, kind, base_name):
    """Где на листах траекторий окажется каждая скважина: {well: (sheet, first_row0, last_row0)}."""
    pos, sheets = {}, []
    sheet, row = base_name, 1
    sheets.append(sheet)
    for res in results:
        w = res["row"]["Скважина"]
        n = len(res["traj"][kind]["md"])
        if row + n > XL_MAX_ROWS:
            sheet = f"{base_name} ({len(sheets) + 1})"
            sheets.append(sheet)
            row = 1
        pos[w] = (sheet, row, row + n - 1)
        row += n
    return pos, sheets


def write_result_book(results: list[dict], errors: list[dict], out_path, cfg, sources: dict):
    import xlsxwriter

    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    results = sorted(results, key=lambda r: str(r["row"]["Скважина"]))
    with_traj = bool(cfg.get("traj_sheets", True)) and all("traj" in r for r in results)
    tol = float(cfg.get("ddi_tol", 0.1))
    b1, b2, b3 = cfg.get("ddi_bounds", [6.0, 6.4, 6.8])
    method = cfg.get("ddi_method", "oag")
    closure = method == "closure"
    method_name = tc.DDI_METHODS.get(method, method)

    wb = xlsxwriter.Workbook(str(out_path), {"constant_memory": True, "nan_inf_to_errors": True})
    F = lambda **k: wb.add_format({"font_name": "Calibri", "font_size": 10, **k})
    f_title = F(bold=True, font_size=14, font_color="#1F5AA6")
    f_sec = F(bold=True, font_size=12, font_color="#1F5AA6")
    f_bold = F(bold=True)
    f_hdr = F(bold=True, font_color="#FFFFFF", bg_color="#1F5AA6", text_wrap=True, valign="vcenter",
              align="center", border=1, border_color="#D0D0D0")
    f_txt = F(border=1, border_color="#D0D0D0")
    f_num = F(border=1, border_color="#D0D0D0", num_format="0.00")
    f_ddi = F(border=1, border_color="#D0D0D0", num_format="0.000")
    f_frm = F(border=1, border_color="#D0D0D0", num_format="0.00", bg_color="#EEF3FB")   # ячейки с формулами
    f_frm3 = F(border=1, border_color="#D0D0D0", num_format="0.000", bg_color="#EEF3FB", bold=True)
    f_frmt = F(border=1, border_color="#D0D0D0", bg_color="#EEF3FB")
    f_link = F(font_color="#0563C1", underline=1)
    f_linkb = F(font_color="#0563C1", underline=1, bold=True, border=1, border_color="#D0D0D0")
    f_left = F(align="left")
    f_note = F(italic=True, font_color="#666666")
    VFMT = {k: F(border=1, border_color="#D0D0D0", bg_color="#" + v) for k, v in tc.VERDICT_FILL.items()}
    VTITLE = {k: F(bold=True, font_size=14, bg_color="#" + v) for k, v in tc.VERDICT_FILL.items()}
    CF = {k: wb.add_format({"bg_color": c}) for k, c in {
        "Низкая": "#C6EFCE", "Средняя": "#FFEB9C", "Высокая": "#F8CBAD", "Экстремальная": "#FF9B9B",
        "РНТ проще": "#C6EFCE", "сопоставимы": "#EDEDED", "РНТ сложнее": "#F8CBAD"}.items()}

    def cf_text(ws, rng):
        for k, fmt in CF.items():
            ws.conditional_format(rng, {"type": "cell", "criteria": "==", "value": f'"{k}"', "format": fmt})

    def num(v):
        if isinstance(v, (np.floating, float)):
            return None if not np.isfinite(v) else float(v)
        if isinstance(v, np.integer):
            return int(v)
        return v

    def put(ws, r, c, v, fmt=None):
        v = num(v)
        if v is None:
            ws.write_blank(r, c, None, fmt or f_txt)
        elif isinstance(v, (int, float)):
            ws.write_number(r, c, v, fmt or f_num)
        else:
            ws.write_string(r, c, str(v), fmt or f_txt)

    def put_table(ws, r0, df: pd.DataFrame, fmts=None):
        ws.set_row(r0, 32)
        for j, col in enumerate(df.columns):
            ws.write_string(r0, j, str(col), f_hdr)
        for i, row in enumerate(df.itertuples(index=False), 1):
            for j, v in enumerate(row):
                put(ws, r0 + i, j, v, (fmts or {}).get(j))
        return r0 + len(df) + 1

    # ---------- имена листов и раскладка ----------
    used = {"сводная", "ошибки", "траектории рн-пд", "траектории рнт"}
    names = {r["row"]["Скважина"]: _sheet_name(r["row"]["Скважина"], used) for r in results}
    if with_traj:
        pos_r, sh_r = _layout_traj(results, "РН-ПД", "Траектории РН-ПД")
        pos_a, sh_a = _layout_traj(results, "РНТ", "Траектории РНТ")
        POS = {"РН-ПД": pos_r, "РНТ": pos_a}
    # Блок «Расчёт DDI» на листе скважины — фиксированные ячейки (0-based)
    R_DDI_HDR, R_AHD, R_MD, R_TORT, R_TVD, R_DDI, R_CAT, R_CMP = 5, 6, 7, 8, 9, 10, 11, 12

    # ======================= Сводная =======================
    ws = wb.add_worksheet("Сводная")
    rows = [dict(r["row"]) for r in results]
    for rw in rows:
        rw.pop("Время, с", None); rw.pop("Сложность РН-ПД", None); rw.pop("Сложность РНТ", None)
    df = pd.DataFrame(rows)
    first = ["Скважина", "Вердикт", "Причины", "Сравнение по DDI", "DDI РН-ПД", "DDI РНТ", "ΔDDI (РНТ−РН-ПД)"]
    if len(df):
        df = df[[c for c in first if c in df.columns] + [c for c in df.columns if c not in first + ["РН-ПД", "РНТ"]]]
    ncol = len(df.columns)
    n = len(df)
    vc = df["Вердикт"].value_counts() if n else pd.Series(dtype=int)

    ws.write_string(0, 0, "Сравнение траекторий скважин: РН-ПД vs РНТ", f_title)
    T0 = 18                      # строка заголовка таблицы (0-based)
    first_data, last_data = T0 + 1, T0 + n
    col = {c: j for j, c in enumerate(df.columns)}

    def colrng(name):
        j = col[name]
        return f"{_cell(first_data, j, True, True)}:{_cell(max(last_data, first_data), j, True, True)}"

    stats = [
        ("РН-ПД (книга)", Path(sources.get("ref", "")).name, None),
        ("РНТ (CSV)", Path(sources.get("alt", "")).name, None),
        ("Сравнено скважин", n, None),
        ("Сходятся", int(vc.get("Сходятся", 0)), None),
        ("Условно сходятся", int(vc.get("Условно сходятся", 0)), None),
        ("Не сходятся", int(vc.get("Не сходятся", 0)), None),
        ("Проблем (без пары / ошибки)", len(errors), None),
        ("Допуски", f"цель {cfg['target_tol']:g} м, устье {cfg['wellhead_tol']:g} м, "
                    f"стволы {cfg['sep_tol']:g} / {cfg['sep_warn']:g} м", None),
        ("Методика DDI", method_name, None),
    ]
    if n:
        dc = df["Сравнение по DDI"].value_counts()
        stats += [
            ("Средний DDI РН-ПД", float(df["DDI РН-ПД"].mean()), f"=AVERAGE({colrng('DDI РН-ПД')})"),
            ("Средний DDI РНТ", float(df["DDI РНТ"].mean()), f"=AVERAGE({colrng('DDI РНТ')})"),
            ("РНТ проще по DDI", int(dc.get("РНТ проще", 0)), f'=COUNTIF({colrng("Сравнение по DDI")},"РНТ проще")'),
            ("РНТ сложнее по DDI", int(dc.get("РНТ сложнее", 0)), f'=COUNTIF({colrng("Сравнение по DDI")},"РНТ сложнее")'),
            ("DDI сопоставимы", int(dc.get("сопоставимы", 0)), f'=COUNTIF({colrng("Сравнение по DDI")},"сопоставимы")'),
            ("Порог «сопоставимы», |ΔDDI| <", tol, None),
        ]
    TOL_CELL = None
    for i, (k, v, fm) in enumerate(stats, 2):
        ws.write_string(i, 0, k, f_bold)
        fmt = VFMT.get(k, f_left)
        if fm:
            ws.write_formula(i, 1, fm, F(align="left", num_format="0.000" if "DDI" in k and "Средн" in k else "0",
                                         bg_color="#EEF3FB"), num(v))
        elif isinstance(v, (int, float)):
            ws.write_number(i, 1, v, F(align="left", bg_color="#FFF2CC") if k.startswith("Порог") else
                            (VFMT[k] if k in VFMT else f_left))
            if k.startswith("Порог"):
                TOL_CELL = f"'Сводная'!{_cell(i, 1, True, True)}"
                ws.write_string(i, 2, "← можно изменить: пересчитается колонка «Сравнение по DDI»", f_note)
        else:
            ws.write_string(i, 1, str(v), fmt)
        if k.startswith("Проблем") and errors:
            ws.write_url(i, 2, "internal:'Ошибки'!A1", f_link, "→ см. лист «Ошибки»")
    ws.write_string(T0 - 1, 0, "Ячейки с голубой заливкой — формулы; "
                    + ("DDI = log10(M·D·T/V), D — смещение от устья на забое, T — извилистость; "
                       if closure else "DDI = log10(AHD·MD·Tortuosity/TVD); ")
                    + "шкала: <6,0 низкая, 6,0–6,4 средняя, 6,4–6,8 высокая, ≥6,8 экстремальная", f_note)

    if n:
        ws.set_row(T0, 32)
        for j, c in enumerate(df.columns):
            ws.write_string(T0, j, c, f_hdr)
        for i, rec in enumerate(df.itertuples(index=False)):
            r = T0 + 1 + i
            d = dict(zip(df.columns, rec))
            w = d["Скважина"]
            sn = _q(names[w])
            for j, c in enumerate(df.columns):
                v = d[c]
                if c == "Скважина":
                    ws.write_url(r, j, f"internal:{sn}!A1", f_linkb, str(w))
                elif c == "Вердикт":
                    ws.write_string(r, j, str(v), VFMT.get(v, f_txt))
                elif c == "DDI РН-ПД":
                    ws.write_formula(r, j, f"={sn}!{_cell(R_DDI, 1)}", f_frm3, num(v))
                elif c == "DDI РНТ":
                    ws.write_formula(r, j, f"={sn}!{_cell(R_DDI, 2)}", f_frm3, num(v))
                elif c == "ΔDDI (РНТ−РН-ПД)":
                    ws.write_formula(r, j, f"={_cell(r, col['DDI РНТ'])}-{_cell(r, col['DDI РН-ПД'])}", f_frm3, num(v))
                elif c == "Сравнение по DDI":
                    ws.write_formula(r, j, f"={sn}!{_cell(R_CMP, 1)}", f_frmt, str(v))
                else:
                    put(ws, r, j, v)
        cf_text(ws, colrng("Сравнение по DDI"))
        ws.autofilter(T0, 0, last_data, ncol - 1)
        ws.freeze_panes(T0 + 1, 1)
    ws.set_column(0, 0, 32)
    ws.set_column(1, 1, 22)
    ws.set_column(2, 2, 44)
    ws.set_column(3, 3, 16)
    ws.set_column(4, max(ncol, 5), 13)
    ws.set_landscape(); ws.fit_to_pages(1, 0)

    # ======================= Листы скважин =======================
    for k, res in enumerate(results):
        row = res["row"]
        w = row["Скважина"]
        sh = wb.add_worksheet(names[w])
        v = row["Вердикт"]
        sh.merge_range(0, 0, 0, 3, f"Скважина {w} — {v}", VTITLE.get(v, f_title))
        # строки пишем строго по порядку (режим constant_memory)
        sh.write_url(1, 0, "internal:'Сводная'!A1", f_link, "← к сводной")
        sh.write_string(1, 1, f"Источник РН-ПД: {res['ref_src']}")
        if with_traj:
            tsheet, a0, a1 = POS["РН-ПД"][w]
            sh.write_url(2, 0, f"internal:{_q(tsheet)}!A{a0 + 1}", f_link,
                         f"→ траектория РН-ПД (строки {a0 + 1}–{a1 + 1})", f"Лист «{tsheet}»")
        sh.write_string(2, 1, f"Источник РНТ: {res['alt_src']}")
        if with_traj:
            tsheet, a0, a1 = POS["РНТ"][w]
            sh.write_url(3, 0, f"internal:{_q(tsheet)}!A{a0 + 1}", f_link,
                         f"→ траектория РНТ (строки {a0 + 1}–{a1 + 1})", f"Лист «{tsheet}»")
        sh.write_string(4, 0, ("Расчёт DDI (формулы по листам траекторий)" if with_traj else "Расчёт DDI"), f_sec)
        sh.write_string(4, 1, method_name, f_note)

        # --- блок DDI с формулами ---
        for j, h in enumerate(["Показатель", "РН-ПД", "РНТ", "Разница (РНТ−РН-ПД)"]):
            sh.write_string(R_DDI_HDR, j, h, f_hdr)
        if closure:
            labels = {R_AHD: "D — смещение от устья на забое, м", R_MD: "M — глубина по стволу на забое, м",
                      R_TORT: "T — извилистость от устья до забоя, град", R_TVD: "V — глубина по вертикали на забое, м",
                      R_DDI: "DDI = log10(M·D·T/V)", R_CAT: "Категория сложности по DDI",
                      R_CMP: "Сравнение по DDI"}
        else:
            labels = {R_AHD: "AHD — путь в плане вдоль ствола, м", R_MD: "MD — длина ствола, м",
                      R_TORT: "Tortuosity — сумма углов искривления, град", R_TVD: "TVD забоя от устья, м",
                      R_DDI: "DDI = log10(AHD·MD·Tortuosity/TVD)", R_CAT: "Категория сложности по DDI",
                      R_CMP: "Сравнение по DDI"}
        cache = {}
        for jj, kind in ((1, "РН-ПД"), (2, "РНТ")):
            t = res["traj"][kind] if "traj" in res else None
            if t is not None and closure:
                alpha, dx, dy, dcl, tort, ddi_c = _traj_calc_closure(t)
                md = np.asarray(t["md"], float); tvd = np.asarray(t["tvd"], float)
                cache[kind] = {R_AHD: dcl[-1], R_MD: md[-1], R_TORT: tort[-1], R_TVD: tvd[-1], R_DDI: ddi_c[-1]}
            elif t is not None:
                dmd, beta, dahd, ahd, tort, ddi_c = _traj_calc(t)
                md = np.asarray(t["md"], float); tvd = np.asarray(t["tvd"], float)
                cache[kind] = {R_AHD: ahd[-1], R_MD: md[-1] - md[0], R_TORT: tort[-1],
                               R_TVD: tvd[-1] - tvd[0], R_DDI: ddi_c[-1]}
            else:
                cache[kind] = {R_DDI: row.get(f"DDI {kind}")}
        for rr in (R_AHD, R_MD, R_TORT, R_TVD, R_DDI):
            sh.write_string(rr, 0, labels[rr], f_bold if rr == R_DDI else f_txt)
            for jj, kind in ((1, "РН-ПД"), (2, "РНТ")):
                fmt = f_frm3 if rr == R_DDI else f_frm
                cv = num(cache[kind].get(rr))
                if with_traj:
                    tsh, a0, a1 = POS[kind][w]
                    q = _q(tsh)
                    if closure:
                        fm = {R_AHD: f"={q}!G{a1 + 1}", R_MD: f"={q}!B{a1 + 1}",
                              R_TORT: f"={q}!H{a1 + 1}", R_TVD: f"={q}!C{a1 + 1}"}.get(rr)
                    else:
                        fm = {R_AHD: f"={q}!M{a1 + 1}",
                              R_MD: f"={q}!B{a1 + 1}-{q}!B{a0 + 1}",
                              R_TORT: f"={q}!N{a1 + 1}",
                              R_TVD: f"={q}!E{a1 + 1}-{q}!E{a0 + 1}"}.get(rr)
                else:
                    fm = None
                if rr == R_DDI:
                    c_ = lambda r_: _cell(r_, jj)
                    fm = f'=IFERROR(LOG10({c_(R_AHD)}*{c_(R_MD)}*{c_(R_TORT)}/{c_(R_TVD)}),"")' if with_traj else None
                if fm:
                    sh.write_formula(rr, jj, fm, fmt, cv if cv is not None else "")
                else:
                    put(sh, rr, jj, cv, fmt)
            dv = None
            if cache["РН-ПД"].get(rr) is not None and cache["РНТ"].get(rr) is not None:
                dv = num(cache["РНТ"][rr] - cache["РН-ПД"][rr])
            sh.write_formula(rr, 3, f'=IFERROR({_cell(rr, 2)}-{_cell(rr, 1)},"")', f_frm3 if rr == R_DDI else f_frm,
                             dv if dv is not None else "")
        # категория и сравнение — тоже формулами
        sh.write_string(R_CAT, 0, labels[R_CAT], f_txt)
        for jj, kind in ((1, "РН-ПД"), (2, "РНТ")):
            d = _cell(R_DDI, jj)
            dv = cache[kind].get(R_DDI)
            cat = "—" if dv is None or not np.isfinite(dv) else tc.DDI_CATS[int(np.searchsorted([b1, b2, b3], dv, side="right"))]
            sh.write_formula(R_CAT, jj, f'=IF({d}="","—",IF({d}<{b1},"Низкая",IF({d}<{b2},"Средняя",'
                                         f'IF({d}<{b3},"Высокая","Экстремальная"))))', f_frmt, cat)
        sh.write_blank(R_CAT, 3, None, f_txt)
        sh.write_string(R_CMP, 0, labels[R_CMP], f_bold)
        dlt = _cell(R_DDI, 3)
        tref = TOL_CELL or str(tol)
        sh.merge_range(R_CMP, 1, R_CMP, 3, "", f_frmt)
        sh.write_formula(R_CMP, 1, f'=IF({dlt}="","нет данных",IF(ABS({dlt})<{tref},"сопоставимы",'
                                   f'IF({dlt}<0,"РНТ проще","РНТ сложнее")))', f_frmt, row.get("Сравнение по DDI", ""))
        cf_text(sh, f"B{R_CAT + 1}:C{R_CMP + 1}")

        # --- итог сравнения (без DDI — он выше) ---
        r = R_CMP + 2
        sh.write_string(r, 0, "Итог сравнения траекторий", f_sec)
        items = pd.DataFrame([{"Показатель": a, "Значение": b} for a, b in row.items()
                              if a not in ("Скважина", "РН-ПД", "РНТ", "Время, с")
                              and not any(a.startswith(x) for x in DDI_KEYS)])
        r = put_table(sh, r + 1, items)
        r += 1
        sh.write_string(r, 0, "Параметры траекторий", f_sec)
        summ = res["summary"][~res["summary"]["Параметр"].isin(SUMMARY_DROP)].reset_index(drop=True)
        r = put_table(sh, r + 1, summ)

        img_bottom = 5
        if res.get("png"):
            from PIL import Image
            wpx = Image.open(io.BytesIO(res["png"])).size
            scale = 1150 / wpx[0]
            sh.insert_image(4, 5, "plot.png", {"image_data": io.BytesIO(res["png"]), "x_scale": scale,
                                               "y_scale": scale, "object_position": 3})
            img_bottom = 4 + int(wpx[1] * scale / 20) + 2
        r = max(r + 1, img_bottom)
        sh.write_string(r, 0, "Цели и реперные точки", f_sec)
        r = put_table(sh, r + 1, res["targets"])
        r += 1
        step = cfg.get("profile_step", 25.0)
        p = res["profile"]
        md = p.iloc[:, 0].values
        keep = np.r_[True, np.diff(np.floor(md / step)) > 0]
        keep[-1] = True
        sh.write_string(r, 0, f"Расхождение стволов по глубине (шаг {step:g} м)", f_sec)
        put_table(sh, r + 1, p[keep].reset_index(drop=True))

        sh.set_column(0, 0, 44)
        sh.set_column(1, 3, 19)
        sh.set_column(4, 11, 13)
        sh.set_zoom(90)
        sh.set_landscape(); sh.fit_to_pages(1, 0)
        if (k + 1) % 100 == 0:
            log.info("  записано листов скважин: %d", k + 1)

    # ======================= Листы траекторий =======================
    if with_traj and closure:
        for kind, sheets, pos in (("РН-ПД", sh_r, pos_r), ("РНТ", sh_a, pos_a)):
            wss = {}
            for s_ in sheets:
                t_ws = wb.add_worksheet(s_)
                t_ws.set_row(0, 58)
                for j, (h, wdt) in enumerate(TRAJ_COLS_CLOSURE):
                    t_ws.write_string(0, j, h, f_hdr)
                    t_ws.set_column(j, j, wdt)
                t_ws.freeze_panes(1, 2)
                t_ws.write_comment(0, 3, "α — пространственная интенсивность на интервале (i−1, i], "
                                         "по методу минимальной кривизны; в первой строке 0")
                t_ws.write_comment(0, 6, "D = √(ΔX² + ΔY²) — суммарное смещение от устья")
                t_ws.write_comment(0, 7, "T(i) = T(i−1) + α/10 · (M(i) − M(i−1)); в первой строке 0")
                t_ws.write_comment(0, 8, "DDI = LOG10(M · D · T / V)")
                wss[s_] = t_ws
            f_wl = F(font_color="#0563C1", underline=1, bold=True, top=1)
            f2v = F(num_format="0.00")
            f3v = F(num_format="0.000")
            f2 = F(num_format="0.00", bg_color="#EEF3FB")
            f3 = F(num_format="0.000", bg_color="#EEF3FB")
            f3b = F(num_format="0.000", bg_color="#EEF3FB", bold=True)
            for res in results:
                w = res["row"]["Скважина"]
                t = res["traj"][kind]
                s_, a0, a1 = pos[w]
                t_ws = wss[s_]
                alpha, dx, dy, dcl, tort, ddi_c = _traj_calc_closure(t)
                last = len(t["md"]) - 1
                for i in range(len(t["md"])):
                    r = a0 + i
                    R, P = r + 1, r
                    if i == 0:
                        t_ws.write_url(r, 0, f"internal:{_q(names[w])}!A1", f_wl, str(w))
                    else:
                        t_ws.write_string(r, 0, str(w))
                    put(t_ws, r, 1, t["md"][i], f2v)
                    put(t_ws, r, 2, t["tvd"][i], f2v)
                    put(t_ws, r, 3, alpha[i], f3v)
                    put(t_ws, r, 4, dx[i], f2v)
                    put(t_ws, r, 5, dy[i], f2v)
                    t_ws.write_formula(r, 6, f"=SQRT(E{R}^2+F{R}^2)", f2, num(dcl[i]))
                    if i == 0:
                        t_ws.write_number(r, 7, 0.0, f3)
                    else:
                        t_ws.write_formula(r, 7, f"=H{P}+D{R}/10*(B{R}-B{P})", f3, num(tort[i]))
                    dv = num(ddi_c[i])
                    t_ws.write_formula(r, 8, f'=IF(AND(B{R}>0,G{R}>0,H{R}>0,C{R}>0),LOG10(B{R}*G{R}*H{R}/C{R}),"")',
                                       f3b if i == last else f3, dv if dv is not None else "")
                    for j, key in ((9, "inc"), (10, "azi"), (11, "z"), (12, "n"), (13, "e")):
                        val = num(t[key][i])
                        if val is None:
                            t_ws.write_blank(r, j, None)
                        else:
                            t_ws.write_number(r, j, val)
                    if t["comment"][i]:
                        t_ws.write_string(r, 14, str(t["comment"][i]))
            log.info("  лист(ы) траекторий %s записаны", kind)
    elif with_traj:
        for kind, sheets, pos in (("РН-ПД", sh_r, pos_r), ("РНТ", sh_a, pos_a)):
            wss = {}
            for s in sheets:
                t_ws = wb.add_worksheet(s)
                t_ws.set_row(0, 45)
                for j, (h, wdt) in enumerate(TRAJ_COLS):
                    t_ws.write_string(0, j, h, f_hdr)
                    t_ws.set_column(j, j, wdt)
                t_ws.freeze_panes(1, 2)
                t_ws.write_comment(0, 10, "β = DEGREES(ACOS(cos(I2−I1) − sin I1·sin I2·(1 − cos(A2−A1)))) "
                                          "— пространственный угол между соседними точками")
                t_ws.write_comment(0, 14, "DDI = LOG10(AHD · (MD−MD устья) · Tortuosity / (TVD−TVD устья))")
                wss[s] = t_ws
            f_wl = F(font_color="#0563C1", underline=1, bold=True, top=1)
            f2 = F(num_format="0.00", bg_color="#EEF3FB")
            f3 = F(num_format="0.000", bg_color="#EEF3FB")
            f3b = F(num_format="0.000", bg_color="#EEF3FB", bold=True)
            FMT = {9: f2, 10: f3, 11: f2, 12: f2, 13: f3, 14: f3}
            for res in results:
                w = res["row"]["Скважина"]
                t = res["traj"][kind]
                s, a0, a1 = pos[w]
                t_ws = wss[s]
                dmd, beta, dahd, ahd, tort, ddi_c = _traj_calc(t)
                cols = [t["md"], t["inc"], t["azi"], t["tvd"], t["z"], t["n"], t["e"]]
                for i in range(len(t["md"])):
                    r = a0 + i
                    R = r + 1                       # номер строки в Excel
                    if i == 0:
                        t_ws.write_url(r, 0, f"internal:{_q(names[w])}!A1", f_wl, str(w))
                    else:
                        t_ws.write_string(r, 0, str(w))
                    for j, arr in enumerate(cols, 1):
                        val = num(arr[i])
                        if val is None:
                            t_ws.write_blank(r, j, None)
                        else:
                            t_ws.write_number(r, j, val)
                    if t["comment"][i]:
                        t_ws.write_string(r, 8, str(t["comment"][i]))
                    if i == 0:
                        t_ws.write_number(r, 9, 0.0, f2)
                        t_ws.write_number(r, 10, 0.0, f3)
                        t_ws.write_number(r, 11, 0.0, f2)
                        t_ws.write_formula(r, 12, f"=L{R}", f2, 0.0)
                        t_ws.write_formula(r, 13, f"=K{R}", f3, 0.0)
                        t_ws.write_blank(r, 14, None, f3)
                    else:
                        P = R - 1
                        t_ws.write_formula(r, 9, f"=B{R}-B{P}", f2, num(dmd[i]))
                        t_ws.write_formula(
                            r, 10,
                            f"=DEGREES(ACOS(MAX(-1,MIN(1,COS(RADIANS(C{R}-C{P}))-SIN(RADIANS(C{P}))*SIN(RADIANS(C{R}))"
                            f"*(1-COS(RADIANS(D{R}-D{P})))))))", f3, num(beta[i]))
                        t_ws.write_formula(r, 11, f"=SQRT((G{R}-G{P})^2+(H{R}-H{P})^2)", f2, num(dahd[i]))
                        t_ws.write_formula(r, 12, f"=M{P}+L{R}", f2, num(ahd[i]))
                        t_ws.write_formula(r, 13, f"=N{P}+K{R}", f3, num(tort[i]))
                        S = a0 + 1
                        dv = num(ddi_c[i])
                        t_ws.write_formula(
                            r, 14, f'=IF(AND(M{R}>0,N{R}>0,B{R}-$B${S}>0,E{R}-$E${S}>0),'
                                   f'LOG10(M{R}*(B{R}-$B${S})*N{R}/(E{R}-$E${S})),"")',
                            f3b if i == len(t["md"]) - 1 else f3, dv if dv is not None else "")
            log.info("  лист(ы) траекторий %s записаны", kind)

    # ======================= Ошибки =======================
    es = wb.add_worksheet("Ошибки")
    edf = pd.DataFrame(errors or [{"Скважина": "", "Проблема": "нет", "Подробности": ""}])
    edf = edf[[c for c in ["Скважина", "Проблема", "Подробности", "РН-ПД", "РНТ"] if c in edf.columns]]
    put_table(es, 0, edf)
    for j, wdt in enumerate((22, 30, 80, 30, 30)):
        es.set_column(j, j, wdt)
    es.write_url(0, 6, "internal:'Сводная'!A1", f_link, "← к сводной")

    log.info("Сохраняю файл …")
    wb.close()
    return out_path


# ----------------------------------------------------------------------------
# Запуск
# ----------------------------------------------------------------------------
def run_workbook(ref_xlsx, alt_csv, out_xlsx, cfg=None, workers=1, progress=None, cancel=None):
    cfg = dict(tc.DEFAULTS, **(cfg or {}))
    t0 = time.time()
    book = load_reference_book(ref_xlsx)
    alts = load_alternatives(alt_csv)
    pairs, no_alt, no_ref = match_wells(list(book), list(alts))
    log.info("Пар: %d; РН-ПД без РНТ: %d; РНТ без РН-ПД: %d",
             len(pairs), len(no_alt), len(no_ref))

    errors = [{"Скважина": w, "Проблема": "Нет CSV РНТ", "Подробности":
               "лист есть в книге РН-ПД, подходящий CSV не найден", "РН-ПД": f"лист «{w}»"} for w in no_alt]
    jobs = []
    for r, a in pairs:
        src, raw = alts[a]
        if isinstance(raw, Exception):
            errors.append({"Скважина": r, "Проблема": "CSV не читается", "Подробности": str(raw), "РНТ": src})
            continue
        jobs.append((r, book[r], f"{Path(ref_xlsx).name} / лист «{r}»", raw, src, cfg))
    for a in no_ref:
        src, raw = alts[a]
        errors.append({"Скважина": a, "Проблема": "Нет листа в книге РН-ПД" if not isinstance(raw, Exception)
                       else "CSV не читается", "Подробности": "" if not isinstance(raw, Exception) else str(raw),
                       "РНТ": src})

    results = []
    n = len(jobs)

    def handle(i, ok, well, data):
        if ok:
            results.append(data)
            r = data["row"]
        else:
            errors.append(data)
            r = {"Скважина": well, "Ошибка": data["Подробности"]}
        if progress:
            progress(i, n, r)

    if workers > 1 and n > 1:
        with ProcessPoolExecutor(workers) as ex:
            futs = [ex.submit(_job, j) for j in jobs]
            for i, f in enumerate(as_completed(futs), 1):
                handle(i, *f.result())
                if cancel and cancel():
                    for g in futs:
                        g.cancel()
                    break
                if i % 25 == 0 or i == n:
                    log.info("  посчитано %d/%d (%.0f с)", i, n, time.time() - t0)
    else:
        for i, j in enumerate(jobs, 1):
            if cancel and cancel():
                break
            handle(i, *_job(j))

    log.info("Записываю итоговую книгу …")
    out = write_result_book(results, errors, out_xlsx, cfg, {"ref": ref_xlsx, "alt": alt_csv})
    log.info("Готово: %s — скважин %d, проблем %d, %.0f с", out.name, len(results), len(errors), time.time() - t0)
    return results, errors, out
