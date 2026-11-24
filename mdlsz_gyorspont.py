#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
===============================================================================
 Project     : MDLSZ_Gyorspont
 File        : mdlsz_gyorspont.py
 Version     : 1.0.3
 Author      : Csaba Csanaki (cscsanaki)
 Repository  : https://github.com/cscsanaki/MDLSZ_Gyorspont
 Description : MDLSZ Gyorspont lovészeti verseny nyilvántartó és
               eredményrögzítő alkalmazás (Tkinter + openpyxl).
===============================================================================
"""

import copy
import os
import shutil
import sys
import zipfile
import tkinter as tk
from datetime import datetime
from tkinter import filedialog, messagebox, ttk
import openpyxl


DIVISIONS = [
    "KKPI - Nyílt",
    "KKPI - Optika",
    "NKPI - Nyílt",
    "NKPI - Optika",
    "KKPU - Nyílt",
    "KKPU - Optika",
    "NKPU - Nyílt",
    "NKPU - Optika",
    "Sörétes - Nyílt",
    "Sörétes - Optika",
]

LICENSE_SHEET = "Versenyengedelyek"
DATA_START_ROW = 6
DATA_END_ROW = 55  # 50 alapértelmezett sor, de dinamikusan bővíthető


import io
import re

INT_XML_ATTRS = {
    "activeTab", "firstSheet", "tabRatio", "zoomScale", "zoomScaleNormal",
    "zoomScalePageLayoutView", "zoomScaleSheetLayoutView", "colorId", "theme",
    "tint", "applyFont", "applyFill", "applyBorder", "applyAlignment",
    "applyProtection", "numFmtId", "fontId", "fillId", "borderId", "xfId",
    "outlineLevelRow", "outlineLevelCol", "showGridLines", "tabSelected",
    "xWindow", "yWindow", "windowWidth", "windowHeight", "maxDigitWidth",
    "xSplit", "ySplit", "activeCellId", "count", "uniqueCount", "sheetId",
}


def _to_int_str(raw_val):
    cleaned = str(raw_val).strip().replace(",", ".")
    return str(int(round(float(cleaned))))


def _sanitize_xlsx_bytes(raw_bytes, extra_exact_replacements=None):
    exact_map = dict(extra_exact_replacements or {})
    out_buf = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(raw_bytes), "r") as zin, zipfile.ZipFile(out_buf, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.endswith(".xml") or item.filename.endswith(".rels"):
                txt = data.decode("utf-8", errors="ignore")

                def fix_attr(match):
                    attr_name = match.group(1)
                    attr_val = match.group(2)
                    if attr_name == "version":
                        return match.group(0)
                    val_stripped = attr_val.strip()
                    if val_stripped in exact_map:
                        return f'{attr_name}="{exact_map[val_stripped]}"'
                    if re.fullmatch(r"[+-]?[0-9]+\.[0-9]+([eE][+-]?[0-9]+)?", val_stripped):
                        try:
                            return f'{attr_name}="{_to_int_str(val_stripped)}"'
                        except Exception:
                            return f'{attr_name}="0"'
                    return match.group(0)

                txt = re.sub(r'\b([A-Za-z0-9_:]+)="([^"]*)"', fix_attr, txt)
                for bad_val, good_val in exact_map.items():
                    txt = txt.replace(f'>{bad_val}<', f'>{good_val}<')
                data = txt.encode("utf-8")
            zout.writestr(item, data)
    return out_buf.getvalue()


def _extract_bad_literal(err_msg):
    m = re.search(r"got '([^']+)'", err_msg)
    if m:
        return m.group(1)
    m2 = re.search(r'got "([^"]+)"', err_msg)
    if m2:
        return m2.group(1)
    return None


def _open_workbook_resilient(path):
    try:
        return openpyxl.load_workbook(path, data_only=False)
    except Exception as exc:
        err_msg = str(exc)
        low = err_msg.lower()
        is_int_err = (
            "expected integer" in low
            or "expected <class 'int'>" in low
            or "invalid literal for int" in low
            or isinstance(exc, TypeError)
        )
        if not is_int_err:
            raise

        with open(path, "rb") as f:
            cur_bytes = f.read()

        exact_reps = {}
        bad_lit = _extract_bad_literal(err_msg)
        if bad_lit:
            try:
                exact_reps[bad_lit] = _to_int_str(bad_lit)
            except Exception:
                exact_reps[bad_lit] = "0"

        last_exc = exc
        for _ in range(5):
            cur_bytes = _sanitize_xlsx_bytes(cur_bytes, extra_exact_replacements=exact_reps)
            try:
                return openpyxl.load_workbook(io.BytesIO(cur_bytes), data_only=False)
            except Exception as inner_exc:
                last_exc = inner_exc
                inner_msg = str(inner_exc)
                lit = _extract_bad_literal(inner_msg)
                if lit and lit not in exact_reps:
                    try:
                        exact_reps[lit] = _to_int_str(lit)
                    except Exception:
                        exact_reps[lit] = "0"
                else:
                    break
        raise last_exc


def normalize_license(val):
    if val is None:
        return ""
    s = str(val).strip()
    if s.endswith(".0"):
        s = s[:-2]
    # Kezdő nullák eltávolítása számjegyekből (pl. "04241" -> "4241"), hogy egyezzen az Excelben tárolt értékkel
    if s.isdigit():
        s = s.lstrip("0") or "0"
    return s


HIT_TIERS = (13, 12, 11, 10, 9)


def _hit_tuple(hit_counts):
    hc = hit_counts or {}
    return tuple(-int(hc.get(k, 0) or 0) for k in HIT_TIERS)


def parse_score_for_sort(val, hit_counts=None):
    """Csökkenő rendezési kulcs:
    - Érvényes numerikus pontszám: (0, -float(val), -13db, -12db, -11db, -10db, -9db) -> legmagasabb elöl, holtversenyben a 13..9 találatok száma dönt
    - 'fegyverhiba': (1, 0, 0, 0, 0, 0, 0) -> pontszámok után
    - Üres / Regisztrált: (2, 0, 0, 0, 0, 0, 0) -> legalul
    """
    ht = _hit_tuple(hit_counts)
    if val is None:
        return (2, 0.0) + ht
    s = str(val).strip()
    if not s or s.lower() == "regisztrált":
        return (2, 0.0) + ht
    if s.lower() == "fegyverhiba":
        return (1, 0.0) + ht
    try:
        return (0, -float(s.replace(",", "."))) + ht
    except ValueError:
        return (2, 0.0) + ht


def validate_score_input(raw_text):
    """Visszaadja a normalizált Excel értéket (int/float vagy 'fegyverhiba' vagy None),
    vagy ValueError-t dob érvénytelen adat esetén."""
    s = raw_text.strip()
    if s == "":
        return None
    if s.lower() == "fegyverhiba":
        return "fegyverhiba"
    try:
        num = float(s.replace(",", "."))
    except ValueError:
        raise ValueError("Csak 0 és 195 közötti szám, vagy a 'fegyverhiba' szöveg adható meg!")
    if num < 0 or num > 195:
        raise ValueError("A pontszám csak 0 és 195 között lehet!")
    if num.is_integer():
        return int(num)
    return round(num, 2)



def _make_circle_photo(active: bool, size: int = 24):
    """Nagyméretű (24x24 px) kör / célkör PhotoImage a könnyű kattintáshoz és jól láthatósághoz."""
    img = tk.PhotoImage(width=size, height=size)
    cx = cy = (size - 1) / 2.0
    r_outer = size * 0.44
    r_mid = size * 0.27
    r_inner = size * 0.12

    bg = "#FFFFFF"
    outer_col = "#B91C1C" if active else "#64748B"
    mid_col = "#FEE2E2" if active else "#F1F5F9"
    inner_col = "#DC2626" if active else "#94A3B8"

    for y in range(size):
        row_pixels = []
        for x in range(size):
            d = ((x - cx) ** 2 + (y - cy) ** 2) ** 0.5
            if abs(d - r_outer) <= 1.35:
                row_pixels.append(outer_col)
            elif abs(d - r_mid) <= 1.15:
                row_pixels.append(mid_col if active else outer_col)
            elif d <= r_inner:
                row_pixels.append(inner_col)
            elif d < r_outer:
                row_pixels.append("#FEF2F2" if active else bg)
            else:
                row_pixels.append(bg)
        img.put("{" + " ".join(row_pixels) + "}", to=(0, y))
    return img


class LoterModernApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("MDLSZ_Gyorspont v1.0.3")
        self.geometry("1280x820")
        self.minsize(1060, 700)
        self.configure(bg="#F1F5F9")

        self.workbook_path = None
        self.mode_var = tk.StringVar(value="Eredménybevitel")
        self.comp_type_var = tk.StringVar(value="Gyorspont")

        # Adatstruktúrák
        # licenses: { "4241": {"name": "...", "club": "..."} }
        self.licenses = {}
        # registrations: { div_name: [ {"lic": "4241", "score": None | int | "fegyverhiba", "hits": None, "note": None, "unlicensed_name": None}, ... ] }
        self.registrations = {d: [] for d in DIVISIONS}
        self.meta_info = {}

        self.current_lic = None
        self.recent_competitors = []
        self.recent_actions = {}
        self.sort_reverse = {}
        self.img_reg = _make_circle_photo(True, size=24)
        self.img_unreg = _make_circle_photo(False, size=24)

        self._setup_styles()
        self._build_start_screen()

    def _setup_styles(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        style.configure("TFrame", background="#F1F5F9")
        style.configure("Card.TFrame", background="#FFFFFF", relief="flat")
        style.configure(
            "Treeview",
            background="#FFFFFF",
            foreground="#0F172A",
            rowheight=36,
            fieldbackground="#FFFFFF",
            font=("Segoe UI", 10),
            borderwidth=0,
        )
        style.configure(
            "Treeview.Heading",
            background="#0F172A",
            foreground="#FFFFFF",
            font=("Segoe UI", 10, "bold"),
            relief="flat",
            padding=(8, 6),
        )
        style.map(
            "Treeview",
            background=[("selected", "#0F766E")],
            foreground=[("selected", "#FFFFFF")],
        )
        style.map(
            "Treeview.Heading",
            background=[("active", "#1E293B")],
        )

    # ------------------------------------------------------------------
    # 1. KEZDŐKÉPERNYŐ (MDLSZ Gyorspont + XLS/XLSX betöltés)
    # ------------------------------------------------------------------
    def _build_start_screen(self):
        for w in self.winfo_children():
            w.destroy()

        outer = tk.Frame(self, bg="#F1F5F9")
        outer.pack(fill="both", expand=True)

        card = tk.Frame(
            outer,
            bg="#FFFFFF",
            padx=42,
            pady=40,
            highlightbackground="#CBD5E1",
            highlightthickness=1,
        )
        card.place(relx=0.5, rely=0.5, anchor="center", width=560)

        tk.Label(
            card,
            text="MDLSZ_Gyorspont",
            font=("Segoe UI", 26, "bold"),
            fg="#0F172A",
            bg="#FFFFFF",
            anchor="center",
        ).pack(fill="x", pady=(0, 26))

        tk.Button(
            card,
            text="XLS betöltése",
            font=("Segoe UI", 12, "bold"),
            fg="#FFFFFF",
            bg="#0F766E",
            activebackground="#115E59",
            activeforeground="#FFFFFF",
            pady=14,
            relief="flat",
            cursor="hand2",
            command=self._open_excel_dialog,
        ).pack(fill="x")

    # ------------------------------------------------------------------
    # 2. EXCEL BEOLVASÁS ÉS MENTÉS
    # ------------------------------------------------------------------
    def _open_excel_dialog(self):
        path = filedialog.askopenfilename(
            title="Verseny Excel fájl megnyitása",
            filetypes=[("Excel munkafüzet (*.xlsx, *.xls)", "*.xlsx *.xls"), ("Minden fájl (*.*)", "*.*")],
        )
        if not path:
            return

        try:
            self._load_workbook(path)
        except Exception as exc:
            messagebox.showerror("Hiba az Excel megnyitásakor", f"Nem sikerült beolvasni a fájlt:\n{exc}")
            return

        self._build_main_workspace()

    def _load_workbook(self, path):
        wb = _open_workbook_resilient(path)
        if LICENSE_SHEET not in wb.sheetnames:
            raise ValueError(f"Hiányzik a kötelező rejtett munkalap: '{LICENSE_SHEET}'!")

        # Rejtett versenyengedélyek beolvasása
        ws_lic = wb[LICENSE_SHEET]
        licenses = {}
        for r in range(1, ws_lic.max_row + 1):
            lic = normalize_license(ws_lic.cell(r, 1).value)
            name = ws_lic.cell(r, 2).value
            club = ws_lic.cell(r, 3).value
            if lic and name:
                licenses[lic] = {
                    "name": str(name).strip(),
                    "club": str(club).strip() if club is not None else "",
                }

        if not licenses:
            raise ValueError("A 'Versenyengedelyek' munkalap üres, nem található benne versenyző.")

        # Versenyszámok beolvasása
        registrations = {d: [] for d in DIVISIONS}
        meta_info = {}
        for div in DIVISIONS:
            if div not in wb.sheetnames:
                continue
            ws = wb[div]
            if not meta_info:
                found_comp_no = ws.cell(1, 2).value
                for r_idx in range(1, 6):
                    for c_idx in range(1, 6):
                        lbl_txt = str(ws.cell(r_idx, c_idx).value or "").strip().lower()
                        if "verseny száma" in lbl_txt or "versenyszáma" in lbl_txt or "verseny szám" in lbl_txt:
                            cand = ws.cell(r_idx, c_idx + 1).value
                            if cand is not None and str(cand).strip() != "":
                                found_comp_no = cand
                                break
                meta_info = {
                    "verseny_szam": str(found_comp_no).strip() if found_comp_no is not None else "",
                    "szervezo": ws.cell(2, 2).value or "",
                    "helyszin": ws.cell(3, 2).value or "",
                    "datum": ws.cell(4, 2).value or "",
                }

            for r in range(DATA_START_ROW, ws.max_row + 1):
                lic = normalize_license(ws.cell(r, 3).value)
                score_raw = ws.cell(r, 5).value
                unlic_name = ws.cell(r, 6).value
                hits = ws.cell(r, 7).value
                note = ws.cell(r, 8).value

                if lic:
                    score_val = None
                    if score_raw is not None and str(score_raw).strip() != "":
                        s_str = str(score_raw).strip()
                        if s_str.lower() == "fegyverhiba":
                            score_val = "fegyverhiba"
                        elif s_str.lower() == "regisztrált":
                            score_val = None
                        else:
                            try:
                                f_val = float(s_str.replace(",", "."))
                                score_val = int(f_val) if f_val.is_integer() else f_val
                            except ValueError:
                                score_val = s_str
                    parsed_hc = {k: 0 for k in HIT_TIERS}
                    if note:
                        for m in re.finditer(r"\b(13|12|11|10|9)\s*[:x×*]\s*(\d+)", str(note)):
                            parsed_hc[int(m.group(1))] = int(m.group(2))
                    registrations[div].append(
                        {
                            "lic": lic,
                            "score": score_val,
                            "unlicensed_name": unlic_name,
                            "hits": hits,
                            "note": note,
                            "hit_counts": parsed_hc,
                        }
                    )

        # Készítsünk biztonsági másolatot első megnyitáskor
        backup_path = path + ".bak"
        if not os.path.exists(backup_path):
            try:
                shutil.copy2(path, backup_path)
            except OSError:
                pass

        self.workbook_path = path
        self.licenses = licenses
        self.registrations = registrations
        self.meta_info = meta_info

        # Kezdő lista feltöltése a munkafüzetben már meglévő utolsó 10 versenyzővel
        initial_recent = []
        for div in DIVISIONS:
            for entry in registrations[div]:
                if entry["lic"] not in initial_recent:
                    initial_recent.append(entry["lic"])
        self.recent_competitors = initial_recent[:10]

        # Induláskor rendezzük és szinkronizáljuk az Excel-t
        self._save_to_excel(silent=True)

    def _save_to_excel(self, silent=False):
        """Minden versenyszámot csökkenő sorrendbe rendez pontszám alapján,
        és azonnal visszaírja a megnyitott XLSX fájlba a VLOOKUP képletek megtartásával."""
        if not self.workbook_path:
            return

        # Csökkenő sorrendbe rendezés minden versenyszámnál
        for div in DIVISIONS:
            self.registrations[div].sort(
                key=lambda item: parse_score_for_sort(item["score"], item.get("hit_counts"))
            )

        try:
            wb = openpyxl.load_workbook(self.workbook_path)
            for div in DIVISIONS:
                if div not in wb.sheetnames:
                    ws = wb.create_sheet(title=div)
                    ws.cell(5, 1, "Sorszám")
                    ws.cell(5, 2, "Név")
                    ws.cell(5, 3, "V.eng.")
                    ws.cell(5, 4, "Egyesület")
                    ws.cell(5, 5, "Eredmény")
                    ws.cell(5, 6, "Rajtengedély nélküli versenyző neve")
                    ws.cell(5, 7, "Találatok száma")
                    ws.cell(5, 8, "Megjegyzés")
                else:
                    ws = wb[div]

                entries = self.registrations[div]
                max_target_row = max(DATA_END_ROW, DATA_START_ROW + len(entries) - 1, ws.max_row)

                # Sablonstílus a 6. sorból (ha létezik)
                template_styles = {}
                if ws.max_row >= DATA_START_ROW:
                    for col_idx in range(1, 9):
                        src = ws.cell(DATA_START_ROW, col_idx)
                        template_styles[col_idx] = {
                            "font": copy.copy(src.font),
                            "fill": copy.copy(src.fill),
                            "border": copy.copy(src.border),
                            "alignment": copy.copy(src.alignment),
                            "number_format": src.number_format,
                        }

                for idx, row_num in enumerate(range(DATA_START_ROW, max_target_row + 1)):
                    ws.cell(row_num, 1, idx + 1)
                    # VLOOKUP képletek a rejtett Versenyengedelyek munkalapra
                    ws.cell(row_num, 2, f'=IF(C{row_num}="","",VLOOKUP(C{row_num}, Versenyengedelyek!A1:C10000,2, 0))')
                    ws.cell(row_num, 4, f'=IF(C{row_num}="","",VLOOKUP(C{row_num}, Versenyengedelyek!A1:C10000,3, 0))')

                    if idx < len(entries):
                        entry = entries[idx]
                        lic_val = int(entry["lic"]) if entry["lic"].isdigit() else entry["lic"]
                        ws.cell(row_num, 3, lic_val)
                        ws.cell(row_num, 5, entry["score"])
                        ws.cell(row_num, 6, entry.get("unlicensed_name"))
                        ws.cell(row_num, 7, entry.get("hits"))
                        ws.cell(row_num, 8, entry.get("note"))
                    else:
                        ws.cell(row_num, 3, None)
                        ws.cell(row_num, 5, None)
                        ws.cell(row_num, 6, None)
                        ws.cell(row_num, 7, None)
                        ws.cell(row_num, 8, None)

                    # Cellastílusok átvitele új sorokra is
                    if template_styles:
                        for col_idx, st in template_styles.items():
                            c = ws.cell(row_num, col_idx)
                            c.font = copy.copy(st["font"])
                            c.fill = copy.copy(st["fill"])
                            c.border = copy.copy(st["border"])
                            c.alignment = copy.copy(st["alignment"])
                            c.number_format = st["number_format"]

            # Rejtett lap biztosan rejtett maradjon
            if LICENSE_SHEET in wb.sheetnames:
                wb[LICENSE_SHEET].sheet_state = "hidden"

            wb.save(self.workbook_path)
            if hasattr(self, "lbl_save_status"):
                now_str = datetime.now().strftime("%H:%M:%S")
                self.lbl_save_status.configure(text=f"Automatikusan mentve: {now_str}", fg="#10B981")
        except PermissionError:
            if not silent:
                messagebox.showerror(
                    "Fájl zárolva!",
                    "Az Excel fájl jelenleg meg van nyitva egy másik programban (pl. Microsoft Excel).\n"
                    "Zárd be a fájlt az Excelben, hogy a program menteni tudja a módosításokat!",
                )
            if hasattr(self, "lbl_save_status"):
                self.lbl_save_status.configure(text="Hiba: Excel fájl zárolva!", fg="#F87171")
        except Exception as exc:
            if not silent:
                messagebox.showerror("Mentési hiba", f"Hiba történt az Excel mentésekor:\n{exc}")

    # ------------------------------------------------------------------
    # 3. FŐ MUNKAABLAK ÉS FÜLEK
    # ------------------------------------------------------------------
    def _build_main_workspace(self):
        for w in self.winfo_children():
            w.destroy()

        # Felső navigációs sáv (sötét fejléc a képernyőképek stílusában)
        topbar = tk.Frame(self, bg="#0F172A", height=48)
        topbar.pack(fill="x", side="top")
        topbar.pack_propagate(False)

        raw_comp_no = self.meta_info.get("verseny_szam", "").strip()
        if raw_comp_no:
            comp_no_disp = raw_comp_no if raw_comp_no.lower().startswith("verseny száma") else f"Verseny száma: {raw_comp_no}"
        else:
            comp_no_disp = "Verseny száma: —"

        tk.Label(
            topbar,
            text=comp_no_disp,
            font=("Segoe UI", 10, "bold"),
            fg="#5EEAD4",
            bg="#0F172A",
            padx=14,
        ).pack(side="left", padx=(6, 12))

        self.nav_buttons = {}
        tabs = [
            ("eredmenybevitel", "Eredménybevitel"),
            ("versenyzok", "Versenyzők"),
            ("statisztika", "Statisztika"),
            ("eredmenyek", "Eredmények"),
        ]

        for key, label in tabs:
            btn = tk.Button(
                topbar,
                text=label,
                font=("Segoe UI", 10, "bold"),
                fg="#CBD5E1",
                bg="#0F172A",
                activebackground="#1E293B",
                activeforeground="#FFFFFF",
                bd=0,
                padx=18,
                pady=10,
                cursor="hand2",
                command=lambda k=key: self._switch_tab(k),
            )
            btn.pack(side="left", padx=(6, 0))
            self.nav_buttons[key] = btn

        # Konténer a 4 fülnek
        self.content_area = tk.Frame(self, bg="#F1F5F9")
        self.content_area.pack(fill="both", expand=True)

        self.tab_frames = {}
        self._build_tab_eredmenybevitel()
        self._build_tab_versenyzok()
        self._build_tab_statisztika()
        self._build_tab_eredmenyek()

        default_tab = "eredmenyek" if self.mode_var.get() == "Eredményjelző" else "eredmenybevitel"
        self._switch_tab(default_tab)

    def _switch_tab(self, tab_key):
        for k, frame in self.tab_frames.items():
            frame.pack_forget()
        for k, btn in self.nav_buttons.items():
            if k == tab_key:
                btn.configure(fg="#FFFFFF", bg="#1E293B")
            else:
                btn.configure(fg="#CBD5E1", bg="#0F172A")

        self.tab_frames[tab_key].pack(fill="both", expand=True)
        self.update_idletasks()

        if tab_key == "eredmenybevitel":
            self._refresh_overview_table()
            if self.current_lic:
                self._refresh_competitor_divisions()
        elif tab_key == "versenyzok":
            self._refresh_versenyzok_tab()
        elif tab_key == "statisztika":
            self._refresh_statisztika_tab()
        elif tab_key == "eredmenyek":
            self._refresh_eredmenyek_tab()

    # ------------------------------------------------------------------
    # FÜL 1: EREDMÉNYBEVITEL (Nevezés + Pontszám bevitel + Áttekintés)
    # ------------------------------------------------------------------
    def _build_tab_eredmenybevitel(self):
        frame = tk.Frame(self.content_area, bg="#F1F5F9", padx=14, pady=10)
        self.tab_frames["eredmenybevitel"] = frame

        # --- FELSŐ RÉSZEK: Fix magasságú felső sáv (Bal: Keresés + Jobb: Divíziók és pontbevitel) ---
        top_pane = tk.Frame(frame, bg="#F1F5F9", height=340)
        top_pane.pack(side="top", fill="x", pady=(0, 10))
        top_pane.pack_propagate(False)

        # BAL KÁRTYA: Versenyző keresés
        left_card = tk.Frame(
            top_pane,
            bg="#FFFFFF",
            padx=18,
            pady=14,
            highlightbackground="#CBD5E1",
            highlightthickness=1,
            width=370,
        )
        left_card.pack(side="left", fill="y", padx=(0, 12))
        left_card.pack_propagate(False)

        tk.Label(
            left_card,
            text="Versenyző",
            font=("Segoe UI", 15, "bold"),
            fg="#0F172A",
            bg="#FFFFFF",
            anchor="w",
        ).pack(fill="x", pady=(0, 8))

        tk.Label(
            left_card,
            text="Versenyengedély",
            font=("Segoe UI", 10),
            fg="#475569",
            bg="#FFFFFF",
            anchor="w",
        ).pack(fill="x", pady=(0, 4))

        search_row = tk.Frame(left_card, bg="#FFFFFF")
        search_row.pack(fill="x", pady=(0, 10))

        self.entry_lic = tk.Entry(
            search_row,
            font=("Segoe UI", 13),
            relief="solid",
            bd=1,
            bg="#FFFFFF",
            fg="#0F172A",
        )
        self.entry_lic.pack(side="left", fill="x", expand=True, ipady=5)
        self.entry_lic.bind("<Return>", lambda _e: self._search_competitor())

        tk.Button(
            search_row,
            text="Keresés",
            font=("Segoe UI", 11, "bold"),
            fg="#FFFFFF",
            bg="#0F766E",
            activebackground="#115E59",
            activeforeground="#FFFFFF",
            relief="flat",
            padx=14,
            pady=6,
            cursor="hand2",
            command=self._search_competitor,
        ).pack(side="left", padx=(8, 0))

        # Találati kártya
        self.comp_info_box = tk.Frame(
            left_card,
            bg="#F0FDF4",
            padx=14,
            pady=12,
            highlightbackground="#BBF7D0",
            highlightthickness=1,
        )
        self.comp_info_box.pack(fill="both", expand=True)

        self.lbl_comp_name_title = tk.Label(
            self.comp_info_box,
            text="Név",
            font=("Segoe UI", 9),
            fg="#475569",
            bg="#F0FDF4",
            anchor="w",
        )
        self.lbl_comp_name_title.pack(fill="x")

        self.lbl_comp_name = tk.Label(
            self.comp_info_box,
            text="Írj be egy versenyengedély számot",
            font=("Segoe UI", 13, "bold"),
            fg="#0F172A",
            bg="#F0FDF4",
            anchor="w",
            wraplength=300,
            justify="left",
        )
        self.lbl_comp_name.pack(fill="x", pady=(2, 10))

        self.lbl_comp_club_title = tk.Label(
            self.comp_info_box,
            text="Egyesület",
            font=("Segoe UI", 9),
            fg="#475569",
            bg="#F0FDF4",
            anchor="w",
        )
        self.lbl_comp_club_title.pack(fill="x")

        self.lbl_comp_club = tk.Label(
            self.comp_info_box,
            text="—",
            font=("Segoe UI", 11),
            fg="#0F172A",
            bg="#F0FDF4",
            anchor="w",
            wraplength=300,
            justify="left",
        )
        self.lbl_comp_club.pack(fill="x", pady=(2, 0))

        # JOBB KÁRTYA: Versenyző divíziói + Pontszám bevitel
        right_card = tk.Frame(
            top_pane,
            bg="#FFFFFF",
            padx=18,
            pady=12,
            highlightbackground="#CBD5E1",
            highlightthickness=1,
        )
        right_card.pack(side="left", fill="both", expand=True)

        header_r = tk.Frame(right_card, bg="#FFFFFF")
        header_r.pack(fill="x", pady=(0, 6))

        tk.Label(
            header_r,
            text="Versenyző divíziói",
            font=("Segoe UI", 15, "bold"),
            fg="#0F172A",
            bg="#FFFFFF",
        ).pack(side="left")

        tk.Button(
            header_r,
            text="Kiválasztott törlése",
            font=("Segoe UI", 10, "bold"),
            fg="#FFFFFF",
            bg="#B91C1C",
            activebackground="#991B1B",
            activeforeground="#FFFFFF",
            relief="flat",
            padx=14,
            pady=5,
            cursor="hand2",
            command=self._remove_selected_division,
        ).pack(side="right")

        # Pontbevitel sáv
        score_bar = tk.Frame(
            right_card,
            bg="#F8FAFC",
            padx=10,
            pady=6,
            highlightbackground="#E2E8F0",
            highlightthickness=1,
        )
        score_bar.pack(fill="x", pady=(0, 6))

        tk.Label(
            score_bar,
            text="Kijelölt versenyszám:",
            font=("Segoe UI", 9, "bold"),
            fg="#334155",
            bg="#F8FAFC",
        ).pack(side="left")

        self.lbl_selected_div = tk.Label(
            score_bar,
            text="Nincs kiválasztva",
            font=("Segoe UI", 10, "bold"),
            fg="#0F766E",
            bg="#F8FAFC",
            width=15,
            anchor="w",
        )
        self.lbl_selected_div.pack(side="left", padx=(4, 8))

        tk.Label(
            score_bar,
            text="Pontszám (0–195 vagy fegyverhiba):",
            font=("Segoe UI", 9),
            fg="#334155",
            bg="#F8FAFC",
        ).pack(side="left")

        self.entry_score = tk.Entry(score_bar, font=("Segoe UI", 11), width=11, relief="solid", bd=1)
        self.entry_score.pack(side="left", padx=(6, 6), ipady=2)
        self.entry_score.bind("<Return>", lambda _e: self._save_score_for_selected())

        tk.Button(
            score_bar,
            text="Eredmény mentése",
            font=("Segoe UI", 9, "bold"),
            fg="#FFFFFF",
            bg="#0F766E",
            activebackground="#115E59",
            activeforeground="#FFFFFF",
            relief="flat",
            padx=12,
            pady=4,
            cursor="hand2",
            command=self._save_score_for_selected,
        ).pack(side="left", padx=(0, 6))

        tk.Button(
            score_bar,
            text="Fegyverhiba",
            font=("Segoe UI", 9, "bold"),
            fg="#991B1B",
            bg="#FEE2E2",
            activebackground="#FECACA",
            relief="flat",
            padx=10,
            pady=4,
            cursor="hand2",
            command=self._set_weapon_malfunction,
        ).pack(side="left")

        # Versenyző divíziói táblázat
        div_table_wrap = tk.Frame(right_card, bg="#FFFFFF")
        div_table_wrap.pack(fill="both", expand=True)

        cols = ("divizio", "eredmeny", "megjegyzes")
        self.tree_divs = ttk.Treeview(div_table_wrap, columns=cols, show="tree headings", height=5, selectmode="browse")
        self.tree_divs.heading("#0", text="Nevezve")
        self.tree_divs.heading("divizio", text="Divízió")
        self.tree_divs.heading("eredmeny", text="Eredmény")
        self.tree_divs.heading("megjegyzes", text="Állapot / Megjegyzés")

        self.tree_divs.column("#0", width=115, anchor="center", stretch=False)
        self.tree_divs.column("divizio", width=220, anchor="w")
        self.tree_divs.column("eredmeny", width=145, anchor="center")
        self.tree_divs.column("megjegyzes", width=220, anchor="w")

        vsb_div = ttk.Scrollbar(div_table_wrap, orient="vertical", command=self.tree_divs.yview)
        self.tree_divs.configure(yscrollcommand=vsb_div.set)
        self.tree_divs.pack(side="left", fill="both", expand=True)
        vsb_div.pack(side="right", fill="y")

        self.tree_divs.tag_configure("registered", foreground="#B91C1C", font=("Segoe UI", 12, "bold"))
        self.tree_divs.tag_configure("unregistered", foreground="#475569", font=("Segoe UI", 12))

        self.tree_divs.bind("<<TreeviewSelect>>", self._on_div_select)
        self.tree_divs.bind("<Button-1>", self._on_div_click)
        self.tree_divs.bind("<Double-1>", self._on_div_double_click)

        # --- ALSÓ KÁRTYA: Utolsó 10 versenyző áttekintése ---
        bottom_card = tk.Frame(
            frame,
            bg="#FFFFFF",
            padx=16,
            pady=12,
            highlightbackground="#CBD5E1",
            highlightthickness=1,
        )
        bottom_card.pack(side="top", fill="both", expand=True)

        bottom_hdr = tk.Frame(bottom_card, bg="#FFFFFF")
        bottom_hdr.pack(fill="x", pady=(0, 8))

        tk.Label(
            bottom_hdr,
            text="Utolsó 10 versenyző — regisztrált állapot és pontszámok",
            font=("Segoe UI", 14, "bold"),
            fg="#0F172A",
            bg="#FFFFFF",
        ).pack(side="left")

        self.lbl_overview_badge = tk.Label(
            bottom_hdr,
            text="0 versenyző, 0 nevezés, 0 hiányzó eredmény",
            font=("Segoe UI", 10, "bold"),
            fg="#0F766E",
            bg="#F0FDF4",
            padx=14,
            pady=4,
            highlightbackground="#BBF7D0",
            highlightthickness=1,
        )
        self.lbl_overview_badge.pack(side="right")

        ov_wrap = tk.Frame(bottom_card, bg="#FFFFFF")
        ov_wrap.pack(fill="both", expand=True)

        ov_cols = ["veng", "nev", "egyesulet", "allapot", "nevezesek", "hianyzik"] + [f"div_{i}" for i in range(len(DIVISIONS))]
        self.tree_overview = ttk.Treeview(ov_wrap, columns=ov_cols, show="headings", height=8, selectmode="browse")

        self.tree_overview.heading("veng", text="V.eng.")
        self.tree_overview.heading("nev", text="Név")
        self.tree_overview.heading("egyesulet", text="Egyesület")
        self.tree_overview.heading("allapot", text="Állapot")
        self.tree_overview.heading("nevezesek", text="Nevezés")
        self.tree_overview.heading("hianyzik", text="Hiány")

        self.tree_overview.tag_configure("has_reg", foreground="#DC2626", font=("Segoe UI", 10, "bold"))
        self.tree_overview.tag_configure("normal_reg", foreground="#0F172A", font=("Segoe UI", 10))

        self.tree_overview.column("veng", width=68, anchor="center", stretch=False)
        self.tree_overview.column("nev", width=170, anchor="w", stretch=False)
        self.tree_overview.column("egyesulet", width=200, anchor="w", stretch=False)
        self.tree_overview.column("allapot", width=185, anchor="w", stretch=False)
        self.tree_overview.column("nevezesek", width=80, anchor="center", stretch=False)
        self.tree_overview.column("hianyzik", width=70, anchor="center", stretch=False)

        for i, d in enumerate(DIVISIONS):
            col_id = f"div_{i}"
            self.tree_overview.heading(col_id, text=d)
            self.tree_overview.column(col_id, width=115, anchor="center", stretch=False)

        vsb_ov = ttk.Scrollbar(ov_wrap, orient="vertical", command=self.tree_overview.yview)
        hsb_ov = ttk.Scrollbar(ov_wrap, orient="horizontal", command=self.tree_overview.xview)
        self.tree_overview.configure(yscrollcommand=vsb_ov.set, xscrollcommand=hsb_ov.set)

        vsb_ov.pack(side="right", fill="y")
        hsb_ov.pack(side="bottom", fill="x")
        self.tree_overview.pack(side="left", fill="both", expand=True)

        self.tree_overview.bind("<Double-1>", self._on_overview_double_click)

        self._populate_empty_divisions()

    def _record_recent_competitor(self, lic, action_text=None):
        if not lic:
            return
        if lic in self.recent_competitors:
            self.recent_competitors.remove(lic)
        self.recent_competitors.insert(0, lic)
        if len(self.recent_competitors) > 10:
            self.recent_competitors = self.recent_competitors[:10]
        if action_text:
            self.recent_actions[lic] = action_text

    def _populate_empty_divisions(self):
        for item in self.tree_divs.get_children():
            self.tree_divs.delete(item)
        for div in DIVISIONS:
            self.tree_divs.insert("", "end", iid=div, image=self.img_unreg, values=(div, "—", "Nincs nevezve"), tags=("unregistered",))

    def _search_competitor(self, lic_override=None):
        raw = lic_override if lic_override is not None else self.entry_lic.get()
        lic = normalize_license(raw)
        if not lic:
            messagebox.showwarning("Hiányzó adat", "Add meg a versenyengedély számát!")
            return

        if lic not in self.licenses:
            self.current_lic = None
            self.comp_info_box.configure(bg="#FEF2F2", highlightbackground="#FECACA")
            self.lbl_comp_name_title.configure(text="Státusz", bg="#FEF2F2", fg="#991B1B")
            self.lbl_comp_name.configure(
                text="Adott versenyző nem indulhat!",
                bg="#FEF2F2",
                fg="#B91C1C",
            )
            self.lbl_comp_club_title.configure(text="Indoklás", bg="#FEF2F2", fg="#991B1B")
            self.lbl_comp_club.configure(
                text=f"A(z) '{lic}' versenyengedély szám nem található a nyilvántartásban.",
                bg="#FEF2F2",
                fg="#7F1D1D",
            )
            self._populate_empty_divisions()
            self.lbl_selected_div.configure(text="Nincs kiválasztva")
            self.entry_score.delete(0, "end")
            messagebox.showerror("Nem indulhat", f"Adott versenyző nem indulhat!\n(Versenyengedély: {lic} nem található.)")
            return

        self.current_lic = lic
        self._record_recent_competitor(lic, "Lekérdezve")
        info = self.licenses[lic]
        self.comp_info_box.configure(bg="#F0FDF4", highlightbackground="#BBF7D0")
        self.lbl_comp_name_title.configure(text="Név", bg="#F0FDF4", fg="#475569")
        self.lbl_comp_name.configure(text=info["name"], bg="#F0FDF4", fg="#0F172A")
        self.lbl_comp_club_title.configure(text="Egyesület", bg="#F0FDF4", fg="#475569")
        self.lbl_comp_club.configure(text=info["club"], bg="#F0FDF4", fg="#0F172A")

        self._refresh_competitor_divisions()
        self._refresh_overview_table()

    def _find_entry(self, div, lic):
        for entry in self.registrations[div]:
            if entry["lic"] == lic:
                return entry
        return None

    def _refresh_competitor_divisions(self):
        selected_before = self.tree_divs.selection()
        sel_id = selected_before[0] if selected_before else None

        for item in self.tree_divs.get_children():
            self.tree_divs.delete(item)

        if not self.current_lic:
            self._populate_empty_divisions()
            return

        for div in DIVISIONS:
            entry = self._find_entry(div, self.current_lic)
            if entry is not None:
                sc = entry["score"]
                if sc is None:
                    score_disp = "Regisztrált"
                    status_disp = "Pontszámra vár"
                elif str(sc).lower() == "fegyverhiba":
                    score_disp = "fegyverhiba"
                    status_disp = "Fegyverhiba rögzítve"
                else:
                    score_disp = f"{sc} pont"
                    status_disp = "Eredmény rögzítve"
                self.tree_divs.insert(
                    "",
                    "end",
                    iid=div,
                    image=self.img_reg,
                    values=(div, score_disp, status_disp),
                    tags=("registered",),
                )
            else:
                self.tree_divs.insert(
                    "",
                    "end",
                    iid=div,
                    image=self.img_unreg,
                    values=(div, "—", "Kattints a nevezéshez"),
                    tags=("unregistered",),
                )

        if sel_id and sel_id in DIVISIONS:
            self.tree_divs.selection_set(sel_id)
            self._on_div_select()

    def _toggle_division_registration(self, div):
        if not self.current_lic:
            messagebox.showinfo("Nincs versenyző", "Előbb keress rá egy érvényes versenyengedély számra!")
            return

        self._record_recent_competitor(self.current_lic)
        existing = self._find_entry(div, self.current_lic)
        if existing is None:
            # Új nevezés felvétele pontszám nélkül -> "Regisztrált"
            self.registrations[div].append(
                {
                    "lic": self.current_lic,
                    "score": None,
                    "unlicensed_name": None,
                    "hits": None,
                    "note": None,
                    "hit_counts": {k: 0 for k in HIT_TIERS},
                }
            )
            self._record_recent_competitor(self.current_lic, f"Regisztrálva: {div}")
            self._save_to_excel()
            self._refresh_competitor_divisions()
            self._refresh_overview_table()
            self.tree_divs.selection_set(div)
            self._on_div_select()
        else:
            if existing["score"] is not None:
                if not messagebox.askyesno(
                    "Nevezés törlése",
                    f"A versenyzőnek már van rögzített eredménye ({existing['score']}) a(z) '{div}' versenyszámban.\n"
                    "Biztosan törlöd a nevezést és az eredményt?",
                ):
                    return
            self.registrations[div] = [e for e in self.registrations[div] if e["lic"] != self.current_lic]
            self._record_recent_competitor(self.current_lic, f"Törölve: {div}")
            self._save_to_excel()
            self._refresh_competitor_divisions()
            self._refresh_overview_table()

    def _on_div_click(self, event):
        if not self.current_lic:
            return
        region = self.tree_divs.identify("region", event.x, event.y)
        col = self.tree_divs.identify_column(event.x)
        row_id = self.tree_divs.identify_row(event.y)
        if region in ("cell", "tree") and row_id in DIVISIONS and col == "#0":
            self._toggle_division_registration(row_id)

    def _on_div_double_click(self, event):
        if not self.current_lic:
            return
        row_id = self.tree_divs.identify_row(event.y)
        if row_id not in DIVISIONS:
            return
        entry = self._find_entry(row_id, self.current_lic)
        if entry is None:
            # Ha még nem volt nevezve, dupla klikkre nevezzük be és fókuszáljunk a pont mezőre
            self._toggle_division_registration(row_id)
        self.tree_divs.selection_set(row_id)
        self._on_div_select()
        self.entry_score.focus_set()
        self.entry_score.select_range(0, "end")

    def _on_div_select(self, _event=None):
        sel = self.tree_divs.selection()
        if not sel:
            self.lbl_selected_div.configure(text="Nincs kiválasztva")
            return
        div = sel[0]
        self.lbl_selected_div.configure(text=div)
        self.entry_score.delete(0, "end")
        if self.current_lic:
            entry = self._find_entry(div, self.current_lic)
            if entry and entry["score"] is not None:
                self.entry_score.insert(0, str(entry["score"]))

    def _remove_selected_division(self):
        if not self.current_lic:
            return
        sel = self.tree_divs.selection()
        if not sel:
            messagebox.showinfo("Nincs kijelölés", "Jelölj ki egy versenyszámot a törléshez!")
            return
        div = sel[0]
        entry = self._find_entry(div, self.current_lic)
        if entry is None:
            messagebox.showinfo("Nincs nevezve", f"A versenyző nincs nevezve a(z) '{div}' versenyszámba.")
            return
        self._toggle_division_registration(div)

    def _save_score_for_selected(self):
        if self.mode_var.get() == "Eredményjelző":
            messagebox.showwarning("Csak olvasható mód", "Eredményjelző módban az eredmények nem módosíthatók.")
            return
        if not self.current_lic:
            messagebox.showwarning("Nincs versenyző", "Előbb keress rá egy versenyzőre a versenyengedély száma alapján!")
            return
        sel = self.tree_divs.selection()
        if not sel:
            messagebox.showwarning("Nincs versenyszám", "Jelöld ki azt a versenyszámot, amelyikbe a pontszámot beírod!")
            return

        div = sel[0]
        entry = self._find_entry(div, self.current_lic)
        if entry is None:
            messagebox.showwarning(
                "Nincs regisztrálva",
                f"A versenyző nincs regisztrálva a(z) '{div}' versenyszámba!\n"
                "Előbb regisztráld a versenyszámra (kattints a célkeresztre a Nevezve oszlopban).",
            )
            return

        try:
            validated = validate_score_input(self.entry_score.get())
        except ValueError as err:
            messagebox.showerror("Érvénytelen pontszám", str(err))
            return

        entry["score"] = validated
        sc_label = "Regisztrált" if validated is None else ("fegyverhiba" if str(validated).lower() == "fegyverhiba" else f"{validated} pont")
        self._record_recent_competitor(self.current_lic, f"{div}: {sc_label}")
        self._save_to_excel()
        self._refresh_competitor_divisions()
        self._refresh_overview_table()

    def _set_weapon_malfunction(self):
        if not self.current_lic:
            return
        sel = self.tree_divs.selection()
        if not sel:
            messagebox.showwarning("Nincs versenyszám", "Jelöld ki a versenyszámot a listában!")
            return
        self.entry_score.delete(0, "end")
        self.entry_score.insert(0, "fegyverhiba")
        self._save_score_for_selected()

    def _on_overview_double_click(self, _event):
        sel = self.tree_overview.selection()
        if not sel:
            return
        lic = sel[0]
        self.entry_lic.delete(0, "end")
        self.entry_lic.insert(0, lic)
        self._search_competitor(lic)

    def _collect_registered_competitors(self):
        """Összegyűjti a versenyre regisztrált összes egyedi versenyzőt és adataikat."""
        comp_map = {}
        for div in DIVISIONS:
            for entry in self.registrations[div]:
                lic = entry["lic"]
                if lic not in comp_map:
                    info = self.licenses.get(lic, {"name": f"Ismeretlen ({lic})", "club": ""})
                    comp_map[lic] = {
                        "lic": lic,
                        "name": info["name"],
                        "club": info["club"],
                        "divisions": {},
                    }
                comp_map[lic]["divisions"][div] = entry["score"]
        return comp_map

    def _refresh_overview_table(self):
        for item in self.tree_overview.get_children():
            self.tree_overview.delete(item)

        comp_map = self._collect_registered_competitors()
        total_regs = 0
        missing_count = 0
        for comp in comp_map.values():
            reg_divs = comp["divisions"]
            total_regs += len(reg_divs)
            missing_count += sum(1 for sc in reg_divs.values() if sc is None)

        # Az utolsó 10 versenyző összeállítása (legutóbb kezelt/keresett elöl, kiegészítve a munkafüzetben már szereplőkkel)
        ordered_lics = []
        for lic in self.recent_competitors:
            if lic in self.licenses and lic not in ordered_lics:
                ordered_lics.append(lic)
        for lic in comp_map.keys():
            if lic not in ordered_lics:
                ordered_lics.append(lic)

        last_10_lics = ordered_lics[:10]

        for lic in last_10_lics:
            if lic in comp_map:
                comp = comp_map[lic]
            else:
                info = self.licenses.get(lic, {"name": f"Ismeretlen ({lic})", "club": ""})
                comp = {
                    "lic": lic,
                    "name": info["name"],
                    "club": info["club"],
                    "divisions": {},
                }

            reg_divs = comp["divisions"]
            n_reg = len(reg_divs)
            n_miss = sum(1 for sc in reg_divs.values() if sc is None)

            div_cells = []
            for d in DIVISIONS:
                if d not in reg_divs:
                    div_cells.append("")
                else:
                    sc = reg_divs[d]
                    if sc is None:
                        div_cells.append("Regisztrált")
                    elif str(sc).lower() == "fegyverhiba":
                        div_cells.append("fegyverhiba")
                    else:
                        div_cells.append(f"{sc} pont")

            if n_reg == 0:
                status_txt = "Nincs regisztrálva"
            elif n_miss == 0:
                status_txt = "Minden rögzítve"
            else:
                status_txt = f"{n_miss} hiány"

            row_tag = "has_reg" if n_miss > 0 else "normal_reg"
            row_vals = [comp["lic"], comp["name"], comp["club"], status_txt, n_reg, n_miss] + div_cells
            self.tree_overview.insert("", "end", iid=comp["lic"], values=row_vals, tags=(row_tag,))

        self.lbl_overview_badge.configure(
            text=f"Megjelenítve: utolsó {len(last_10_lics)} fő | Összesen: {len(comp_map)} versenyző, {total_regs} nevezés, {missing_count} hiányzó eredmény"
        )

    # ------------------------------------------------------------------
    # FÜL 2: VERSENYZŐK (Regisztrált felhasználók + rendezhető oszlopok)
    # ------------------------------------------------------------------
    def _build_tab_versenyzok(self):
        frame = tk.Frame(self.content_area, bg="#F1F5F9", padx=16, pady=14)
        self.tab_frames["versenyzok"] = frame

        card = tk.Frame(
            frame,
            bg="#FFFFFF",
            padx=20,
            pady=16,
            highlightbackground="#CBD5E1",
            highlightthickness=1,
        )
        card.pack(fill="both", expand=True)

        hdr = tk.Frame(card, bg="#FFFFFF")
        hdr.pack(fill="x", pady=(0, 12))

        tk.Label(
            hdr,
            text="Versenyzők (regisztrált indulók)",
            font=("Segoe UI", 16, "bold"),
            fg="#0F172A",
            bg="#FFFFFF",
        ).pack(side="left")

        self.lbl_reg_users_badge = tk.Label(
            hdr,
            text="● Regisztrált versenyzők: 0",
            font=("Segoe UI", 10, "bold"),
            fg="#0F766E",
            bg="#F0FDF4",
            padx=14,
            pady=5,
            highlightbackground="#BBF7D0",
            highlightthickness=1,
        )
        self.lbl_reg_users_badge.pack(side="right")

        # Keresőmezők oszloponként
        filter_row = tk.Frame(card, bg="#FFFFFF")
        filter_row.pack(fill="x", pady=(0, 12))

        self.filter_lic_var = tk.StringVar()
        self.filter_name_var = tk.StringVar()
        self.filter_club_var = tk.StringVar()

        for var in (self.filter_lic_var, self.filter_name_var, self.filter_club_var):
            var.trace_add("write", lambda *_a: self._refresh_versenyzok_tab())

        tk.Label(filter_row, text="V.eng.:", font=("Segoe UI", 9, "bold"), fg="#475569", bg="#FFFFFF").pack(side="left")
        tk.Entry(filter_row, textvariable=self.filter_lic_var, font=("Segoe UI", 10), width=14, relief="solid", bd=1).pack(
            side="left", padx=(4, 14), ipady=4
        )

        tk.Label(filter_row, text="Név:", font=("Segoe UI", 9, "bold"), fg="#475569", bg="#FFFFFF").pack(side="left")
        tk.Entry(filter_row, textvariable=self.filter_name_var, font=("Segoe UI", 10), width=26, relief="solid", bd=1).pack(
            side="left", padx=(4, 14), ipady=4
        )

        tk.Label(filter_row, text="Egyesület:", font=("Segoe UI", 9, "bold"), fg="#475569", bg="#FFFFFF").pack(side="left")
        tk.Entry(filter_row, textvariable=self.filter_club_var, font=("Segoe UI", 10), width=32, relief="solid", bd=1).pack(
            side="left", padx=(4, 14), ipady=4
        )

        tk.Label(
            filter_row,
            text="(Kattints bármelyik oszlopfejlécre a rendezéshez)",
            font=("Segoe UI", 9),
            fg="#64748B",
            bg="#FFFFFF",
        ).pack(side="right")

        tbl_wrap = tk.Frame(card, bg="#FFFFFF")
        tbl_wrap.pack(fill="both", expand=True)

        cols = ("veng", "nev", "egyesulet", "nevezes_db", "versenyszamok", "allapot")
        self.tree_users = ttk.Treeview(tbl_wrap, columns=cols, show="headings", selectmode="browse")

        headers = {
            "veng": "V.eng.",
            "nev": "Név",
            "egyesulet": "Egyesület",
            "nevezes_db": "Nevezések",
            "versenyszamok": "Regisztrált versenyszámok és eredmények",
            "allapot": "Státusz",
        }
        widths = {
            "veng": 90,
            "nev": 210,
            "egyesulet": 260,
            "nevezes_db": 100,
            "versenyszamok": 390,
            "allapot": 130,
        }
        anchors = {
            "veng": "center",
            "nev": "w",
            "egyesulet": "w",
            "nevezes_db": "center",
            "versenyszamok": "w",
            "allapot": "center",
        }

        for col in cols:
            self.tree_users.heading(
                col,
                text=headers[col],
                command=lambda c=col: self._sort_treeview_column(self.tree_users, c, numeric=(c in ("veng", "nevezes_db"))),
            )
            self.tree_users.column(col, width=widths[col], anchor=anchors[col])

        vsb = ttk.Scrollbar(tbl_wrap, orient="vertical", command=self.tree_users.yview)
        self.tree_users.configure(yscrollcommand=vsb.set)
        self.tree_users.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        self.tree_users.bind("<Double-1>", self._on_user_row_double_click)

    def _sort_treeview_column(self, tree, col, numeric=False):
        key = (id(tree), col)
        reverse = self.sort_reverse.get(key, False)

        items = [(tree.set(k, col), k) for k in tree.get_children("")]
        if numeric:
            def num_key(pair):
                try:
                    return float(str(pair[0]).strip())
                except ValueError:
                    return -1.0
            items.sort(key=num_key, reverse=reverse)
        else:
            items.sort(key=lambda pair: str(pair[0]).lower(), reverse=reverse)

        for idx, (_val, k) in enumerate(items):
            tree.move(k, "", idx)

        self.sort_reverse[key] = not reverse

    def _on_user_row_double_click(self, _event):
        sel = self.tree_users.selection()
        if not sel:
            return
        lic = sel[0]
        self._switch_tab("eredmenybevitel")
        self.entry_lic.delete(0, "end")
        self.entry_lic.insert(0, lic)
        self._search_competitor(lic)

    def _refresh_versenyzok_tab(self):
        for item in self.tree_users.get_children():
            self.tree_users.delete(item)

        comp_map = self._collect_registered_competitors()
        self.lbl_reg_users_badge.configure(
            text=f"● Regisztrált versenyzők: {len(comp_map)} fő"
        )

        f_lic = self.filter_lic_var.get().strip().lower()
        f_name = self.filter_name_var.get().strip().lower()
        f_club = self.filter_club_var.get().strip().lower()

        for comp in sorted(comp_map.values(), key=lambda c: c["name"].lower()):
            if f_lic and f_lic not in comp["lic"].lower():
                continue
            if f_name and f_name not in comp["name"].lower():
                continue
            if f_club and f_club not in comp["club"].lower():
                continue

            parts = []
            missing = 0
            for d in DIVISIONS:
                if d in comp["divisions"]:
                    sc = comp["divisions"][d]
                    if sc is None:
                        parts.append(f"{d} (Regisztrált)")
                        missing += 1
                    elif str(sc).lower() == "fegyverhiba":
                        parts.append(f"{d}: fegyverhiba")
                    else:
                        parts.append(f"{d}: {sc} pont")

            status = "Teljesítve" if missing == 0 else f"{missing} függőben"
            self.tree_users.insert(
                "",
                "end",
                iid=comp["lic"],
                values=(
                    comp["lic"],
                    comp["name"],
                    comp["club"],
                    len(comp["divisions"]),
                    ", ".join(parts),
                    status,
                ),
            )

    # ------------------------------------------------------------------
    # FÜL 3: STATISZTIKA
    # ------------------------------------------------------------------
    def _build_tab_statisztika(self):
        frame = tk.Frame(self.content_area, bg="#F1F5F9", padx=16, pady=14)
        self.tab_frames["statisztika"] = frame

        # Felső 3 KPI kártya
        kpi_row = tk.Frame(frame, bg="#F1F5F9")
        kpi_row.pack(fill="x", pady=(0, 14))
        for i in range(3):
            kpi_row.columnconfigure(i, weight=1)

        self.kpi_labels = {}
        kpi_defs = [
            ("reg_competitors", "Regisztrált versenyzők száma", "#0F172A"),
            ("total_entries", "Összes nevezés (indulás)", "#1D4ED8"),
            ("pending_entries", "Regisztrált, eredmény nélküli versenyszámok", "#B45309"),
        ]

        for idx, (key, title, accent) in enumerate(kpi_defs):
            box = tk.Frame(
                kpi_row,
                bg="#FFFFFF",
                padx=18,
                pady=16,
                highlightbackground="#CBD5E1",
                highlightthickness=1,
            )
            box.grid(row=0, column=idx, sticky="nsew", padx=(0 if idx == 0 else 8, 0))

            tk.Label(
                box,
                text=title,
                font=("Segoe UI", 10, "bold"),
                fg="#475569",
                bg="#FFFFFF",
                anchor="w",
                wraplength=250,
                justify="left",
            ).pack(fill="x")

            val_lbl = tk.Label(
                box,
                text="0",
                font=("Segoe UI", 26, "bold"),
                fg=accent,
                bg="#FFFFFF",
                anchor="w",
            )
            val_lbl.pack(fill="x", pady=(8, 0))
            self.kpi_labels[key] = val_lbl

        # Alsó kártya: Versenyszámonkénti bontás
        card = tk.Frame(
            frame,
            bg="#FFFFFF",
            padx=20,
            pady=16,
            highlightbackground="#CBD5E1",
            highlightthickness=1,
        )
        card.pack(fill="both", expand=True)

        tk.Label(
            card,
            text="Versenyszámonkénti statisztika",
            font=("Segoe UI", 15, "bold"),
            fg="#0F172A",
            bg="#FFFFFF",
            anchor="w",
        ).pack(fill="x", pady=(0, 10))

        cols = ("divizio", "indulok", "ertekelt", "hianyzo", "fegyverhiba", "legjobb")
        self.tree_stats = ttk.Treeview(card, columns=cols, show="headings", height=11)

        self.tree_stats.heading("divizio", text="Versenyszám")
        self.tree_stats.heading("indulok", text="Indulók száma (fő)")
        self.tree_stats.heading("ertekelt", text="Eredménnyel rendelkezik")
        self.tree_stats.heading("hianyzo", text="Regisztrált (eredmény nélkül)")
        self.tree_stats.heading("fegyverhiba", text="Fegyverhiba")
        self.tree_stats.heading("legjobb", text="Jelenlegi legjobb eredmény")

        self.tree_stats.column("divizio", width=220, anchor="w")
        self.tree_stats.column("indulok", width=160, anchor="center")
        self.tree_stats.column("ertekelt", width=180, anchor="center")
        self.tree_stats.column("hianyzo", width=210, anchor="center")
        self.tree_stats.column("fegyverhiba", width=130, anchor="center")
        self.tree_stats.column("legjobb", width=220, anchor="center")

        self.tree_stats.pack(fill="both", expand=True)

    def _refresh_statisztika_tab(self):
        comp_map = self._collect_registered_competitors()
        total_entries = 0
        pending_entries = 0
        active_divisions = 0

        for item in self.tree_stats.get_children():
            self.tree_stats.delete(item)

        for div in DIVISIONS:
            entries = self.registrations[div]
            n_all = len(entries)
            if n_all > 0:
                active_divisions += 1
            n_pending = sum(1 for e in entries if e["score"] is None)
            n_malf = sum(1 for e in entries if str(e["score"]).lower() == "fegyverhiba")
            n_scored = n_all - n_pending

            total_entries += n_all
            pending_entries += n_pending

            numeric_scores = [
                float(e["score"])
                for e in entries
                if e["score"] is not None and str(e["score"]).lower() != "fegyverhiba"
            ]
            best_str = f"{int(max(numeric_scores)) if max(numeric_scores).is_integer() else max(numeric_scores)} pont" if numeric_scores else "—"

            self.tree_stats.insert(
                "",
                "end",
                values=(div, n_all, n_scored, n_pending, n_malf, best_str),
            )

        self.kpi_labels["reg_competitors"].configure(text=f"{len(comp_map)} fő")
        self.kpi_labels["total_entries"].configure(text=f"{total_entries} db")
        self.kpi_labels["pending_entries"].configure(text=f"{pending_entries} db")

    # ------------------------------------------------------------------
    # FÜL 4: EREDMÉNYEK (Versenyszámonként az első három versenyző)
    # ------------------------------------------------------------------
    def _build_tab_eredmenyek(self):
        frame = tk.Frame(self.content_area, bg="#F1F5F9", padx=16, pady=14)
        self.tab_frames["eredmenyek"] = frame

        hdr = tk.Frame(frame, bg="#F1F5F9")
        hdr.pack(fill="x", pady=(0, 10))

        tk.Label(
            hdr,
            text="Eredmények — Dobogós helyezettek versenyszámonként (Top 3)",
            font=("Segoe UI", 16, "bold"),
            fg="#0F172A",
            bg="#F1F5F9",
        ).pack(side="left")

        tk.Button(
            hdr,
            text="Frissítés",
            font=("Segoe UI", 10, "bold"),
            fg="#FFFFFF",
            bg="#0F766E",
            activebackground="#115E59",
            relief="flat",
            padx=14,
            pady=6,
            cursor="hand2",
            command=self._refresh_eredmenyek_tab,
        ).pack(side="right")

        # Görgethető rács a 10 versenyszám dobogós kártyáihoz
        canvas = tk.Canvas(frame, bg="#F1F5F9", highlightthickness=0)
        vsb = ttk.Scrollbar(frame, orient="vertical", command=canvas.yview)
        self.podium_container = tk.Frame(canvas, bg="#F1F5F9")

        self.podium_container.bind(
            "<Configure>",
            lambda _e: canvas.configure(scrollregion=canvas.bbox("all")),
        )
        canvas_win = canvas.create_window((0, 0), window=self.podium_container, anchor="nw")
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(canvas_win, width=e.width))
        canvas.configure(yscrollcommand=vsb.set)

        canvas.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        self.podium_container.columnconfigure(0, weight=1)
        self.podium_container.columnconfigure(1, weight=1)

    def _refresh_eredmenyek_tab(self):
        for w in self.podium_container.winfo_children():
            w.destroy()

        medal_colors = {
            1: ("1. hely", "#FEF3C7", "#B45309"),
            2: ("2. hely", "#F1F5F9", "#334155"),
            3: ("3. hely", "#FFEDD5", "#9A3412"),
        }

        for idx, div in enumerate(DIVISIONS):
            r = idx // 2
            c = idx % 2

            card = tk.Frame(
                self.podium_container,
                bg="#FFFFFF",
                padx=16,
                pady=12,
                highlightbackground="#CBD5E1",
                highlightthickness=1,
            )
            card.grid(row=r, column=c, sticky="nsew", padx=(0 if c == 0 else 8, 0), pady=(0, 10))

            top_row = tk.Frame(card, bg="#0F172A", padx=10, pady=6)
            top_row.pack(fill="x", pady=(0, 8))

            tk.Label(
                top_row,
                text=div,
                font=("Segoe UI", 11, "bold"),
                fg="#FFFFFF",
                bg="#0F172A",
            ).pack(side="left")

            entries = self.registrations[div]
            tk.Label(
                top_row,
                text=f"Indulók: {len(entries)} fő",
                font=("Segoe UI", 9),
                fg="#CBD5E1",
                bg="#0F172A",
            ).pack(side="right")

            # Csak a számszerű pontszámmal rendelkezők kerülnek a dobogóra, pontszám + (13..9 találat) szerint rendezve
            scored = [
                e
                for e in entries
                if e["score"] is not None and str(e["score"]).lower() != "fegyverhiba"
            ]
            scored.sort(
                key=lambda item: (
                    -float(item["score"]),
                    _hit_tuple(item.get("hit_counts")),
                )
            )
            top3 = scored[:3]

            # Pontazonosság vizsgálata az első három helyen (vagy azonos pontszámú dobogós/esélyes versenyzőknél)
            top3_scores = [float(x["score"]) for x in top3]
            has_top3_tie = len(top3_scores) != len(set(top3_scores)) or any(
                float(x["score"]) in top3_scores for x in scored[3:]
            )

            for place in (1, 2, 3):
                label_txt, bg_badge, fg_badge = medal_colors[place]
                row = tk.Frame(card, bg="#FFFFFF", pady=3)
                row.pack(fill="x")

                tk.Label(
                    row,
                    text=label_txt,
                    font=("Segoe UI", 9, "bold"),
                    fg=fg_badge,
                    bg=bg_badge,
                    width=8,
                    pady=2,
                ).pack(side="left")

                if place <= len(top3):
                    e = top3[place - 1]
                    info = self.licenses.get(e["lic"], {"name": e["lic"], "club": ""})
                    hc = e.get("hit_counts") or {}
                    hc_parts = [f"{k}:{hc.get(k, 0)}" for k in HIT_TIERS if int(hc.get(k, 0) or 0) > 0]
                    hc_suffix = f"  [{', '.join(hc_parts)}]" if hc_parts else ""

                    tk.Label(
                        row,
                        text=f"  {info['name']} ({info['club']}){hc_suffix}",
                        font=("Segoe UI", 10, "bold" if place == 1 else "normal"),
                        fg="#0F172A",
                        bg="#FFFFFF",
                        anchor="w",
                    ).pack(side="left", fill="x", expand=True)

                    tk.Label(
                        row,
                        text=f"{e['score']} pont",
                        font=("Segoe UI", 10, "bold"),
                        fg="#0F766E",
                        bg="#FFFFFF",
                    ).pack(side="right")
                else:
                    tk.Label(
                        row,
                        text="  — Még nincs értékelhető eredmény —",
                        font=("Segoe UI", 9, "italic"),
                        fg="#94A3B8",
                        bg="#FFFFFF",
                        anchor="w",
                    ).pack(side="left", fill="x", expand=True)

            if has_top3_tie:
                tie_bar = tk.Frame(card, bg="#FEF3C7", padx=10, pady=6)
                tie_bar.pack(fill="x", pady=(8, 0))
                tk.Label(
                    tie_bar,
                    text="Pontazonosság az élmezőnyben — 13–9 találatok megadása szükséges",
                    font=("Segoe UI", 9, "bold"),
                    fg="#92400E",
                    bg="#FEF3C7",
                ).pack(side="left")
                tk.Button(
                    tie_bar,
                    text="Találatok (13–9) beírása",
                    font=("Segoe UI", 9, "bold"),
                    fg="#FFFFFF",
                    bg="#B45309",
                    activebackground="#92400E",
                    relief="flat",
                    padx=10,
                    pady=3,
                    cursor="hand2",
                    command=lambda d=div: self._open_tiebreak_dialog(d),
                ).pack(side="right")



    def _open_tiebreak_dialog(self, div):
        entries = [
            e for e in self.registrations[div]
            if e["score"] is not None and str(e["score"]).lower() != "fegyverhiba"
        ]
        entries.sort(key=lambda item: (-float(item["score"]), _hit_tuple(item.get("hit_counts"))))

        # Azonos pontszámú csoportok az első 3 hely környékén
        top_scores = {float(e["score"]) for e in entries[:3]}
        score_counts = {}
        for e in entries:
            sc_f = float(e["score"])
            score_counts[sc_f] = score_counts.get(sc_f, 0) + 1

        tied_entries = [e for e in entries if float(e["score"]) in top_scores and score_counts[float(e["score"])] > 1]
        if not tied_entries:
            tied_entries = entries[:3]

        win = tk.Toplevel(self)
        win.title(f"Pontazonosság rendezése — {div}")
        win.geometry("760x380")
        win.configure(bg="#F8FAFC")
        win.transient(self)
        win.grab_set()

        tk.Label(
            win,
            text=f"{div} — 13, 12, 11, 10, 9 találatok számossága",
            font=("Segoe UI", 13, "bold"),
            fg="#0F172A",
            bg="#F8FAFC",
            padx=16,
            pady=10,
        ).pack(anchor="w")

        grid_frame = tk.Frame(win, bg="#FFFFFF", padx=14, pady=12, highlightbackground="#CBD5E1", highlightthickness=1)
        grid_frame.pack(fill="both", expand=True, padx=16, pady=(0, 12))

        headers = ["Versenyző", "Pont", "13 db", "12 db", "11 db", "10 db", "9 db"]
        for col_i, h_txt in enumerate(headers):
            tk.Label(
                grid_frame,
                text=h_txt,
                font=("Segoe UI", 10, "bold"),
                fg="#0F172A",
                bg="#FFFFFF",
                padx=6,
                pady=6,
            ).grid(row=0, column=col_i, sticky="w" if col_i == 0 else "ew")

        entry_vars = {}
        for row_i, ent in enumerate(tied_entries, start=1):
            info = self.licenses.get(ent["lic"], {"name": ent["lic"], "club": ""})
            tk.Label(
                grid_frame,
                text=f"{info['name']} ({info['club']})",
                font=("Segoe UI", 10),
                fg="#0F172A",
                bg="#FFFFFF",
                padx=6,
                pady=4,
            ).grid(row=row_i, column=0, sticky="w")

            tk.Label(
                grid_frame,
                text=f"{ent['score']}",
                font=("Segoe UI", 10, "bold"),
                fg="#0F766E",
                bg="#FFFFFF",
                padx=6,
            ).grid(row=row_i, column=1)

            hc = ent.get("hit_counts") or {}
            row_map = {}
            for tier_idx, tier in enumerate(HIT_TIERS, start=2):
                var = tk.StringVar(value=str(int(hc.get(tier, 0) or 0)))
                sp = tk.Entry(grid_frame, textvariable=var, width=6, justify="center", font=("Segoe UI", 10), relief="solid", bd=1)
                sp.grid(row=row_i, column=tier_idx, padx=4, pady=4, ipady=2)
                row_map[tier] = var
            entry_vars[ent["lic"]] = row_map

        def _apply_tiebreak():
            for ent in tied_entries:
                new_hc = {}
                for tier in HIT_TIERS:
                    raw_v = entry_vars[ent["lic"]][tier].get().strip()
                    if raw_v == "":
                        val_i = 0
                    else:
                        if not raw_v.isdigit():
                            messagebox.showerror("Érvénytelen érték", f"A(z) {tier} találatos mezőben csak nemnegatív egész szám adható meg!", parent=win)
                            return
                        val_i = int(raw_v)
                    new_hc[tier] = val_i
                ent["hit_counts"] = new_hc
                hc_summary = ", ".join(f"{k}:{new_hc[k]}" for k in HIT_TIERS if new_hc[k] > 0)
                if hc_summary:
                    base_note = str(ent.get("note") or "").strip()
                    ent["note"] = f"{base_note} | {hc_summary}".strip(" |") if base_note else hc_summary

            self._save_to_excel()
            self._refresh_eredmenyek_tab()
            win.destroy()

        btn_row = tk.Frame(win, bg="#F8FAFC", padx=16, pady=8)
        btn_row.pack(fill="x")

        tk.Button(
            btn_row,
            text="Mégse",
            font=("Segoe UI", 10),
            fg="#334155",
            bg="#E2E8F0",
            relief="flat",
            padx=14,
            pady=6,
            cursor="hand2",
            command=win.destroy,
        ).pack(side="right", padx=(8, 0))

        tk.Button(
            btn_row,
            text="Mentés és sorrendezés",
            font=("Segoe UI", 10, "bold"),
            fg="#FFFFFF",
            bg="#0F766E",
            activebackground="#115E59",
            relief="flat",
            padx=16,
            pady=6,
            cursor="hand2",
            command=_apply_tiebreak,
        ).pack(side="right")


if __name__ == "__main__":
    app = LoterModernApp()
    app.mainloop()
