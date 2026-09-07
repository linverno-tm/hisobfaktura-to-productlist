#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Hisob-faktura (.xls, aslida HTML) fayllarini SmartPOS savdo saytiga
yuklanadigan "product_list" Excel shabloniga aylantiradigan desktop dastur.

Ustunlar POZITSIYASI bo'yicha emas, balki har bir faylning o'z sarlavha
matniga qarab DINAMIK aniqlanadi — chunki turli ta'minotchilarning hisob-
fakturalarida ustunlar soni farq qilishi mumkin (masalan, markировкasiz
tovarlarda "Маркировка коди" ustuni umuman bo'lmaydi).

Bu fayl GitHub'da saqlanadi va launcher.py orqali har ishga tushganda
avtomatik yangilanadi — bu yerni tahrirlash = barcha foydalanuvchilarning
dasturi keyingi ochilishda yangilanadi degani.
"""
__version__ = "2026-09-07.1"

import os
import re
import sys
import threading
import queue
import datetime

import openpyxl
from openpyxl.worksheet.datavalidation import DataValidation
from bs4 import BeautifulSoup

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# ---------------------------------------------------------------------------
# Shablon konstantalari
# ---------------------------------------------------------------------------

HEADERS = [
    "Название товара", "Штрих код", "Ед. измерения", "Цена", "НДС %",
    "Избранное", "Маркировка", "ИКПУ", "Код ед. измерения тасниф",
]

UNITS = [
    "упаковка", "килограмм", "миллилитр", "блок", "коробка", "куб", "ампула",
    "часы", "тюбик", "флакон", "пластиковые банки / бутылки", "проценты",
    "комплект", "куб.м", "тонна", "пачка", "кв. дециметр", "комплект / набор",
    "банка", "мешок", "км", "порция", "сум", "пара", "сотих", "гектар",
    "кв. м.", "грамм", "пг. м.", "сутка", "кВа", "метр", "фут", "блистер",
    "литр", "штук", "человек",
]
VATS = ["Без НДС", "6.00", "0.00", "12.00"]
FAVOURITES = ["Да", "Нет"]

DEFAULT_BARCODE = "1111111111111"  # shtrix kodi yo'q tovarlar uchun (13 ta "1")

# hisob-fakturadagi lotin/o'zbekcha o'lchov birliklarini sayt kutgan
# rus tilidagi nomlarga moslash. Kerak bo'lsa shu ro'yxatga qo'shib boring.
UNIT_MAP = {
    "dona": "штук",
    "kg": "килограмм", "kilogramm": "килограмм",
    "litr": "литр", "l": "литр",
    "metr": "метр", "m": "метр",
    "komplekt": "комплект", "to'plam": "комплект", "toplam": "комплект",
    "quti": "коробка", "karobka": "коробка", "karobka/quti": "коробка",
    "blok": "блок",
    "para": "пара", "juft": "пара",
    "tonna": "тонна",
    "gramm": "грамм", "g": "грамм",
    "flakon": "флакон",
    "tyubik": "тюбик",
    "banka": "банка",
    "paket": "мешок", "meshok": "мешок", "qop": "мешок",
    "blister": "блистер",
    "soat": "часы", "chas": "часы",
    "kub": "куб", "kub.m": "куб.м", "kubm": "куб.м",
    "ampula": "ампула",
    "pachka": "пачка",
    "sum": "сум",
    "km": "км",
    "porsiya": "порция",
    "sutka": "сутка",
    "foiz": "проценты", "protsent": "проценты",
    "inson": "человек", "kishi": "человек", "chelovek": "человек",
    "gektar": "гектар",
    "sotix": "сотих", "sotih": "сотих",
}


# ---------------------------------------------------------------------------
# Hisob-fakturani o'qish (dinamik ustun aniqlash)
# ---------------------------------------------------------------------------

def _colspan(cell):
    raw = cell.get("colspan", "1")
    try:
        return max(1, int(float(raw)))
    except (TypeError, ValueError):
        return 1


def _parse_header_row(tr):
    """tr ichidagi har bir katakning MATNI -> boshlang'ich ustun indeksi.

    colspan hisobga olinadi, shu bilan pastdagi data qatorlaridagi haqiqiy
    katak indekslari bilan mos keladi.
    """
    mapping = {}
    col = 0
    for cell in tr.find_all(["td", "th"], recursive=False):
        label = cell.get_text(strip=True)
        if label and label not in mapping:
            mapping[label] = col
        col += _colspan(cell)
    return mapping


def _find_header_map(thead):
    for tr in thead.find_all("tr"):
        text = tr.get_text().lower()
        if "маҳсулот" in text or "наименование" in text:
            return _parse_header_row(tr)
    return None


def _resolve_field_indices(header_map):
    """Sarlavha matnlariga qarab har bir maydonning ustun-indeksini topadi."""
    idx = {}

    def claim(field, col):
        idx.setdefault(field, col)

    for label, col in header_map.items():
        low = label.lower()
        if "маҳсулот" in low or "наименование" in low:
            claim("name", col)
        elif "маркировка" in low:
            claim("marking", col)
        elif "идентификация" in low or "икпу" in low or "каталог" in low:
            claim("ikpu", col)
        elif "штрих" in low:
            claim("barcode", col)
        elif "ўлчов" in low or "измерен" in low:
            claim("unit", col)
        elif "миқдор" in low or "количество" in low:
            claim("qty", col)
        elif "нарҳ" in low or "нарх" in low or low.strip() == "цена":
            claim("price", col)
        elif "ққс" in low or "ндс" in low:
            claim("vat", col)
    return idx


def _find_item_table(soup):
    """Hisob-faktura jadvalini (thead+tbody) qidiradi."""
    for table in soup.find_all("table"):
        thead = table.find("thead")
        tbody = table.find("tbody")
        if thead is None or tbody is None:
            continue
        header_map = _find_header_map(thead)
        if header_map:
            return header_map, tbody
    return None, None


def map_unit(raw, warnings, name):
    base = re.sub(r"\(.*?\)", "", raw).strip().lower()
    if base in UNIT_MAP:
        return UNIT_MAP[base]
    mapped = raw.strip()
    if mapped not in UNITS:
        warnings.append(
            f"  - '{name[:60]}': o'lchov birligi '{raw}' saytning ro'yxatida yo'q, "
            f"'{mapped}' deb yozildi — qo'lda tekshiring yoki UNIT_MAP ga qo'shing"
        )
    return mapped


def parse_price(text):
    cleaned = text.replace("\xa0", "").replace(" ", "").replace(",", ".")
    try:
        value = float(cleaned)
    except ValueError:
        return None
    return format(value, ".2f")  # doim nuqta bilan, vergul bilan emas


def parse_vat(text):
    text = text.strip()
    if "без" in text.lower():
        return "Без НДС"
    m = re.search(r"[\d.,]+", text)
    if not m:
        return text
    num = m.group(0).replace(",", ".")
    try:
        return format(float(num), ".2f")
    except ValueError:
        return text


def extract_ikpu(text):
    m = re.search(r"\d{15,17}", text)
    if m:
        return m.group(0)
    return text.split("\n")[0].strip()


def parse_invoice(path):
    with open(path, encoding="utf-8", errors="ignore") as fh:
        html = fh.read()
    soup = BeautifulSoup(html, "html.parser")
    header_map, tbody = _find_item_table(soup)
    if tbody is None:
        raise RuntimeError(
            "Hisob-faktura jadvali topilmadi. Fayl kutilgan '.xls (HTML)' "
            "hisobvaraq-faktura formatida emas ko'rinadi."
        )

    idx = _resolve_field_indices(header_map)
    missing_required = [f for f in ("name", "price") if f not in idx]
    if missing_required:
        raise RuntimeError(
            "Jadval sarlavhasida quyidagi ustunlar topilmadi: "
            + ", ".join(missing_required)
            + ". Fayl tuzilishi kutilganidan farq qiladi — namuna sifatida yuboring."
        )

    def get(cells, field, default=""):
        i = idx.get(field)
        if i is None or i >= len(cells):
            return default
        return cells[i]

    items = []
    warnings = []
    for tr in tbody.find_all("tr"):
        cells = [c.get_text(strip=True) for c in tr.find_all(["td", "th"])]
        if not cells or not cells[0].strip().isdigit():
            continue  # "Жами" jami qatori yoki bo'sh qator

        name = get(cells, "name")
        if not name:
            continue
        marking_raw = get(cells, "marking")
        ikpu_raw = get(cells, "ikpu")
        barcode = get(cells, "barcode").strip()
        unit_raw = get(cells, "unit")
        price_raw = get(cells, "price")
        vat_raw = get(cells, "vat")

        items.append({
            "name": name,
            "barcode": barcode or DEFAULT_BARCODE,
            "unit": map_unit(unit_raw, warnings, name) if unit_raw else "штук",
            "price": parse_price(price_raw),
            "vat": parse_vat(vat_raw) if vat_raw else "12.00",
            "favourite": "Да",
            "marking": "Да" if marking_raw.strip() == "Маркировкаланган" else "Нет",
            "ikpu": extract_ikpu(ikpu_raw) if ikpu_raw else "",
            "classifier_code": None,
        })

    if not idx.get("unit"):
        warnings.append("  DIQQAT: bu faylda 'Ўлчов бирлиги' ustuni topilmadi, hammasiga 'штук' qo'yildi.")
    if not idx.get("vat"):
        warnings.append("  DIQQAT: bu faylda 'ҚҚС' ustuni topilmadi, hammasiga '12.00' qo'yildi.")
    if not idx.get("ikpu"):
        warnings.append("  DIQQAT: bu faylda ИКПУ ustuni topilmadi — bo'sh qoldirildi, saytga yuklashdan oldin tekshiring.")

    return items, warnings


# ---------------------------------------------------------------------------
# Excel yozish
# ---------------------------------------------------------------------------

def write_product_list(items, out_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Продукты"
    ws.append(HEADERS)
    for it in items:
        ws.append([
            it["name"], it["barcode"], it["unit"], it["price"], it["vat"],
            it["favourite"], it["marking"], it["ikpu"], it["classifier_code"],
        ])

    ref = wb.create_sheet("Справочник")
    for i, u in enumerate(UNITS, start=1):
        ref.cell(row=i, column=1, value=u)
    for i, v in enumerate(VATS, start=1):
        ref.cell(row=i, column=2, value=v)
    for i, f in enumerate(FAVOURITES, start=1):
        ref.cell(row=i, column=3, value=f)

    wb.defined_names["Units"] = openpyxl.workbook.defined_name.DefinedName(
        "Units", attr_text=f"Справочник!$A$1:$A${len(UNITS)}")
    wb.defined_names["Vats"] = openpyxl.workbook.defined_name.DefinedName(
        "Vats", attr_text=f"Справочник!$B$1:$B${len(VATS)}")
    wb.defined_names["Favourites"] = openpyxl.workbook.defined_name.DefinedName(
        "Favourites", attr_text=f"Справочник!$C$1:$C${len(FAVOURITES)}")

    dv_units = DataValidation(type="list", formula1="=Units", allow_blank=True)
    dv_vats = DataValidation(type="list", formula1="=Vats", allow_blank=True)
    dv_fav = DataValidation(type="list", formula1="=Favourites", allow_blank=True)
    ws.add_data_validation(dv_units)
    ws.add_data_validation(dv_vats)
    ws.add_data_validation(dv_fav)
    dv_units.add("C2:C1001")
    dv_vats.add("E2:E1001")
    dv_fav.add("F2:G1001")

    wb.save(out_path)


def unique_path(out_dir, base_name):
    """base_name.xlsx band bo'lsa _2, _3 ... qo'shib beradi."""
    stem, ext = os.path.splitext(base_name)
    candidate = os.path.join(out_dir, base_name)
    i = 2
    while os.path.exists(candidate):
        candidate = os.path.join(out_dir, f"{stem}_{i}{ext}")
        i += 1
    return candidate


def default_out_name(in_path):
    base = os.path.splitext(os.path.basename(in_path))[0]
    m = re.search(r"(\d+)", base)
    num = m.group(1) if m else base
    return f"product_list_from_invoice_{num}.xlsx"


# ---------------------------------------------------------------------------
# GUI
# ---------------------------------------------------------------------------

class App:
    def __init__(self, root):
        self.root = root
        root.title(f"Hisob-faktura → Product list (v{__version__})")
        root.geometry("820x600")
        root.minsize(700, 480)

        self.selected_files = []
        self.output_dir = None
        self.combine_var = tk.BooleanVar(value=True)
        self.log_queue = queue.Queue()

        self._build_ui()
        self._poll_log_queue()
        self._log(f"Dastur versiyasi: {__version__}")
        if os.environ.get("HF2PL_UPDATE_STATUS"):
            self._log(os.environ["HF2PL_UPDATE_STATUS"])

        # dastur doim ko'rinadigan, oldingi planda ochilsin — fonda yashirin
        # ishlamasligi uchun.
        root.deiconify()
        root.lift()
        root.attributes("-topmost", True)
        root.after(300, lambda: root.attributes("-topmost", False))
        root.focus_force()

    # ---------- UI qurish ----------
    def _build_ui(self):
        pad = {"padx": 10, "pady": 6}

        step1 = ttk.LabelFrame(self.root, text="1) Hisob-faktura fayllarini tanlang")
        step1.pack(fill="x", **pad)

        btn_row = ttk.Frame(step1)
        btn_row.pack(fill="x", padx=8, pady=6)
        ttk.Button(btn_row, text="Fayl(lar) tanlash...", command=self.choose_files).pack(side="left")
        ttk.Button(btn_row, text="Ro'yxatni tozalash", command=self.clear_files).pack(side="left", padx=6)
        self.files_count_label = ttk.Label(btn_row, text="0 ta fayl tanlandi")
        self.files_count_label.pack(side="left", padx=10)

        self.files_listbox = tk.Listbox(step1, height=6, selectmode="extended")
        self.files_listbox.pack(fill="x", padx=8, pady=(0, 8))

        step2 = ttk.LabelFrame(self.root, text="2) Natijalarni qayerga saqlash")
        step2.pack(fill="x", **pad)
        row2 = ttk.Frame(step2)
        row2.pack(fill="x", padx=8, pady=6)
        ttk.Button(row2, text="Papka tanlash...", command=self.choose_output_dir).pack(side="left")
        self.output_label = ttk.Label(row2, text="(tanlanmadi, fayl turgan joyga saqlanadi)")
        self.output_label.pack(side="left", padx=10)

        ttk.Checkbutton(
            step2, text="Barcha tanlangan fayllarni BITTA Excelga birlashtirish",
            variable=self.combine_var,
        ).pack(anchor="w", padx=8, pady=(0, 6))

        step3 = ttk.Frame(self.root)
        step3.pack(fill="x", **pad)
        self.start_btn = ttk.Button(step3, text="Boshlash", command=self.start_conversion)
        self.start_btn.pack(side="left")
        self.open_folder_btn = ttk.Button(step3, text="Papkani ochish", command=self.open_output_folder, state="disabled")
        self.open_folder_btn.pack(side="left", padx=6)

        log_frame = ttk.LabelFrame(self.root, text="Jarayon / natija")
        log_frame.pack(fill="both", expand=True, **pad)
        self.log_text = tk.Text(log_frame, wrap="word", state="disabled")
        scroll = ttk.Scrollbar(log_frame, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scroll.set)
        self.log_text.pack(side="left", fill="both", expand=True, padx=(8, 0), pady=8)
        scroll.pack(side="right", fill="y", pady=8, padx=(0, 8))

    # ---------- Fayl tanlash ----------
    def choose_files(self):
        paths = filedialog.askopenfilenames(
            title="Hisob-faktura fayllarini tanlang (bir nechtasini birga tanlash mumkin)",
            filetypes=[("Hisob-faktura fayllari", "*.xls *.xlsx *.html *.htm"),
                       ("Barcha fayllar", "*.*")],
        )
        for p in paths:
            if p not in self.selected_files:
                self.selected_files.append(p)
                self.files_listbox.insert("end", os.path.basename(p))
        self.files_count_label.config(text=f"{len(self.selected_files)} ta fayl tanlandi")

    def clear_files(self):
        self.selected_files = []
        self.files_listbox.delete(0, "end")
        self.files_count_label.config(text="0 ta fayl tanlandi")

    def choose_output_dir(self):
        d = filedialog.askdirectory(title="Natijalarni qayerga saqlash kerak?")
        if d:
            self.output_dir = d
            self.output_label.config(text=d)
            self._log(f"Saqlash papkasi tanlandi: {d}")
        else:
            self.output_dir = None
            self.output_label.config(text="(tanlanmadi, fayl turgan joyga saqlanadi)")

    def open_output_folder(self):
        d = self.output_dir or (os.path.dirname(self.selected_files[0]) if self.selected_files else None)
        if d and os.path.isdir(d):
            os.startfile(d)

    # ---------- Log ----------
    def _log(self, text):
        self.log_queue.put(text)

    def _poll_log_queue(self):
        try:
            while True:
                text = self.log_queue.get_nowait()
                self.log_text.configure(state="normal")
                self.log_text.insert("end", text + "\n")
                self.log_text.see("end")
                self.log_text.configure(state="disabled")
        except queue.Empty:
            pass
        self.root.after(150, self._poll_log_queue)

    # ---------- Konvertatsiya ----------
    def start_conversion(self):
        if not self.selected_files:
            messagebox.showwarning("Diqqat", "Avval hisob-faktura fayl(lar)ini tanlang.")
            return
        self.start_btn.config(state="disabled", text="Ishlanmoqda...")
        self.open_folder_btn.config(state="disabled")
        thread = threading.Thread(target=self._run_conversion, daemon=True)
        thread.start()

    def _run_conversion(self):
        combine = self.combine_var.get()
        combined_items = []
        saved_paths = []
        error_count = 0

        for path in self.selected_files:
            base = os.path.basename(path)
            self._log(f"\n--- {base} ---")
            try:
                self._log("Fayl o'qilmoqda...")
                items, warnings = parse_invoice(path)
                self._log(f"{len(items)} ta mahsulot topildi.")
                for w in warnings:
                    self._log(w)

                if combine:
                    combined_items.extend(items)
                else:
                    out_dir = self.output_dir or os.path.dirname(path) or "."
                    out_path = unique_path(out_dir, default_out_name(path))
                    write_product_list(items, out_path)
                    saved_paths.append(out_path)
                    self._log(f"Saqlandi: {out_path}")
            except Exception as exc:  # noqa: BLE001
                error_count += 1
                self._log(f"XATOLIK: {exc}")

        if combine and combined_items:
            out_dir = self.output_dir or os.path.dirname(self.selected_files[0]) or "."
            stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M")
            out_path = unique_path(out_dir, f"product_list_combined_{stamp}.xlsx")
            write_product_list(combined_items, out_path)
            saved_paths.append(out_path)
            self._log(f"\n=== Barchasi birlashtirildi: {len(combined_items)} ta mahsulot ===")
            self._log(f"Saqlandi: {out_path}")

        self._log(f"\nTayyor. {len(self.selected_files)} ta fayl ishlandi"
                   f"{f', {error_count} tasida xatolik' if error_count else ''}.")

        self.root.after(0, self._on_conversion_done, saved_paths)

    def _on_conversion_done(self, saved_paths):
        self.start_btn.config(state="normal", text="Boshlash")
        if saved_paths:
            self.open_folder_btn.config(state="normal")
        self.root.lift()
        messagebox.showinfo(
            "Tayyor",
            f"Konvertatsiya tugadi.\n\nSaqlangan fayllar soni: {len(saved_paths)}"
            + ("\n\n" + "\n".join(saved_paths[:10]) if saved_paths else ""),
        )


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
