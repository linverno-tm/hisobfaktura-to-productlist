#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
HisobFaktura2ProductList — kichik ishga tushiruvchi (launcher).

Har safar ishga tushganda GitHub'dagi eng so'nggi hisobfaktura_core.py
faylini yuklab olishga harakat qiladi (4 soniya ichida). Muvaffaqiyatli
bo'lsa — o'sha eng yangi versiyani ishga tushiradi va keshga saqlaydi.
Internet bo'lmasa yoki server javob bermasa:
  1) avval yuklab olingan (keshdagi) nusxa ishlatiladi,
  2) u ham bo'lmasa, dastur ichiga o'ralgan zaxira nusxa ishlatiladi.
Shu tufayli dastur HECH QACHON butunlay ishlamay qolmaydi, lekin internet
bor joyda doim eng so'nggi tuzatishlar bilan ochiladi — foydalanuvchi
faylni qayta yuklab olishi shart emas.
"""
import os
import sys
import types
import tempfile
import urllib.request

RAW_URL = (
    "https://raw.githubusercontent.com/linverno-tm/"
    "hisobfaktura-to-productlist/main/hisobfaktura_core.py"
)
APP_DIR_NAME = "HisobFaktura2ProductList"
FETCH_TIMEOUT = 4  # soniya


def _cache_dir():
    base = os.environ.get("LOCALAPPDATA") or tempfile.gettempdir()
    d = os.path.join(base, APP_DIR_NAME)
    os.makedirs(d, exist_ok=True)
    return d


def _bundled_fallback_path():
    # PyInstaller --add-data bilan .exe ichiga qo'shilgan zaxira nusxa
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, "hisobfaktura_core.py")


def _load_from_url():
    with urllib.request.urlopen(RAW_URL, timeout=FETCH_TIMEOUT) as resp:
        data = resp.read()
    text = data.decode("utf-8")
    if "class App" not in text or "def main" not in text:
        raise ValueError("yuklab olingan fayl noto'g'ri ko'rinadi")
    return text


def load_core_source():
    """(source_code, holat_xabari) qaytaradi."""
    cache_path = os.path.join(_cache_dir(), "hisobfaktura_core.py")

    try:
        text = _load_from_url()
        with open(cache_path, "w", encoding="utf-8") as f:
            f.write(text)
        return text, "✓ Internetdan eng so'nggi versiya yuklab olindi."
    except Exception:
        pass

    if os.path.exists(cache_path):
        try:
            with open(cache_path, encoding="utf-8") as f:
                text = f.read()
            return text, "⚠ Internetga ulanib bo'lmadi — avval yuklab olingan versiya ishlatildi."
        except Exception:
            pass

    fallback = _bundled_fallback_path()
    with open(fallback, encoding="utf-8") as f:
        text = f.read()
    return text, "⚠ Internetga ulanib bo'lmadi — dastur ichidagi zaxira versiya ishlatildi."


def _exec_module(source):
    module = types.ModuleType("hisobfaktura_core")
    module.__file__ = "hisobfaktura_core.py"
    exec(compile(source, "hisobfaktura_core.py", "exec"), module.__dict__)
    return module


def main():
    source, status = load_core_source()
    os.environ["HF2PL_UPDATE_STATUS"] = status

    try:
        module = _exec_module(source)
    except Exception as exc:  # yuklangan kod buzuq bo'lsa — zaxiraga qaytamiz
        try:
            fallback = _bundled_fallback_path()
            with open(fallback, encoding="utf-8") as f:
                source = f.read()
            os.environ["HF2PL_UPDATE_STATUS"] = (
                f"⚠ Yangi versiyada xatolik topildi ({exc}) — zaxira versiya ishlatildi."
            )
            module = _exec_module(source)
        except Exception as exc2:
            import tkinter as tk
            from tkinter import messagebox
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror("Xatolik", f"Dasturni ishga tushirib bo'lmadi:\n{exc2}")
            return

    module.main()


if __name__ == "__main__":
    main()
