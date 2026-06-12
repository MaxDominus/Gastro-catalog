# -*- coding: utf-8 -*-
"""
build.py — Збирає catalog_ready.html з вбудованими даними з Excel
Використання: python build.py
Потрібно: Excel-файл + catalog.html (шаблон) + папка photos/
Результат: catalog_ready.html  — копіюйте на телефон разом з photos/
"""

import os, json, sys
import pandas as pd
from PIL import Image

# ─── НАЛАШТУВАННЯ ──────────────────────────────────────────────
EXCEL_FILE    = "confectionery_catalog_50_positions.xlsx"
TEMPLATE_HTML = "catalog.html"
OUTPUT_HTML   = "catalog_ready.html"
PHOTOS_DIR    = "photos"
PHOTO_MAX_PX  = 900
PHOTO_QUALITY = 82
PHOTO_MAX_KB  = 200
# ───────────────────────────────────────────────────────────────

def compress_photo(path):
    if not os.path.exists(path):
        return
    size_kb = os.path.getsize(path) / 1024
    try:
        with Image.open(path) as img:
            if img.width <= PHOTO_MAX_PX and img.height <= PHOTO_MAX_PX and size_kb < PHOTO_MAX_KB:
                return
            img.thumbnail((PHOTO_MAX_PX, PHOTO_MAX_PX), Image.Resampling.LANCZOS)
            if img.mode in ("RGBA", "P", "LA"):
                img = img.convert("RGB")
            img.save(path, "JPEG", quality=PHOTO_QUALITY, optimize=True)
            new_kb = os.path.getsize(path) / 1024
            print(f"   ✓ {os.path.basename(path)}: {size_kb:.0f}KB → {new_kb:.0f}KB")
    except Exception as e:
        print(f"   ⚠ Не вдалося стиснути {path}: {e}")

def photo_path(pid):
    num = f"{pid:02d}" if pid < 100 else str(pid)
    return f"{PHOTOS_DIR}/{num}.jpg"

def build():
    print("=" * 55)
    print("  GASTRO CATALOG — збірка catalog_ready.html")
    print("=" * 55)

    for f, label in [(EXCEL_FILE, "Excel"), (TEMPLATE_HTML, "Шаблон HTML")]:
        if not os.path.exists(f):
            print(f"\n❌ Не знайдено: {f}")
            sys.exit(1)

    # 1. Читаємо Excel
    print(f"\n1. Читаємо {EXCEL_FILE}...")
    df = pd.read_excel(EXCEL_FILE, sheet_name="Продукція", skiprows=1, header=0)
    col_map = {
        df.columns[0]:"ID", df.columns[1]:"Артикул", df.columns[2]:"Назва",
        df.columns[3]:"Виробник", df.columns[4]:"Категорія", df.columns[5]:"Підкатегорія",
        df.columns[6]:"Теги", df.columns[7]:"Ціна", df.columns[8]:"Стара ціна",
        df.columns[9]:"Одиниця", df.columns[10]:"Наявність",
        df.columns[11]:"Хар1 Назва", df.columns[12]:"Хар1 Значення",
        df.columns[13]:"Хар2 Назва", df.columns[14]:"Хар2 Значення",
        df.columns[15]:"Фото", df.columns[16]:"Опис", df.columns[17]:"Примітки",
    }
    df.rename(columns=col_map, inplace=True)
    df = df[pd.notna(df["ID"]) & pd.notna(df["Назва"])].copy()
    df["ID"] = df["ID"].astype(int)
    print(f"   Знайдено позицій: {len(df)}")

    # 2. Стискаємо фото
    print(f"\n2. Оптимізація фото у '{PHOTOS_DIR}/'...")
    compressed = 0
    for _, row in df.iterrows():
        p = photo_path(int(row["ID"]))
        before = os.path.getsize(p) if os.path.exists(p) else 0
        compress_photo(p)
        after = os.path.getsize(p) if os.path.exists(p) else 0
        if after < before:
            compressed += 1
    print(f"   Оптимізовано: {compressed} фото" if compressed else "   Всі фото вже оптимальні.")

    # 3. Формуємо масив продуктів
    print(f"\n3. Формуємо дані продуктів...")
    products = []
    for _, row in df.iterrows():
        pid = int(row["ID"])
        specs = {}
        if pd.notna(row.get("Хар1 Назва")) and pd.notna(row.get("Хар1 Значення")):
            specs[str(row["Хар1 Назва"]).strip()] = str(row["Хар1 Значення"]).strip()
        if pd.notna(row.get("Хар2 Назва")) and pd.notna(row.get("Хар2 Значення")):
            specs[str(row["Хар2 Назва"]).strip()] = str(row["Хар2 Значення"]).strip()

        tags_raw = str(row.get("Теги", "") or "")
        tags = [t.strip() for t in tags_raw.split(",") if t.strip() and tags_raw != "nan"]

        price     = round(float(row["Ціна"]), 2) if pd.notna(row.get("Ціна")) else 0.0
        price_old = round(float(row["Стара ціна"]), 2) if pd.notna(row.get("Стара ціна")) else None
        stock     = int(row["Наявність"]) if pd.notna(row.get("Наявність")) else 0
        desc_val  = row.get("Опис", "")
        desc      = str(desc_val).strip() if pd.notna(desc_val) and str(desc_val) != "nan" else ""

        products.append({
            "id": pid, "article": str(row.get("Артикул","")).strip(),
            "name": str(row["Назва"]).strip(), "brand": str(row.get("Виробник","")).strip(),
            "category": str(row.get("Категорія","")).strip(), "sub": str(row.get("Підкатегорія","")).strip(),
            "tags": tags, "price": price, "price_old": price_old,
            "unit": str(row.get("Одиниця","шт")).strip(), "stock": stock,
            "specs": specs, "photo": photo_path(pid), "desc": desc,
        })

    # 4. Вшиваємо в HTML
    print(f"\n4. Вшиваємо дані в {TEMPLATE_HTML}...")
    with open(TEMPLATE_HTML, "r", encoding="utf-8") as f:
        html = f.read()

    json_data = json.dumps(products, ensure_ascii=False)
    # Замінюємо рядок завантаження fetch на вбудований масив
    old = "loadData();"
    new = f"ALL_PRODUCTS = {json_data};\ndocument.getElementById('headerCount').textContent = ALL_PRODUCTS.length + ' позицій';\nbuildFilters();\nrender();"
    if old not in html:
        print("❌ Маркер 'loadData();' не знайдено в шаблоні!")
        sys.exit(1)
    html = html.replace(old, new)

    with open(OUTPUT_HTML, "w", encoding="utf-8") as f:
        f.write(html)

    size_kb = os.path.getsize(OUTPUT_HTML) / 1024
    print(f"""
{'=' * 55}
  ✅ ГОТОВО!

  Файл: {OUTPUT_HTML} ({size_kb:.0f} KB)

  Копіюйте на телефон:
    📄 {OUTPUT_HTML}
    📁 {PHOTOS_DIR}/

  Відкривайте {OUTPUT_HTML} файловим менеджером
  або браузером — працює БЕЗ інтернету!
{'=' * 55}
""")

if __name__ == "__main__":
    build()
