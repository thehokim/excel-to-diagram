import os
import re
import math
import pandas as pd
import matplotlib
matplotlib.use("Agg")  # headless режим
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
from datetime import datetime, date
from tqdm import tqdm  # прогресс-бар (pip install tqdm)

# ====== НАСТРОЙКИ ======
FILE_PATH = "Exelll.xls"           # <-- укажите ваш файл
ID_COL = "ID"                                # имя столбца с ID
SKIP_COLS = {ID_COL, "PL_maydon"}            # какие столбцы НЕ являются датами/значениями
OUT_DIR = "charts"                           # папка с PNG
PDF_PATH = "all_charts.pdf"                  # общий PDF
FIGSIZE = (9, 5)
DPI = 130
# =======================

# Определяем движок
ext = os.path.splitext(FILE_PATH)[1].lower()
engine = "openpyxl" if ext == ".xlsx" else "xlrd" if ext == ".xls" else None

# Читаем Excel
df = pd.read_excel(FILE_PATH, engine=engine)

# Выявляем столбцы-дату (всё, что не в SKIP_COLS)
candidate_cols = [c for c in df.columns if c not in SKIP_COLS]

def strip_time_from_label(val) -> str:
    """
    Возвращает строковую метку ДАТЫ без времени.
    Поддерживает pandas.Timestamp, datetime/date, а также строки:
    - '2024-07-01 12:30:00' -> '2024-07-01'
    - '01.07.2024 12:30'    -> '01.07.2024'
    - '10-May 14:05'        -> '10-May'
    Если не удаётся распарсить — пробуем отрезать всё после пробела/«T».
    """
    # pandas.Timestamp / datetime / date
    if hasattr(val, "date"):
        try:
            return val.date().isoformat()
        except Exception:
            pass

    s = str(val).strip()
    # Быстрый срез по пробелу/«T» если есть явная часть времени
    if " " in s:
        s = s.split(" ", 1)[0]
    if "T" in s and len(s) > 10:
        s = s.split("T", 1)[0]

    # Попытка распарсить как дату в популярных форматах и отдать «только дату»
    known_formats = [
        "%Y-%m-%d",
        "%d.%m.%Y",
        "%m/%d/%Y",
        "%d/%m/%Y",
        "%d-%b",        # '10-May'
        "%d-%b-%Y",     # '10-May-2024'
        "%b-%d-%Y",     # 'May-10-2024'
        "%d.%b.%Y",     # '10.May.2024'
    ]
    for fmt in known_formats:
        try:
            dt = datetime.strptime(s, fmt)
            # Если в формате нет года — оставим как было (например '10-May')
            if "%Y" in fmt:
                return dt.date().isoformat()
            else:
                return dt.strftime("%d-%b")  # компактная дата без времени
        except Exception:
            continue

    return s  # вернуть как есть, если ничего не подошло

def normalize_label(col_name):
    """Нормализуем заголовок столбца (убираем время)."""
    return strip_time_from_label(col_name)

# Подписи оси X без времени
date_labels = [normalize_label(c) for c in candidate_cols]

def fix_id_value(raw):
    """
    Приводим ID к аккуратной строке:
    - float '8.0' -> '8'
    - int 8 -> '8'
    - строки очищаем: оставляем только [0-9A-Za-z_-], пробелы -> '_'
    - если пусто/NaN -> 'unknown'
    """
    if pd.isna(raw):
        return "unknown"
    # Числовые типы
    if isinstance(raw, (int,)):
        return str(raw)
    if isinstance(raw, float):
        # Если целое в виде float -> без .0
        if raw.is_integer():
            return str(int(raw))
        return str(raw).replace(".", "_")

    s = str(raw).strip()
    if s == "":
        return "unknown"
    s = s.replace(" ", "_")
    s = re.sub(r"[^0-9A-Za-z_\-]", "", s)
    return s if s else "unknown"

os.makedirs(OUT_DIR, exist_ok=True)

# PDF для всех графиков
with PdfPages(PDF_PATH) as pdf:
    for _, row in tqdm(df.iterrows(), total=len(df), desc="Building charts"):
        id_val = fix_id_value(row.get(ID_COL, "unknown"))

        # значения по датам, пропускаем NaN
        x = []
        y = []
        for col, label in zip(candidate_cols, date_labels):
            val = row[col]
            if pd.notna(val):
                x.append(label)
                try:
                    y.append(float(val))
                except Exception:
                    # если в ячейке мусор — пропускаем
                    continue

        if not y:
            continue  # нечего рисовать

        plt.figure(figsize=FIGSIZE)
        plt.plot(x, y, marker="o", label=f"ID: {id_val}")

        # подписи каждой точки в виде процентов
        for xi, yi in zip(x, y):
            plt.text(xi, yi, f"{yi:.2f}%", fontsize=8, ha="center", va="bottom")

        plt.title(f"Chiziqli Diagramma: ID {id_val}")
        plt.xlabel("Sana")   # подпись оси X (без времени в метках)
        plt.ylabel("Foiz")
        plt.grid(True, alpha=0.35)
        plt.legend(loc="best")
        plt.xticks(rotation=0)  # не наклоняем подписи (здесь они уже без времени)

        # PNG
        png_path = os.path.join(OUT_DIR, f"chart_ID_{id_val}.png")
        plt.savefig(png_path, dpi=DPI, bbox_inches="tight")

        # в общий PDF
        pdf.savefig(bbox_inches="tight")

        plt.close()

print(f"✅ PNG сохранены в: {OUT_DIR}")
print(f"✅ Многостраничный PDF: {PDF_PATH}")
