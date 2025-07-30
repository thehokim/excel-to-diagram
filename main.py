import os
import re
import math
import pandas as pd
import matplotlib
matplotlib.use("Agg")  # headless режим
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
from datetime import datetime
from tqdm import tqdm  # прогресс-бар (pip install tqdm)

# ====== НАСТРОЙКИ ======
FILE_PATH = "Exelll.xls"           # <-- укажите ваш файл
ID_COL = "ID"                                # имя столбца с ID
SKIP_COLS = {ID_COL, "PL_maydon"}            # какие столбцы НЕ являются датами/значениями
OUT_DIR = "charts"                            # папка с PNG
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

# Нормализуем заголовки дат в удобные метки (например "10-May" -> "10-May")
def normalize_label(s):
    s = str(s).strip()
    return s

date_labels = [normalize_label(c) for c in candidate_cols]

os.makedirs(OUT_DIR, exist_ok=True)

# PDF для всех графиков
with PdfPages(PDF_PATH) as pdf:
    for _, row in tqdm(df.iterrows(), total=len(df), desc="Building charts"):
        id_val = row[ID_COL]

        # значения по датам, пропускаем NaN
        x = []
        y = []
        for col, label in zip(candidate_cols, date_labels):
            val = row[col]
            if pd.notna(val):
                x.append(label)
                y.append(float(val))

        if not y:
            continue  # нечего рисовать

        plt.figure(figsize=FIGSIZE)
        plt.plot(x, y, marker="o", label=f"ID: {id_val}")

        # подписи каждой точки в виде процентов
        for xi, yi in zip(x, y):
            plt.text(xi, yi, f"{yi:.2f}%", fontsize=8, ha="center", va="bottom")

        plt.title(f"Chiziqli Diagramma: ID {id_val}")
        plt.xlabel("Sana")
        plt.ylabel("Foiz")
        plt.grid(True, alpha=0.35)
        plt.legend(loc="best")

        # PNG
        png_path = os.path.join(OUT_DIR, f"chart_ID_{id_val}.png")
        plt.savefig(png_path, dpi=DPI, bbox_inches="tight")

        # в общий PDF
        pdf.savefig(bbox_inches="tight")

        plt.close()

print(f"✅ PNG сохранены в: {OUT_DIR}")
print(f"✅ Многостраничный PDF: {PDF_PATH}")
