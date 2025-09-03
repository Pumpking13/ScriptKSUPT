import re
import sys
import pandas as pd
import numpy as np
from pathlib import Path
import subprocess


# НАСТРОЙКИ / ПУТИ
# =====================
BASE_FOLDER = Path("/Users/mikhailsokolov/Desktop/МГТ/Рейсы")
OUTPUT_FILE = BASE_FOLDER / "ЭП" / "ЭП_итог.xlsx"
MARKS_FILE = BASE_FOLDER / "Отметки выхода июль.xlsx"
SHEET_NAME = "Выпуск и рейсы КСУПТ"
# =====================


GK_SUFFIX_RE = re.compile(r'\s*/\s*гк(?:\s*-\s*[\w\-а-яё\d]+)?', re.IGNORECASE)

def strip_gk_suffix(value: str) -> str:
    return GK_SUFFIX_RE.sub('', str(value)).strip()

def normalize_route_series(series: pd.Series) -> pd.Series:
    return (
        series.fillna("")
              .astype(str)
              .str.strip()
              .str.replace("_", "", regex=False)
              .str.lower()
              .apply(strip_gk_suffix)
              .str.replace("_", "", regex=False)
    )

def detect_vehicle_type_series(s: pd.Series) -> pd.Series:
    """
    Усиленное определение типа по тексту из 'Вид ТС'.
    Понимает: 'электробус', 'электро автобус', 'электрический автобус', 'эл автобус' и т.п.
    """
    s = (
        s.fillna("")
         .astype(str)
         .str.lower()
         .str.replace("\xa0", " ", regex=False)
         .str.replace(r"[^\w\sа-яё-]", " ", regex=True)
         .str.replace(r"\s+", " ", regex=True)
         .str.strip()
    )

    is_tram = s.str.contains(r"трамва")
    is_trol = s.str.contains(r"тролл")

    is_electric = (
        s.str.contains(r"элект\w*бус") |                 # электробус, электро...бус
        s.str.contains(r"электр\w*\s*авто?бус") |        # электрический автобус
        s.str.contains(r"\bэл\b\s*авто?бус") |           # эл автобус (после очистки "." уже нет)
        s.str.contains(r"\bэлектроавто?бус\b")           # электроавтобус слитно
    )

    is_bus = s.str.contains(r"\bавто?бус\b")             # автобус

    return np.select(
        [is_tram, is_trol, is_electric, is_bus],
        ["трамвай", "троллейбус", "электробус", "автобус"],
        default=None,
    )

def main():
    print("[INFO] Запуск SCRIPT2.py")

    for f, name in [(OUTPUT_FILE, "ЭП_итог"), (MARKS_FILE, "Отметки")]:
        if not f.exists():
            print(f"[ERROR] Не найден файл {name}: {f}")
            sys.exit(1)

    print(f"[OK] Найдены файлы:\n  - {OUTPUT_FILE}\n  - {MARKS_FILE}")

    df = pd.read_excel(MARKS_FILE)
    required_cols = {"Дата", "Маршрут", "ТП", "Вид ТС", "Территория", "Факт рейсов"}
    missing = required_cols - set(df.columns)
    if missing:
        print(f"[ERROR] Отсутствуют колонки: {missing}")
        sys.exit(1)

    print(f"[INFO] Прочитано строк: {len(df)}")

    df["Дата"] = pd.to_datetime(df["Дата"], errors="coerce", dayfirst=True).dt.strftime("%d.%m.%Y")
    df["Маршрут"] = df["Маршрут"].fillna("").astype(str).str.strip()
    df["ТП"] = df["ТП"].fillna("").astype(str).str.strip()
    df["Территория"] = df["Территория"].fillna("").astype(str).str.strip()

    df = df[(df["Дата"].notna()) & (df["Маршрут"] != "") & (df["ТП"] != "")]
    print(f"[INFO] Осталось после очистки: {len(df)} строк")

    df["Маршрут_norm"] = normalize_route_series(df["Маршрут"])
    df["Ключ 2"] = df[["Дата", "Маршрут_norm"]].agg(" ".join, axis=1).str.strip()

    df["__type"] = detect_vehicle_type_series(df["Вид ТС"])
    df = df[df["__type"].isin(["автобус", "электробус"])].copy()
    print(f"[INFO] После фильтрации по типу осталось {len(df)} строк")

    releases = pd.read_excel(OUTPUT_FILE)  
    if not {"Ключ 2", "ТипТС"}.issubset(set(releases.columns)):
        print("[WARNING] В ЭП_итог.xlsx не найдено 'Ключ 2' и 'ТипТС'. Буду использовать только распознавание по 'Вид ТС'.")
        map_df = pd.DataFrame(columns=["Ключ 2", "ТипТС"])
    else:
        map_df = releases[["Ключ 2", "ТипТС"]].dropna().copy()

    if not map_df.empty:
        counts = map_df.groupby("Ключ 2")["ТипТС"].nunique()
        ambiguous_keys = set(counts[counts > 1].index)
        if ambiguous_keys:
            print(f"[INFO] Неоднозначных ключей из Sheet1: {len(ambiguous_keys)} (для них оставляем тип по 'Вид ТС')")
        map_df = (map_df[~map_df["Ключ 2"].isin(ambiguous_keys)]
                         .drop_duplicates(subset=["Ключ 2"], keep="first"))

    df = df.merge(map_df, on="Ключ 2", how="left", suffixes=("", "_from_rel"))

    abbr_map = {"Автобус": "Авт", "Электробус": "Эл"}

    df["Авт/Эл_from_rel"] = df["ТипТС"].map(abbr_map)
    df["Авт/Эл"] = np.where(
        df["Авт/Эл_from_rel"].notna(),
        df["Авт/Эл_from_rel"],
        np.where(df["__type"].eq("электробус"), "Эл", "Авт")
    )

    df["Вид ТС"] = np.where(
        df["ТипТС"].notna(),
        df["ТипТС"],
        np.where(df["Авт/Эл"].eq("Эл"), "Электробус", "Автобус")
    )

    df["Филиал_clean"] = df["ТП"].str.replace(r"\s*\(.*?\)", "", regex=True).str.strip()
    df["Территория_clean"] = df["Территория"].str.replace(r"\s*\(.*?\)", "", regex=True).str.strip()
    df["Филиал"] = df["Филиал_clean"]
    df["Площадка"] = df["Территория_clean"]

    df["Ключ 3"] = df[["Ключ 2", "Площадка"]].agg(" ".join, axis=1).str.strip()
    df["Ключ 4"] = df[["Ключ 2", "Филиал", "Авт/Эл"]].agg(" ".join, axis=1).str.strip()
    df["Ключ 5"] = df[["Ключ 2", "Филиал", "Авт/Эл", "Площадка"]].agg(" ".join, axis=1).str.strip()

    df["Не ноль рейсов"] = np.where(df["Факт рейсов"].fillna(0) > 0, "ПРАВДА", "ЛОЖЬ")

    before = len(df)
    bad_fili = (
        df["Филиал"].isna()
        | (df["Филиал"].str.strip() == "")
        | df["Филиал"].str.fullmatch(r"(?i)nan|none|null|без\s*филиала")
    )
    df = df[~bad_fili].copy()
    allowed_fili = {"ФСЗ", "ФСВ", "ФЮ", "ФЮЗ"}
    df = df[df["Филиал"].isin(allowed_fili)].copy()
    removed = before - len(df)
    if removed > 0:
        print(f"[INFO] Удалено строк с недопустимым филиалом: {removed}")

    df = df.sort_values(["Дата", "Маршрут_norm", "Филиал", "Авт/Эл", "Площадка"]).copy()

    used_from_rel = df["ТипТС"].notna().sum()
    total = len(df)
    n_el = (df["Авт/Эл"] == "Эл").sum()
    print(f"[OK] Тип ТС подтянут из Sheet1 для {used_from_rel}/{total} строк")
    print(f"[OK] Электробусов (Эл) в результате: {n_el} из {total}")

    df = df.drop(columns=["__type", "Филиал_clean", "Территория_clean", "Маршрут_norm", "Авт/Эл_from_rel"], errors="ignore").fillna("")

    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        df.to_excel(writer, sheet_name=SHEET_NAME, index=False)

    print(f"[OK] Лист '{SHEET_NAME}' обновлён в {OUTPUT_FILE}")
    print("[DONE] SCRIPT2.py завершил работу ✅")

if __name__ == "__main__":
    main()
    answer = input("Запустить script3.py для создания ЭкспПоказ? (yes/no): ").strip().lower()
    if answer == "yes":
        try:
            subprocess.run([sys.executable, str(BASE_FOLDER / "script3.py")], check=True)
        except Exception as e:
            print(f"[ERROR] Не удалось запустить script3.py: {e}")
    else:
        print("[INFO] script3.py не запущен")
