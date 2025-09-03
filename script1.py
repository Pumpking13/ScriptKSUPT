import re
import sys
from pathlib import Path
from typing import Optional, Tuple, List
import numpy as np
import pandas as pd



# НАСТРОЙКИ / ПУТИ
# =====================
SOURCE_FOLDER = Path("/Users/mikhailsokolov/Desktop/МГТ/Рейсы")
OUTPUT_FOLDER = SOURCE_FOLDER / "ЭП"
OUTPUT_FILE = OUTPUT_FOLDER / "ЭП_итог.xlsx"
# =====================


TRANSPORT_MARKERS = {
    "Автобус": ["автобус"],
    "Трамвай": ["трамвай"],
    "Троллейбус": ["троллейбус"],
    "Электробус": ["электр"],
}

ALLOWED_TTYPES = {"Автобус", "Электробус"}

TRANSPORT_ABBR = {
    "Автобус": "Авт",
    "Электробус": "Эл",
}

ALLOWED_BRANCHES = {"ЮЗ", "СВ", "СЗ", "Ю"}

GK_SUFFIX_RE = re.compile(r'\s*/\s*гк(?:\s*-\s*[\w\-а-яё\d]+)?', re.IGNORECASE)

def has_gk_suffix(value) -> bool:
    return bool(GK_SUFFIX_RE.search(str(value)))

def strip_gk_suffix(value: str) -> str:
    return GK_SUFFIX_RE.sub('', str(value)).strip()

def read_excel_auto(file_path: Path) -> pd.DataFrame:
    ext = file_path.suffix.lower()
    if ext in (".xlsx", ".xlsm"):
        return pd.read_excel(file_path, header=None, engine="openpyxl")
    elif ext == ".xls":
        return pd.read_excel(file_path, header=None, engine="xlrd")
    raise RuntimeError(f"Unsupported format: {ext}")

def normalize_branch(name: Optional[str]) -> str:
    if not name or str(name).strip().lower() in ("nan", "none"):
        return ""
    norm_name = str(name).strip().lower()

    patterns = [
        (r'\bфсв\b', None, "ФСВ"),
        ('юго', 'зап', "ЮЗ"),
        ('юго', 'вост', "ЮВ"),
        ('север', 'вост', "СВ"),
        ('север', 'зап', "СЗ"),
        ('южн', None, "Ю"),
    ]
    for p1, p2, code in patterns:
        if p2:
            if p1 in norm_name and p2 in norm_name:
                return code
        else:
            if re.search(p1, norm_name):
                return code

    tokens = re.findall(r'[а-яё]+', norm_name)
    abbrev_map = {"юз": "ЮЗ", "юв": "ЮВ", "св": "СВ", "сз": "СЗ", "ю": "Ю"}
    for token, code in abbrev_map.items():
        if token in tokens:
            return code

    match = re.match(r'^([^\(]+)', norm_name)
    return match.group(1).strip() if match else norm_name

def to_number(value) -> float:
    try:
        if value is None or (isinstance(value, float) and pd.isna(value)):
            return 0.0
        s = str(value).replace('\xa0', '').replace(' ', '').replace(',', '.')
        s = re.sub(r'[^0-9\.\-]', '', s)
        return float(s) if s not in ("", "-", ".") else 0.0
    except:
        return 0.0

def is_route_cell_valid(value) -> bool:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return False
    s = str(value).strip()
    if s == "":
        return False
    sl = s.lower()
    return not (re.search(r'№\s*м-?та', sl)
                or 'маршрут' in sl
                or 'справка' in sl
                or 'план' in sl
                or 'факт' in sl
                or 'итого' in sl
                or 'всего' in sl)

def row_text(df: pd.DataFrame, i: int) -> str:
    return " ".join(str(x) for x in df.iloc[i].tolist() if pd.notna(x)).lower()

def find_transport_headers(df: pd.DataFrame) -> List[Tuple[str, int]]:
    headers = []
    for idx in range(len(df)):
        txt = row_text(df, idx)
        for ttype, keywords in TRANSPORT_MARKERS.items():
            if any(k in txt for k in keywords):
                headers.append((ttype, idx))
                break
    return headers

def detect_branch_in_row(row: pd.Series) -> Optional[str]:
    val = row.get(1, None)
    candidates = []
    if pd.notna(val):
        candidates.append(str(val))
    s = " ".join(str(x) for x in row.tolist() if pd.notna(x))
    candidates.append(s)

    for cand in candidates:
        code = normalize_branch(cand)
        if code in ALLOWED_BRANCHES:
            return code
    return np.nan

def extract_date_from_filename(filename: str) -> str:
    for pattern in [r'(\d{2}\.\d{2}\.\d{4})', r'(\d{2}-\d{2}-\d{4})']:
        match = re.search(pattern, filename)
        if match:
            date_str = match.group(1).replace('-', '.')
            try:
                return pd.to_datetime(date_str, dayfirst=True, errors="coerce").strftime("%d.%m.%Y")
            except Exception:
                return date_str
    return ""

def process_release_file(file_path: Path) -> pd.DataFrame:
    try:
        table = read_excel_auto(file_path)
    except Exception as e:
        print(f"[WARNING] Error reading {file_path.name}: {e}")
        return pd.DataFrame()

    date = extract_date_from_filename(file_path.name)
    headers = find_transport_headers(table)
    if not headers:
        return pd.DataFrame()

    ranges = []
    for i, (ttype, idx) in enumerate(headers):
        end = (headers[i + 1][1] - 1) if i + 1 < len(headers) else len(table) - 1
        ranges.append((ttype, idx + 1, end))

    all_rows = []

    for ttype, start, end in ranges:

        if ttype not in ALLOWED_TTYPES:
            continue

        df_block = table.iloc[start:end + 1, :].copy()
        df_block = df_block.rename(columns={
            1: "Филиал_raw",
            2: "Маршрут_raw",
            3: "ПланВыпуск",
            8: "ФактВыпуск",
            13: "ПланРейсы",
            14: "ФактРейсы",
            15: "Потери"
        })

        branch_series = df_block.apply(detect_branch_in_row, axis=1)
        df_block["Филиал"] = branch_series.ffill()

        mask_route = df_block["Маршрут_raw"].apply(is_route_cell_valid) if "Маршрут_raw" in df_block else pd.Series(False, index=df_block.index)
        mask_branch = df_block["Филиал"].isin(ALLOWED_BRANCHES)
        df_block = df_block[mask_route & mask_branch].copy()

        df_block["Маршрут"] = (
            df_block["Маршрут_raw"].astype(str)
            .str.strip()
            .str.replace("_", "", regex=False)
            .str.lower()
        )

        df_block["КТР"] = df_block["Маршрут"].apply(lambda s: "КТР" if has_gk_suffix(s) else "не КТР")
        df_block["Маршрут"] = df_block["Маршрут"].apply(strip_gk_suffix).str.replace("_", "", regex=False)

        for col in ["ПланРейсы", "ФактРейсы", "Потери", "ПланВыпуск", "ФактВыпуск"]:
            if col in df_block:
                df_block[col] = df_block[col].apply(to_number)
            else:
                df_block[col] = 0.0

        df_block["Дата"] = date
        df_block["Имя_файла"] = file_path.name
        df_block["ТипТС"] = ttype

        keep_cols = ["Дата", "Маршрут", "Филиал", "ТипТС", "КТР",
                     "ПланВыпуск", "ФактВыпуск", "ПланРейсы", "ФактРейсы", "Потери"]
        df_block = df_block[[c for c in keep_cols if c in df_block.columns]]
        all_rows.append(df_block)

    if not all_rows:
        return pd.DataFrame()

    df = pd.concat(all_rows, ignore_index=True)

    if df["Филиал"].isna().any() or (df["Филиал"] == "").any():
        known_map = (df[df["Филиал"].isin(ALLOWED_BRANCHES)]
                     .groupby(["Дата", "Маршрут"])["Филиал"]
                     .agg(lambda s: s.mode().iat[0] if not s.mode().empty else s.iloc[0])
                     .to_dict())
        missing_mask = df["Филиал"].isna() | (df["Филиал"] == "")
        if missing_mask.any():
            df.loc[missing_mask, "Филиал"] = df.loc[missing_mask].apply(
                lambda r: known_map.get((r["Дата"], r["Маршрут"]), r["Филиал"]), axis=1
            )

    return df

def main():
    if not SOURCE_FOLDER.exists():
        print("[ERROR] No source folder:", SOURCE_FOLDER)
        sys.exit(1)
    OUTPUT_FOLDER.mkdir(parents=True, exist_ok=True)

    release_files = sorted(SOURCE_FOLDER.glob("Выпуск*.xls*"))
    if not release_files:
        print("[ERROR] No 'Выпуск*.xls*' files found in folder:", SOURCE_FOLDER)
        sys.exit(0)

    frames = []
    for file in release_files:
        df_part = process_release_file(file)
        if not df_part.empty:
            frames.append(df_part)

    if not frames:
        print("[ERROR] No data extracted from release files")
        sys.exit(1)

    df_releases = pd.concat(frames, ignore_index=True)

  

    df_releases = df_releases.rename(columns={"Маршрут": "№\nм-та"})
    df_releases["Ключ 2"] = df_releases["Дата"].astype(str) + " " + df_releases["№\nм-та"].astype(str)
    df_releases["Ключ 4"] = (
        df_releases["Дата"].astype(str) + " " +
        df_releases["№\nм-та"].astype(str) + " " +
        "Ф" + df_releases["Филиал"].astype(str) + " " +
        df_releases["ТипТС"].map(TRANSPORT_ABBR).fillna(df_releases["ТипТС"])
    )

    final_cols = [
        "Дата", "№\nм-та", "Филиал", "ТипТС", "КТР",
        "ПланВыпуск", "ФактВыпуск", "ПланРейсы", "ФактРейсы", "Потери",
        "Ключ 2", "Ключ 4"
    ]
    df_releases = df_releases[[c for c in final_cols if c in df_releases.columns]]

    df_releases.to_excel(OUTPUT_FILE, index=False)
    print("[OK] Result saved:", OUTPUT_FILE)

if __name__ == "__main__":
    main()
    answer = input("Запустить второй скрипт для КСУПТ? (yes/no): ").strip().lower()
    if answer == "yes":
        import sys
        import subprocess
        try:
            subprocess.run([sys.executable, "script2.py"], check=True)
        except Exception as e:
            print(f"[ERROR] Не удалось запустить script2.py: {e}")
