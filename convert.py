# convert.py — читает СТАРЫЙ .xls через xlrd
import xlrd, json, re, io

XLS = "4_uif.xls"
GROUPS = ["3301", "3302"]
PAIRS = ["1", "2", "3", "4", "5"]

DAY_MAP = {
    "ПОНЕДЕЛЬНИК": "Понедельник",
    "ВТОРНИК":     "Вторник",
    "СРЕДА":       "Среда",
    "ЧЕТВЕРГ":     "Четверг",
    "ПЯТНИЦА":     "Пятница",
    "СУББОТА":     "Суббота",
}

def clean(v):
    if v is None: return ""
    # Числовые 0 (int/float) — пусто
    if isinstance(v, (int, float)) and float(v) == 0:
        return ""
    s = str(v).strip()
    if s in ("", "0", "0.0", "0,0"):
        return ""
    s = re.sub(r'<br\s*/?>', ' ', s)
    s = re.sub(r'\s+', ' ', s).strip()
    return "" if s in ("0", "0.0") else s

def norm_num(v):
    """'3301.0' → '3301'"""
    s = str(v).strip()
    if s.endswith(".0"):
        s = s[:-2]
    return s

def parse_date_cell(val, datemode):
    """Распознаёт дату Excel: число 46265, '46265', '46265.0',
       ISO '2026-08-31' или '31.08.2026'."""
    if val is None or val == "":
        return None
    # Число Excel
    try:
        num = float(str(val).strip())
        dt = xlrd.xldate_as_datetime(num, datemode)
        if 2020 <= dt.year <= 2100:
            return dt.strftime("%Y-%m-%d")
    except (ValueError, TypeError):
        pass
    # ISO-строка
    m = re.match(r'(\d{4})-(\d{2})-(\d{2})', str(val))
    if m:
        return m.group(0)
    # ДД.ММ.ГГГГ
    m = re.match(r'(\d{2})\.(\d{2})\.(\d{4})', str(val))
    if m:
        return f"{m.group(3)}-{m.group(2)}-{m.group(1)}"
    return None

wb = xlrd.open_workbook(XLS)
ws = wb.sheet_by_index(0)
rows = [[ws.cell_value(r, c) for c in range(ws.ncols)] for r in range(ws.nrows)]

data = {g: {} for g in GROUPS}
current_day, current_dates, current_pair = None, [], None

for row in rows:
    # День недели
    if row[0]:
        cell = str(row[0]).replace(" ", "").strip().upper()
        if cell in DAY_MAP:
            current_day = DAY_MAP[cell]
            current_pair = None

    # Строка "Числа месяца"
    if row[0] and "Числа" in str(row[0]):
        current_dates = []
        for j in range(4, min(21, len(row))):
            iso = parse_date_cell(row[j], wb.datemode)
            if iso:
                current_dates.append(iso)
        continue

    # Строки с парами
    group_val = norm_num(row[2]) if row[2] else ""
    if group_val in GROUPS:
        g = group_val
        pv = norm_num(row[1]) if row[1] else ""
        if pv in PAIRS:
            current_pair = pv
        if current_pair and current_day and current_dates:
            for j, iso in enumerate(current_dates):
                idx = 4 + j
                txt = clean(row[idx] if idx < len(row) else None)
                if not txt: continue
                if iso not in data[g]:
                    y, m_, d = iso.split("-")
                    data[g][iso] = {
                        "date": f"{d}.{m_}.{y}",
                        "iso": iso,
                        "day": current_day,
                        "lessons": {p: "" for p in PAIRS},
                    }
                data[g][iso]["lessons"][current_pair] = txt

out = {g: sorted(data[g].values(), key=lambda x: x["iso"]) for g in GROUPS}

with io.open("new_data.js", "w", encoding="utf-8") as f:
    f.write("const DATA = " + json.dumps(out, ensure_ascii=False) + ";\n")
    f.write("const GROUPS = " + json.dumps(GROUPS, ensure_ascii=False) + ";\n")

print("OK → new_data.js")
print("Записей:", {g: len(v) for g, v in out.items()})
# Показать первую запись для проверки
for g in GROUPS:
    if out[g]:
        first = out[g][0]
        print(f"Первая запись {g}: {first['date']} ({first['day']}), пар: "
              f"{sum(1 for v in first['lessons'].values() if v)}")