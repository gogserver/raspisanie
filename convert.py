# convert.py — с эвристикой: общие пары копируются обеим группам
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
    if isinstance(v, (int, float)) and float(v) == 0:
        return ""
    s = str(v).strip()
    if s in ("", "0", "0.0", "0,0"):
        return ""
    s = re.sub(r'<br\s*/?>', ' ', s)
    s = re.sub(r'\s+', ' ', s).strip()
    return "" if s in ("0", "0.0") else s

def norm_num(v):
    s = str(v).strip()
    return s[:-2] if s.endswith(".0") else s

def parse_date_cell(val, datemode):
    if val is None or val == "": return None
    try:
        num = float(str(val).strip())
        dt = xlrd.xldate_as_datetime(num, datemode)
        if 2020 <= dt.year <= 2100:
            return dt.strftime("%Y-%m-%d")
    except (ValueError, TypeError):
        pass
    m = re.match(r'(\d{4})-(\d{2})-(\d{2})', str(val))
    if m: return m.group(0)
    m = re.match(r'(\d{2})\.(\d{2})\.(\d{4})', str(val))
    if m: return f"{m.group(3)}-{m.group(2)}-{m.group(1)}"
    return None

wb = xlrd.open_workbook(XLS)
ws = wb.sheet_by_index(0)
nrows, ncols = ws.nrows, ws.ncols

def cell(r, c):
    return ws.cell_value(r, c) if r < nrows and c < ncols else ""

data = {g: {} for g in GROUPS}
current_day, current_dates, current_pair = None, [], None

for r in range(nrows):
    v0 = cell(r, 0)
    if v0:
        t = str(v0).replace(" ", "").strip().upper()
        if t in DAY_MAP:
            current_day = DAY_MAP[t]
            current_pair = None

    if v0 and "Числа" in str(v0):
        current_dates = []
        for c in range(4, ncols):
            iso = parse_date_cell(cell(r, c), wb.datemode)
            if iso:
                current_dates.append((c, iso))
        continue

    g = norm_num(cell(r, 2)) if cell(r, 2) else ""
    if g in GROUPS:
        pv = norm_num(cell(r, 1)) if cell(r, 1) else ""
        row_pair = None
        if pv in PAIRS:
            row_pair = pv
            current_pair = pv
        else:
            TIME_TO_PAIR = {"8.30":"1","10.00":"2","11.30":"3","14.20":"4","15.50":"5","17.20":"6"}
            for prefix, num in TIME_TO_PAIR.items():
                if pv.startswith(prefix):
                    row_pair = num
                    break

        if row_pair and current_day and current_dates:
            for col, iso in current_dates:
                txt = clean(cell(r, col))
                if not txt:
                    continue
                if iso not in data[g]:
                    y, m_, d = iso.split("-")
                    data[g][iso] = {
                        "date": f"{d}.{m_}.{y}", "iso": iso, "day": current_day,
                        "lessons": {p: "" for p in PAIRS},
                    }
                data[g][iso]["lessons"][row_pair] = txt

# === ЭВРИСТИКА ОБЩИХ ПАР ===
# Для каждой даты и пары: если у одной группы есть, а у другой пусто → копируем
for iso in set().union(*[set(data[g].keys()) for g in GROUPS]):
    for p in PAIRS:
        vals = {g: data[g].get(iso, {}).get("lessons", {}).get(p, "") for g in GROUPS}
        filled = {g: v for g, v in vals.items() if v}
        if len(filled) == 1:
            src_g, src_v = next(iter(filled.items()))
            for dst_g in GROUPS:
                if dst_g == src_g: continue
                if iso not in data[dst_g]:
                    y, m_, d = iso.split("-")
                    data[dst_g][iso] = {
                        "date": f"{d}.{m_}.{y}", "iso": iso,
                        "day": data[src_g][iso]["day"],
                        "lessons": {pp: "" for pp in PAIRS},
                    }
                data[dst_g][iso]["lessons"][p] = src_v

out = {g: sorted(data[g].values(), key=lambda x: x["iso"]) for g in GROUPS}

with io.open("new_data.js", "w", encoding="utf-8") as f:
    f.write("const DATA = " + json.dumps(out, ensure_ascii=False) + ";\n")
    f.write("const GROUPS = " + json.dumps(GROUPS, ensure_ascii=False) + ";\n")

print("OK → new_data.js")
print("Дней:", {g: len(v) for g, v in out.items()})
print("Пар всего:", {g: sum(1 for e in v for pp in e['lessons'] if e['lessons'][pp]) for g, v in out.items()})

# Проверка 17.09
for g in GROUPS:
    for e in out[g]:
        if e['iso'] == '2026-09-17':
            print(f"\n{g} {e['date']} ({e['day']}):")
            for pp in PAIRS:
                print(f"  пара {pp}: {e['lessons'].get(pp, '')[:70]}")