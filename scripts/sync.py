"""Download the shared Excel workbook from OneDrive and write members.csv."""
import base64
import csv
import io
import os
import re
import sys
import urllib.request

import openpyxl

SHARE_URL = os.environ["EXCEL_SHARE_URL"]
OUT = os.environ.get("OUT_FILE", "members.csv")

HEADERS = {
    "department": ["department", "الإدارة", "الادارة", "إدارة", "ادارة", "القسم"],
    "member": ["member", "name", "الاسم", "اسم الموظف", "الموظف", "اسم العضو", "العضو"],
    "course": ["course", "الدورة", "الدوره", "اسم الدورة", "اسم الدوره", "الدورات"],
    "progress": ["progress", "نسبة الإكمال", "نسبة الاكمال", "نسبة التقدم", "النسبة", "نسبه الاكمال", "نسبة الانجاز", "نسبة الإنجاز", "الإنجاز"],
}


def norm(s):
    s = re.sub(r"[ً-ْـ]", "", str(s or "")).strip().lower()
    return s.replace("أ", "ا").replace("إ", "ا").replace("آ", "ا").replace("ة", "ه").replace("ى", "ي")


def download(url):
    token = base64.urlsafe_b64encode(url.encode()).decode().rstrip("=")
    api = f"https://api.onedrive.com/v1.0/shares/u!{token}/root/content"
    req = urllib.request.Request(api, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, timeout=60) as r:
        return r.read()


def find_columns(row):
    cells = [norm(c) for c in row]
    found = {}
    for key, names in HEADERS.items():
        wanted = [norm(n) for n in names]
        for i, c in enumerate(cells):
            if c and any(c == w or w in c for w in wanted):
                found.setdefault(key, i)
                break
    return found


def to_pct(v):
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        return round(v * 100, 1) if 0 < v <= 1 and not float(v).is_integer() else round(float(v), 1)
    s = str(v).replace("%", "").replace("٪", "").strip()
    s = s.translate(str.maketrans("٠١٢٣٤٥٦٧٨٩", "0123456789"))
    try:
        return round(float(s), 1)
    except ValueError:
        return None


def main():
    wb = openpyxl.load_workbook(io.BytesIO(download(SHARE_URL)), data_only=True, read_only=True)
    for ws in wb.worksheets:
        rows = list(ws.iter_rows(values_only=True))
        for h, row in enumerate(rows[:10]):
            cols = find_columns(row)
            if {"department", "member", "progress"} <= cols.keys():
                out = []
                for r in rows[h + 1:]:
                    get = lambda k: r[cols[k]] if k in cols and cols[k] < len(r) else None
                    name, dept, pct = get("member"), get("department"), to_pct(get("progress"))
                    if not name or not dept or pct is None:
                        continue
                    out.append([str(dept).strip(), str(name).strip(), str(get("course") or "").strip(), int(pct) if float(pct).is_integer() else pct])
                buf = io.StringIO()
                w = csv.writer(buf, lineterminator="\n")
                w.writerow(["department", "member", "course", "progress"])
                w.writerows(out)
                new = buf.getvalue()
                old = open(OUT, encoding="utf-8").read() if os.path.exists(OUT) else ""
                if new != old:
                    open(OUT, "w", encoding="utf-8").write(new)
                    print(f"updated {OUT}: {len(out)} rows from sheet '{ws.title}'")
                else:
                    print("no changes")
                return
    sys.exit("Could not find columns: الإدارة، الاسم، نسبة الإكمال")


if __name__ == "__main__":
    main()
