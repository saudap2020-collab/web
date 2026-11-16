"""Download the shared Excel workbook from OneDrive and write members.csv."""
import base64
import csv
import io
import json
import os
import re
import sys
import urllib.parse
import urllib.request

import openpyxl

SHARE_URL = os.environ["EXCEL_SHARE_URL"]
OUT = os.environ.get("OUT_FILE", "members.csv")

HEADERS = {
    "department": ["department", "الإدارة", "الادارة", "إدارة", "ادارة", "القسم"],
    "member": ["member", "المتدرب", "اسم المتدرب", "المتدربين", "name", "الاسم", "اسم الموظف", "الموظف", "اسم العضو", "العضو"],
    "course": ["course", "الدورة", "الدوره", "اسم الدورة", "اسم الدوره", "الدورات"],
    "progress": ["progress", "حالة الإكمال", "حالة الاكمال", "الحالة", "نسبة الإكمال", "نسبة الاكمال", "نسبة التقدم", "النسبة", "نسبه الاكمال", "نسبة الانجاز", "نسبة الإنجاز", "الإنجاز"],
}


def norm(s):
    s = re.sub(r"[ً-ْـ]", "", str(s or "")).strip().lower()
    return s.replace("أ", "ا").replace("إ", "ا").replace("آ", "ا").replace("ة", "ه").replace("ى", "ي")


UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"}


def fetch(url, headers=None):
    req = urllib.request.Request(url, headers={**UA, **(headers or {})})
    with urllib.request.urlopen(req, timeout=60) as r:
        return r.read(), r.geturl()


def badger_download(url):
    """Anonymous access used by the OneDrive web app for new-style share links."""
    req = urllib.request.Request(
        "https://api-badgerp.svc.ms/v1.0/token",
        data=b'{"appId":"5cbed6ac-a083-4e14-b191-b4ba07653de2"}',
        headers={**UA, "Content-Type": "application/json"},
        method="POST",
    )
    with urllib.request.urlopen(req, timeout=60) as r:
        tok = json.loads(r.read())["token"]
    share = "u!" + base64.urlsafe_b64encode(url.encode()).decode().rstrip("=")
    for host in ("https://my.microsoftpersonalcontent.com/_api/v2.0", "https://api.onedrive.com/v1.0"):
        try:
            data, _ = fetch(f"{host}/shares/{share}/driveitem/content", {"Authorization": f"Badger {tok}", "Prefer": "autoredeem"})
            print(f"[badger {host.split('/')[2]}] {len(data)} bytes, xlsx={data[:2] == b'PK'}")
            if data[:2] == b"PK":
                return data
        except Exception as e:  # noqa: BLE001
            print(f"[badger {host.split('/')[2]}] failed: {e}")
    return None


def candidates(url):
    token = base64.urlsafe_b64encode(url.encode()).decode().rstrip("=")
    yield "shares-api", f"https://api.onedrive.com/v1.0/shares/u!{token}/root/content"
    yield "share+download", url + ("&" if "?" in url else "?") + "download=1"
    try:
        _, final = fetch(url)
        print("redirected to:", final.split("?")[0])
        yield "final+download", final + ("&" if "?" in final else "?") + "download=1"
        q = urllib.parse.parse_qs(urllib.parse.urlparse(final).query)
        resid = (q.get("resid") or q.get("id") or [None])[0]
        redeem = (q.get("redeem") or [None])[0]
        cid = (q.get("cid") or [None])[0]
        if redeem:
            yield "redeem-api", f"https://api.onedrive.com/v1.0/shares/{redeem}/root/content"
            yield "redeem-u", f"https://api.onedrive.com/v1.0/shares/u!{redeem}/root/content"
        if resid:
            cid = cid or resid.split("!")[0]
            yield "resid-download", f"https://onedrive.live.com/download?cid={cid}&resid={urllib.parse.quote(resid)}"
            yield "my-personal", f"https://my.microsoftpersonalcontent.com/personal/{cid.lower()}/_layouts/15/download.aspx?UniqueId={urllib.parse.quote(resid.split('!')[-1])}"
    except Exception as e:  # noqa: BLE001
        print("redirect probe failed:", e)


def download(url):
    try:
        data = badger_download(url)
        if data:
            return data
    except Exception as e:  # noqa: BLE001
        print("[badger] failed:", e)
    for name, c in candidates(url):
        try:
            data, _ = fetch(c)
            ok = data[:2] == b"PK"
            print(f"[{name}] {len(data)} bytes, xlsx={ok}")
            if ok:
                return data
        except Exception as e:  # noqa: BLE001
            print(f"[{name}] failed: {e}")
    sys.exit("Could not download the workbook. Make sure the link is shared as 'Anyone with the link can view'.")


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
        pass
    n = norm(s)
    if any(w in n for w in ("لم يبدا", "لم تبدا", "غير مكتمل", "لم يكتمل")):
        return 0.0 if "لم يبدا" in n or "لم تبدا" in n else None
    if "مكتمل" in n or "منجز" in n or "تم" == n:
        return 100.0
    return None


def main():
    wb = openpyxl.load_workbook(io.BytesIO(download(SHARE_URL)), data_only=True, read_only=True)
    for ws in wb.worksheets:
        rows = list(ws.iter_rows(values_only=True))
        for h, row in enumerate(rows[:10]):
            cols = find_columns(row)
            if {"department", "member", "progress"} <= cols.keys():
                out, unknown = [], set()
                for r in rows[h + 1:]:
                    get = lambda k: r[cols[k]] if k in cols and cols[k] < len(r) else None
                    name, dept, pct = get("member"), get("department"), to_pct(get("progress"))
                    if name and dept and pct is None:
                        unknown.add(str(get("progress")))
                    if not name or not dept or pct is None:
                        continue
                    out.append([str(dept).strip(), str(name).strip(), str(get("course") or "").strip(), int(pct) if float(pct).is_integer() else pct])
                if unknown:
                    print("unrecognised progress values:", sorted(unknown)[:20])
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
    for ws in wb.worksheets:
        for i, row in enumerate(ws.iter_rows(values_only=True, max_row=4)):
            print(f"sheet '{ws.title}' row {i + 1}:", [c for c in row if c is not None][:12] if i == 0 else len([c for c in row if c is not None]), "cells")
    sys.exit("Could not find columns: الإدارة، الاسم، نسبة الإكمال")


if __name__ == "__main__":
    main()
