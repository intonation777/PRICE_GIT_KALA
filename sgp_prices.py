#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
دریافت قیمت محصولات برندهای SGP و ساخت یک فایل اکسل فارسی و مرتب برای هر برند.

نصب (یک‌بار):   pip install requests openpyxl pillow
اجرا:
    python sgp_prices.py                  همه برندها (به‌صورت موازی، سریع)
    python sgp_prices.py --brands 7 5     فقط برندهای انتخابی
    python sgp_prices.py --images         عکس هر محصول هم داخل اکسل گذاشته شود (کندتر)
    python sgp_prices.py --raw            ذخیره پاسخ خام JSON برای عیب‌یابی
    python sgp_prices.py --debug          ذخیره خلاصه ساختار پاسخ هر برند

خروجی:
    prices/<تاریخ_ساعت>/<نام برند>.xlsx
    prices/history.csv   (تاریخچه همه اجراها، هر بار به انتها اضافه می‌شود)
"""
import argparse
import base64
import csv
import hashlib
import io
import json
import re
import sys
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from pathlib import Path

import requests
from openpyxl import Workbook
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.properties import PageSetupProperties

try:
    from openpyxl.drawing.image import Image as XLImage
    from PIL import Image as PILImage

    HAS_PIL = True
except ImportError:  # عکس‌ها اختیاری هستند
    HAS_PIL = False

try:
    sys.stdout.reconfigure(errors="replace")
    sys.stderr.reconfigure(errors="replace")
except Exception:  # noqa: BLE001
    pass

URL = "https://demoback.sgpco.com/api/excelGrid/GetData"
OUT_DIR = Path("prices")

# شناسه برند: (نام، رنگ سازمانی) — رنگ‌ها از خود سایت برداشته شده
BRANDS = {
    4: ("آذین لوله", "0054A6"),
    5: ("نیوپایپ", "D61A20"),
    6: ("دینا پلیمر", "51B848"),
    7: ("نیوفلکس", "29BDBE"),
    8: ("سپهرفوم", "FDAF17"),
    10: ("SGP", "D71920"),
    12: ("باتیس", "D71920"),
    14: ("مولتی پایپ", "482F92"),
}

HEADERS = {
    "accept": "application/json, text/plain, */*",
    "accept-language": "en-US,en;q=0.9",
    "content-type": "application/json",
    "origin": "https://www.gitikala.com",
    "referer": "https://www.gitikala.com/",
    "user-agent": (
        "Mozilla/5.0 (Linux; Android 16; Pixel 10) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/156.0.0.0 Mobile Safari/537.36"
    ),
}

PREFERRED = ["کد", "سایز", "تعداد_در_کارتن", "قیمتریال"]
TITLES = {
    "کد": "کد کالا",
    "سایز": "سایز",
    "تعداد_در_کارتن": "تعداد در کارتن",
    "قیمتریال": "قیمت (ریال)",
}
SKIP_KEYS = {"ردیف"}  # شماره ردیف را خودمان می‌سازیم
HINTS = {"کد", "سایز", "قیمتریال", "تعداد_در_کارتن", "code", "size", "price"}
HEADER_ROW = 4

# ----------------------------------------------------------------- تاریخ شمسی
FA_DIGITS = str.maketrans("0123456789", "۰۱۲۳۴۵۶۷۸۹")
AR_TO_EN = str.maketrans("۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩", "01234567890123456789")
MONTHS = ["فروردین", "اردیبهشت", "خرداد", "تیر", "مرداد", "شهریور",
          "مهر", "آبان", "آذر", "دی", "بهمن", "اسفند"]


def fa(s):
    return str(s).translate(FA_DIGITS)


def gregorian_to_jalali(gy, gm, gd):
    g_d_m = [0, 31, 59, 90, 120, 151, 181, 212, 243, 273, 304, 334]
    gy2 = gy + 1 if gm > 2 else gy
    days = (355666 + 365 * gy + (gy2 + 3) // 4 - (gy2 + 99) // 100
            + (gy2 + 399) // 400 + gd + g_d_m[gm - 1])
    jy = -1595 + 33 * (days // 12053)
    days %= 12053
    jy += 4 * (days // 1461)
    days %= 1461
    if days > 365:
        jy += (days - 1) // 365
        days = (days - 1) % 365
    if days < 186:
        jm, jd = 1 + days // 31, 1 + days % 31
    else:
        jm, jd = 7 + (days - 186) // 30, 1 + (days - 186) % 30
    return jy, jm, jd


def jalali_numeric(dt):
    jy, jm, jd = gregorian_to_jalali(dt.year, dt.month, dt.day)
    return f"{jy}/{jm:02d}/{jd:02d}"


def jalali_text(dt):
    jy, jm, jd = gregorian_to_jalali(dt.year, dt.month, dt.day)
    return fa(f"{jd} {MONTHS[jm - 1]} {jy}")


# ----------------------------------------------------------------- دریافت داده
def fetch_brand(brand_id, retries=3):
    payload = {"brandId": brand_id, "page": 1, "pageSize": 1000}
    session = requests.Session()
    session.headers.update(HEADERS)
    last = None
    for attempt in range(1, retries + 1):
        try:
            r = session.post(URL, json=payload, timeout=(10, 120))
            if r.status_code != 200:
                raise RuntimeError(f"کد HTTP {r.status_code}: {r.text[:150]!r}")
            if not r.text.strip():
                raise RuntimeError("پاسخ سرور خالی بود")
            try:
                return r.json()
            except ValueError:
                raise RuntimeError(f"پاسخ JSON نبود: {r.text[:150]!r}")
        except Exception as e:  # noqa: BLE001
            last = e
            time.sleep(attempt)
    raise RuntimeError(str(last))


# ----------------------------------------------------------------- پردازش پاسخ
def iter_tables(o, path=""):
    """همه لیست‌های دیکشنری (جدول‌های احتمالی) را در هر عمقی پیدا می‌کند."""
    if isinstance(o, list):
        if o and all(isinstance(x, dict) for x in o):
            yield path, o
        for i, x in enumerate(o):
            if isinstance(x, (dict, list)):
                yield from iter_tables(x, f"{path}[{i}]")
    elif isinstance(o, dict):
        for k, v in o.items():
            if isinstance(v, (dict, list)):
                yield from iter_tables(v, f"{path}.{k}")


def table_hits(rows):
    keys = set()
    for r in rows[:20]:
        keys.update(str(k).lower() if k.isascii() else k for k in r)
    return len(keys & HINTS)


def flatten(row):
    """رکورد تو در تو را تخت می‌کند؛ عکس‌های data:image جدا برمی‌گردند."""
    out, imgs = {}, {}

    def is_img(v):
        return isinstance(v, str) and v.startswith("data:image")

    def walk(v, key):
        if isinstance(v, dict):
            for k, x in v.items():
                walk(x, f"{key}.{k}" if key else str(k))
        elif isinstance(v, list):
            texts = []
            for i, x in enumerate(v):
                if is_img(x):
                    imgs[f"{key}[{i}]"] = x
                elif isinstance(x, (dict, list)):
                    walk(x, f"{key}[{i}]")
                elif x is not None:
                    texts.append(str(x))
            if texts:
                out[key] = "، ".join(texts)
        elif is_img(v):
            imgs[key] = v
        elif isinstance(v, str):
            v = ILLEGAL_CHARACTERS_RE.sub("", v).strip()
            out[key] = v[:2000] + "…" if len(v) > 2000 else v
        else:
            out[key] = v

    walk(row, "")
    return out, imgs


def is_empty_row(flat):
    known = [k for k in ("کد", "سایز", "قیمتریال") if k in flat]
    if known:
        return all(flat.get(k) in (None, "") for k in known)
    return all(v in (None, "") for v in flat.values())


def to_number(v):
    if isinstance(v, bool):
        return v
    if isinstance(v, (int, float)):
        return v
    if isinstance(v, str):
        s = v.translate(AR_TO_EN).replace(",", "").replace("٬", "").replace(" ", "").strip()
        if re.fullmatch(r"-?\d+(\.\d+)?", s):
            f = float(s)
            return int(f) if f.is_integer() else f
    return v


def extract_rows(data):
    tables = list(iter_tables(data))
    hinted = [rows for _, rows in tables if table_hits(rows) > 0]
    if hinted:
        raw_rows = [r for rows in hinted for r in rows]
    elif tables:
        raw_rows = max((rows for _, rows in tables), key=lambda rs: (len(rs), len(rs[0])))
    else:
        raw_rows = []

    rows, imgs, seen = [], [], set()
    for r in raw_rows:
        flat, im = flatten(r)
        if is_empty_row(flat):
            continue
        sig = json.dumps(flat, ensure_ascii=False, sort_keys=True, default=str)
        if sig in seen:
            continue
        seen.add(sig)
        if "قیمتریال" in flat:
            flat["قیمتریال"] = to_number(flat["قیمتریال"])
        if "تعداد_در_کارتن" in flat:
            flat["تعداد_در_کارتن"] = to_number(flat["تعداد_در_کارتن"])
        if flat.get("کد") not in (None, ""):
            flat["کد"] = str(flat["کد"])
        rows.append(flat)
        imgs.append(next(iter(im.values()), None))
    return rows, imgs


def split_images(imgs):
    """اگر همه ردیف‌ها یک عکس یکسان دارند، آن لوگوی برند است، نه عکس محصول."""
    present = [i for i in imgs if i]
    if not present:
        return None, [None] * len(imgs)
    hashes = {hashlib.md5(i.encode()).hexdigest() for i in present}
    if len(hashes) == 1:
        return present[0], [None] * len(imgs)
    return None, imgs


def describe(o, depth=0, key="", lines=None):
    """خلاصه ساختار JSON (رشته‌های بلند کوتاه می‌شوند)."""
    lines = [] if lines is None else lines
    pad = "  " * depth
    if isinstance(o, dict):
        lines.append(f"{pad}{key} dict({len(o)})")
        for k, v in list(o.items())[:25]:
            describe(v, depth + 1, k, lines)
    elif isinstance(o, list):
        lines.append(f"{pad}{key} list({len(o)})")
        for i, x in enumerate(o[:2]):
            describe(x, depth + 1, f"[{i}]", lines)
    else:
        lines.append(f"{pad}{key} = {str(o)[:70]!r}")
    return lines


# ----------------------------------------------------------------- اکسل
def _rgb(h):
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))


def darken(h, f):
    r, g, b = _rgb(h)
    return "%02X%02X%02X" % (int(r * f), int(g * f), int(b * f))


def text_color(h):
    r, g, b = _rgb(h)
    return "1F1F1F" if 0.299 * r + 0.587 * g + 0.114 * b > 170 else "FFFFFF"


def make_thumb(data_uri, max_h):
    raw = base64.b64decode(data_uri.split(",", 1)[1])
    im = PILImage.open(io.BytesIO(raw)).convert("RGBA")
    im.thumbnail((max_h * 4, max_h))
    buf = io.BytesIO()
    im.save(buf, "PNG")
    buf.seek(0)
    return XLImage(buf)


def column_keys(rows):
    keys = list(dict.fromkeys(k for r in rows for k in r))
    ordered = [k for k in PREFERRED if k in keys] + [
        k for k in keys if k not in PREFERRED and k not in SKIP_KEYS
    ]
    return [k for k in ordered if any(r.get(k) not in (None, "") for r in rows)]


def build_workbook(path, name, color, rows, logo, row_imgs, when):
    keys = column_keys(rows)
    with_imgs = HAS_PIL and any(row_imgs)

    spec = [("ردیف", None, "idx")]
    if with_imgs:
        spec.append(("تصویر", None, "img"))
    for k in keys:
        kind = "price" if k == "قیمتریال" else "int" if k == "تعداد_در_کارتن" else "text"
        spec.append((TITLES.get(k, k.replace("_", " ")), k, kind))
        if k == "قیمتریال":
            spec.append(("قیمت (تومان)", k, "toman"))
    ncols = len(spec)
    last = get_column_letter(ncols)
    n = len(rows)

    wb = Workbook()
    ws = wb.active
    ws.title = name[:31]
    ws.sheet_view.rightToLeft = True
    ws.sheet_view.showGridLines = False

    base = color
    dark = darken(color, 0.78)
    font = "Tahoma"
    side = Side(style="thin", color="D0D5DB")
    border = Border(left=side, right=side, top=side, bottom=side)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # --- عنوان
    use_logo = HAS_PIL and logo
    title_end = ncols - 1 if (use_logo and ncols > 2) else ncols
    for c in range(1, ncols + 1):
        ws.cell(row=1, column=c).fill = PatternFill("solid", fgColor=base)
        ws.cell(row=2, column=c).fill = PatternFill("solid", fgColor="EEF1F5")
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=title_end)
    t = ws.cell(row=1, column=1, value=f"لیست قیمت {name}")
    t.font = Font(name=font, size=17, bold=True, color=text_color(base))
    t.alignment = center
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=ncols)
    sub = (f"تاریخ: {jalali_text(when)}   |   {when:%Y-%m-%d}   |   "
           f"ساعت {fa(when.strftime('%H:%M'))}   |   تعداد محصولات: {fa(n)}")
    s = ws.cell(row=2, column=1, value=sub)
    s.font = Font(name=font, size=10, color="44546A")
    s.alignment = center
    ws.row_dimensions[1].height = 44
    ws.row_dimensions[2].height = 22
    ws.row_dimensions[3].height = 6

    if use_logo:
        try:
            img = make_thumb(logo, 46)
            img.anchor = f"{last}1"
            ws.add_image(img)
        except Exception:  # noqa: BLE001
            pass

    # --- سرستون
    hfill = PatternFill("solid", fgColor=dark)
    for c, (title, _, _) in enumerate(spec, start=1):
        cell = ws.cell(row=HEADER_ROW, column=c, value=title)
        cell.font = Font(name=font, size=11, bold=True, color=text_color(dark))
        cell.fill = hfill
        cell.alignment = center
        cell.border = border
    ws.row_dimensions[HEADER_ROW].height = 30

    # --- داده‌ها
    zebra = PatternFill("solid", fgColor="F5F7FA")
    col_vals = [[] for _ in spec]
    thumb_cache = {}
    for i, r in enumerate(rows, start=1):
        rr = HEADER_ROW + i
        for c, (_, key, kind) in enumerate(spec, start=1):
            if kind == "idx":
                v = i
            elif kind == "img":
                v = None
            elif kind == "toman":
                p = r.get(key)
                v = None
                if isinstance(p, (int, float)) and not isinstance(p, bool):
                    v = int(p // 10) if p % 10 == 0 else round(p / 10, 1)
            else:
                v = r.get(key)
            cell = ws.cell(row=rr, column=c, value=v)
            cell.font = Font(name=font, size=10)
            cell.alignment = center
            cell.border = border
            if kind in ("price", "toman", "int"):
                cell.number_format = "#,##0"
            if i % 2 == 0:
                cell.fill = zebra
            if v is not None:
                col_vals[c - 1].append(v)
        if with_imgs:
            ws.row_dimensions[rr].height = 40
            data_uri = row_imgs[i - 1]
            if data_uri:
                try:
                    h = hashlib.md5(data_uri.encode()).hexdigest()
                    if h not in thumb_cache:
                        thumb_cache[h] = data_uri
                    img = make_thumb(data_uri, 36)
                    img.anchor = f"B{rr}"
                    ws.add_image(img)
                except Exception:  # noqa: BLE001
                    pass
        else:
            ws.row_dimensions[rr].height = 21

    # --- عرض ستون‌ها
    for c, (title, _, kind) in enumerate(spec, start=1):
        if kind == "idx":
            w = 7
        elif kind == "img":
            w = 14
        else:
            vals = [len(f"{v:,}") if isinstance(v, (int, float)) else len(str(v))
                    for v in col_vals[c - 1][:300]]
            w = min(max([len(title)] + vals) + 5, 48)
            w = max(w, 12)
        ws.column_dimensions[get_column_letter(c)].width = w

    # --- فیلتر، فریز، چاپ
    ws.freeze_panes = f"A{HEADER_ROW + 1}"
    ws.auto_filter.ref = f"A{HEADER_ROW}:{last}{HEADER_ROW + n}"
    ws.print_title_rows = f"{HEADER_ROW}:{HEADER_ROW}"
    ws.page_setup.orientation = "portrait"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.oddFooter.center.text = "صفحه &P از &N"

    wb.save(path)


# ----------------------------------------------------------------- تاریخچه
HIST_HEADER = ["تاریخ", "تاریخ شمسی", "ساعت", "برند", "شناسه برند",
               "کد کالا", "سایز", "تعداد در کارتن", "قیمت ریال", "سایر"]


def append_history(path, name, bid, rows, when):
    new = not path.exists()
    with path.open("a", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        if new:
            w.writerow(HIST_HEADER)
        for r in rows:
            other = {k: v for k, v in r.items() if k not in PREFERRED and k not in SKIP_KEYS}
            w.writerow([
                f"{when:%Y-%m-%d}", jalali_numeric(when), f"{when:%H:%M:%S}", name, bid,
                r.get("کد", ""), r.get("سایز", ""), r.get("تعداد_در_کارتن", ""),
                r.get("قیمتریال", ""),
                json.dumps(other, ensure_ascii=False) if other else "",
            ])


# ----------------------------------------------------------------- اجرا
def process_brand(bid, args, run_dir):
    t0 = time.time()
    data = fetch_brand(bid)
    if args.raw:
        (run_dir / "raw").mkdir(exist_ok=True)
        (run_dir / "raw" / f"raw_{bid}.json").write_text(
            json.dumps(data, ensure_ascii=False), encoding="utf-8")
    rows, imgs = extract_rows(data)
    if args.debug or len(rows) < 5:
        (run_dir / "debug").mkdir(exist_ok=True)
        (run_dir / "debug" / f"structure_{bid}.txt").write_text(
            "\n".join(describe(data)), encoding="utf-8")
    logo, row_imgs = split_images(imgs)
    if not args.images:
        row_imgs = [None] * len(rows)
    return {"rows": rows, "logo": logo, "row_imgs": row_imgs, "secs": time.time() - t0}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--brands", nargs="*", type=int, help="شناسه برندها (پیش‌فرض: همه)")
    ap.add_argument("--images", action="store_true", help="عکس محصولات داخل اکسل")
    ap.add_argument("--raw", action="store_true", help="ذخیره پاسخ خام JSON")
    ap.add_argument("--debug", action="store_true", help="ذخیره خلاصه ساختار پاسخ")
    ap.add_argument("--workers", type=int, default=4, help="تعداد درخواست همزمان")
    args = ap.parse_args()

    if args.images and not HAS_PIL:
        print("برای عکس داخل اکسل لازم است:  pip install pillow", file=sys.stderr)

    ids = args.brands or list(BRANDS)
    when = datetime.now()
    run_dir = OUT_DIR / when.strftime("%Y-%m-%d_%H%M")
    run_dir.mkdir(parents=True, exist_ok=True)
    t_start = time.time()
    print(f"دریافت {len(ids)} برند ...")

    results, errors = {}, {}
    with ThreadPoolExecutor(max_workers=max(1, min(args.workers, len(ids)))) as ex:
        futs = {ex.submit(process_brand, bid, args, run_dir): bid for bid in ids}
        for fut in as_completed(futs):
            bid = futs[fut]
            name = BRANDS.get(bid, (f"برند {bid}", ""))[0]
            try:
                results[bid] = fut.result()
                print(f"  ✓ {name}: {len(results[bid]['rows'])} محصول")
            except Exception as e:  # noqa: BLE001
                errors[bid] = str(e)
                print(f"  ✗ {name}: {e}", file=sys.stderr)

    made = []
    hist = OUT_DIR / "history.csv"
    for bid in ids:
        res = results.get(bid)
        if not res or not res["rows"]:
            if res is not None:
                print(f"  - برند {bid}: داده‌ای پیدا نشد (فایل debug را ببینید)")
            continue
        name, color = BRANDS.get(bid, (f"برند {bid}", "305496"))
        safe = re.sub(r'[\\/:*?"<>|]', "_", name)
        path = run_dir / f"{safe}.xlsx"
        build_workbook(path, name, color, res["rows"], res["logo"], res["row_imgs"], when)
        append_history(hist, name, bid, res["rows"], when)
        made.append((name, len(res["rows"]), path))

    print(f"\nانجام شد در {time.time() - t_start:.1f} ثانیه")
    for name, n, path in made:
        print(f"  {name}: {n} محصول -> {path}")
    if made:
        print(f"تاریخچه: {hist}")
    if errors:
        print(f"{len(errors)} برند خطا داشت: {sorted(errors)}", file=sys.stderr)
    if not made:
        sys.exit(1)


if __name__ == "__main__":
    main()
