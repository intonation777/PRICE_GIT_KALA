#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
قیمت همه برندها را از API می‌گیرد و فایل docs/data/prices.json را می‌سازد.

- اگر دریافت یک برند خطا بدهد، قیمت قبلی همان برند حفظ می‌شود (با نشان «قدیمی»).
- تغییر قیمت نسبت به اجرای قبلی (قیمت قبلی، تاریخ تغییر) و محصولات جدید ثبت می‌شود.
- اگر هیچ برندی دریافت نشود، فایل دست‌نخورده می‌ماند و برنامه با کد خطا تمام می‌شود.

اجرا:  python update_data.py
"""
import base64
import json
import re
import sys
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timedelta, timezone
from pathlib import Path

import requests

import sgp_prices as sp

ROOT = Path(__file__).resolve().parent
DOCS = ROOT / "docs"
DATA_FILE = DOCS / "data" / "prices.json"
LOGO_DIR = DOCS / "logos"

PAGE_SIZE = 1000
MAX_PAGES = 10
KEEP_DAYS = 14  # نشان «جدید» و «تغییر قیمت» تا این‌قدر روز می‌ماند
TEHRAN = timezone(timedelta(hours=3, minutes=30))

LOGO_FALLBACK = {
    4: "https://svcpool.sgpco.com/file/SgpProduct/Brand/br-4.png",
    5: "https://svcpool.sgpco.com/file/SgpProduct/Brand/br-5.PNG",
    6: "https://svcpool.sgpco.com/file/SgpProduct/Brand/br-6.png",
    7: "https://svcpool.sgpco.com/file/SgpProduct/Brand/br-7.png",
    8: "https://svcpool.sgpco.com/file/SgpProduct/Brand/br-8.PNG",
    10: "https://svcpool.sgpco.com/file/SgpProduct/Brand/br-10.png",
    12: "https://svcpool.sgpco.com/file/SgpProduct/Brand/br-12.png",
    14: "https://svcpool.sgpco.com/file/SgpProduct/Brand/br-14.png",
}


# ----------------------------------------------------------------- دریافت
def fetch_page(bid, page):
    payload = {"brandId": bid, "page": page, "pageSize": PAGE_SIZE}
    last = None
    for attempt in range(1, 4):
        try:
            r = requests.post(sp.URL, json=payload, headers=sp.HEADERS, timeout=(10, 120))
            if r.status_code == 204:
                return None  # سرور می‌گوید برای این برند داده‌ای ندارد؛ تلاش دوباره فایده‌ای ندارد
            if r.status_code != 200:
                if 400 <= r.status_code < 500 and r.status_code != 429:
                    raise PermissionError(f"HTTP {r.status_code}: {r.text[:120]!r}")
                raise RuntimeError(f"HTTP {r.status_code}: {r.text[:120]!r}")
            if not r.text.strip():
                raise RuntimeError("پاسخ خالی")
            try:
                return r.json()
            except ValueError:
                raise RuntimeError(f"پاسخ JSON نبود: {r.text[:120]!r}")
        except PermissionError as e:  # خطای ۴xx: تکرار فایده ندارد
            raise RuntimeError(str(e))
        except Exception as e:  # noqa: BLE001
            last = e
            time.sleep(2 * attempt)
    raise RuntimeError(str(last))


def fetch_brand_all(bid):
    """همه صفحه‌های یک برند؛ صفحه بعد فقط وقتی خوانده می‌شود که صفحه قبل پر بوده باشد."""
    rows_all, imgs_all, seen = [], [], set()
    for page in range(1, MAX_PAGES + 1):
        try:
            data = fetch_page(bid, page)
        except RuntimeError:
            if page == 1:
                raise
            break
        if data is None:
            break
        rows, imgs = sp.extract_rows(data)
        new = 0
        for r, im in zip(rows, imgs):
            sig = json.dumps(r, ensure_ascii=False, sort_keys=True, default=str)
            if sig in seen:
                continue
            seen.add(sig)
            rows_all.append(r)
            imgs_all.append(im)
            new += 1
        if new == 0 or len(rows) < PAGE_SIZE * 0.9:
            break
    return rows_all, imgs_all


# ----------------------------------------------------------------- تبدیل
def to_product(r):
    extra = {
        k: v for k, v in r.items()
        if k not in sp.PREFERRED and k not in sp.SKIP_KEYS and v not in (None, "")
    }
    p = {
        "code": r.get("کد", ""),
        "size": r.get("سایز", ""),
        "carton": r.get("تعداد_در_کارتن"),
        "price": r.get("قیمتریال"),
    }
    if extra:
        p["extra"] = extra
    return p


def save_logo(bid, data_uri):
    m = re.match(r"data:image/([\w+.-]+)", data_uri)
    ext = {"jpeg": "jpg", "svg+xml": "svg"}.get(m.group(1), m.group(1)) if m else "png"
    if ext not in {"png", "jpg", "gif", "webp", "svg"}:
        return None
    LOGO_DIR.mkdir(parents=True, exist_ok=True)
    (LOGO_DIR / f"brand-{bid}.{ext}").write_bytes(base64.b64decode(data_uri.split(",", 1)[1]))
    return f"logos/brand-{bid}.{ext}"


def merge_with_old(products, old_products, today):
    """قیمت قبلی، تاریخ تغییر و نشان «جدید» را از اجرای قبل منتقل یا محاسبه می‌کند."""
    old = {(p.get("code"), p.get("size")): p for p in old_products}
    cutoff = (today - timedelta(days=KEEP_DAYS)).isoformat()
    for p in products:
        o = old.get((p["code"], p["size"]))
        if o is None:
            if old_products:
                p["added"] = today.isoformat()
            continue
        if o.get("added", "") >= cutoff and "added" in o:
            p["added"] = o["added"]
        op, np_ = o.get("price"), p.get("price")
        num = (int, float)
        if isinstance(op, num) and isinstance(np_, num) and op != np_:
            p["prev"] = op
            p["changed"] = today.isoformat()
        elif "prev" in o and o.get("changed", "") >= cutoff:
            p["prev"] = o["prev"]
            p["changed"] = o["changed"]
    return products


def dump(obj):
    """JSON با هر محصول در یک خط، تا تفاوت‌ها در گیت خوانا بماند."""
    head = {k: v for k, v in obj.items() if k != "brands"}
    parts = []
    for b in obj["brands"]:
        meta = {k: v for k, v in b.items() if k != "products"}
        prods = ",\n".join(
            json.dumps(p, ensure_ascii=False, separators=(",", ":")) for p in b["products"]
        )
        parts.append(
            json.dumps(meta, ensure_ascii=False, separators=(",", ":"))[:-1]
            + ',"products":[\n' + prods + "\n]}"
        )
    head_s = json.dumps(head, ensure_ascii=False, separators=(",", ":"))[:-1]
    return head_s + ',"brands":[\n' + ",\n".join(parts) + "\n]}\n"


def load_old():
    try:
        old = json.loads(DATA_FILE.read_text(encoding="utf-8"))
    except Exception:  # noqa: BLE001
        return {}
    return {} if old.get("sample") else old  # داده نمونه را مبنای مقایسه نگیر


# ----------------------------------------------------------------- اجرا
def main():
    now = datetime.now(TEHRAN)
    today = now.date()
    old = load_old()
    old_brands = {b["id"]: b for b in old.get("brands", [])}

    results, errors = {}, {}
    with ThreadPoolExecutor(max_workers=4) as ex:
        futs = {ex.submit(fetch_brand_all, bid): bid for bid in sp.BRANDS}
        for fut in as_completed(futs):
            bid = futs[fut]
            name = sp.BRANDS[bid][0]
            try:
                results[bid] = fut.result()
                n = len(results[bid][0])
                print(f"✓ {name}: {n} محصول" if n else f"○ {name}: سرور داده‌ای برنگرداند (۲۰۴ یا خالی)")
            except Exception as e:  # noqa: BLE001
                errors[bid] = str(e)
                print(f"✗ {name}: {e}", file=sys.stderr)

    brands_out, updated = [], 0
    for bid, (name, color) in sp.BRANDS.items():
        res = results.get(bid)
        old_b = old_brands.get(bid)
        if res and res[0]:
            rows, imgs = res
            products = merge_with_old(
                [to_product(r) for r in rows], (old_b or {}).get("products", []), today)
            logo_data, _ = sp.split_images(imgs)
            logo = None
            if logo_data:
                try:
                    logo = save_logo(bid, logo_data)
                except Exception:  # noqa: BLE001
                    logo = None
            logo = logo or (old_b or {}).get("logo") or LOGO_FALLBACK.get(bid)
            brands_out.append({
                "id": bid, "name": name, "color": "#" + color, "logo": logo,
                "updated": today.isoformat(), "stale": False,
                "count": len(products), "products": products,
            })
            updated += 1
        elif old_b:
            keep = dict(old_b)
            keep["stale"] = True
            brands_out.append(keep)
            print(f"! {name}: قیمت قبلی حفظ شد", file=sys.stderr)
        else:
            print(f"- {name}: داده‌ای نیست", file=sys.stderr)

    if updated == 0:
        print("هیچ برندی دریافت نشد؛ فایل قیمت‌ها تغییر نکرد.", file=sys.stderr)
        return 1

    out = {
        "sample": False,
        "generated_at": now.isoformat(timespec="seconds"),
        "date_gregorian": today.isoformat(),
        "date_jalali": sp.jalali_numeric(now),
        "date_jalali_text": sp.jalali_text(now),
        "time": now.strftime("%H:%M"),
        "brands": brands_out,
    }
    DATA_FILE.parent.mkdir(parents=True, exist_ok=True)
    DATA_FILE.write_text(dump(out), encoding="utf-8")
    total = sum(b["count"] for b in brands_out)
    print(f"\nانجام شد: {total} محصول در {len(brands_out)} برند -> {DATA_FILE}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
