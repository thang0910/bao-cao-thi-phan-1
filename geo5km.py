# -*- coding: utf-8 -*-
"""
geo5km.py
Phan tich thi phan theo ban kinh 5km: lay moi cua hang CellphoneS (CPS) lam tam,
dem so cua hang MWG (TGDD + Dien May Xanh) trong vong 5km, so sanh doanh so
trung binh/cua hang.

Nguon du lieu:
  - Vi tri cua hang: data/postgres_public_stores_cps.csv, .._mwg.csv
    (cot: ten, dia chi, "POINT (lng lat)", ma_cua_hang)
  - Doanh so: lay tu all_df da parse (cot R_DoanhSo = CPS, Y_DoanhSo = MWG)
    o cap shop (TenShop). CPS join truc tiep theo ma shop;
    MWG dung trung binh cap tinh lam proxy (ma shop doanh so khac he voi file vi tri).
"""
import csv, re, math
from pathlib import Path

R_EARTH_KM = 6371.0
RADIUS_KM = 5.0


def _parse_point(p):
    m = re.search(r"POINT \(([-\d.]+) ([-\d.]+)\)", str(p))
    if not m:
        return None, None
    return float(m.group(1)), float(m.group(2))  # lng, lat


def _load_locs(path):
    rows = []
    with open(path, newline="", encoding="utf-8") as f:
        for r in csv.reader(f):
            if len(r) < 4:
                continue
            name, addr, point, code = r[0], r[1], r[-2], r[-1]
            lng, lat = _parse_point(point)
            if lat is None:
                continue
            rows.append({"name": name, "addr": addr, "lng": lng, "lat": lat, "code": code})
    return rows


def _find_csv(data_dirs, needle):
    for d in data_dirs:
        d = Path(d)
        if not d.exists():
            continue
        hits = sorted(d.glob(f"*stores_{needle}*.csv"))
        if hits:
            return hits[0]
    return None


def _cps_prov(code):
    m = re.match(r"CPS-([A-Z]{3})-", str(code))
    return m.group(1) if m else None


def build_geo5km(all_df, n_months, data_dirs):
    """Tra ve dict payload cho section 5km, hoac None neu thieu file vi tri."""
    cps_csv = _find_csv(data_dirs, "cps")
    mwg_csv = _find_csv(data_dirs, "mwg")
    if not cps_csv or not mwg_csv:
        print("  [geo5km] Khong tim thay file vi tri stores_cps/stores_mwg trong data/ -> bo qua section 5km")
        return None
    n_months = max(int(n_months or 1), 1)

    cps = _load_locs(cps_csv)
    mwg = _load_locs(mwg_csv)
    for m in mwg:
        m["brand"] = "TGDD" if "Thế Giới Di Động" in m["name"] else ("DMX" if "Điện Máy Xanh" in m["name"] else "OTHER")

    # ----- Doanh so tu all_df (cap shop) -----
    df = all_df
    shop = df[(df["TenShop"].notna()) & (df["TenShop"].astype(str) != "Total") &
              (df["NganhHang"].astype(str) == "Total")].copy()
    shop["TenShop"] = shop["TenShop"].astype(str)
    cps_rev = {}
    yprov = {}
    for s, R, Y in zip(shop["TenShop"], shop["R_DoanhSo"].fillna(0), shop["Y_DoanhSo"].fillna(0)):
        R = float(R or 0); Y = float(Y or 0)
        if s.startswith("CPS-") and R > 0:
            cps_rev[s] = cps_rev.get(s, 0.0) + R
        elif Y > 0:
            mp = re.match(r"([A-Z]{3})_", s)
            if mp:
                pv = mp.group(1)
                d = yprov.setdefault(pv, {"rev": 0.0, "shops": set()})
                d["rev"] += Y; d["shops"].add(s)
    prov_mwg_avg = {pv: (d["rev"] / n_months) / len(d["shops"]) for pv, d in yprov.items() if d["shops"]}

    # ----- Haversine 5km -----
    mlat = [math.radians(m["lat"]) for m in mwg]
    mlng = [math.radians(m["lng"]) for m in mwg]
    mbr = [m["brand"] for m in mwg]
    stores = []
    for s in cps:
        la, lo = math.radians(s["lat"]), math.radians(s["lng"])
        n_all = n_t = n_d = 0
        for i in range(len(mwg)):
            dlat = mlat[i] - la; dlng = mlng[i] - lo
            h = math.sin(dlat / 2) ** 2 + math.cos(la) * math.cos(mlat[i]) * math.sin(dlng / 2) ** 2
            if 2 * R_EARTH_KM * math.asin(math.sqrt(h)) <= RADIUS_KM:
                n_all += 1
                if mbr[i] == "TGDD": n_t += 1
                elif mbr[i] == "DMX": n_d += 1
        pv = _cps_prov(s["code"])
        cr = cps_rev.get(s["code"]); cr = (cr / n_months) if cr else None
        ma = prov_mwg_avg.get(pv)
        dev = (cr - ma) if (cr is not None and ma is not None) else None
        dp = round(dev / ma * 100, 1) if (dev is not None and ma) else None
        stores.append({
            "n": s["name"].split(" - ", 1)[-1], "a": s["addr"], "pv": pv,
            "lat": round(s["lat"], 6), "lng": round(s["lng"], 6),
            "m": n_all, "t": n_t, "d": n_d,
            "cr": round(cr) if cr is not None else None,
            "ma": round(ma) if ma is not None else None,
            "dp": dp,
        })

    mwg_pts = [[round(m["lat"], 5), round(m["lng"], 5), 0 if m["brand"] == "TGDD" else 1,
                m["name"].split(" - ", 1)[-1]] for m in mwg if m["brand"] in ("TGDD", "DMX")]

    valid = [s for s in stores if s["cr"] is not None and s["ma"] is not None]
    n = len(stores)
    def avg(key):
        return round(sum(s[key] for s in stores) / n, 1) if n else 0
    summary = {
        "n_cps": n,
        "n_cps_with_rev": sum(1 for s in stores if s["cr"] is not None),
        "avg_mwg": avg("m"), "avg_tgdd": avg("t"), "avg_dmx": avg("d"),
        "max_mwg": max((s["m"] for s in stores), default=0),
        "no_mwg": sum(1 for s in stores if s["m"] == 0),
        "mean_cps_rev": round(sum(s["cr"] for s in valid) / len(valid)) if valid else None,
        "mean_mwg_avg": round(sum(s["ma"] for s in valid) / len(valid)) if valid else None,
        "mean_dev_pct": round(sum(s["dp"] for s in valid) / len(valid), 1) if valid else None,
        "n_months": n_months,
    }
    print(f"  [geo5km] {n} CPS / {len(mwg_pts)} MWG pts | TB {summary['avg_mwg']} MWG trong 5km | "
          f"khop doanh so {summary['n_cps_with_rev']}/{n}")
    return {"summary": summary, "stores": stores, "mwg": mwg_pts}
