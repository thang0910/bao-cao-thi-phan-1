#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
build_report.py
Script tong hop du lieu thi truong R (CellphoneS) vs Y (MWG)
Ho tro nhieu nganh hang (LOA, TIVI, PHU KIEN, DIEN THOAI...).
Moi file Excel dat trong folder data/ duoc tu nhan dien nganh hang
va gop lai thanh 1 bao cao HTML co bo loc dong.
"""
import json
import sys
import warnings
from datetime import datetime
from pathlib import Path

import io
import pandas as pd

from product_normalize import normalize_canonical, load_manual_mapping
import os, io, urllib.request
from vn_region import lookup_mien
import geo5km

warnings.filterwarnings("ignore")

SCRIPT_DIR = Path(__file__).resolve().parent
DATA_DIRS = [SCRIPT_DIR / "data", SCRIPT_DIR]
OUTPUT_HTML = SCRIPT_DIR / "index.html"
TEMPLATE_HTML = SCRIPT_DIR / "template.html"
MAPPING_FILE = SCRIPT_DIR / "product_mapping.xlsx"
REVIEW_FILE = SCRIPT_DIR / "product_review.xlsx"

# === Google Sheet dinh danh ten san pham (published CSV) ===
# Link "Publish to web" cua sheet Mapping. Co the override bang bien moi truong GSHEET_CSV_URL.
GSHEET_CSV_URL = os.environ.get(
    "GSHEET_CSV_URL",
    "https://docs.google.com/spreadsheets/d/e/2PACX-1vTzPo4gzB6LvyseHT7V9JpgzO8OGg92v7jzeGLMq6gQ46TNOcBrR2c9uuZs8VCWC5snZr0k9jnbXElg/pub?gid=513085925&single=true&output=csv",
)


def load_gsheet_mapping():
    """Doc mapping {original: ten_chung} tu Google Sheet (CSV publish-to-web)."""
    if not GSHEET_CSV_URL:
        return {}
    try:
        req = urllib.request.Request(GSHEET_CSV_URL, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=30) as r:
            raw = r.read()
        df = pd.read_csv(io.BytesIO(raw))
    except Exception as e:
        print(f"  WARN: Khong tai duoc Google Sheet CSV: {e}")
        return {}
    df.columns = [str(c).strip().lower() for c in df.columns]
    if "original" not in df.columns or "ten_chung" not in df.columns:
        print("  WARN: CSV thieu cot 'original' hoac 'ten_chung'.")
        return {}
    m = {}
    for _, row in df.iterrows():
        o = str(row["original"]).strip()
        c = str(row["ten_chung"]).strip()
        if o and c and o.lower() != "nan" and c.lower() != "nan":
            m[o] = c
    print(f"  Doc {len(m)} mapping tu Google Sheet (CSV)")
    return m

# Format day du (16 cot, co Mien va Quan Huyen)
DIMENSION_COLS_FULL = [
    "Nam_Thang", "Nam_Tuan", "Date", "Mien", "TinhThanh", "QuanHuyen", "TenShop",
    "NganhHang", "ThuongHieu", "Model", "TenSanPham", "HinhThucXuat",
]
# Format goi (14 cot, KHONG co Mien va Quan Huyen - tu suy ra)
DIMENSION_COLS_COMPACT = [
    "Nam_Thang", "Nam_Tuan", "Date", "TinhThanh", "TenShop",
    "NganhHang", "ThuongHieu", "Model", "TenSanPham", "HinhThucXuat",
]
# DIMENSION_COLS standard sau normalize (luon co Mien, QuanHuyen)
DIMENSION_COLS = DIMENSION_COLS_FULL


def find_excel_files():
    files, seen = [], set()
    for d in DATA_DIRS:
        if not d.exists():
            continue
        for f in sorted(d.glob("*.xlsx")):
            if f.name.startswith("~$"):
                continue
            # Bo qua cac file output do chinh script sinh ra (tranh doc nham la data)
            if f.name in ("product_review.xlsx", "product_mapping.xlsx"):
                continue
            rp = f.resolve()
            if rp in seen:
                continue
            seen.add(rp)
            files.append(f)
    return files


def _read_xlsx_safe(path):
    """Doc xlsx, strip trailing null bytes neu co (de fix file MWG export bi padding)."""
    with open(path, "rb") as f:
        data = f.read()
    end_marker = b"PK\x05\x06"
    pos = data.rfind(end_marker)
    if pos > 0:
        comment_len = int.from_bytes(data[pos + 20:pos + 22], "little")
        truncate_at = pos + 22 + comment_len
        if truncate_at < len(data):
            data = data[:truncate_at]
    return io.BytesIO(data)


# Cot cua dinh dang phang moi (BI export): header o dong 3, moi dong 1 giao dich
NEW_FLAT_COLS = [
    "Nam_Thang", "Nam_Tuan", "Date", "Mien", "TinhThanh", "QuanHuyen", "TenShop",
    "NganhHang", "ThuongHieu", "Model", "TenSanPham", "HinhThucXuat",
    "DoanhSo", "PctCty", "Cty",
]


def read_flat_file(path):
    """Doc dinh dang phang moi: 15 cot, cot 'Cong ty' = R hoac Y, moi dong 1 giao dich.
    Gop thanh leaf wide rows (R_DoanhSo, Y_DoanhSo) tuong duong format cu (level 12)."""
    df = pd.read_excel(_read_xlsx_safe(path), sheet_name=0, header=2)
    if df.shape[1] != 15:
        raise ValueError(f"File {path.name} (flat) co {df.shape[1]} cot, ky vong 15.")
    df.columns = NEW_FLAT_COLS
    df["Cty"] = df["Cty"].astype(str).str.upper().str.strip()
    df = df[df["Cty"].isin(["R", "Y"])].copy()
    df["DoanhSo"] = pd.to_numeric(df["DoanhSo"], errors="coerce").fillna(0)
    df["R_DoanhSo"] = df["DoanhSo"].where(df["Cty"] == "R", 0.0)
    df["Y_DoanhSo"] = df["DoanhSo"].where(df["Cty"] == "Y", 0.0)
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    leaf = df.groupby(DIMENSION_COLS_FULL, as_index=False, dropna=False)[["R_DoanhSo", "Y_DoanhSo"]].sum()
    leaf["R_Pct"] = 0.0
    leaf["Y_Pct"] = 0.0
    leaf = leaf[DIMENSION_COLS_FULL + ["R_DoanhSo", "R_Pct", "Y_DoanhSo", "Y_Pct"]]
    leaf["_level"] = 12
    leaf["_source_file"] = path.name
    print(f"    Format: phang 15 cot (BI export) - {len(leaf):,} dong leaf")
    return leaf


def read_one_file(path):
    print(f"  Doc: {path.name}")
    # Phat hien dinh dang phang moi: peek 4 dong dau, tim chu 'Cong ty'
    try:
        peek = pd.read_excel(_read_xlsx_safe(path), sheet_name=0, header=None, nrows=4)
        is_flat = peek.astype(str).apply(lambda c: c.str.contains("Công ty", na=False)).to_numpy().any()
    except Exception:
        is_flat = False
    if is_flat:
        return read_flat_file(path)

    df = pd.read_excel(_read_xlsx_safe(path), sheet_name=0, header=[0, 1])
    n_cols = df.shape[1]

    if n_cols == 16:
        # Format day du
        df.columns = DIMENSION_COLS_FULL + ["R_DoanhSo", "R_Pct", "Y_DoanhSo", "Y_Pct"]
        print(f"    Format: 16 cot (full)")
    elif n_cols == 14:
        # Format goi: thieu Mien va QuanHuyen
        df.columns = DIMENSION_COLS_COMPACT + ["R_DoanhSo", "R_Pct", "Y_DoanhSo", "Y_Pct"]
        # Suy ra Mien tu TinhThanh
        df.insert(3, "Mien", df["TinhThanh"].map(lambda t: lookup_mien(t) if isinstance(t, str) else None))
        # Them QuanHuyen placeholder
        # - "(N/A)" khi TenShop co gia tri thuc (leaf-level rows)
        # - "Total" khi TenShop = Total (subtotal rows) -> de level numbering khop format full
        df.insert(5, "QuanHuyen", "Total")
        mask_shop_real = df["TenShop"].notna() & (df["TenShop"].astype(str) != "Total")
        df.loc[mask_shop_real, "QuanHuyen"] = "(N/A)"
        # Cho subtotal rows (TinhThanh = 'Total'), set Mien = 'Total'
        df.loc[df["TinhThanh"] == "Total", "Mien"] = "Total"
        # Cac tinh khong tim duoc mien -> 'Mien khac'
        mask_unknown = df["TinhThanh"].notna() & (df["TinhThanh"] != "Total") & df["Mien"].isna()
        df.loc[mask_unknown, "Mien"] = "Mien khac"
        print(f"    Format: 14 cot (compact) - tu suy Mien tu Tinh thanh")
    else:
        raise ValueError(f"File {path.name} co {n_cols} cot, ky vong 14 hoac 16.")

    df = df[~df["Nam_Thang"].astype(str).str.startswith("Applied filters", na=False)]
    df = df[~df["Nam_Thang"].astype(str).str.startswith("Exported", na=False)]
    is_leaf = pd.DataFrame(
        {f: (df[f].notna() & (df[f].astype(str) != "Total")) for f in DIMENSION_COLS_FULL}
    )
    cum = is_leaf.astype(int).cumprod(axis=1)
    df["_level"] = cum.sum(axis=1)
    df["_source_file"] = path.name
    for c in ["R_DoanhSo", "Y_DoanhSo"]:
        df[c] = df[c].fillna(0)
    return df


def detect_file_categories(df):
    cats = df[df["_level"] == 8]["NganhHang"].dropna().unique().tolist()
    if not cats:  # dinh dang phang: khong co level 8 -> lay tu leaf
        cats = df[df["_level"] == 12]["NganhHang"].dropna().unique().tolist()
    cats = sorted([c for c in cats if c != "Total"])
    return cats


def extract_monthly_by_category(df, file_cats):
    if len(file_cats) == 1:
        cat = file_cats[0]
        m = df[df["_level"] == 1][["Nam_Thang", "R_DoanhSo", "Y_DoanhSo"]].copy()
        m["NganhHang"] = cat
        return m.rename(columns={"Nam_Thang": "month"})[["month", "NganhHang", "R_DoanhSo", "Y_DoanhSo"]]
    else:
        sub = df[df["_level"] == 8][["Date", "NganhHang", "R_DoanhSo", "Y_DoanhSo"]].copy()
        sub["Date"] = pd.to_datetime(sub["Date"])
        sub["month"] = sub["Date"].dt.strftime("%Y-%m")
        return sub.groupby(["month", "NganhHang"], as_index=False)[["R_DoanhSo", "Y_DoanhSo"]].sum()


def extract_weekly_by_category(df, file_cats):
    if len(file_cats) == 1:
        cat = file_cats[0]
        w = df[df["_level"] == 2][["Nam_Tuan", "R_DoanhSo", "Y_DoanhSo"]].copy()
        w = w.groupby("Nam_Tuan", as_index=False)[["R_DoanhSo", "Y_DoanhSo"]].sum()
        w["NganhHang"] = cat
        return w.rename(columns={"Nam_Tuan": "week"})[["week", "NganhHang", "R_DoanhSo", "Y_DoanhSo"]]
    else:
        sub = df[(df["_level"] == 8) & df["Nam_Tuan"].notna() & (df["Nam_Tuan"] != "Total")][
            ["Nam_Tuan", "NganhHang", "R_DoanhSo", "Y_DoanhSo"]].copy()
        return sub.groupby(["Nam_Tuan", "NganhHang"], as_index=False)[["R_DoanhSo", "Y_DoanhSo"]].sum().rename(
            columns={"Nam_Tuan": "week"})


def extract_leaf(df):
    leaf = df[df["_level"] == 12].copy()
    leaf = leaf[(leaf["R_DoanhSo"] != 0) | (leaf["Y_DoanhSo"] != 0)]
    key_cols = ["Date", "Mien", "TinhThanh", "QuanHuyen", "TenShop",
                "NganhHang", "ThuongHieu", "Model", "TenSanPham", "HinhThucXuat"]
    leaf = leaf.drop_duplicates(subset=key_cols, keep="first")
    leaf["Date"] = pd.to_datetime(leaf["Date"])
    return leaf.sort_values("Date")


def sub_group(df, level, cols):
    if isinstance(cols, str):
        cols = [cols]
    s = df[df["_level"] == level]
    if s.empty:  # dinh dang phang khong co dong Total -> tinh tu leaf (level 12)
        s = df[df["_level"] == 12]
    return s.groupby(cols, as_index=False)[["R_DoanhSo", "Y_DoanhSo"]].sum()


def build_data_json(dfs_per_file):
    all_df = pd.concat([df for df, _ in dfs_per_file], ignore_index=True)

    monthly_records, weekly_records = [], []
    for df, file_cats in dfs_per_file:
        if not file_cats:
            continue
        monthly_records.append(extract_monthly_by_category(df, file_cats))
        weekly_records.append(extract_weekly_by_category(df, file_cats))

    monthly_df = pd.concat(monthly_records, ignore_index=True) if monthly_records else pd.DataFrame()
    if not monthly_df.empty:
        monthly_df["_tot"] = monthly_df["R_DoanhSo"] + monthly_df["Y_DoanhSo"]
        monthly_df = monthly_df.sort_values("_tot", ascending=False).drop_duplicates(["month", "NganhHang"])
        monthly_df = monthly_df.sort_values(["month", "NganhHang"]).drop(columns=["_tot"])

    weekly_df = pd.concat(weekly_records, ignore_index=True) if weekly_records else pd.DataFrame()
    if not weekly_df.empty:
        weekly_df["_tot"] = weekly_df["R_DoanhSo"] + weekly_df["Y_DoanhSo"]
        weekly_df = weekly_df.sort_values("_tot", ascending=False).drop_duplicates(["week", "NganhHang"])
        weekly_df = weekly_df.sort_values(["week", "NganhHang"]).drop(columns=["_tot"])

    # === Dinh dang phang (chi co leaf, khong co dong Total) ===
    # Tinh monthly/weekly truc tiep tu leaf de tranh loi gop file p1/p2 bi trung thang.
    is_flat = (not all_df.empty) and (int((all_df["_level"] < 12).sum()) == 0)
    if is_flat:
        Lf = all_df[all_df["_level"] == 12]
        monthly_df = Lf.groupby(["Nam_Thang", "NganhHang"], as_index=False)[["R_DoanhSo", "Y_DoanhSo"]].sum().rename(columns={"Nam_Thang": "month"})
        weekly_df = Lf.groupby(["Nam_Tuan", "NganhHang"], as_index=False)[["R_DoanhSo", "Y_DoanhSo"]].sum().rename(columns={"Nam_Tuan": "week"})

    grand_R = float(monthly_df["R_DoanhSo"].sum()) if not monthly_df.empty else 0
    grand_Y = float(monthly_df["Y_DoanhSo"].sum()) if not monthly_df.empty else 0

    cat_sub_df = monthly_df.groupby("NganhHang", as_index=False)[["R_DoanhSo", "Y_DoanhSo"]].sum() if not monthly_df.empty else pd.DataFrame()
    monthly_total_df = monthly_df.groupby("month", as_index=False)[["R_DoanhSo", "Y_DoanhSo"]].sum() if not monthly_df.empty else pd.DataFrame()
    weekly_total_df = weekly_df.groupby("week", as_index=False)[["R_DoanhSo", "Y_DoanhSo"]].sum() if not weekly_df.empty else pd.DataFrame()

    leaf_full = extract_leaf(all_df)

    # === Top Shop: tinh san tu du lieu day du (truoc khi rut gon) ===
    _ts = leaf_full.groupby(["Mien", "TinhThanh", "TenShop"], as_index=False)[["R_DoanhSo", "Y_DoanhSo"]].sum()
    _ts["_t"] = _ts["R_DoanhSo"] + _ts["Y_DoanhSo"]
    _ts = _ts[_ts["_t"] > 0].sort_values("_t", ascending=False).head(50)
    top_shops = [{"mien": str(r.Mien), "tinh": str(r.TinhThanh), "shop": str(r.TenShop),
                  "R": float(r.R_DoanhSo), "Y": float(r.Y_DoanhSo)} for r in _ts.itertuples()]

    # === Rut gon leaf cho payload rows de giam dung luong file ===
    # Bo cap Shop/Quan + gom theo TUAN (giu 1 ngay dai dien moi tuan) -> nhe ~8-10x.
    # Van giu Mien/Tinh/Brand/Model/SanPham/HinhThuc -> moi bo loc & bao cao van chay;
    # bieu do thang/tuan dung; bieu do ngay/thu chi con do phan giai cap tuan.
    _gcols = ["Nam_Tuan", "Mien", "TinhThanh", "NganhHang", "ThuongHieu", "Model", "TenSanPham", "HinhThucXuat"]
    leaf = leaf_full.groupby(_gcols, as_index=False, dropna=False).agg(
        R_DoanhSo=("R_DoanhSo", "sum"), Y_DoanhSo=("Y_DoanhSo", "sum"), Date=("Date", "min"))
    leaf["QuanHuyen"] = "(gộp)"
    leaf["TenShop"] = "(gộp)"
    print(f"  Rut gon rows: {len(leaf_full):,} -> {len(leaf):,} dong (bo shop/quan, gom theo tuan)")

    brand_sub = sub_group(all_df, 9, "ThuongHieu")
    modelgroup_sub = sub_group(all_df, 10, "Model")
    tinh_sub = sub_group(all_df, 5, ["Mien", "TinhThanh"])
    mien_sub = sub_group(all_df, 4, "Mien")
    # Fallback: voi compact format khong co level 4 (Mien subtotal),
    # tu derive tu level 5 (Tinh subtotal) bang cach gop theo Mien
    if mien_sub.empty and not tinh_sub.empty:
        mien_sub = tinh_sub.groupby("Mien", as_index=False)[["R_DoanhSo", "Y_DoanhSo"]].sum()

    def enc(series):
        cats = sorted(series.fillna("").astype(str).unique().tolist())
        return cats, {v: i for i, v in enumerate(cats)}

    nganh_cats, nganh_idx = enc(leaf["NganhHang"])
    mien_cats, mien_idx = enc(leaf["Mien"])
    tinh_cats, tinh_idx = enc(leaf["TinhThanh"])
    quan_cats, quan_idx = enc(leaf["QuanHuyen"])
    shop_cats, shop_idx = enc(leaf["TenShop"])
    brand_cats, brand_idx = enc(leaf["ThuongHieu"])
    model_cats, model_idx = enc(leaf["Model"])
    product_cats, product_idx = enc(leaf["TenSanPham"])
    hinh_cats, hinh_idx = enc(leaf["HinhThucXuat"])

    # === Canonical product names: gom cac variant cua cung 1 SP ===
    # Buoc 1: doc manual mapping (neu co)
    manual_map = load_gsheet_mapping()
    if not manual_map:
        manual_map = load_manual_mapping(MAPPING_FILE)

    # Buoc 2: voi moi product, xac dinh brand chinh (brand co doanh so cao nhat cho product do)
    prod_brand_map = {}
    pb = leaf.groupby(["TenSanPham", "ThuongHieu"], as_index=False)[["R_DoanhSo", "Y_DoanhSo"]].sum()
    pb["_tot"] = pb["R_DoanhSo"] + pb["Y_DoanhSo"]
    pb = pb.sort_values("_tot", ascending=False).drop_duplicates("TenSanPham")
    for _, r in pb.iterrows():
        prod_brand_map[str(r["TenSanPham"])] = str(r["ThuongHieu"])

    # Buoc 3: tinh canonical cho moi product
    product_to_canonical = {}  # product_name -> (key, display)
    for pname in product_cats:
        if pname in manual_map:
            # Ten chuan do user dat
            disp = manual_map[pname]
            key = disp.upper().strip()
            product_to_canonical[pname] = (key, disp)
        else:
            brand = prod_brand_map.get(pname, "")
            key, disp = normalize_canonical(pname, brand)
            product_to_canonical[pname] = (key, disp)

    # Buoc 4: tao dict canonical (unique keys), va mapping product_idx -> canonical_idx
    canonical_keys = sorted({v[0] for v in product_to_canonical.values() if v[0]})
    canonical_key_to_idx = {k: i for i, k in enumerate(canonical_keys)}
    # Display ten dep nhat cho moi key (chon display dau tien gap)
    canonical_display = [""] * len(canonical_keys)
    for pname, (key, disp) in product_to_canonical.items():
        if not key:
            continue
        i = canonical_key_to_idx[key]
        if not canonical_display[i]:
            canonical_display[i] = disp
    # product_idx -> canonical_idx (theo thu tu product_cats)
    product_canonical_idx = []
    for pname in product_cats:
        key = product_to_canonical.get(pname, ("", ""))[0]
        product_canonical_idx.append(canonical_key_to_idx.get(key, -1))

    L = leaf.copy()
    L["Ni"] = L["NganhHang"].fillna("").astype(str).map(nganh_idx)
    L["Mi"] = L["Mien"].fillna("").astype(str).map(mien_idx)
    L["Ti"] = L["TinhThanh"].fillna("").astype(str).map(tinh_idx)
    L["Qi"] = L["QuanHuyen"].fillna("").astype(str).map(quan_idx)
    L["Si"] = L["TenShop"].fillna("").astype(str).map(shop_idx)
    L["Bi"] = L["ThuongHieu"].fillna("").astype(str).map(brand_idx)
    L["Moi"] = L["Model"].fillna("").astype(str).map(model_idx)
    L["Pi"] = L["TenSanPham"].fillna("").astype(str).map(product_idx)
    L["Hi"] = L["HinhThucXuat"].fillna("").astype(str).map(hinh_idx)
    L["DateStr"] = L["Date"].dt.strftime("%Y-%m-%d")

    rows = []
    for r in L.itertuples(index=False):
        rows.append([
            r.DateStr,
            int(r.Ni),
            int(r.Mi), int(r.Ti), int(r.Qi), int(r.Si),
            int(r.Bi), int(r.Moi), int(r.Pi), int(r.Hi),
            round(float(r.R_DoanhSo), 2),
            round(float(r.Y_DoanhSo), 2),
        ])

    all_categories = sorted({c for _, cats in dfs_per_file for c in cats})

    payload = {
        "build_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "source_files": sorted({str(s) for s in all_df["_source_file"].unique()}),
        "categories": all_categories,
        "grand": {"R": grand_R, "Y": grand_Y},
        "monthly": [{"month": r.month, "R": float(r.R_DoanhSo), "Y": float(r.Y_DoanhSo)}
                    for r in monthly_total_df.itertuples()] if not monthly_total_df.empty else [],
        "weekly": [{"week": r.week, "R": float(r.R_DoanhSo), "Y": float(r.Y_DoanhSo)}
                   for r in weekly_total_df.itertuples()] if not weekly_total_df.empty else [],
        "monthly_by_category": [{"month": r.month, "category": r.NganhHang,
                                  "R": float(r.R_DoanhSo), "Y": float(r.Y_DoanhSo)}
                                 for r in monthly_df.itertuples()] if not monthly_df.empty else [],
        "weekly_by_category": [{"week": r.week, "category": r.NganhHang,
                                 "R": float(r.R_DoanhSo), "Y": float(r.Y_DoanhSo)}
                                for r in weekly_df.itertuples()] if not weekly_df.empty else [],
        "subtotals": {
            "category": [{"category": r.NganhHang, "R": float(r.R_DoanhSo), "Y": float(r.Y_DoanhSo)}
                         for r in cat_sub_df.itertuples()] if not cat_sub_df.empty else [],
            "mien": [{"mien": str(r.Mien), "R": float(r.R_DoanhSo), "Y": float(r.Y_DoanhSo)}
                     for r in mien_sub.itertuples()],
            "tinh": [{"mien": str(r.Mien), "tinh": str(r.TinhThanh),
                      "R": float(r.R_DoanhSo), "Y": float(r.Y_DoanhSo)}
                     for r in tinh_sub.itertuples()],
            "brand": [{"brand": str(r.ThuongHieu), "R": float(r.R_DoanhSo), "Y": float(r.Y_DoanhSo)}
                      for r in brand_sub.itertuples()],
            "model_group": [{"model": str(r.Model), "R": float(r.R_DoanhSo), "Y": float(r.Y_DoanhSo)}
                            for r in modelgroup_sub.itertuples()],
        },
        "dict": {
            "nganhhang": nganh_cats,
            "mien": mien_cats, "tinh": tinh_cats, "quan": quan_cats, "shop": shop_cats,
            "brand": brand_cats, "model": model_cats, "product": product_cats, "hinhthuc": hinh_cats,
            "product_canonical": canonical_display,
        },
        "product_to_canonical": product_canonical_idx,
        "top_shops": top_shops,
        "rows": rows,
        "row_schema": ["date", "nganhhang", "mien", "tinh", "quan", "shop",
                       "brand", "model", "product", "hinhthuc", "R", "Y"],
        "stats": {
            "n_leaf_rows": len(rows),
            "n_files": len(dfs_per_file),
            "n_categories": len(all_categories),
            "date_min": leaf["Date"].min().strftime("%Y-%m-%d") if len(leaf) else None,
            "date_max": leaf["Date"].max().strftime("%Y-%m-%d") if len(leaf) else None,
            "leaf_R_total": float(leaf["R_DoanhSo"].sum()),
            "leaf_Y_total": float(leaf["Y_DoanhSo"].sum()),
            "coverage_R": float(leaf["R_DoanhSo"].sum()) / grand_R if grand_R else 0,
            "coverage_Y": float(leaf["Y_DoanhSo"].sum()) / grand_Y if grand_Y else 0,
        },
    }
    return payload


def render_html(payload):
    if not TEMPLATE_HTML.exists():
        raise FileNotFoundError("Khong tim thay template.html canh script.")
    html = TEMPLATE_HTML.read_text(encoding="utf-8")
    data_str = json.dumps(payload, ensure_ascii=False, separators=(",", ":"))
    return html.replace("__DATA_PLACEHOLDER__", data_str)


def export_product_review(payload, all_df):
    """Xuat file product_review.xlsx de user kiem tra/sua mapping."""
    try:
        product_cats = payload["dict"]["product"]
        canonical_disp = payload["dict"]["product_canonical"]
        product_to_can = payload["product_to_canonical"]

        leaf = all_df[all_df["_level"] == 12]
        prod_sales = leaf.groupby("TenSanPham", as_index=False)[["R_DoanhSo", "Y_DoanhSo"]].sum()
        sales_map = {str(r["TenSanPham"]): (float(r["R_DoanhSo"]), float(r["Y_DoanhSo"]))
                     for _, r in prod_sales.iterrows()}

        rows = []
        for i, pname in enumerate(product_cats):
            ci = product_to_can[i]
            cname = canonical_disp[ci] if 0 <= ci < len(canonical_disp) else ""
            r_sale, y_sale = sales_map.get(pname, (0, 0))
            rows.append({
                "original": pname,
                "canonical_auto": cname,
                "canonical_override": "",
                "R_doanhso": r_sale,
                "Y_doanhso": y_sale,
                "total": r_sale + y_sale,
            })

        df_review = pd.DataFrame(rows)
        cnt = df_review.groupby("canonical_auto").size().to_dict()
        df_review["n_variants"] = df_review["canonical_auto"].map(cnt)
        df_review = df_review.sort_values(["canonical_auto", "total"], ascending=[True, False])
        df_review = df_review[["canonical_auto", "n_variants", "original",
                               "canonical_override", "R_doanhso", "Y_doanhso", "total"]]
        df_review.to_excel(REVIEW_FILE, index=False)
        print(f"  Xuat file review: {REVIEW_FILE.name} ({len(df_review)} dong, {len(cnt)} nhom canonical)")
    except Exception as e:
        print(f"  WARN: Khong xuat duoc product_review.xlsx: {e}")


def main():
    print("Build Report - R (CellphoneS) vs Y (MWG) - Multi-category")
    files = find_excel_files()
    if not files:
        print("KHONG TIM THAY file .xlsx nao. Dat file vao folder data/")
        sys.exit(1)
    print(f"Tim thay {len(files)} file:")
    for f in files:
        print(f"  - {f.name}  ({f.stat().st_size / 1024 / 1024:.2f} MB)")

    dfs_per_file = []
    for f in files:
        df = read_one_file(f)
        cats = detect_file_categories(df)
        print(f"  Nganh hang trong {f.name}: {cats}")
        dfs_per_file.append((df, cats))

    print("\nTong hop...")
    payload = build_data_json(dfs_per_file)
    s = payload["stats"]
    print(f"  Nganh hang: {payload['categories']}")
    print(f"  Thoi gian (leaf): {s['date_min']} -> {s['date_max']}")
    print(f"  So dong chi tiet: {s['n_leaf_rows']:,}")
    print(f"  Grand R: {payload['grand']['R']:,.0f}")
    print(f"  Grand Y: {payload['grand']['Y']:,.0f}")
    print(f"  Coverage R: {s['coverage_R'] * 100:.1f}%")
    print(f"  Coverage Y: {s['coverage_Y'] * 100:.1f}%")
    n_prod = len(payload["dict"]["product"])
    n_canon = len(payload["dict"]["product_canonical"])
    print(f"  San pham: {n_prod} ten goc -> {n_canon} nhom canonical (gom {n_prod-n_canon} bien the)")

    all_df = pd.concat([df for df, _ in dfs_per_file], ignore_index=True)

    # Section 5km: thi phan CPS vs MWG theo ban kinh dia ly
    try:
        payload["geo5km"] = geo5km.build_geo5km(all_df, len(files), DATA_DIRS)
    except Exception as e:
        print(f"  WARN: Khong tinh duoc geo5km: {e}")
        payload["geo5km"] = None

    export_product_review(payload, all_df)

    html = render_html(payload)
    OUTPUT_HTML.write_text(html, encoding="utf-8")
    print(f"OK {OUTPUT_HTML.name} ({OUTPUT_HTML.stat().st_size / 1024:.1f} KB)")


if __name__ == "__main__":
    main()
