#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
build_report.py
===============
Script tu dong tong hop du lieu thi truong (R - CellphoneS vs Y - MWG)
tu nhieu file Excel va sinh ra bao cao HTML co bo loc dong.

Cach dung:
    1. Dat cac file Excel (.xlsx) cung format vao folder ./data/
       (Hoac chinh thu muc chua script nay)
    2. Chay:
           python build_report.py
    3. Mo file index.html sinh ra trong cung thu muc.

Yeu cau:
    pip install pandas openpyxl
"""
import json
import sys
import warnings
from datetime import datetime
from pathlib import Path

import pandas as pd

warnings.filterwarnings("ignore")

SCRIPT_DIR = Path(__file__).resolve().parent
DATA_DIRS = [SCRIPT_DIR / "data", SCRIPT_DIR]
OUTPUT_HTML = SCRIPT_DIR / "index.html"
TEMPLATE_HTML = SCRIPT_DIR / "template.html"

DIMENSION_COLS = [
    "Nam_Thang", "Nam_Tuan", "Date", "Mien", "TinhThanh", "QuanHuyen", "TenShop",
    "NganhHang", "ThuongHieu", "Model", "TenSanPham", "HinhThucXuat",
]


def find_excel_files():
    files = []
    seen = set()
    for d in DATA_DIRS:
        if not d.exists():
            continue
        for f in sorted(d.glob("*.xlsx")):
            if f.name.startswith("~$"):
                continue
            rp = f.resolve()
            if rp in seen:
                continue
            seen.add(rp)
            files.append(f)
    return files


def read_one_file(path: Path) -> pd.DataFrame:
    print(f"  Doc: {path.name}")
    df = pd.read_excel(path, sheet_name=0, header=[0, 1])
    if df.shape[1] != 16:
        raise ValueError(
            f"File {path.name} co {df.shape[1]} cot, ky vong 16. Kiem tra lai format."
        )
    df.columns = DIMENSION_COLS + ["R_DoanhSo", "R_Pct", "Y_DoanhSo", "Y_Pct"]

    df = df[~df["Nam_Thang"].astype(str).str.startswith("Applied filters", na=False)]
    df = df[~df["Nam_Thang"].astype(str).str.startswith("Exported", na=False)]

    # Vectorized level calculation
    is_leaf = pd.DataFrame(
        {f: (df[f].notna() & (df[f].astype(str) != "Total")) for f in DIMENSION_COLS}
    )
    cum = is_leaf.astype(int).cumprod(axis=1)
    df["_level"] = cum.sum(axis=1)
    df["_source_file"] = path.name

    for c in ["R_DoanhSo", "Y_DoanhSo"]:
        df[c] = df[c].fillna(0)

    return df


def merge_files(dfs):
    if not dfs:
        return pd.DataFrame()
    return pd.concat(dfs, ignore_index=True)


def extract_monthly(df):
    m = df[df["_level"] == 1][["Nam_Thang", "R_DoanhSo", "Y_DoanhSo", "_source_file"]].copy()
    m["_tot"] = m["R_DoanhSo"] + m["Y_DoanhSo"]
    m = m.sort_values("_tot", ascending=False).drop_duplicates("Nam_Thang")
    m = m.sort_values("Nam_Thang")
    return m.drop(columns=["_tot"])


def extract_weekly(df):
    w = df[df["_level"] == 2][["Nam_Thang", "Nam_Tuan", "R_DoanhSo", "Y_DoanhSo"]].copy()
    w = w.groupby("Nam_Tuan", as_index=False)[["R_DoanhSo", "Y_DoanhSo"]].sum()
    w = w.sort_values("Nam_Tuan")
    return w


def extract_leaf(df):
    leaf = df[df["_level"] == 12].copy()
    leaf = leaf[(leaf["R_DoanhSo"] != 0) | (leaf["Y_DoanhSo"] != 0)]
    key_cols = ["Date", "Mien", "TinhThanh", "QuanHuyen", "TenShop",
                "NganhHang", "ThuongHieu", "Model", "TenSanPham", "HinhThucXuat"]
    leaf = leaf.drop_duplicates(subset=key_cols, keep="first")
    leaf["Date"] = pd.to_datetime(leaf["Date"])
    leaf = leaf.sort_values("Date")
    return leaf


def extract_subtotal_brand(df):
    b = df[df["_level"] == 9].copy()
    return b.groupby("ThuongHieu", as_index=False)[["R_DoanhSo", "Y_DoanhSo"]].sum()


def extract_subtotal_modelgroup(df):
    m = df[df["_level"] == 10].copy()
    return m.groupby("Model", as_index=False)[["R_DoanhSo", "Y_DoanhSo"]].sum()


def extract_subtotal_tinh(df):
    t = df[df["_level"] == 5].copy()
    return t.groupby(["Mien", "TinhThanh"], as_index=False)[["R_DoanhSo", "Y_DoanhSo"]].sum()


def extract_subtotal_mien(df):
    m = df[df["_level"] == 4].copy()
    return m.groupby("Mien", as_index=False)[["R_DoanhSo", "Y_DoanhSo"]].sum()


def build_data_json(merged):
    monthly = extract_monthly(merged)
    weekly = extract_weekly(merged)
    leaf = extract_leaf(merged)

    brand_sub = extract_subtotal_brand(merged)
    modelgroup_sub = extract_subtotal_modelgroup(merged)
    tinh_sub = extract_subtotal_tinh(merged)
    mien_sub = extract_subtotal_mien(merged)

    grand_R = float(monthly["R_DoanhSo"].sum())
    grand_Y = float(monthly["Y_DoanhSo"].sum())

    def enc(series):
        cats = series.fillna("").astype(str).unique().tolist()
        cats_sorted = sorted(cats)
        idx = {v: i for i, v in enumerate(cats_sorted)}
        return cats_sorted, idx

    mien_cats, mien_idx = enc(leaf["Mien"])
    tinh_cats, tinh_idx = enc(leaf["TinhThanh"])
    quan_cats, quan_idx = enc(leaf["QuanHuyen"])
    shop_cats, shop_idx = enc(leaf["TenShop"])
    brand_cats, brand_idx = enc(leaf["ThuongHieu"])
    model_cats, model_idx = enc(leaf["Model"])
    product_cats, product_idx = enc(leaf["TenSanPham"])
    hinh_cats, hinh_idx = enc(leaf["HinhThucXuat"])

    # Vectorized row encoding
    rows = []
    leaf_R = leaf[["Date", "Mien", "TinhThanh", "QuanHuyen", "TenShop",
                   "ThuongHieu", "Model", "TenSanPham", "HinhThucXuat",
                   "R_DoanhSo", "Y_DoanhSo"]].copy()
    leaf_R["Mien_i"] = leaf_R["Mien"].fillna("").astype(str).map(mien_idx)
    leaf_R["Tinh_i"] = leaf_R["TinhThanh"].fillna("").astype(str).map(tinh_idx)
    leaf_R["Quan_i"] = leaf_R["QuanHuyen"].fillna("").astype(str).map(quan_idx)
    leaf_R["Shop_i"] = leaf_R["TenShop"].fillna("").astype(str).map(shop_idx)
    leaf_R["Brand_i"] = leaf_R["ThuongHieu"].fillna("").astype(str).map(brand_idx)
    leaf_R["Model_i"] = leaf_R["Model"].fillna("").astype(str).map(model_idx)
    leaf_R["Product_i"] = leaf_R["TenSanPham"].fillna("").astype(str).map(product_idx)
    leaf_R["Hinh_i"] = leaf_R["HinhThucXuat"].fillna("").astype(str).map(hinh_idx)
    leaf_R["DateStr"] = leaf_R["Date"].dt.strftime("%Y-%m-%d")

    for r in leaf_R.itertuples(index=False):
        rows.append([
            r.DateStr,
            int(r.Mien_i), int(r.Tinh_i), int(r.Quan_i), int(r.Shop_i),
            int(r.Brand_i), int(r.Model_i), int(r.Product_i), int(r.Hinh_i),
            round(float(r.R_DoanhSo), 2),
            round(float(r.Y_DoanhSo), 2),
        ])

    monthly_records = [
        {"month": str(r.Nam_Thang), "R": float(r.R_DoanhSo), "Y": float(r.Y_DoanhSo)}
        for r in monthly.itertuples()
    ]
    weekly_records = [
        {"week": str(r.Nam_Tuan), "R": float(r.R_DoanhSo), "Y": float(r.Y_DoanhSo)}
        for r in weekly.itertuples()
    ]
    sub_brand = [
        {"brand": str(r.ThuongHieu), "R": float(r.R_DoanhSo), "Y": float(r.Y_DoanhSo)}
        for r in brand_sub.itertuples()
    ]
    sub_modelgroup = [
        {"model": str(r.Model), "R": float(r.R_DoanhSo), "Y": float(r.Y_DoanhSo)}
        for r in modelgroup_sub.itertuples()
    ]
    sub_tinh = [
        {"mien": str(r.Mien), "tinh": str(r.TinhThanh),
         "R": float(r.R_DoanhSo), "Y": float(r.Y_DoanhSo)}
        for r in tinh_sub.itertuples()
    ]
    sub_mien = [
        {"mien": str(r.Mien), "R": float(r.R_DoanhSo), "Y": float(r.Y_DoanhSo)}
        for r in mien_sub.itertuples()
    ]

    payload = {
        "build_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "source_files": sorted({str(s) for s in merged["_source_file"].unique()}),
        "grand": {"R": grand_R, "Y": grand_Y},
        "monthly": monthly_records,
        "weekly": weekly_records,
        "subtotals": {
            "mien": sub_mien,
            "tinh": sub_tinh,
            "brand": sub_brand,
            "model_group": sub_modelgroup,
        },
        "dict": {
            "mien": mien_cats,
            "tinh": tinh_cats,
            "quan": quan_cats,
            "shop": shop_cats,
            "brand": brand_cats,
            "model": model_cats,
            "product": product_cats,
            "hinhthuc": hinh_cats,
        },
        "rows": rows,
        "row_schema": ["date", "mien", "tinh", "quan", "shop",
                       "brand", "model", "product", "hinhthuc", "R", "Y"],
        "stats": {
            "n_leaf_rows": len(rows),
            "n_files": len(merged["_source_file"].unique()),
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
        raise FileNotFoundError(
            f"Khong tim thay template.html canh script. Vui long tai lai bo file."
        )
    html = TEMPLATE_HTML.read_text(encoding="utf-8")
    data_str = json.dumps(payload, ensure_ascii=False, separators=(",", ":"))
    html = html.replace("__DATA_PLACEHOLDER__", data_str)
    return html


def main():
    print("=" * 60)
    print("Build Report - R (CellphoneS) vs Y (MWG)")
    print("=" * 60)

    files = find_excel_files()
    if not files:
        print("\nKHONG TIM THAY file .xlsx nao trong:")
        for d in DATA_DIRS:
            print(f"  - {d}")
        print("\nVui long dat file du lieu vao ./data/ va chay lai.")
        sys.exit(1)

    print(f"\nTim thay {len(files)} file du lieu:")
    for f in files:
        print(f"  - {f.name}  ({f.stat().st_size / 1024 / 1024:.2f} MB)")

    print("\nDoc & xu ly...")
    dfs = [read_one_file(f) for f in files]
    merged = merge_files(dfs)
    print(f"\nTong dong sau khi gop: {len(merged):,}")

    print("\nXuat du lieu...")
    payload = build_data_json(merged)
    stats = payload["stats"]
    print(f"  Khoang thoi gian (leaf): {stats['date_min']} -> {stats['date_max']}")
    print(f"  So dong chi tiet: {stats['n_leaf_rows']:,}")
    print(f"  Grand R: {payload['grand']['R']:,.0f}")
    print(f"  Grand Y: {payload['grand']['Y']:,.0f}")
    print(f"  Coverage R: {stats['coverage_R'] * 100:.1f}%")
    print(f"  Coverage Y: {stats['coverage_Y'] * 100:.1f}%")

    print("\nSinh HTML...")
    html = render_html(payload)
    OUTPUT_HTML.write_text(html, encoding="utf-8")
    size_kb = OUTPUT_HTML.stat().st_size / 1024
    print(f"  OK {OUTPUT_HTML.name} ({size_kb:.1f} KB)")
    print("\nHoan tat! Mo index.html bang trinh duyet de xem bao cao.")


if __name__ == "__main__":
    main()
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                       