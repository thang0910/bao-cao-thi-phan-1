#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
build_report.py
Script tong hop du lieu thi truong R (CellphoneS) vs Y (MWG)
tu nhieu file Excel va sinh ra bao cao HTML co bo loc dong.
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


def read_one_file(path):
    print(f"  Doc: {path.name}")
    df = pd.read_excel(path, sheet_name=0, header=[0, 1])
    if df.shape[1] != 16:
        raise ValueError(f"File {path.name} co {df.shape[1]} cot, ky vong 16.")
    df.columns = DIMENSION_COLS + ["R_DoanhSo", "R_Pct", "Y_DoanhSo", "Y_Pct"]
    df = df[~df["Nam_Thang"].astype(str).str.startswith("Applied filters", na=False)]
    df = df[~df["Nam_Thang"].astype(str).str.startswith("Exported", na=False)]
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
    m = df[df["_level"] == 1][["Nam_Thang", "R_DoanhSo", "Y_DoanhSo"]].copy()
    m["_tot"] = m["R_DoanhSo"] + m["Y_DoanhSo"]
    m = m.sort_values("_tot", ascending=False).drop_duplicates("Nam_Thang")
    return m.sort_values("Nam_Thang").drop(columns=["_tot"])


def extract_weekly(df):
    w = df[df["_level"] == 2][["Nam_Tuan", "R_DoanhSo", "Y_DoanhSo"]].copy()
    w = w.groupby("Nam_Tuan", as_index=False)[["R_DoanhSo", "Y_DoanhSo"]].sum()
    return w.sort_values("Nam_Tuan")


def extract_leaf(df):
    leaf = df[df["_level"] == 12].copy()
    leaf = leaf[(leaf["R_DoanhSo"] != 0) | (leaf["Y_DoanhSo"] != 0)]
    key_cols = ["Date", "Mien", "TinhThanh", "QuanHuyen", "TenShop",
                "NganhHang", "ThuongHieu", "Model", "TenSanPham", "HinhThucXuat"]
    leaf = leaf.drop_duplicates(subset=key_cols, keep="first")
    leaf["Date"] = pd.to_datetime(leaf["Date"])
    return leaf.sort_values("Date")


def sub_group(df, level, cols):
    s = df[df["_level"] == level].copy()
    return s.groupby(cols, as_index=False)[["R_DoanhSo", "Y_DoanhSo"]].sum()


def build_data_json(merged):
    monthly = extract_monthly(merged)
    weekly = extract_weekly(merged)
    leaf = extract_leaf(merged)
    brand_sub = sub_group(merged, 9, "ThuongHieu")
    modelgroup_sub = sub_group(merged, 10, "Model")
    tinh_sub = sub_group(merged, 5, ["Mien", "TinhThanh"])
    mien_sub = sub_group(merged, 4, "Mien")
    grand_R = float(monthly["R_DoanhSo"].sum())
    grand_Y = float(monthly["Y_DoanhSo"].sum())

    def enc(series):
        cats = sorted(series.fillna("").astype(str).unique().tolist())
        return cats, {v: i for i, v in enumerate(cats)}

    mien_cats, mien_idx = enc(leaf["Mien"])
    tinh_cats, tinh_idx = enc(leaf["TinhThanh"])
    quan_cats, quan_idx = enc(leaf["QuanHuyen"])
    shop_cats, shop_idx = enc(leaf["TenShop"])
    brand_cats, brand_idx = enc(leaf["ThuongHieu"])
    model_cats, model_idx = enc(leaf["Model"])
    product_cats, product_idx = enc(leaf["TenSanPham"])
    hinh_cats, hinh_idx = enc(leaf["HinhThucXuat"])

    L = leaf.copy()
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
            int(r.Mi), int(r.Ti), int(r.Qi), int(r.Si),
            int(r.Bi), int(r.Moi), int(r.Pi), int(r.Hi),
            round(float(r.R_DoanhSo), 2),
            round(float(r.Y_DoanhSo), 2),
        ])

    payload = {
        "build_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "source_files": sorted({str(s) for s in merged["_source_file"].unique()}),
        "grand": {"R": grand_R, "Y": grand_Y},
        "monthly": [{"month": str(r.Nam_Thang), "R": float(r.R_DoanhSo), "Y": float(r.Y_DoanhSo)}
                    for r in monthly.itertuples()],
        "weekly": [{"week": str(r.Nam_Tuan), "R": float(r.R_DoanhSo), "Y": float(r.Y_DoanhSo)}
                   for r in weekly.itertuples()],
        "subtotals": {
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
            "mien": mien_cats, "tinh": tinh_cats, "quan": quan_cats, "shop": shop_cats,
            "brand": brand_cats, "model": model_cats, "product": product_cats, "hinhthuc": hinh_cats,
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
        raise FileNotFoundError("Khong tim thay template.html canh script.")
    html = TEMPLATE_HTML.read_text(encoding="utf-8")
    data_str = json.dumps(payload, ensure_ascii=False, separators=(",", ":"))
    return html.replace("__DATA_PLACEHOLDER__", data_str)


def main():
    print("Build Report - R (CellphoneS) vs Y (MWG)")
    files = find_excel_files()
    if not files:
        print("KHONG TIM THAY file .xlsx nao. Dat file vao folder data/")
        sys.exit(1)
    print(f"Tim thay {len(files)} file:")
    for f in files:
        print(f"  - {f.name}  ({f.stat().st_size / 1024 / 1024:.2f} MB)")
    dfs = [read_one_file(f) for f in files]
    merged = merge_files(dfs)
    print(f"Tong dong sau khi gop: {len(merged):,}")
    payload = build_data_json(merged)
    stats = payload["stats"]
    print(f"  Thoi gian (leaf): {stats['date_min']} -> {stats['date_max']}")
    print(f"  So dong chi tiet: {stats['n_leaf_rows']:,}")
    print(f"  Grand R: {payload['grand']['R']:,.0f}")
    print(f"  Grand Y: {payload['grand']['Y']:,.0f}")
    print(f"  Coverage R: {stats['coverage_R'] * 100:.1f}%")
    print(f"  Coverage Y: {stats['coverage_Y'] * 100:.1f}%")
    html = render_html(payload)
    OUTPUT_HTML.write_text(html, encoding="utf-8")
    print(f"OK {OUTPUT_HTML.name} ({OUTPUT_HTML.stat().st_size / 1024:.1f} KB)")


if __name__ == "__main__":
    main()
