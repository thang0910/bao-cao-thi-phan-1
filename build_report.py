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
    files, seen = [], set()
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


def detect_file_categories(df):
    cats = df[df["_level"] == 8]["NganhHang"].dropna().unique().tolist()
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
    s = df[df["_level"] == level].copy()
    if isinstance(cols, str):
        cols = [cols]
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

    grand_R = float(monthly_df["R_DoanhSo"].sum()) if not monthly_df.empty else 0
    grand_Y = float(monthly_df["Y_DoanhSo"].sum()) if not monthly_df.empty else 0

    cat_sub_df = monthly_df.groupby("NganhHang", as_index=False)[["R_DoanhSo", "Y_DoanhSo"]].sum() if not monthly_df.empty else pd.DataFrame()
    monthly_total_df = monthly_df.groupby("month", as_index=False)[["R_DoanhSo", "Y_DoanhSo"]].sum() if not monthly_df.empty else pd.DataFrame()
    weekly_total_df = weekly_df.groupby("week", as_index=False)[["R_DoanhSo", "Y_DoanhSo"]].sum() if not weekly_df.empty else pd.DataFrame()

    leaf = extract_leaf(all_df)

    brand_sub = sub_group(all_df, 9, "ThuongHieu")
    modelgroup_sub = sub_group(all_df, 10, "Model")
    tinh_sub = sub_group(all_df, 5, ["Mien", "TinhThanh"])
    mien_sub = sub_group(all_df, 4, "Mien")

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
        },
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
    s = payload["stats"]
    print(f"  Nganh hang: {payload['categories']}")
    print(f"  Thoi gian (leaf): {s['date_min']} -> {s['date_max']}")
    print(f"  So dong chi tiet: {s['n_leaf_rows']:,}")
    print(f"  Grand R: {payload['grand']['R']:,.0f}")
    print(f"  Grand Y: {payload['grand']['Y']:,.0f}")
    print(f"  Coverage R: {s['coverage_R'] * 100:.1f}%")
    print(f"  Coverage Y: {s['coverage_Y'] * 100:.1f}%")

    html = render_html(payload)
    OUTPUT_HTML.write_text(html, encoding="utf-8")
    print(f"OK {OUTPUT_HTML.name} ({OUTPUT_HTML.stat().st_size / 1024:.1f} KB)")


if __name__ == "__main__":
    main()
