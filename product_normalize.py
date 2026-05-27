# -*- coding: utf-8 -*-
"""product_normalize.py - Normalize ten san pham."""
import re
import unicodedata


PREFIXES = [
    "TAI NGHE BLUETOOTH", "TAI NGHE CHUP TAI", "TAI NGHE",
    "LOA BLUETOOTH", "LOA KARAOKE", "LOA SOUNDBAR",
    "LOA VI TINH", "LOA VI TÍNH",
    "LOA KIEM AM", "LOA KIỂM ÂM",
    "LOA TRO GIANG", "LOA TRỢ GIẢNG",
    "MIC KARAOKE", "LOA SUB", "LOA",
]

COLOR_WORDS = {
    "BLK", "BLACK", "DEN",
    "WHT", "WHITE", "TRANG",
    "RED", "DO",
    "BLU", "BLUE", "XANH",
    "NAVY", "NAVYBLUE",
    "GRY", "GREY", "GRAY", "XAM",
    "YEL", "YELLOW", "VANG",
    "PNK", "PINK", "HONG",
    "SLV", "SILVER", "BAC",
    "GRN", "GREEN", "XANHLA",
    "ORG", "ORANGE", "CAM",
    "PUR", "PURPLE", "TIM",
    "BRN", "BROWN", "NAU",
    "KHAKI", "BEIGE", "VINTAGE",
    "GOLD", "VANGGOLD",
}

SUFFIX_PATTERNS = [
    re.compile(r"\s*-\s*IMEI\s*$", re.IGNORECASE),
    re.compile(r"\s+IMEI\s*$", re.IGNORECASE),
]


def strip_accents(s):
    if not isinstance(s, str):
        return ""
    # Replace Đ/đ first
    s = s.replace("Đ", "D").replace("đ", "d")
    nfkd = unicodedata.normalize("NFD", s)
    return "".join(c for c in nfkd if unicodedata.category(c) != "Mn")


def is_sku_token(token):
    if len(token) < 6:
        return False
    has_letter = any(c.isalpha() and c.isascii() for c in token)
    has_digit = any(c.isdigit() for c in token)
    return has_letter and has_digit


def normalize_canonical(name, brand=""):
    if not isinstance(name, str) or not name.strip():
        return ("", str(name) if name else "")

    s = name.upper().strip()
    for pat in SUFFIX_PATTERNS:
        s = pat.sub("", s)

    for p in sorted(PREFIXES, key=len, reverse=True):
        if s.startswith(p + " "):
            s = s[len(p) + 1:].strip()
            break

    brand_up = (brand or "").upper().strip()
    if brand_up and s.startswith(brand_up + " "):
        s = s[len(brand_up) + 1:].strip()

    s_no_punct = re.sub(r"[,;()\[\]/\\]", " ", s)
    tokens = s_no_punct.split()

    keep = []
    for t in tokens:
        t_no_dash = t.replace("-", "").replace("_", "")
        t_no_acc = strip_accents(t_no_dash).upper()
        if t_no_acc in COLOR_WORDS:
            continue
        if is_sku_token(t_no_dash):
            continue
        keep.append(t)

    base = " ".join(keep).strip()
    if not base:
        base = s

    canonical_display = (brand_up + " " + base).strip() if brand_up else base
    canonical_key = strip_accents(canonical_display).upper()
    canonical_key = re.sub(r"\s+", " ", canonical_key).strip()
    canonical_key = re.sub(r"[^A-Z0-9 ]", "", canonical_key)
    return (canonical_key, canonical_display)


def load_manual_mapping(path):
    import pandas as pd
    if not path.exists():
        return {}
    try:
        if path.suffix.lower() == ".csv":
            df = pd.read_csv(path)
        else:
            df = pd.read_excel(path)
    except Exception as e:
        print(f"  WARN: Khong doc duoc mapping file: {e}")
        return {}

    cols_lower = {c.lower(): c for c in df.columns}
    orig_col = cols_lower.get("original") or cols_lower.get("ten goc") or cols_lower.get("tên gốc")
    canon_col = cols_lower.get("canonical") or cols_lower.get("ten chuan") or cols_lower.get("tên chuẩn")
    if not orig_col or not canon_col:
        print(f"  WARN: Mapping file thieu cot 'original' va 'canonical'.")
        return {}

    mapping = {}
    for _, r in df.iterrows():
        o = str(r[orig_col]).strip()
        c = str(r[canon_col]).strip()
        if o and c and o != "nan" and c != "nan":
            mapping[o] = c
    print(f"  Doc {len(mapping)} mapping thu cong tu {path.name}")
    return mapping
