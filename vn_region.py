# -*- coding: utf-8 -*-
"""
vn_region.py
Mapping tinh thanh Viet Nam -> Mien theo phan loai cua MWG/CPS:
- Mien Bac: Bac + Trung (toi Da Nang)
- Mien Nam: Trung Nam + Tay Nguyen + Nam Bo
"""

TINH_TO_MIEN = {}

# Mien Bac (Bac + Trung tu Da Nang ve Bac)
_BAC = [
    "HÀ NỘI", "HÀ GIANG", "CAO BẰNG", "BẮC KẠN", "TUYÊN QUANG",
    "LÀO CAI", "ĐIỆN BIÊN", "LAI CHÂU", "SƠN LA", "YÊN BÁI",
    "HÒA BÌNH", "THÁI NGUYÊN", "LẠNG SƠN", "QUẢNG NINH", "BẮC GIANG",
    "PHÚ THỌ", "VĨNH PHÚC", "BẮC NINH", "HẢI DƯƠNG", "HẢI PHÒNG",
    "HƯNG YÊN", "THÁI BÌNH", "HÀ NAM", "NAM ĐỊNH", "NINH BÌNH",
    "THANH HÓA", "NGHỆ AN", "HÀ TĨNH", "QUẢNG BÌNH", "QUẢNG TRỊ",
    "THỪA THIÊN - HUẾ", "THỪA THIÊN HUẾ", "HUẾ",
    "ĐÀ NẴNG", "QUẢNG NAM", "QUẢNG NGÃI",
]

# Mien Nam (Trung Nam + Tay Nguyen + Nam Bo)
_NAM = [
    "BÌNH ĐỊNH", "PHÚ YÊN", "KHÁNH HÒA", "NINH THUẬN", "BÌNH THUẬN",
    "KON TUM", "GIA LAI", "ĐẮK LẮK", "ĐẮK NÔNG", "LÂM ĐỒNG",
    "BÌNH PHƯỚC", "TÂY NINH", "BÌNH DƯƠNG", "ĐỒNG NAI",
    "BÀ RỊA - VŨNG TÀU", "BÀ RỊA VŨNG TÀU", "VŨNG TÀU",
    "HỒ CHÍ MINH", "TP HỒ CHÍ MINH", "TPHCM", "HCM", "HCMC", "SÀI GÒN",
    "LONG AN", "TIỀN GIANG", "BẾN TRE", "TRÀ VINH", "VĨNH LONG",
    "ĐỒNG THÁP", "AN GIANG", "KIÊN GIANG", "CẦN THƠ", "HẬU GIANG",
    "SÓC TRĂNG", "BẠC LIÊU", "CÀ MAU",
]

for tinh in _BAC:
    TINH_TO_MIEN[tinh] = "Miền Bắc"
for tinh in _NAM:
    TINH_TO_MIEN[tinh] = "Miền Nam"


def lookup_mien(tinh_name):
    """Tra mien tu ten tinh thanh. Return 'Miền Bắc'/'Miền Nam'/None."""
    if not isinstance(tinh_name, str):
        return None
    t = tinh_name.strip().upper()
    if t in TINH_TO_MIEN:
        return TINH_TO_MIEN[t]
    # Try fuzzy: remove "TỈNH " or "THÀNH PHỐ "
    for prefix in ["TỈNH ", "THÀNH PHỐ ", "TP. ", "TP "]:
        if t.startswith(prefix):
            tt = t[len(prefix):].strip()
            if tt in TINH_TO_MIEN:
                return TINH_TO_MIEN[tt]
    return None
