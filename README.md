# Báo cáo Thị phần R (CellphoneS) vs Y (MWG)

Báo cáo HTML tương tác phân tích thị phần ngành LOA giữa CellphoneS và MWG, có bộ lọc động theo thời gian / địa lý / thương hiệu / nhóm sản phẩm / shop.

**Demo:** sau khi setup, link sẽ là `https://<username>.github.io/<repo-name>/`

---

## 1. Cấu trúc thư mục

```
.
├── data/                       <-- Đặt các file Excel (.xlsx) vào đây
│   └── Chi tiết thị trường.xlsx
├── build_report.py             <-- Script Python: gộp data → sinh HTML
├── template.html               <-- Template HTML (không sửa trừ khi cần)
├── index.html                  <-- File báo cáo sinh ra tự động
├── requirements.txt            <-- Thư viện Python cần thiết
├── .github/workflows/build.yml <-- GitHub Actions tự động build + deploy
├── .gitignore
└── README.md
```

## 2. Chạy báo cáo trên máy tính (không cần GitHub)

```bash
# Bước 1: Cài Python 3.10+ (https://www.python.org/downloads/)
# Bước 2: Cài thư viện
pip install -r requirements.txt

# Bước 3: Đặt file Excel vào folder data/
# Bước 4: Chạy
python build_report.py

# Bước 5: Mở index.html bằng trình duyệt (Chrome, Edge, Firefox)
```

## 3. Thêm dữ liệu mới (nối file)

Bạn có nhiều file Excel cùng format (ví dụ: tháng 1, tháng 2, tháng 3...):

1. Copy tất cả file `.xlsx` vào folder `data/`
2. Chạy lại `python build_report.py`
3. Script sẽ tự động:
   - Đọc tất cả file
   - Loại bỏ dòng trùng (theo ngày + tỉnh + shop + thương hiệu + sản phẩm + hình thức xuất)
   - Gộp thành dataset tổng
   - Sinh lại `index.html`

> Lưu ý: Tên file Excel có thể bất kỳ (`Q1_2026.xlsx`, `January.xlsx`...), miễn là **format pivot 2 dòng header giống file gốc** (16 cột: 12 chiều + R Chỉ số + R % cty + Y Chỉ số + Y % cty).

---

## 4. Đẩy lên GitHub & bật GitHub Pages (cho người mới)

### Bước 1: Tạo tài khoản GitHub

1. Truy cập **https://github.com/signup**
2. Đăng ký bằng email công ty (`thang.nguyen@cellphones.com.vn`)
3. Xác minh email
4. Đặt username (sẽ xuất hiện trong link, ví dụ: `thangnguyen`)

### Bước 2: Cài Git trên máy

- **Windows:** tải tại https://git-scm.com/download/win → cài đặt mặc định
- Mở **PowerShell** và kiểm tra:
  ```bash
  git --version
  ```

### Bước 3: Tạo repository (kho chứa code) trên GitHub

1. Đăng nhập GitHub → bấm dấu **+** ở góc trên phải → **New repository**
2. Đặt tên: `bao-cao-thi-phan` (hoặc tên bạn muốn)
3. Chọn **Public** (để dùng Pages miễn phí). Nếu cần Private, GitHub Pages cho repo private chỉ có ở gói trả phí.
4. **KHÔNG** tick "Add a README" (vì repo của bạn đã có README rồi)
5. Bấm **Create repository**

### Bước 4: Upload code lên GitHub

GitHub vừa hiện trang trắng với hướng dẫn. Bạn có 2 cách:

#### Cách A: Dùng web (dễ nhất, không cần Git)

1. Trên trang repo trống vừa tạo, bấm **uploading an existing file**
2. Kéo thả TOÀN BỘ thư mục này (`build_report.py`, `template.html`, `index.html`, folder `data/`, folder `.github/`, `requirements.txt`, `.gitignore`, `README.md`)
3. Cuộn xuống, gõ commit message: `Initial commit`
4. Bấm **Commit changes**

> **Lưu ý quan trọng:** Khi upload qua web, folder `.github` đôi khi bị ẩn. Bạn cần đảm bảo file `.github/workflows/build.yml` được upload (nếu thiếu, kéo riêng file `build.yml` rồi tạo folder `.github/workflows/` trong giao diện web).

#### Cách B: Dùng dòng lệnh (Git)

Mở PowerShell trong thư mục này:

```bash
git init
git branch -M main
git add .
git commit -m "Initial commit"
git remote add origin https://github.com/<username>/bao-cao-thi-phan.git
git push -u origin main
```

(Thay `<username>` bằng username GitHub của bạn)

GitHub sẽ hỏi đăng nhập — dùng tài khoản GitHub. Nếu hỏi password, bạn cần dùng **Personal Access Token** thay vì password thật:
- Vào https://github.com/settings/tokens → Generate new token (classic) → tick `repo` → Generate → copy chuỗi → dùng nó làm password

### Bước 5: Bật GitHub Pages

1. Vào repo trên GitHub → tab **Settings**
2. Menu trái → **Pages**
3. Phần **Build and deployment**:
   - **Source:** chọn **GitHub Actions**
4. Bấm **Save** (nếu có)

### Bước 6: Chờ Actions chạy

1. Vào tab **Actions** trên repo
2. Bạn sẽ thấy workflow **"Build & Deploy Report"** đang chạy (icon xoay vàng)
3. Đợi ~2-3 phút cho đến khi icon thành xanh ✅
4. Sau khi xanh, vào lại **Settings → Pages** sẽ thấy:
   > Your site is live at `https://<username>.github.io/bao-cao-thi-phan/`
5. Bấm vào link đó để xem báo cáo!

---

## 5. Cập nhật báo cáo sau này

Mỗi khi có file Excel mới, bạn chỉ cần:

### Qua web GitHub
1. Vào repo → folder `data/`
2. Bấm **Add file → Upload files** → kéo file mới vào → Commit
3. Vào tab **Actions** → đợi workflow chạy xong (~2 phút)
4. Báo cáo trên Pages tự cập nhật

### Qua Git (dòng lệnh)
```bash
# Trong thư mục dự án
# 1. Copy file mới vào data/
# 2. Push lên GitHub:
git add data/
git commit -m "Add new data file"
git push
```

GitHub Actions sẽ tự chạy `build_report.py`, sinh lại `index.html`, và deploy lên Pages.

---

## 6. Bộ lọc trên báo cáo

Báo cáo có **bộ lọc động** ở phần đầu:

- **Từ ngày / Đến ngày** — chọn khoảng thời gian
- **Preset:** 7 ngày gần nhất, 30 ngày, Tháng này, Tháng trước, Toàn bộ
- **Miền** — Bắc / Nam
- **Tỉnh thành** — 63 tỉnh, có ô tìm kiếm
- **Thương hiệu** — JBL, Marshall, Bose...
- **Nhóm sản phẩm** — Bluetooth, Karaoke, Vi tính...
- **Shop** — 2500+ shop, có tìm kiếm
- **Xoá tất cả filter** — reset về toàn bộ

Tất cả biểu đồ + bảng tự cập nhật khi bạn đổi filter.

---

## 7. Lưu ý về dữ liệu nguồn

File Excel xuất từ hệ thống nội bộ là dạng **pivot có subtotal**. Nếu file có cảnh báo "Exported data exceeded the allowed volume" ở footer, nghĩa là một số dòng chi tiết bị cắt do giới hạn export.

- **Tổng tháng & Grand total**: chính xác 100% (Excel lưu sẵn các subtotal này)
- **Dữ liệu chi tiết (leaf)**: có thể không đầy đủ → khi áp filter sẽ cho số thấp hơn

**Khuyến nghị:** Để tránh truncation, export Excel theo từng tháng (file nhỏ hơn). Script sẽ tự gộp tất cả file lại.

---

## 8. Chia sẻ báo cáo

Sau khi bật Pages, link `https://<username>.github.io/bao-cao-thi-phan/` có thể chia sẻ cho bất kỳ ai có link (kể cả ngoài công ty). Họ không cần tài khoản GitHub.

**Nếu data nhạy cảm:** nên dùng repo private. Tuy nhiên GitHub Pages cho repo private cần gói GitHub Pro (4 USD/tháng) hoặc Enterprise. Phương án thay thế:
- Mở repo private và chia sẻ file `index.html` qua email/Drive
- Dùng dịch vụ host private khác (Netlify, Vercel — đều có free tier)

---

## 9. Khắc phục sự cố

| Lỗi | Cách khắc phục |
|---|---|
| `KHONG TIM THAY file .xlsx nao` | Đảm bảo file Excel nằm trong folder `data/` |
| `co X cot, ky vong 16` | File Excel sai format - cần đúng pivot 2-dòng-header với 16 cột |
| GitHub Actions báo lỗi đỏ | Vào tab Actions → bấm vào lần chạy lỗi → xem log để biết lý do |
| Pages link không lên | Đợi 3-5 phút sau khi Actions xong. Kiểm tra Settings → Pages → Source = "GitHub Actions" |
| Báo cáo trống/lỗi JS | Mở DevTools (F12) trong trình duyệt → tab Console → xem lỗi |

---

## Bản quyền

Báo cáo nội bộ - chỉ dùng cho mục đích phân tích thị trường.
