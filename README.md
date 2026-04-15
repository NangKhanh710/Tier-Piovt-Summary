# Tier Pivot Report — Streamlit App

## Cấu trúc files

```
tier_app/
├── app.py                    ← Main app
├── requirements.txt          ← Python dependencies
├── packages.txt              ← System packages (ODBC driver)
└── .streamlit/
    └── secrets.toml          ← DB + Email credentials (KHÔNG commit lên GitHub)
```

---

## Deploy lên Streamlit Cloud

### Bước 1: Push lên GitHub
```bash
# Tạo repo mới trên GitHub (private)
git init
git add app.py requirements.txt packages.txt
# KHÔNG add secrets.toml
git commit -m "initial"
git remote add origin https://github.com/your-username/tier-pivot-app.git
git push -u origin main
```

### Bước 2: Deploy trên Streamlit Cloud
1. Vào https://share.streamlit.io
2. Click **New app**
3. Chọn repo vừa tạo → branch `main` → file `app.py`
4. Click **Advanced settings** → **Secrets**
5. Paste nội dung file `.streamlit/secrets.toml` vào

### Bước 3: Embed link vào Power BI
1. Lấy URL app từ Streamlit Cloud (ví dụ: `https://your-app.streamlit.app`)
2. Trong Power BI, thêm **Button** → Action type: **Web URL**
3. Paste URL vào

---

## Lưu ý quan trọng

### ODBC Driver trên Streamlit Cloud
Streamlit Cloud dùng Linux — ODBC Driver 17 cần được cài qua `packages.txt`.
Nếu kết nối DB bị lỗi, thêm vào `packages.txt`:
```
msodbcsql17
```

### Email Office 365
Nếu công ty dùng MFA, cần tạo **App Password** thay vì dùng password thông thường:
- Vào https://account.microsoft.com → Security → App passwords

### Security
- KHÔNG bao giờ commit `secrets.toml` lên GitHub
- Thêm `.streamlit/secrets.toml` vào `.gitignore`

---

## Tính năng app

| Tính năng | Chi tiết |
|---|---|
| Filter | Year, Month, Region, Channel, Territory |
| Tier logic | Dynamic theo filter, CumulativeSum ≤51%/81%/96% |
| Pivot table | Tier × Month, đơn vị USD Millions |
| Export | Download Excel trực tiếp |
| Email | Gửi file Excel qua Outlook/Office 365 |
| Cache | Data cache 1 giờ, giảm tải DB |
