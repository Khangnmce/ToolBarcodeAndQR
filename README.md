# QR & Barcode Bulk Generator

Web app tao QR Code va Barcode hang loat tu file Excel/CSV, xuat ZIP.

**Live demo:** https://toolbarcodeandqr-pbv92vy4sdarxhhqgunsw9.streamlit.app/ **

---

## Tinh nang

- Upload file Excel (`.xlsx`, `.xls`, `.xlsm`) hoac CSV
- Tu dong phat hien cot chua ma code
- Tao **QR Code** hoac **Barcode Code 128** hang loat
- Xuat tat ca anh vao 1 file **ZIP** tai ve truc tiep
- Xu ly da luong (16 threads), khong ghi ra o cung
- Giao dien Dark Mode hien dai

## Chay local

```bash
pip install -r requirements.txt
streamlit run streamlit_app.py
```

## Deploy len Streamlit Cloud

1. Push code len GitHub repo
2. Vao [share.streamlit.io](https://share.streamlit.io)
3. Chon repo, branch `main`, file `streamlit_app.py`
4. Click **Deploy**

## Cong nghe

- [Streamlit](https://streamlit.io)
- [pandas](https://pandas.pydata.org)
- [qrcode](https://github.com/lincolnloop/python-qrcode)
- [python-barcode](https://github.com/WhyNotHugo/python-barcode)
- [Pillow](https://python-pillow.org)
