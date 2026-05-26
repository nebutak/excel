# Lệnh chạy và đóng gói

## 1. Tạo môi trường ảo

```bash
python -m venv venv
```

Windows:

```bash
venv\Scripts\activate
```

macOS/Linux:

```bash
source venv/bin/activate
```

## 2. Cài thư viện

```bash
pip install -r requirements.txt
```

## 3. Chạy app

```bash
python main.py
```

## 4. Đóng gói exe

```bash
pyinstaller --onefile --windowed main.py
```

File `.exe` sẽ nằm trong thư mục:

```text
dist/
```
