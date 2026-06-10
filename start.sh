#!/bin/bash

# start.sh
# Tự động kiểm tra port, khởi động Flask Server và mở trình duyệt cho Web Tool V2

# 1. Định vị thư mục chứa script
DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" >/dev/null 2>&1 && pwd )"
cd "$DIR"

echo "=================================================="
echo "   Khởi động Auto Glossary Extractor (V2)         "
echo "=================================================="

# 2. Giải phóng port 5001 nếu đang bị chiếm dụng bởi phiên cũ
PORT=5001
PID=$(lsof -t -i:$PORT)
if [ ! -z "$PID" ]; then
    echo "[*] Phát hiện tiến trình cũ chạy trên port $PORT (PID: $PID). Đang dọn dẹp..."
    kill -9 $PID
    sleep 1
fi

# 3. Kiểm tra môi trường Python & các thư viện cần thiết
echo "[*] Đang kiểm tra môi trường Python..."
PYTHON_CMD=""

if python3 -c "import docx, fitz, pandas, openpyxl, flask" 2>/dev/null; then
    PYTHON_CMD="python3"
elif [ -f "venv/bin/python" ] && ./venv/bin/python -c "import docx, fitz, pandas, openpyxl, flask" 2>/dev/null; then
    PYTHON_CMD="./venv/bin/python"
else
    echo "[!] Thiếu thư viện hoặc môi trường chưa cài đặt đầy đủ."
    if [ -f "venv/bin/activate" ]; then
        echo "[*] Đang kích hoạt venv và cài đặt các thư viện thiếu..."
        source venv/bin/activate
        pip install python-docx PyMuPDF pandas openpyxl flask werkzeug
        PYTHON_CMD="python"
    else
        echo "[*] Đang cài đặt thư viện..."
        pip3 install python-docx PyMuPDF pandas openpyxl flask werkzeug
        PYTHON_CMD="python3"
    fi
fi

# 4. Khởi chạy Flask Server dưới dạng chạy nền (Background process)
echo "[*] Đang khởi động Flask Server..."
nohup $PYTHON_CMD app.py > server.log 2>&1 &

# Đợi server khởi động
sleep 1.5

# 5. Mở trình duyệt web
echo "[*] Đang mở trình duyệt: http://127.0.0.1:5001"
open "http://127.0.0.1:5001"

echo "=================================================="
echo "   Đã khởi động thành công!                      "
echo "   - Server chạy ngầm, bạn có thể tắt terminal.  "
echo "   - Nhật ký hoạt động ghi tại: server.log       "
echo "=================================================="
