import os
import sys

# Đảm bảo đường dẫn import api.index
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from api.index import app

if __name__ == "__main__":
    print("Khởi động Ứng Dụng Quản Lý Tài Chính Gia Đình (Hòa & Quỳnh - 27.000k)...")
    print("Đơn vị tính: nghìn đồng (k VNĐ)")
    print("Mở trình duyệt tại: http://localhost:5001")
    app.run(host="0.0.0.0", port=5001, debug=True)
