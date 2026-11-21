import os
import re
import csv
import json
import datetime
import subprocess
from flask import Flask, render_template, request, jsonify
import openpyxl
import requests

current_dir = os.path.dirname(os.path.abspath(__file__))
root_dir = os.path.abspath(os.path.join(current_dir, ".."))
template_dir = os.path.join(root_dir, "templates")
public_dir = os.path.join(root_dir, "public")

app = Flask(__name__, template_folder=template_dir)

SPREADSHEET_ID = os.environ.get("SPREADSHEET_ID", "1OdHwBFJ5L5JO5_Q_pk-_F_acBHUbpMsdG0fd7pyb_HA")
ACTUAL_SHEET = "Chi Tiêu Thực Tế"
DEFAULT_GSHEETS_URL = "https://script.google.com/macros/s/AKfycbxQbmWKoui8wSXBtHQkJrC4tdEm54geUDQ8FvvXysdX9zMTWN-REFWkLLQ6r-5GN8aTEg/exec"
_env_url = os.environ.get("GSHEETS_URL", "")
if not _env_url or "AKfycbxqg9NL" in _env_url or "AKfycbxiHxuH" in _env_url:
    GSHEETS_URL = DEFAULT_GSHEETS_URL
else:
    GSHEETS_URL = _env_url

FALLBACK_EXPENSES = []

BUDGET_PLAN = [
    {"id": 1, "group": "1. Cố Định Con Cái", "name": "Học Sam (chính khóa)", "owner": "Vợ (Quỳnh)", "budget": 3500},
    {"id": 2, "group": "1. Cố Định Con Cái", "name": "Tiền học thêm con (120k x 9 buổi)", "owner": "Vợ (Quỳnh)", "budget": 1080},
    {"id": 3, "group": "1. Cố Định Con Cái", "name": "Sách truyện, học liệu con", "owner": "Vợ (Quỳnh)", "budget": 350},
    {"id": 4, "group": "1. Cố Định Con Cái", "name": "Quỹ học Sam (2tr/năm)", "owner": "Vợ (Quỳnh)", "budget": 200},
    {"id": 5, "group": "1. Cố Định Con Cái", "name": "Quà cô giáo Sam (2,5tr/năm)", "owner": "Vợ (Quỳnh)", "budget": 210},
    {"id": 6, "group": "2. Cố Định Bếp Núc", "name": "Ăn uống, tiền chợ hàng ngày", "owner": "Vợ (Quỳnh)", "budget": 9000},
    {"id": 7, "group": "2. Cố Định Bếp Núc", "name": "Bỉm sữa, gạo", "owner": "Vợ (Quỳnh)", "budget": 2100},
    {"id": 8, "group": "3. Cố Định Hạ Tầng", "name": "Tiền điện (bao gồm sạc 2 xe điện)", "owner": "Chồng (Hòa)", "budget": 2500},
    {"id": 9, "group": "3. Cố Định Hạ Tầng", "name": "Nước sinh hoạt", "owner": "Chồng (Hòa)", "budget": 300},
    {"id": 10, "group": "3. Cố Định Hạ Tầng", "name": "Internet cáp quang", "owner": "Chồng (Hòa)", "budget": 280},
    {"id": 11, "group": "3. Cố Định Hạ Tầng", "name": "Điện thoại Hòa ĐT 339", "owner": "Chồng (Hòa)", "budget": 120},
    {"id": 12, "group": "3. Cố Định Hạ Tầng", "name": "Điện thoại ĐT 7999", "owner": "Chồng (Hòa)", "budget": 10},
    {"id": 13, "group": "3. Cố Định Hạ Tầng", "name": "Điện thoại ĐT 1881", "owner": "Chồng (Hòa)", "budget": 10},
    {"id": 14, "group": "3. Cố Định Hạ Tầng", "name": "Điện thoại ĐT 6336", "owner": "Chồng (Hòa)", "budget": 10},
    {"id": 15, "group": "3. Cố Định Hạ Tầng", "name": "Rửa xe, gửi xe", "owner": "Chồng (Hòa)", "budget": 100},
    {"id": 16, "group": "3. Cố Định Hạ Tầng", "name": "Bảo trì thiết bị gia đình", "owner": "Chồng (Hòa)", "budget": 200},
    {"id": 17, "group": "4. Biến Đổi Gia Đình", "name": "Đi siêu thị hàng tuần (*4 tuần)", "owner": "Vợ (Quỳnh)", "budget": 1300},
    {"id": 18, "group": "4. Biến Đổi Gia Đình", "name": "Tiền ăn tối, đi chơi hàng tuần (*4 tuần)", "owner": "Chồng (Hòa)", "budget": 1300},
    {"id": 19, "group": "4. Biến Đổi Gia Đình", "name": "Sam chơi TiNiWorld / Vui chơi", "owner": "Vợ (Quỳnh)", "budget": 300},
    {"id": 20, "group": "5. Biến Đổi Quỳnh", "name": "Chi làm đẹp Quỳnh", "owner": "Vợ (Quỳnh)", "budget": 400},
    {"id": 21, "group": "5. Biến Đổi Quỳnh", "name": "Quần áo, giày dép Quỳnh (TMĐT)", "owner": "Vợ (Quỳnh)", "budget": 400},
    {"id": 22, "group": "6. Biến Đổi Hòa", "name": "Hòa: Ăn sáng, cafe cá nhân", "owner": "Chồng (Hòa)", "budget": 370},
    {"id": 23, "group": "6. Biến Đổi Hòa", "name": "Hòa: Ăn uống tiếp khách, ngoại giao", "owner": "Chồng (Hòa)", "budget": 500},
    {"id": 24, "group": "6. Biến Đổi Hòa", "name": "Quần áo, giày dép Hòa", "owner": "Chồng (Hòa)", "budget": 200},
    {"id": 25, "group": "6. Biến Đổi Hòa", "name": "Hòa: Thiết bị, máy móc, đồ cơ khí", "owner": "Chồng (Hòa)", "budget": 200},
    {"id": 26, "group": "6. Biến Đổi Hòa", "name": "Hòa: Phần mềm AI (Gemini, Claude, GPT)", "owner": "Chồng (Hòa)", "budget": 600},
    {"id": 27, "group": "7. Biến Đổi Y Tế", "name": "Hòa: Khám + Thuốc", "owner": "Chồng (Hòa)", "budget": 700},
    {"id": 28, "group": "7. Biến Đổi Y Tế", "name": "Thuốc Sam + Quỳnh", "owner": "Vợ (Quỳnh)", "budget": 300},
    {"id": 29, "group": "7. Biến Đổi Y Tế", "name": "Cắt tóc (Hòa + Sam)", "owner": "Chồng (Hòa)", "budget": 160},
    {"id": 30, "group": "7. Biến Đổi Y Tế", "name": "Quần áo, giày dép Sam", "owner": "Vợ (Quỳnh)", "budget": 300},
    {"id": 31, "group": "8. Chi Phí Phát Sinh", "name": "Chi Phí Khác (Đột xuất / Chưa liệt kê)", "owner": "Chồng (Hòa)", "budget": 0}
]

CATEGORY_MAP = {c["name"]: c for c in BUDGET_PLAN}

def strip_accents(s):
    import unicodedata
    if not s:
        return ""
    s = unicodedata.normalize('NFD', str(s))
    return ''.join(c for c in s if unicodedata.category(c) != 'Mn').replace('đ', 'd').replace('Đ', 'd').lower()

# Danh mục chuyển đổi cụm từ không dấu từ sao kê sang tiếng Việt chuẩn có dấu và có nghĩa
VIETNAMESE_PHRASE_RULES = [
    # Cụm người & gia đình
    (r'\bhoa:\b', 'Hòa:'),
    (r'\bhoa\b', 'Hòa'),
    (r'\bquynh\b', 'Quỳnh'),
    (r'\bba lan\b', 'Bà Lan'),
    (r'\bco giao\b', 'cô giáo'),
    (r'\bthay co\b', 'thầy cô'),
    (r'\bchau mai\b', 'cháu Mai'),
    (r'\bchau\b', 'cháu'),
    (r'\bong ba\b', 'ông bà'),
    (r'\bbo me\b', 'bố mẹ'),
    (r'\banh chi\b', 'anh chị'),

    # Cụm sự kiện, cỗ bàn
    (r'\ban cuoi\b', 'ăn cưới'),
    (r'\bdam cuoi\b', 'đám cưới'),
    (r'\ban hoi\b', 'ăn hỏi'),
    (r'\bdam gio\b', 'đám giỗ'),
    (r'\bsinh nhat\b', 'sinh nhật'),
    (r'\bmung sinh nhat\b', 'mừng sinh nhật'),
    (r'\bmung tho\b', 'mừng thọ'),
    (r'\bthap huong\b', 'thắp hương'),
    (r'\bmung tuoi\b', 'mừng tuổi'),
    (r'\bli xi\b', 'lì xì'),

    # Cụm ăn uống, chợ búa, đi chợ
    (r'\bchan ngan\b', 'chân ngan'),
    (r'\bthit ngan\b', 'thịt ngan'),
    (r'\bthit bo\b', 'thịt bò'),
    (r'\bthit lon\b', 'thịt lợn'),
    (r'\bthit heo\b', 'thịt heo'),
    (r'\bthit ga\b', 'thịt gà'),
    (r'\bthit vit\b', 'thịt vịt'),
    (r'\bhai san\b', 'hải sản'),
    (r'\btom ca\b', 'tôm cá'),
    (r'\brau cu\b', 'rau củ'),
    (r'\bhoa qua\b', 'hoa quả'),
    (r'\btrai cay\b', 'trái cây'),
    (r'\btien cho\b', 'tiền chợ'),
    (r'\btien an\b', 'tiền ăn'),
    (r'\bdi cho hom\b', 'đi chợ Hôm'),
    (r'\bdi cho\b', 'đi chợ'),
    (r'\bcho hom\b', 'chợ Hôm'),
    (r'\bcho sang\b', 'chợ sáng'),
    (r'\bcho chieu\b', 'chợ chiều'),
    (r'\ban sang\b', 'ăn sáng'),
    (r'\ban trua\b', 'ăn trưa'),
    (r'\ban toi\b', 'ăn tối'),
    (r'\ban uong\b', 'ăn uống'),
    (r'\bpho bo\b', 'phở bò'),
    (r'\bpho ga\b', 'phở gà'),
    (r'\bbun cha\b', 'bún chả'),
    (r'\bbun bo\b', 'bún bò'),
    (r'\bbanh mi\b', 'bánh mì'),
    (r'\bcom rang\b', 'cơm rang'),
    (r'\bcom tam\b', 'cơm tấm'),
    (r'\bcom trua\b', 'cơm trưa'),
    (r'\bca phe\b', 'cà phê'),
    (r'\bcafe sang\b', 'cafe sáng'),
    (r'\btra sua\b', 'trà sữa'),
    (r'\bnuoc ep\b', 'nước ép'),
    (r'\bsinh to\b', 'sinh tố'),
    (r'\bgia vi\b', 'gia vị'),
    (r'\bdau an\b', 'dầu ăn'),
    (r'\bnuoc mam\b', 'nước mắm'),
    (r'\bhat nem\b', 'hạt nêm'),
    (r'\bmi chinh\b', 'mì chính'),
    (r'\bbot ngot\b', 'bột ngọt'),
    (r'\bduong muoi\b', 'đường muối'),

    # Cụm bỉm sữa, gạo, con cái
    (r'\bbim sua\b', 'bỉm sữa'),
    (r'\bta bim\b', 'tã bỉm'),
    (r'\bsua bot\b', 'sữa bột'),
    (r'\bsua tuoi\b', 'sữa tươi'),
    (r'\bsua chua\b', 'sữa chua'),
    (r'\bvang sua\b', 'váng sữa'),
    (r'\bgao st25\b', 'gạo ST25'),
    (r'\bgao tam\b', 'gạo tám'),
    (r'\bgao nep\b', 'gạo nếp'),
    (r'\bgao\b', 'gạo'),
    (r'\btien hoc them\b', 'tiền học thêm'),
    (r'\bhoc them\b', 'học thêm'),
    (r'\bhoc phi\b', 'học phí'),
    (r'\btien hoc\b', 'tiền học'),
    (r'\bhoc sam\b', 'học Sam'),
    (r'\bsam hoc\b', 'Sam học'),
    (r'\btruong sam\b', 'trường Sam'),
    (r'\blop sam\b', 'lớp Sam'),
    (r'\bsach truyen\b', 'sách truyện'),
    (r'\bhoc lieu\b', 'học liệu'),
    (r'\btap to\b', 'tập tô'),
    (r'\bquy lop\b', 'quỹ lớp'),
    (r'\bquy hoc\b', 'quỹ học'),
    (r'\bquy phu huynh\b', 'quỹ phụ huynh'),
    (r'\bdo choi\b', 'đồ chơi'),
    (r'\bvui choi\b', 'vui chơi'),
    (r'\bkhu vui choi\b', 'khu vui chơi'),

    # Cụm hạ tầng, điện thoại, xe cộ
    (r'\bo dien pana\b', 'ổ điện Panasonic'),
    (r'\bo dien\b', 'ổ điện'),
    (r'\bpana\b', 'Panasonic'),
    (r'\btien dien\b', 'tiền điện'),
    (r'\bdien luc\b', 'điện lực'),
    (r'\bsac xe evo\b', 'sạc xe Evo'),
    (r'\bsac xe\b', 'sạc xe'),
    (r'\bxe dien\b', 'xe điện'),
    (r'\btien nuoc\b', 'tiền nước'),
    (r'\bnuoc sinh hoat\b', 'nước sinh hoạt'),
    (r'\bnuoc sach\b', 'nước sạch'),
    (r'\btien mang\b', 'tiền mạng'),
    (r'\bcuoc mang\b', 'cước mạng'),
    (r'\bcap quang\b', 'cáp quang'),
    (r'\bdien thoai\b', 'điện thoại'),
    (r'\bnap the\b', 'nạp thẻ'),
    (r'\bcuoc dt\b', 'cước ĐT'),
    (r'\brua xe\b', 'rửa xe'),
    (r'\bgui xe\b', 'gửi xe'),
    (r'\bdo xang\b', 'đổ xăng'),
    (r'\bxang xe\b', 'xăng xe'),
    (r'\bthay dau xe\b', 'thay dầu xe'),
    (r'\bthay dau\b', 'thay dầu'),
    (r'\bbao duong xe\b', 'bảo dưỡng xe'),
    (r'\bbao duong\b', 'bảo dưỡng'),
    (r'\bsua xe\b', 'sửa xe'),
    (r'\bva sam\b', 'vá săm'),
    (r'\bva lop\b', 'vá lốp'),
    (r'\bthay sam\b', 'thay săm'),
    (r'\bthay lop\b', 'thay lốp'),
    (r'\bbao tri\b', 'bảo trì'),
    (r'\bthiet bi\b', 'thiết bị'),
    (r'\bmay moc\b', 'máy móc'),
    (r'\bdo co khi\b', 'đồ cơ khí'),

    # Cụm y tế, sức khỏe, làm đẹp
    (r'\bthay pin dh quynh\b', 'thay pin đồng hồ Quỳnh'),
    (r'\bthay pin dh\b', 'thay pin đồng hồ'),
    (r'\bthay pin\b', 'thay pin'),
    (r'\bdh quynh\b', 'đồng hồ Quỳnh'),
    (r'\bdh\b', 'đồng hồ'),
    (r'\blam dep\b', 'làm đẹp'),
    (r'\blam toc\b', 'làm tóc'),
    (r'\bcat toc\b', 'cắt tóc'),
    (r'\buon toc\b', 'uốn tóc'),
    (r'\bnhuom toc\b', 'nhuộm tóc'),
    (r'\bgoi dau\b', 'gội đầu'),
    (r'\blam nail\b', 'làm nail'),
    (r'\blam mong\b', 'làm móng'),
    (r'\bmy pham\b', 'mỹ phẩm'),
    (r'\bquan ao\b', 'quần áo'),
    (r'\bgiay dep\b', 'giày dép'),
    (r'\btui xach\b', 'túi xách'),
    (r'\bkham benh\b', 'khám bệnh'),
    (r'\bkham thuoc\b', 'khám thuốc'),
    (r'\bmua thuoc\b', 'mua thuốc'),
    (r'\bthuoc ha sot\b', 'thuốc hạ sốt'),
    (r'\bthuoc\b', 'thuốc'),
    (r'\bha sot\b', 'hạ sốt'),
    (r'\bnho mat\b', 'nhỏ mắt'),
    (r'\bkhau trang\b', 'khẩu trang'),

    # Cụm mua sắm, tài khoản, trả nợ
    (r'\bmua 2 tk google sheet\b', 'Mua 2 tài khoản Google Sheet'),
    (r'\bmua 2 tk\b', 'Mua 2 tài khoản'),
    (r'\bmua tk google sheet\b', 'Mua tài khoản Google Sheet'),
    (r'\bmua tk\b', 'Mua tài khoản'),
    (r'\btk google sheet\b', 'tài khoản Google Sheet'),
    (r'\btk google\b', 'tài khoản Google'),
    (r'\bgoogle sheet\b', 'Google Sheet'),
    (r'\btk\b', 'tài khoản'),
    (r'\bsieu thi\b', 'siêu thị'),
    (r'\btap hoa\b', 'tạp hóa'),
    (r'\btra no\b', 'trả nợ'),
    (r'\btrả no\b', 'trả nợ'),
    (r'\btra nợ\b', 'trả nợ'),
    (r'\bck tra no\b', 'chuyển khoản trả nợ'),
    (r'\bchuyen khoan tra no\b', 'chuyển khoản trả nợ'),
    (r'\bchuyen khoan\b', 'chuyển khoản'),
    (r'\bchuyen tien\b', 'chuyển tiền'),
    (r'\bhoan tien\b', 'hoàn tiền'),
    (r'\bhoan hang\b', 'hoàn hàng'),
    (r'\bhoan tra\b', 'hoàn trả'),
    (r'\blai bank\b', 'lãi ngân hàng'),
    (r'\bgoc bank\b', 'gốc ngân hàng'),
    (r'\bthau chi\b', 'thấu chi'),

    # Từ nối & hành động thường gặp
    (r'\bdi\b', 'đi'),
    (r'\bva\b', 'và'),
    (r'\bvoi\b', 'với'),
    (r'\bcho\b', 'cho'),
    (r'\bmua\b', 'mua'),
    (r'\bphi\b', 'phí'),
    (r'\bthang (\d+)\b', r'tháng \1'),
    (r'\bbuoi\b', 'buổi'),
    (r'(\d+)\s*tr\b', r'\1 triệu'),
    (r'(\d+)\s*k\b', r'\1k')
]

def restore_vietnamese_desc(desc):
    """
    Tự động phục hồi dấu tiếng Việt chuẩn và có nghĩa từ mô tả sao kê ngân hàng
    (Ví dụ: 'Hoa: an cuoi chau mai' -> 'Hòa: ăn cưới cháu Mai', 'chan ngan' -> 'Chân ngan')
    """
    if not desc:
        return ""
    text = str(desc).strip()
    
    # Sắp xếp theo độ dài mẫu giảm dần để thay thế các cụm dài trước
    sorted_rules = sorted(VIETNAMESE_PHRASE_RULES, key=lambda x: len(x[0]), reverse=True)
    for pat, repl in sorted_rules:
        text = re.sub(pat, repl, text, flags=re.IGNORECASE)
        
    # Cam kết bất di bất dịch: Con tên SAM (không bao giờ chuyển thành Sâm)
    text = re.sub(r'\bsam\b', 'Sam', text, flags=re.IGNORECASE)
    
    # Chuẩn hóa khoảng trắng
    text = re.sub(r'\s+', ' ', text).strip()
    
    # Viết hoa chữ cái đầu tiên
    if text:
        text = text[0].upper() + text[1:]
    return text

def auto_categorize(desc):
    t = strip_accents(desc)
    
    if any(k in t for k in ["lai vib", "goc vib", "thau chi", "ba lan", "tra no", "ck tra no", "chuyen khoan tra no"]):
        return None, "Bỏ qua (Trả nợ Bank/Bà Lan - có tài khoản riêng)"
        
    # Nhận diện khoản hoàn tiền mua sắm (Shopee, TikTok, hoàn hàng)
    is_refund = any(k in t for k in ["hoan tien", "refund", "hoan hang", "hoan tra", "tra hang"])
    if is_refund:
        if any(k in t for k in ["shopee", "zalo", "tiktok", "lazada", "tiki"]):
            if any(k in t for k in ["hoa", "quat", "hdd", "may", "co khi", "khoan"]):
                return "Hòa: Thiết bị, máy móc, đồ cơ khí", "Chồng (Hòa)"
            return "Quần áo, giày dép Quỳnh (TMĐT)", "Vợ (Quỳnh)"
        return "Chi Phí Khác (Đột xuất / Chưa liệt kê)", "Chồng (Hòa)"

    if any(k in t for k in ["lai tien gui", "tang tien mat", "tien luong", "thu hoi no"]):
        return None, "Bỏ qua (Khoản thu nhập/Lương thưởng)"
        
    if any(k in t for k in ["hoc sam", "truong sam", "mam non sam", "tieu hoc sam"]):
        return "Học Sam (chính khóa)", "Vợ (Quỳnh)"
    if any(k in t for k in ["hoc them", "tien hoc them"]):
        return "Tiền học thêm con (120k x 9 buổi)", "Vợ (Quỳnh)"
    if any(k in t for k in ["sach", "truyen", "tap to", "hoc lieu"]):
        return "Sách truyện, học liệu con", "Vợ (Quỳnh)"
    if any(k in t for k in ["quy lop", "quy hoc"]):
        return "Quỹ học Sam (2tr/năm)", "Vợ (Quỳnh)"
    if any(k in t for k in ["co giao", "bieu co", "20/10", "20/11", "8/3"]):
        return "Quà cô giáo Sam (2,5tr/năm)", "Vợ (Quỳnh)"
    # Ưu tiên các danh mục cụ thể trước danh mục chợ chung
    if any(k in t for k in ["cat toc", "barber"]):
        return "Cắt tóc (Hòa + Sam)", "Chồng (Hòa)"
        
    if any(k in t for k in ["benh vien", "bv dai hoc y", "dai hoc y", "phong kham", "thuoc hoa"]):
        return "Hòa: Khám + Thuốc", "Chồng (Hòa)"
    if any(k in t for k in ["pharmacity", "long chau", "ha sot", "men vi sinh", "siro", "thuoc sam", "kham sam", "thuoc"]):
        if any(k in t for k in ["sam", "quynh", "ha sot", "siro", "con"]):
            return "Thuốc Sam + Quỳnh", "Vợ (Quỳnh)"
        return "Hòa: Khám + Thuốc", "Chồng (Hòa)"
        
    if any(k in t for k in ["an sang", "cafe", "ca phe", "highlands", "tra da", "pho bo", "pho ga", "bun cha", "xoi", "banh mi"]):
        return "Hòa: Ăn sáng, cafe cá nhân", "Chồng (Hòa)"

    if any(k in t for k in ["chan ngan", "ngan", "tien cho", "tien an", "cho hom", "thit", "ca", "rau", "ga", "bo"]):
        return "Ăn uống, tiền chợ hàng ngày", "Vợ (Quỳnh)"
    if any(k in t for k in ["bim", "sua", "gao", "nan", "meiji", "pampers", "huggies"]):
        return "Bỉm sữa, gạo", "Vợ (Quỳnh)"
    if any(k in t for k in ["tien dien", "evn", "dien luc", "sac xe evo", "sac xe"]):
        return "Tiền điện (bao gồm sạc 2 xe điện)", "Chồng (Hòa)"
    if any(k in t for k in ["tien nuoc", "nuoc sinh hoat", "cap nuoc", "sawaco", "viwaco"]):
        return "Nước sinh hoạt", "Chồng (Hòa)"
    if any(k in t for k in ["internet", "fpt telecom", "vnpt", "cap quang"]):
        return "Internet cáp quang", "Chồng (Hòa)"
    if "339" in t:
        return "Điện thoại Hòa ĐT 339", "Chồng (Hòa)"
    if "7999" in t:
        return "Điện thoại ĐT 7999", "Chồng (Hòa)"
    if "1881" in t:
        return "Điện thoại ĐT 1881", "Chồng (Hòa)"
    if "6336" in t:
        return "Điện thoại ĐT 6336", "Chồng (Hòa)"
    if any(k in t for k in ["rua xe", "gui xe", "ve xe", "bai xe"]):
        return "Rửa xe, gửi xe", "Chồng (Hòa)"
    if any(k in t for k in ["bao duong", "bao tri", "sua may", "loi loc", "sua dieu hoa"]):
        return "Bảo trì thiết bị gia đình", "Chồng (Hòa)"
    if any(k in t for k in ["sieu thi", "winmart", "coopmart", "lotte", "mega market", "bach hoa xanh", "aeon"]):
        return "Đi siêu thị hàng tuần (*4 tuần)", "Vợ (Quỳnh)"
    if any(k in t for k in ["nha hang", "lau", "nuong", "kichi", "gogi", "buffet", "haidilao", "an toi", "di choi"]):
        return "Tiền ăn tối, đi chơi hàng tuần (*4 tuần)", "Chồng (Hòa)"
    if any(k in t for k in ["tiniworld", "tini", "khu vui choi"]):
        return "Sam chơi TiNiWorld / Vui chơi", "Vợ (Quỳnh)"
    if any(k in t for k in ["pin dh quynh", "lam dep", "spa", "nail", "mong", "goi dau", "massage", "skincare"]):
        return "Chi làm đẹp Quỳnh", "Vợ (Quỳnh)"
    if any(k in t for k in ["vay", "quan ao quynh", "dam quynh", "giay quynh"]):
        return "Quần áo, giày dép Quỳnh (TMĐT)", "Vợ (Quỳnh)"
    if any(k in t for k in ["bun", "pho", "banh mi", "xoi", "an sang", "com", "cafe", "ca phe", "highlands", "tra da"]):
        return "Hòa: Ăn sáng, cafe cá nhân", "Chồng (Hòa)"
    if any(k in t for k in ["tiep khach", "gap doi tac", "moi com", "nhau"]):
        return "Hòa: Ăn uống tiếp khách, ngoại giao", "Chồng (Hòa)"
    if any(k in t for k in ["ao so mi", "quan au", "quan ao hoa", "giay hoa"]):
        return "Quần áo, giày dép Hòa", "Chồng (Hòa)"
    if any(k in t for k in ["2 o dien pana", "o dien", "pana", "co khi", "may moc", "khoan", "dung cu", "box hdd", "thiet bi", "diy"]):
        return "Hòa: Thiết bị, máy móc, đồ cơ khí", "Chồng (Hòa)"
    if any(k in t for k in ["chatgpt", "openai", "claude", "gemini", "cursor", "midjourney"]):
        return "Hòa: Phần mềm AI (Gemini, Claude, GPT)", "Chồng (Hòa)"
    if any(k in t for k in ["quan ao sam", "giay sam", "dep sam", "do choi sam"]):
        return "Quần áo, giày dép Sam", "Vợ (Quỳnh)"
    if "shopee" in t:
        if any(k in t for k in ["hoa", "quat", "hdd", "elmich", "khoan"]):
            return "Hòa: Thiết bị, máy móc, đồ cơ khí", "Chồng (Hòa)"
        return "Quần áo, giày dép Quỳnh (TMĐT)", "Vợ (Quỳnh)"
        
    return "Chi Phí Khác (Đột xuất / Chưa liệt kê)", "Chồng (Hòa)"

def fetch_actual_expenses():
    if GSHEETS_URL:
        try:
            r = requests.get(GSHEETS_URL, timeout=8)
            if r.status_code == 200:
                data = r.json()
                return data.get("expenses", [])
        except Exception as e:
            print("Lỗi fetch qua Apps Script:", e)

    try:
        cmd = ["gsheets", "readonly", "read", SPREADSHEET_ID, f"'{ACTUAL_SHEET}'!A2:H500"]
        res = subprocess.run(cmd, capture_output=True, text=True)
        expenses = []
        if res.returncode == 0 and res.stdout.strip():
            lines = res.stdout.strip().split("\n")
            for line in lines:
                parts = line.split("\t")
                if len(parts) >= 3:
                    try:
                        date = parts[0].strip()
                        desc = parts[1].strip()
                        amount_str = parts[2].strip().replace(".", "").replace(",", "")
                        amount = int(float(amount_str)) if amount_str else 0
                        if amount >= 10000:
                            amount = int(amount / 1000)
                        cat = parts[3].strip() if len(parts) > 3 else ""
                        owner = parts[4].strip() if len(parts) > 4 else ""
                        group = parts[5].strip() if len(parts) > 5 else ""
                        source = parts[6].strip() if len(parts) > 6 else ""
                        status = parts[7].strip() if len(parts) > 7 else ""
                        expenses.append({
                            "date": date, "desc": desc, "amount": amount,
                            "category": cat, "owner": owner, "group": group,
                            "source": source, "status": status
                        })
                    except Exception as e:
                        continue
            if expenses:
                return expenses
    except Exception as e:
        pass

    return FALLBACK_EXPENSES

def append_to_sheet(rows):
    global GSHEETS_URL
    url_to_use = GSHEETS_URL
    if url_to_use:
        import json as _json
        payload_str = _json.dumps({"rows": rows})
        
        # Phương án 1: POST JSON trực tiếp (Google Apps Script)
        try:
            r = requests.post(url_to_use, json={"rows": rows}, timeout=20, allow_redirects=True)
            if r.status_code == 200:
                try:
                    res_json = r.json()
                    if res_json.get("success"):
                        return True, "Thành công"
                except Exception:
                    if not ("text/html" in r.headers.get("Content-Type", "").lower() or "<!doctype" in r.text.lower()):
                        return True, "Thành công"
        except Exception:
            pass
        
        # Phương án 2: GET với payload qua query param (fallback)
        try:
            import urllib.parse as _up
            get_url = url_to_use + "?action=append&payload=" + _up.quote(payload_str)
            r2 = requests.get(get_url, timeout=20, allow_redirects=True)
            if r2.status_code == 200:
                try:
                    res_json2 = r2.json()
                    if res_json2.get("success"):
                        return True, "Thành công"
                except Exception:
                    if not ("text/html" in r2.headers.get("Content-Type", "").lower() or "<!doctype" in r2.text.lower()):
                        return True, "Thành công"
            return False, f"Google Apps Script trả về HTTP {r2.status_code}. Vui lòng kiểm tra quyền truy cập Anyone."
        except Exception as e:
            return False, f"Lỗi kết nối tới Google Apps Script: {str(e)}"
            
    try:
        batch_val = "|".join([",".join([str(x).replace(",", " ") for x in row]) for row in rows])
        cmd = ["gsheets", "mutate", "append", SPREADSHEET_ID, f"'{ACTUAL_SHEET}'", batch_val]
        res = subprocess.run(cmd, capture_output=True, text=True)
        if res.returncode == 0:
            return True, "Thành công qua CLI"
    except Exception as e:
        pass

    for r in rows:
        FALLBACK_EXPENSES.append({
            "date": r[0], "group": r[1], "category": r[2], "amount": r[3],
            "owner": r[4], "desc": r[5], "source": r[6], "status": r[7]
        })
    return False, "Chưa kết nối được Google Sheet: Vui lòng kiểm tra lại Web App URL"

def get_html():
    for p in [
        os.path.join(public_dir, "index.html"),
        os.path.join(template_dir, "index.html"),
        os.path.join(root_dir, "index.html"),
        "public/index.html",
        "templates/index.html"
    ]:
        if os.path.exists(p):
            with open(p, "r", encoding="utf-8") as f:
                return f.read()
    return "<h1>So Quan Ly Tai Chinh Hoa Quynh</h1>"

def get_summary_data():
    expenses = fetch_actual_expenses()
    cat_spent = {c["name"]: 0 for c in BUDGET_PLAN}
    for e in expenses:
        cat = e.get("category", "")
        amt = e.get("amount", 0)
        if amt >= 10000:
            amt = int(amt / 1000)
        if cat in cat_spent:
            cat_spent[cat] += amt
        else:
            cat_spent["Chi Phí Khác (Đột xuất / Chưa liệt kê)"] += amt
            
    total_budget = 27000
    wife_budget = 19440
    husband_budget = 7560
    
    total_spent = sum(cat_spent.values())
    wife_spent = 0
    husband_spent = 0
    other_spent = cat_spent["Chi Phí Khác (Đột xuất / Chưa liệt kê)"]
    
    categories_data = []
    alerts = []
    
    for c in BUDGET_PLAN:
        name = c["name"]
        budget = c["budget"]
        spent = cat_spent.get(name, 0)
        remaining = budget - spent if budget > 0 else -spent
        percent = round((spent / budget * 100), 1) if budget > 0 else 0
        
        status = "normal"
        if budget > 0:
            if spent > budget:
                status = "over"
                alerts.append({
                    "type": "danger",
                    "title": f"Vượt Hạn Mức: {name}",
                    "message": f"Đã chi {spent:,}k (Vượt {abs(remaining):,}k so với hạn mức {budget:,}k)!"
                })
            elif percent >= 80:
                status = "warning"
                alerts.append({
                    "type": "warning",
                    "title": f"Cận Ngưỡng: {name}",
                    "message": f"Đã tiêu {percent}% ({spent:,}k / {budget:,}k)."
                })
        else:
            if spent > 0:
                status = "other_alert"
                alerts.append({
                    "type": "info",
                    "title": "Phát Sinh Chi Phí Ngoài Kế Hoạch",
                    "message": f"Đã phát sinh {spent:,}k tiền đột xuất/chưa liệt kê, đã trừ vào số dư tổng!"
                })
                
        if c["owner"] == "Vợ (Quỳnh)":
            wife_spent += spent
        else:
            husband_spent += spent
            
        categories_data.append({
            "id": c["id"],
            "group": c["group"],
            "name": name,
            "owner": c["owner"],
            "budget": budget,
            "spent": spent,
            "remaining": remaining,
            "percent": percent,
            "status": status
        })
        
    today = datetime.date.today()
    day_of_month = today.day
    days_in_month = 30
    month_progress_pct = round((day_of_month / days_in_month) * 100, 1)
    
    husband_spend_pct = round((husband_spent / husband_budget * 100), 1) if husband_budget > 0 else 0
    burn_pace = round(husband_spend_pct / month_progress_pct, 2) if month_progress_pct > 0 else 1.0
    
    pace_status = "safe"
    pace_msg = "Tốc độ chi tiêu an toàn, đúng lộ trình tiết kiệm 7.000k!"
    if burn_pace > 1.2:
        pace_status = "danger"
        pace_msg = f"Cảnh báo: Bạn đang chi tiêu nhanh gấp {burn_pace}x tốc độ trôi qua của tháng!"
        alerts.insert(0, {
            "type": "danger",
            "title": "Cảnh Báo Quá Tay",
            "message": pace_msg
        })
    elif burn_pace > 1.05:
        pace_status = "warning"
        pace_msg = f"Lưu ý: Tốc độ chi tiêu hơi nhanh ({burn_pace}x), cần chú ý kiểm soát."

    return jsonify({
        "unit": "k VNĐ",
        "total_budget": total_budget,
        "total_spent": total_spent,
        "remaining_budget": total_budget - total_spent,
        "wife_budget": wife_budget,
        "wife_spent": wife_spent,
        "husband_budget": husband_budget,
        "husband_spent": husband_spent,
        "other_spent": other_spent,
        "day_of_month": day_of_month,
        "month_progress_pct": month_progress_pct,
        "husband_spend_pct": husband_spend_pct,
        "burn_pace": burn_pace,
        "pace_status": pace_status,
        "pace_msg": pace_msg,
        "categories": categories_data,
        "alerts": alerts,
        "recent_expenses": expenses[-15:][::-1]
    })

def handle_add_manual():
    data = request.json
    date = data.get("date") or datetime.date.today().strftime("%d/%m/%Y")
    desc = restore_vietnamese_desc(data.get("desc", "").strip())
    raw_amount = float(data.get("amount", 0))
    
    is_neg = (raw_amount < 0) or any(k in desc.lower() for k in ["hoàn tiền", "refund", "hoàn hàng", "hoàn trả", "hoan tien"])
    abs_amt = abs(raw_amount)
    if abs_amt >= 1000:
        amount_k = int(abs_amt / 1000)
    else:
        amount_k = int(abs_amt)
    if is_neg:
        amount_k = -amount_k
        
    cat = data.get("category")
    if not cat:
        cat, owner = auto_categorize(desc)
    else:
        owner = CATEGORY_MAP.get(cat, {}).get("owner", "Chồng (Hòa)")
        
    group = CATEGORY_MAP.get(cat, {}).get("group", "8. Chi Phí Phát Sinh")
    status = "Hợp lệ"
    if cat == "Chi Phí Khác (Đột xuất / Chưa liệt kê)":
        status = "Chi phí đột xuất"
        
    row = [date, desc, amount_k, cat, owner, group, "Nhập tay", status]
    success, msg = append_to_sheet([row])
    return jsonify({"success": success, "message": f"Đã lưu thành công: {amount_k:,}k ({desc})"})

def check_is_duplicate(date_str, amt_k, desc_str, existing_expenses):
    norm_desc = re.sub(r'[^a-zA-Z0-9]', '', desc_str.lower())
    for ex in existing_expenses:
        try:
            ex_amt = int(float(ex.get("amount", 0)))
        except:
            ex_amt = 0
        if amt_k != ex_amt:
            continue
        ex_desc = re.sub(r'[^a-zA-Z0-9]', '', str(ex.get("desc", "")).lower())
        ex_date = str(ex.get("date", "")).strip()
        
        # Nếu trùng số tiền và nội dung trùng nhau hoặc chứa nhau
        if norm_desc and ex_desc:
            if norm_desc in ex_desc or ex_desc in norm_desc or (len(norm_desc) >= 8 and norm_desc[:12] == ex_desc[:12]):
                return True
        # Hoặc trùng cả ngày và số tiền
        if date_str == ex_date and amt_k == ex_amt:
            return True
    return False

def handle_suggest_category():
    data = request.json
    desc = data.get("desc", "")
    cat, owner = auto_categorize(desc)
    return jsonify({"category": cat, "owner": owner})

def handle_parse_bank():
    if "file" not in request.files:
        return jsonify({"error": "Không tìm thấy file tải lên"}), 400
    file = request.files["file"]
    filename = file.filename.lower()
    parsed_items = []
    
    # Lấy dữ liệu chi tiêu đã có trên Google Sheets để tự động phát hiện và chống trùng lặp
    existing_expenses = fetch_actual_expenses()
    
    raw_rows = []
    if filename.endswith(".xlsx") or filename.endswith(".xls"):
        wb = openpyxl.load_workbook(file, data_only=True)
        sheet = wb.active
        header_row = 1
        date_col = 1
        desc_col = 2
        amt_col = 3
        
        for r in range(1, min(35, sheet.max_row + 1)):
            row_vals = [str(sheet.cell(r, c).value or "").lower() for c in range(1, min(15, sheet.max_column + 1))]
            for c_idx, val in enumerate(row_vals, 1):
                if any(k in val for k in ["ngày", "date"]):
                    date_col = c_idx
                if any(k in val for k in ["nội dung", "diễn giải", "description", "chi tiết"]):
                    desc_col = c_idx
                if any(k in val for k in ["ghi nợ", "tiền ghi nợ", "số tiền", "debit", "amount"]):
                    amt_col = c_idx
            if any("ngày" in v or "date" in v for v in row_vals) and any("nội dung" in v or "description" in v or "diễn giải" in v for v in row_vals):
                header_row = r
                break
                
        for r in range(header_row + 1, sheet.max_row + 1):
            date_val = sheet.cell(r, date_col).value
            desc_val = sheet.cell(r, desc_col).value
            amt_val = sheet.cell(r, amt_col).value
            if not date_val or not desc_val or amt_val is None:
                continue
            if isinstance(date_val, (datetime.datetime, datetime.date)):
                date_str = date_val.strftime("%d/%m/%Y")
            else:
                date_str = str(date_val).strip()
            desc_str = str(desc_val).strip()
            try:
                if isinstance(amt_val, (int, float)):
                    amt = float(amt_val)
                else:
                    amt_clean = str(amt_val).replace(".", "").replace(",", "").strip()
                    amt = float(amt_clean)
            except:
                continue
            raw_rows.append((date_str, desc_str, amt))
            
    elif filename.endswith(".csv"):
        import csv
        import io
        content = file.read().decode("utf-8", errors="ignore")
        reader = csv.reader(io.StringIO(content))
        csv_rows = list(reader)
        header_row = 0
        date_col = 0
        desc_col = 1
        amt_col = 2
        for r_idx, r in enumerate(csv_rows[:35]):
            low = [str(c).lower() for c in r]
            for c_idx, val in enumerate(low):
                if any(k in val for k in ["ngày", "date"]):
                    date_col = c_idx
                if any(k in val for k in ["nội dung", "diễn giải", "description", "chi tiết"]):
                    desc_col = c_idx
                if any(k in val for k in ["ghi nợ", "tiền ghi nợ", "số tiền", "debit", "amount"]):
                    amt_col = c_idx
            if any("ngày" in v or "date" in v for v in low) and any("nội dung" in v or "description" in v or "diễn giải" in v for v in low):
                header_row = r_idx
                break
        for r in csv_rows[header_row + 1:]:
            if len(r) <= max(date_col, desc_col, amt_col):
                continue
            date_str = r[date_col].strip()
            desc_str = r[desc_col].strip()
            amt_val = r[amt_col].strip()
            if not date_str or not desc_str or not amt_val:
                continue
            try:
                amt = float(amt_val.replace(".", "").replace(",", "").strip())
            except:
                continue
            raw_rows.append((date_str, desc_str, amt))

    matched_existing_indices = set()
    for date_str, desc_str, amt in raw_rows:
        desc_clean = restore_vietnamese_desc(desc_str)
        if amt >= 10000:
            amt_k = int(amt / 1000)
        else:
            amt_k = int(amt)
        is_ref = any(k in desc_str.lower() for k in ["hoàn tiền", "refund", "hoàn hàng", "hoàn trả", "hoan tien"])
        if is_ref and amt_k > 0:
            amt_k = -amt_k
            
        cat, owner = auto_categorize(desc_str)
        
        # Kiểm tra trùng lặp thông minh với Google Sheet (ưu tiên bảo lưu 100% bản anh đã sửa)
        is_dup = False
        norm_desc = re.sub(r'[^a-zA-Z0-9]', '', desc_str.lower())
        
        # Bước 1: Khớp số tiền + ngày + nội dung tương đồng
        for idx, ex in enumerate(existing_expenses):
            if idx in matched_existing_indices:
                continue
            try:
                ex_amt = int(float(ex.get("amount", 0)))
            except:
                ex_amt = 0
            if ex_amt == amt_k and str(ex.get("date", "")).strip() == date_str:
                ex_desc = re.sub(r'[^a-zA-Z0-9]', '', str(ex.get("desc", "")).lower())
                if (norm_desc and ex_desc and (norm_desc in ex_desc or ex_desc in norm_desc)) or (norm_desc[:10] == ex_desc[:10] if len(norm_desc)>=10 else False):
                    matched_existing_indices.add(idx)
                    is_dup = True
                    break
                    
        # Bước 2: Khớp số tiền + ngày (kể cả khi anh đã sửa lại toàn bộ tên nội dung trên Sheet)
        if not is_dup:
            for idx, ex in enumerate(existing_expenses):
                if idx in matched_existing_indices:
                    continue
                try:
                    ex_amt = int(float(ex.get("amount", 0)))
                except:
                    ex_amt = 0
                if ex_amt == amt_k and str(ex.get("date", "")).strip() == date_str:
                    matched_existing_indices.add(idx)
                    is_dup = True
                    break
        
        if is_dup:
            parsed_items.append({
                "date": date_str, "desc": desc_clean, "amount": amt_k,
                "category": cat or "Đã lưu trước đó", "owner": owner or "-", "group": "",
                "selected": False, "is_ignored": False, "is_duplicate": True
            })
        elif cat is None:
            parsed_items.append({
                "date": date_str, "desc": desc_clean, "amount": amt_k,
                "category": owner, "owner": "", "group": "",
                "selected": False, "is_ignored": True, "is_duplicate": False
            })
        else:
            group = CATEGORY_MAP.get(cat, {}).get("group", "8. Chi Phí Phát Sinh")
            parsed_items.append({
                "date": date_str, "desc": desc_clean, "amount": amt_k,
                "category": cat, "owner": owner, "group": group,
                "selected": True, "is_ignored": False, "is_duplicate": False
            })
            
    return jsonify({"items": parsed_items, "total": len(parsed_items)})

def handle_save_batch():
    items = request.json.get("items", [])
    if not items:
        return jsonify({"success": False, "message": "Không có giao dịch nào được chọn"}), 400
    rows = []
    for it in items:
        date = it.get("date")
        desc = it.get("desc", "").replace(",", " ")
        amt_k = it.get("amount")
        cat = it.get("category")
        owner = it.get("owner")
        group = it.get("group")
        status = "Hợp lệ" if cat != "Chi Phí Khác (Đột xuất / Chưa liệt kê)" else "Chi phí đột xuất"
        rows.append([date, desc, amt_k, cat, owner, group, "Sao kê VIB", status])
    success, msg = append_to_sheet(rows)
    return jsonify({"success": success, "count": len(rows), "message": msg})

def handle_test_gsheets():
    global GSHEETS_URL
    url = request.args.get("url") or (request.json.get("url") if request.is_json else None) or GSHEETS_URL
    masked_url = (url[:35] + "..." + url[-15:]) if url and len(url) > 50 else (url or "")
    if not url:
        return jsonify({"success": False, "error": "Chưa cấu hình URL Google Apps Script", "url": ""}), 400
    try:
        r = requests.get(url, timeout=10, allow_redirects=True)
        # Kiểm tra nội dung trả về có phải là JSON của Apps Script không
        is_html = "text/html" in r.headers.get("Content-Type", "").lower() or "<!doctype" in r.text.lower()
        if r.status_code == 200 and not is_html:
            return jsonify({"success": True, "status": 200, "message": "Kết nối Google Apps Script thành công!", "url": masked_url})
        elif r.status_code == 404 or is_html:
            return jsonify({
                "success": False, "status": 404, "url": masked_url,
                "error": "Google báo lỗi Không tìm thấy trang (404). Nguyên nhân: Quyền truy cập chưa chọn 'Bất kỳ ai (Anyone)' hoặc URL Web App chưa đúng."
            })
        else:
            return jsonify({"success": False, "status": r.status_code, "url": masked_url, "error": f"Lỗi HTTP {r.status_code} từ Google."})
    except Exception as e:
        return jsonify({"success": False, "url": masked_url, "error": f"Không thể kết nối: {str(e)}"})

# VERCEL FOOLPROOF CATCH-ALL ROUTER
@app.route("/", defaults={"path": ""}, methods=["GET", "POST"])
@app.route("/<path:path>", methods=["GET", "POST"])
def catch_all(path):
    endpoint = request.args.get("endpoint", "").strip("/").lower()
    p = (request.headers.get("x-matched-path") or 
         request.headers.get("x-vercel-matched-path") or 
         request.headers.get("x-forwarded-uri") or 
         request.path or path or "").split("?")[0].strip("/").lower()
    
    is_summary = (endpoint == "summary" or p in ["api/summary", "summary"] or p.endswith("/summary"))
    is_add_manual = (endpoint == "add-manual" or p in ["api/add-manual", "add-manual"] or p.endswith("/add-manual"))
    is_suggest = (endpoint == "suggest-category" or p in ["api/suggest-category", "suggest-category"] or p.endswith("/suggest-category"))
    is_parse = (endpoint == "parse-bank" or p in ["api/parse-bank", "parse-bank"] or p.endswith("/parse-bank"))
    is_save = (endpoint == "save-batch" or p in ["api/save-batch", "save-batch"] or p.endswith("/save-batch"))
    is_test = (endpoint == "test-gsheets" or p in ["api/test-gsheets", "test-gsheets"] or p.endswith("/test-gsheets"))

    if request.method == "POST":
        if is_add_manual:
            return handle_add_manual()
        elif is_suggest:
            return handle_suggest_category()
        elif is_parse:
            return handle_parse_bank()
        elif is_save:
            return handle_save_batch()
        elif is_test:
            return handle_test_gsheets()

    if request.method == "GET":
        if is_summary:
            return get_summary_data()
        elif is_test:
            return handle_test_gsheets()
    
    if request.method == "POST":
        if p in ["api/add-manual", "add-manual"]:
            return handle_add_manual()
        elif p in ["api/suggest-category", "suggest-category"]:
            return handle_suggest_category()
        elif p in ["api/parse-bank", "parse-bank"]:
            return handle_parse_bank()
        elif p in ["api/save-batch", "save-batch"]:
            return handle_save_batch()
        elif p in ["api/test-gsheets", "test-gsheets"]:
            return handle_test_gsheets()

    if request.method == "GET":
        if p in ["api/summary", "summary"]:
            return get_summary_data()
        elif p in ["api/test-gsheets", "test-gsheets"]:
            return handle_test_gsheets()

    # Mặc định trả về toàn bộ HTML ứng dụng (Không bao giờ bị 404!)
    return get_html(), 200, {"Content-Type": "text/html; charset=utf-8"}

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5001, debug=True)
