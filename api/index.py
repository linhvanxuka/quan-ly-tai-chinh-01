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
GSHEETS_URL = os.environ.get("GSHEETS_URL", "")

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

def normalize_text(text):
    if not text:
        return ""
    return str(text).lower()

def auto_categorize(desc):
    t = normalize_text(desc)
    
    if any(k in t for k in ["lãi vib", "gốc vib", "thấu chi", "bà lan", "trả nợ", "chuyển khoản trả nợ"]):
        return None, "Bỏ qua (Trả nợ Bank/Bà Lan - có tài khoản riêng)"
    if any(k in t for k in ["hoàn tiền", "lãi tiền gửi", "tặng tiền mặt", "tiền lương", "thu hồi nợ"]):
        return None, "Bỏ qua (Khoản thu nhập/Hoàn tiền)"
        
    if any(k in t for k in ["học sam", "trường sam", "mầm non sam", "tiểu học sam"]):
        return "Học Sam (chính khóa)", "Vợ (Quỳnh)"
    if any(k in t for k in ["học thêm", "tiền học thêm"]):
        return "Tiền học thêm con (120k x 9 buổi)", "Vợ (Quỳnh)"
    if any(k in t for k in ["sách", "truyện", "tập tô", "học liệu"]):
        return "Sách truyện, học liệu con", "Vợ (Quỳnh)"
    if any(k in t for k in ["quỹ lớp", "quỹ học"]):
        return "Quỹ học Sam (2tr/năm)", "Vợ (Quỳnh)"
    if any(k in t for k in ["cô giáo", "biếu cô", "20/10", "20/11", "8/3"]):
        return "Quà cô giáo Sam (2,5tr/năm)", "Vợ (Quỳnh)"
    if any(k in t for k in ["tiền chợ", "tiền ăn", "chợ hôm", "thịt", "cá", "rau"]):
        return "Ăn uống, tiền chợ hàng ngày", "Vợ (Quỳnh)"
    if any(k in t for k in ["bỉm", "sữa", "gạo", "nan", "meiji", "pampers", "huggies"]):
        return "Bỉm sữa, gạo", "Vợ (Quỳnh)"
    if any(k in t for k in ["tiền điện", "evn", "điện lực", "sạc xe evo", "sạc xe"]):
        return "Tiền điện (bao gồm sạc 2 xe điện)", "Chồng (Hòa)"
    if any(k in t for k in ["tiền nước", "nước sinh hoạt", "cấp nước", "sawaco", "viwaco"]):
        return "Nước sinh hoạt", "Chồng (Hòa)"
    if any(k in t for k in ["internet", "fpt telecom", "vnpt", "cáp quang"]):
        return "Internet cáp quang", "Chồng (Hòa)"
    if "339" in t:
        return "Điện thoại Hòa ĐT 339", "Chồng (Hòa)"
    if "7999" in t:
        return "Điện thoại ĐT 7999", "Chồng (Hòa)"
    if "1881" in t:
        return "Điện thoại ĐT 1881", "Chồng (Hòa)"
    if "6336" in t:
        return "Điện thoại ĐT 6336", "Chồng (Hòa)"
    if any(k in t for k in ["rửa xe", "gửi xe", "vé xe", "bãi xe"]):
        return "Rửa xe, gửi xe", "Chồng (Hòa)"
    if any(k in t for k in ["bảo dưỡng", "bảo trì", "sửa máy", "lõi lọc", "sửa điều hòa"]):
        return "Bảo trì thiết bị gia đình", "Chồng (Hòa)"
    if any(k in t for k in ["siêu thị", "winmart", "coopmart", "lotte", "mega market", "bách hóa xanh", "aeon"]):
        return "Đi siêu thị hàng tuần (*4 tuần)", "Vợ (Quỳnh)"
    if any(k in t for k in ["nhà hàng", "lẩu", "nướng", "kichi", "gogi", "buffet", "haidilao", "ăn tối", "đi chơi"]):
        return "Tiền ăn tối, đi chơi hàng tuần (*4 tuần)", "Chồng (Hòa)"
    if any(k in t for k in ["tiniworld", "tini", "khu vui chơi"]):
        return "Sam chơi TiNiWorld / Vui chơi", "Vợ (Quỳnh)"
    if any(k in t for k in ["làm đẹp", "spa", "nail", "móng", "gội đầu", "massage", "skincare"]):
        return "Chi làm đẹp Quỳnh", "Vợ (Quỳnh)"
    if any(k in t for k in ["váy", "quần áo quỳnh", "đầm quỳnh", "giày quỳnh"]):
        return "Quần áo, giày dép Quỳnh (TMĐT)", "Vợ (Quỳnh)"
    if any(k in t for k in ["bún", "phở", "bánh mì", "xôi", "ăn sáng", "cơm", "cafe", "cà phê", "highlands", "trà đá"]):
        return "Hòa: Ăn sáng, cafe cá nhân", "Chồng (Hòa)"
    if any(k in t for k in ["tiếp khách", "gặp đối tác", "mời cơm", "nhậu"]):
        return "Hòa: Ăn uống tiếp khách, ngoại giao", "Chồng (Hòa)"
    if any(k in t for k in ["áo sơ mi", "quần âu", "quần áo hòa", "giày hòa"]):
        return "Quần áo, giày dép Hòa", "Chồng (Hòa)"
    if any(k in t for k in ["cơ khí", "máy móc", "khoan", "dụng cụ", "box hdd", "thiết bị", "diy"]):
        return "Hòa: Thiết bị, máy móc, đồ cơ khí", "Chồng (Hòa)"
    if any(k in t for k in ["chatgpt", "openai", "claude", "gemini", "cursor", "midjourney"]):
        return "Hòa: Phần mềm AI (Gemini, Claude, GPT)", "Chồng (Hòa)"
    if any(k in t for k in ["bệnh viện", "bv đại học y", "đại học y", "khám bệnh", "phòng khám", "thuốc hòa"]):
        return "Hòa: Khám + Thuốc", "Chồng (Hòa)"
    if any(k in t for k in ["pharmacity", "long châu", "hạ sốt", "men vi sinh", "siro", "thuốc sam"]):
        return "Thuốc Sam + Quỳnh", "Vợ (Quỳnh)"
    if any(k in t for k in ["cắt tóc", "barber"]):
        return "Cắt tóc (Hòa + Sam)", "Chồng (Hòa)"
    if any(k in t for k in ["quần áo sam", "giày sam", "đồ chơi sam", "bánh sam", "búp bê"]):
        return "Quần áo, giày dép Sam", "Vợ (Quỳnh)"
    if any(k in t for k in ["chuyển khoản nhanh qua zalo", "zalo"]):
        return "Hòa: Thiết bị, máy móc, đồ cơ khí", "Chồng (Hòa)"
    if "shopee" in t:
        if "hòa" in t or "quạt" in t or "hdd" in t or "elmich" in t:
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
                if len(parts) >= 4:
                    try:
                        date = parts[0].strip()
                        group = parts[1].strip()
                        cat = parts[2].strip()
                        amount_str = parts[3].strip().replace(".", "").replace(",", "")
                        amount = int(float(amount_str)) if amount_str else 0
                        if amount >= 10000:
                            amount = int(amount / 1000)
                        owner = parts[4].strip() if len(parts) > 4 else ""
                        desc = parts[5].strip() if len(parts) > 5 else ""
                        source = parts[6].strip() if len(parts) > 6 else ""
                        status = parts[7].strip() if len(parts) > 7 else ""
                        expenses.append({
                            "date": date, "group": group, "category": cat,
                            "amount": amount, "owner": owner, "desc": desc,
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
    if GSHEETS_URL:
        try:
            r = requests.post(GSHEETS_URL, json={"rows": rows}, timeout=10)
            if r.status_code == 200:
                return True, "Thành công"
        except Exception as e:
            print("Lỗi POST qua Apps Script:", e)
            
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
    return True, "Đã lưu tạm thời"

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
    desc = data.get("desc", "").strip()
    raw_amount = float(data.get("amount", 0))
    
    if raw_amount >= 1000:
        amount_k = int(raw_amount / 1000)
    else:
        amount_k = int(raw_amount)
        
    cat = data.get("category")
    if not cat:
        cat, owner = auto_categorize(desc)
    else:
        owner = CATEGORY_MAP.get(cat, {}).get("owner", "Chồng (Hòa)")
        
    group = CATEGORY_MAP.get(cat, {}).get("group", "8. Chi Phí Phát Sinh")
    status = "Hợp lệ"
    if cat == "Chi Phí Khác (Đột xuất / Chưa liệt kê)":
        status = "Chi phí đột xuất"
        
    row = [date, group, cat, amount_k, owner, desc, "Nhập tay", status]
    success, msg = append_to_sheet([row])
    return jsonify({"success": success, "message": f"Đã lưu thành công: {amount_k:,}k ({desc})"})

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
    
    if filename.endswith(".xlsx") or filename.endswith(".xls"):
        wb = openpyxl.load_workbook(file, data_only=True)
        sheet = wb.active
        header_row = 1
        date_col = 1
        desc_col = 2
        amt_col = 3
        
        for r in range(1, min(15, sheet.max_row + 1)):
            row_vals = [str(sheet.cell(r, c).value or "").lower() for c in range(1, min(10, sheet.max_column + 1))]
            for c_idx, val in enumerate(row_vals, 1):
                if "ngày" in val:
                    date_col = c_idx
                if "nội dung" in val or "diễn giải" in val:
                    desc_col = c_idx
                if "số tiền" in val or "tiền ghi nợ" in val:
                    amt_col = c_idx
            if any("ngày" in v for v in row_vals) and any("nội dung" in v for v in row_vals):
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
            if amt >= 10000:
                amt_k = int(amt / 1000)
            else:
                amt_k = int(amt)
            cat, owner = auto_categorize(desc_str)
            if cat is None:
                parsed_items.append({
                    "date": date_str, "desc": desc_str, "amount": amt_k,
                    "category": owner, "owner": "", "group": "",
                    "selected": False, "is_ignored": True
                })
            else:
                group = CATEGORY_MAP.get(cat, {}).get("group", "8. Chi Phí Phát Sinh")
                parsed_items.append({
                    "date": date_str, "desc": desc_str, "amount": amt_k,
                    "category": cat, "owner": owner, "group": group,
                    "selected": True, "is_ignored": False
                })
    return jsonify({"items": parsed_items, "total": len(parsed_items)})

def handle_save_batch():
    items = request.json.get("items", [])
    if not items:
        return jsonify({"success": False, "message": "Không có giao dịch nào được chọn"}), 400
    rows = []
    for it in items:
        date = it.get("date")
        group = it.get("group")
        cat = it.get("category")
        amt_k = it.get("amount")
        owner = it.get("owner")
        desc = it.get("desc", "").replace(",", " ")
        status = "Hợp lệ" if cat != "Chi Phí Khác (Đột xuất / Chưa liệt kê)" else "Chi phí đột xuất"
        rows.append([date, group, cat, amt_k, owner, desc, "Sao kê VIB", status])
    success, msg = append_to_sheet(rows)
    return jsonify({"success": success, "count": len(rows)})

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

    if request.method == "POST":
        if is_add_manual:
            return handle_add_manual()
        elif is_suggest:
            return handle_suggest_category()
        elif is_parse:
            return handle_parse_bank()
        elif is_save:
            return handle_save_batch()

    if request.method == "GET":
        if is_summary:
            return get_summary_data()
    
    if request.method == "POST":
        if p in ["api/add-manual", "add-manual"]:
            return handle_add_manual()
        elif p in ["api/suggest-category", "suggest-category"]:
            return handle_suggest_category()
        elif p in ["api/parse-bank", "parse-bank"]:
            return handle_parse_bank()
        elif p in ["api/save-batch", "save-batch"]:
            return handle_save_batch()

    if request.method == "GET":
        if p in ["api/summary", "summary"]:
            return get_summary_data()

    # Mặc định trả về toàn bộ HTML ứng dụng (Không bao giờ bị 404!)
    return get_html(), 200, {"Content-Type": "text/html; charset=utf-8"}

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5001, debug=True)
