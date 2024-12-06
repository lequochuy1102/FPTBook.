from flask import Flask, request, render_template, redirect, url_for, flash
import func
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
import os
import pandas as pd
from json import JSONDecodeError
import json

app = Flask(__name__)
app.secret_key = "sdc"  # Để sử dụng flash messages

'''Kết Nối GG Sheets và GG Drive'''
# SERVICE_ACCOUNT_FILE = os.path.join("config", "SDC2711 Accountant.json")
SERVICE_ACCOUNT_FILE = json.loads(os.environ['GOOGLE_SHEET_CREDENTIALS'])
# Các quyền (scopes) cần thiết cho Google Sheets và Google Drive
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',  # Quyền đọc/ghi Google Sheets
    'https://www.googleapis.com/auth/drive'         # Quyền truy cập Google Drive
]

# Khởi tạo kết nối với Google API

try:
    creds = Credentials.from_service_account_info(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    print("Thông tin đăng nhập hợp lệ.")
    
    # Kết nối Google Sheets API
    sheets_service = build('sheets', 'v4', credentials=creds)
    
    # Kết nối Google Drive API
    drive_service = build('drive', 'v3', credentials=creds)
    
except Exception as e:
    print("Lỗi trong tệp thông tin đăng nhập:", e)



# Route: Trang chủ
@app.route("/")
def homepage():
    return render_template('homepage/homepage.html')

@app.route("/home", methods=["GET", "POST"])
def home():
    if request.method == "POST":
        # Lấy dữ liệu từ form
        date = request.form['date']
        sheet_id = request.form['sheet_id']
        file = request.files['file']

        if not file or not sheet_id or not date:
            flash("Vui lòng nhập đủ thông tin")
            return redirect(url_for('index'))
        
        # Lưu file Excel tạm thời
        temp_file = os.path.join("static", file.filename)
        file.save(temp_file)

        try:
            # Đọc file Excel
            df = func.process_excel(temp_file)
            # Tính toán doanh thu theo shop
            shop_list = func.get_shop_list(df)
            shop_list_dict = {shop: None for shop in shop_list}
            non_vat_data = func.dt_by_shop_nonVAT(df)
            vat_data = func.dt_by_shop_VAT(df)

            # Kiểm tra DataFrame sau khi đọc file
            print("DataFrame từ file Excel:")
            print(df.head())

            # Kiểm tra danh sách shop
            print("Danh sách Shop:")
            print(type(shop_list))

            # Kiểm tra doanh thu Non-VAT và VAT
            # print("Doanh thu Non-VAT:")
            # print(non_vat_data)
            # print("Doanh thu VAT:")
            # print(vat_data)
            # Đồng bộ hóa dữ liệu


            raw_data = []
            for shop in shop_list:
                # Lấy dữ liệu của từng shop từ non_vat_data và vat_data
                non_vat = non_vat_data.get(shop, 0)  # Mặc định 0 nếu không tìm thấy shop
                vat = vat_data.get(shop, 0)          # Mặc định 0 nếu không tìm thấy shop
                raw_data.append([shop, non_vat, vat])  # Tạo một hàng mới cho mỗi shop

            # Kiểm tra dữ liệu thô
            # print("Raw Data trước khi làm sạch:")
            raw_data = [[shop if pd.notna(shop) else "Unknown", non_vat, vat] for shop, non_vat, vat in raw_data]
            # print(raw_data)


            result_data = func.clean_data(raw_data)

            # Kiểm tra dữ liệu sau khi làm sạch
            # print("Result Data sau khi làm sạch:")
            # print(result_data)
            
            # Tạo DataFrame cho kết quả
            result_df = pd.DataFrame(result_data, columns=['Shop', 'Doanh thu thực (Non-VAT)', 'Doanh thu (VAT)'])

            # Kiểm tra DataFrame cuối cùng
            print("DataFrame kết quả:")
            print(result_df)

            # Tạo tên sheet theo định dạng MM/YY
            sheet_name = pd.to_datetime(date, format="%d/%m/%Y").strftime("%m-%y")
            
            
            # Ghi dữ liệu vào Google Sheets
            # result = func.write_to_google_sheet(sheets_service, sheet_id, sheet_name, date, non_vat_data, vat_data, shop_list)
            if func.check_existing_data(sheets_service, sheet_id, sheet_name, date):
                print("shop_list_dict JSON:", json.dumps(shop_list_dict, indent=4))
                print("non_vat_data JSON:", json.dumps(non_vat_data, indent=4))
                print("vat_data JSON:", json.dumps(vat_data, indent=4))

                if not isinstance(shop_list_dict, dict):
                    shop_list_dict = {shop: None for shop in shop_list}
                
                for shop in shop_list:
                    print(f"Shop name: {shop}, Type: {type(shop)}, Repr: {repr(shop)}")
                    print("Kiểm tra non_vat_data:")
                    print(non_vat_data)
                    print("Kiểm tra vat_data:")
                    print(vat_data)
                # Nếu ngày đã tồn tại, yêu cầu người dùng xác nhận ghi đè
                return render_template("home/confirm_overwrite.html", 
                               sheet_id=sheet_id,
                               sheet_name=sheet_name,
                               date=date,
                               shop_list=json.dumps(shop_list_dict),
                               non_vat_data=json.dumps(non_vat_data),
                               vat_data=json.dumps(vat_data))
            else:
                
                print("Kiểu dữ liệu shop_list:", type(shop_list))
                print("Kiểu dữ liệu non_vat_data:", type(non_vat_data))
                print("Kiểu dữ liệu vat_data:", type(vat_data))
                # Nếu ngày chưa tồn tại, ghi dữ liệu ngay
                result = func.write_to_google_sheet(sheets_service, sheet_id, 
                                                    sheet_name, date, shop_list_dict, 
                                                    non_vat_data, vat_data)
                flash(f"Thông báo: {result}")
                return redirect(url_for('home'))


            flash(f"Successfully updated sheets in Google Sheets!")

        except Exception as e:
            flash(f"Error: {str(e)}")

        finally:
            # Xóa file Excel tạm thời
            os.remove(temp_file)
        
        return redirect(url_for('home'))

    # Lấy danh sách Google Sheets
    sheets = func.get_shared_google_sheets()
    return render_template("home/index.html", sheets=sheets)

from flask import json, request, flash, redirect, url_for

from flask import json, request, flash, redirect, url_for

@app.route("/overwrite_data", methods=["POST"])
def overwrite_data():
    try:
        sheet_id = request.form.get('sheet_id')
        sheet_name = request.form.get('sheet_name')
        date = request.form.get('date')
        shop_list_str = request.form.get('shop_list')
        non_vat_data_str = request.form.get('non_vat_data')
        vat_data_str = request.form.get('vat_data')

        # Kiểm tra xem tất cả dữ liệu bắt buộc có hay không
        if not all([sheet_id, sheet_name, date, shop_list_str, non_vat_data_str, vat_data_str]):
            flash("Thiếu dữ liệu từ form.")
            return redirect(url_for('home'))

        try:
            shop_list = json.loads(shop_list_str)
            non_vat_data = json.loads(non_vat_data_str)
            vat_data = json.loads(vat_data_str)
        except json.JSONDecodeError as e:
            flash(f"Dữ liệu JSON không hợp lệ: {e}")
            return redirect(url_for('home'))

        # Đảm bảo shop_list là dictionary. Nếu không, chuyển đổi nó.
        if not isinstance(shop_list, dict):
            if isinstance(shop_list, list):  # Nếu là list, chuyển thành dict
                shop_list = {shop: None for shop in shop_list}
            else:  # Nếu không phải list hoặc dict, báo lỗi
                flash("Dữ liệu shop_list không hợp lệ. Phải là dictionary hoặc list.")
                return redirect(url_for('home'))

        result = func.write_to_google_sheet(
            sheets_service, sheet_id, sheet_name, date, 
            shop_list, non_vat_data, vat_data, user_decision="overwrite"
        )
        flash(f"Thông báo: {result}")

    except Exception as e:
        flash(f"Lỗi không xác định: {e}")

    return redirect(url_for('home'))

@app.route('/vanhanhsx')
def vanhanhsx():
    return render_template("vanhanhsx/vanhanhsx.html")

# @app.route('homepage')
# def homepage():
#     return render_template('homepage/homepage.html')

@app.route('/nhapthu')
def nhapthu():
    return render_template('ketoan/nhapthu/nhapthu.html')

@app.route('/nhapchi')
def nhapchi():
    return render_template('ketoan/nhapchi/nhapchi.html')

@app.route('/theodoicongno')
def theodoicongno():
    return render_template('banhang/theodoicongno/theodoicongno.html')

@app.route('/danhmucsanpham')
def danhmucsanpham():
    return render_template('banhang/danhmucsanpham/danhmucsanpham.html')

@app.route('/quydinhchung')
def quydinhchung():
    return render_template('quydinh/quydinhchung.html')

@app.route('/quydinhbanhang')
def quydinhbanhang():
    return render_template('quydinh/quydinhbanhang.html')

@app.route('/quydinhxuong')
def quydinhxuong():
    return render_template('quydinh/quydinhxuong.html')

@app.route('/hopxuong')
def hopxuong():
    return render_template('sanxuat/hopxuong.html')

@app.route('/parttime')
def parttime():
    return render_template('sanxuat/parttime.html')

@app.route('/quytrinh')
def quytrinh():
    return render_template('sanxuat/quytrinh.html')

if __name__ == "__main__":
    
    app.run(debug=True, host="0.0.0.0", port=5001)
