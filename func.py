import app
import pandas as pd
from googleapiclient.errors import HttpError
from datetime import datetime
import math
import json
import urllib.parse


def get_shared_google_sheets(keyword="TMDT"):
    query = f"mimeType='application/vnd.google-apps.spreadsheet' and name contains '{keyword}'"
    results = app.drive_service.files().list(
        q=query,
        pageSize=50,
        fields="files(id, name)"
    ).execute()
    files = results.get('files', [])
    return files

# def write_to_google_sheet(spreadsheet_id, sheet_name, dataframe):
#     """
#     Ghi dữ liệu vào Google Sheets
#     """
#     values = [dataframe.columns.tolist()] + dataframe.values.tolist()
#     body = {
#         'values': values
#     }
    
#     try:
#         # Kiểm tra xem sheet có tồn tại không
#         existing_sheets = get_spreadsheet_sheets(spreadsheet_id)
#         if sheet_name not in existing_sheets:
#             # Tạo sheet mới nếu không tồn tại
#             create_sheet(spreadsheet_id, sheet_name)
        
#         # Lấy dữ liệu hiện có của sheet để kiểm tra
#         sheet = app.sheets_service.spreadsheets().values().get(
#             spreadsheetId=spreadsheet_id,
#             range=f"{sheet_name}!A6:A6"  # Kiểm tra ô A6
#         ).execute()
        
#         existing_header = sheet.get('values', [])

#         # Kiểm tra xem ô A6 đã có dữ liệu chưa
#         if not existing_header:  # Nếu không có dữ liệu, điền vào ô A6
#             # Lấy danh sách shop từ dataframe
#             shop_list = dataframe['Shop'].tolist()
#             sheet.update(f"A6:{chr(65 + len(shop_list) - 1)}6", [shop_list])  # Cập nhật danh sách shop vào hàng thứ 6

#         # Ghi dữ liệu vào sheet từ hàng 7 trở đi
#         result = app.sheets_service.spreadsheets().values().update(
#             spreadsheetId=spreadsheet_id,
#             range=f"{sheet_name}!A7",  # Ghi vào từ ô A7 trở đi
#             valueInputOption="RAW",
#             body=body
#         ).execute()
        
#         return result.get('updatedCells')

#     except HttpError as err:
#         return f"Error: {err}"

def sanitize_sheet_name(sheet_name):
    # Mã hóa tên sheet để sử dụng trong phạm vi
    return urllib.parse.quote(sheet_name, safe='') 

from googleapiclient.errors import HttpError
def write_to_google_sheet(sheets_service, spreadsheet_id, sheet_name, date, shop_list, non_vat_data, vat_data, user_decision=None):
    """
    Ghi dữ liệu vào Google Sheets với định dạng:
    - Hàng 6: ["Ngày", "DT", "DTT", danh sách tên shop]
    - Hàng 7 trở đi: [Ngày, Tổng DT, Tổng DTT, doanh thu theo từng shop]
    """
    sheet_name = sanitize_sheet_name(sheet_name)
    try:
        # Kiểm tra xem sheet có tồn tại không
        existing_sheets = get_spreadsheet_sheets(spreadsheet_id)
        if sheet_name not in existing_sheets:
            create_sheet(spreadsheet_id, sheet_name)

        # Lấy dữ liệu hiện tại trong sheet
        sheet_data = sheets_service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=f"{sheet_name}!A7:Z"
        ).execute().get('values', [])

        sheet_header = sheets_service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=f"{sheet_name}!A6:Z6"
        ).execute().get('values', [])

        # Nếu shop_list là dictionary, chuyển đổi thành list của các keys
        if isinstance(shop_list, dict):
            shop_list_keys = list(shop_list.keys())  # Chuyển dictionary thành list của keys
        elif isinstance(shop_list, list):
            shop_list_keys = shop_list  # Nếu đã là list, sử dụng luôn
        else:
            raise TypeError("shop_list must be a list or dictionary")

        # Nếu sheet_header trống, tạo header mới với shop_list
        if not sheet_header:
            header = ["Ngày", "Doanh Thu VAT", "Doanh Thu Thực", "Tổng Doanh Thu Các Shop"] + shop_list_keys
            sheets_service.spreadsheets().values().update(
                spreadsheetId=spreadsheet_id,
                range=f"{sheet_name}!A6:{chr(65 + len(header) - 1)}6",
                valueInputOption="RAW",
                body={'values': [header]}
            ).execute()
            sheet_header = [header]  # Tạo header mới cho sheet trống

        # Cập nhật header nếu có shop mới
        current_header = sheet_header[0]
        updated_header = current_header[:]  # Khởi tạo updated_header từ current_header
        new_shops = [shop for shop in shop_list_keys if shop not in current_header]  # Các shop mới
        if new_shops:
            updated_header.extend(new_shops)  # Cập nhật header với các shop mới

            sheets_service.spreadsheets().values().update(
                spreadsheetId=spreadsheet_id,
                range=f"{sheet_name}!A6:{chr(65 + len(updated_header) - 1)}6",
                valueInputOption="RAW",
                body={'values': [updated_header]}
            ).execute()

        # Tạo dữ liệu cho hàng 7 (dữ liệu mới từ file Excel)
        row_data = [pd.to_datetime(date, format="%d/%m/%Y").strftime("%d/%m/%Y"), sum(vat_data.values()), sum(non_vat_data.values()), sum(non_vat_data.values())]

        # Đảm bảo rằng mỗi shop có doanh thu từ non_vat_data và vat_data
        for shop in updated_header[4:]:  # Lấy tất cả shop từ vị trí thứ 4 trở đi (trong header)
            row_data.append(non_vat_data.get(shop, 0))  # Doanh thu Non-VAT cho shop

        # Kiểm tra số lượng cột trong row_data và header
        if len(row_data) != len(updated_header):
            # Điều chỉnh số lượng cột trong row_data cho phù hợp với header
            missing_columns = len(updated_header) - len(row_data)
            if missing_columns > 0:
                # Nếu thiếu cột, thêm giá trị mặc định vào để khớp với header
                row_data.extend([0] * missing_columns)  # Thêm các cột thiếu vào

        # In ra số lượng cột trong row_data và sheet_header để kiểm tra
        print(f"Row Data Length: {len(row_data)}")
        print(f"Sheet Header Length: {len(updated_header)}")

        # Kiểm tra lại phạm vi cột khi ghi dữ liệu
        range_to_update = f"{sheet_name}!A7:{chr(65 + len(row_data) - 1)}7"
        print(f"Range to update: {range_to_update}")

        # Ghi dữ liệu vào Google Sheets
        if user_decision == "overwrite":
            row_index_to_overwrite = None
            for i, row in enumerate(sheet_data):
                if row and row[0] == date:
                    row_index_to_overwrite = i + 7
                    break

            if row_index_to_overwrite:
                range_to_update = f"{sheet_name}!A{row_index_to_overwrite}:{chr(65 + len(row_data) - 1)}{row_index_to_overwrite}"
                sheets_service.spreadsheets().values().update(
                    spreadsheetId=spreadsheet_id,
                    range=range_to_update,
                    valueInputOption="USER_ENTERED",
                    body={'values': [row_data]}
                ).execute()
                return_message = "Dữ liệu đã được ghi đè thành công!"
            else:
                return "Không tìm thấy dữ liệu của ngày này để ghi đè."
        else:
            # Nếu không ghi đè, thêm dữ liệu mới vào sheet
            # sheets_service.spreadsheets().values().append(
            #     spreadsheetId=spreadsheet_id,
            #     range=f"{sheet_name}!A7",
            #     valueInputOption="RAW",
            #     insertDataOption="INSERT_ROWS",
            #     body={'values': [row_data]}
            # ).execute()
            # return_message = "Dữ liệu đã được ghi thành công vào Google Sheets!"

            next_row = len(sheet_data) + 7
            range_to_update = f"{sheet_name}!A{next_row}:{chr(65 + len(row_data) - 1)}{next_row}"

            sheets_service.spreadsheets().values().update(
                    spreadsheetId=spreadsheet_id,
                    range=range_to_update,
                    valueInputOption="RAW",
                    body={'values': [row_data]}
            ).execute()

            return_message = "Dữ liệu đã được ghi thành công vào Google Sheets!"

        # Sắp xếp dữ liệu theo ngày (cột A)
        sheet_id_to_sort = get_sheet_id_by_name(spreadsheet_id, sanitize_sheet_name(sheet_name))
        if sheet_id_to_sort:
            requests = [{
                "sortRange": {
                    "range": {
                        "sheetId": sheet_id_to_sort,
                        "startRowIndex": 6,
                        "startColumnIndex": 0,
                    },
                    "sortSpecs": [{
                        "dimensionIndex": 0,
                        "sortOrder": "ASCENDING"
                    }]
                }
            }]
            body = {"requests": requests}
            sheets_service.spreadsheets().batchUpdate(
                spreadsheetId=spreadsheet_id, body=body
            ).execute()
            return_message += " Đã sắp xếp dữ liệu."

        return return_message

    except HttpError as err:
        print(f"Lỗi HttpError: {err}")
        return f"Lỗi: {err}"
    except Exception as e:
        print(f"Lỗi khác: {e}")
        return f"Lỗi: {e}"
def get_sheet_id_by_name(spreadsheet_id, sheet_name):
    """Lấy sheetId dựa vào tên sheet."""
    try:
        spreadsheet = app.sheets_service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        for sheet in spreadsheet['sheets']:
            if sheet['properties']['title'] == sheet_name:
                return sheet['properties']['sheetId']
        return None  # Trả về None nếu không tìm thấy
    except HttpError as err:
        print(f"Error getting sheet ID: {err}")
        return None

def check_existing_data(sheets_service, spreadsheet_id, sheet_name, date):
    """
    Kiểm tra xem dữ liệu của ngày tháng trong sheet đã tồn tại chưa.
    """
    sheet_name = sanitize_sheet_name(sheet_name)
    try:
        # Lấy dữ liệu từ Google Sheets, giả sử ngày tháng nằm ở cột A, từ hàng 7 trở đi.
        # Đảm bảo rằng tên sheet được thay đổi để không chứa ký tự '/'
        sheet_name = sheet_name.replace('/', '-')  # Hoặc thay bằng '_'
        
        # Kiểm tra phạm vi hợp lệ, sử dụng cột A hợp lệ
        range_ = f"{sheet_name}!A7:A"  # Kiểm tra cột A từ hàng 7 trở đi (cột ngày)
        
        sheet_data = sheets_service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=range_
        ).execute().get('values', [])

        # Lọc ra danh sách các ngày đã tồn tại trong sheet
        existing_dates = [pd.to_datetime(row[0], format="%d/%m/%Y").strftime("%d/%m/%Y") for row in sheet_data if row]
        return date in existing_dates  # Trả về True nếu ngày tháng đã tồn tại

    except Exception as e:
        print(f"Error checking existing data: {e}")
        return False  # Nếu có lỗi, coi như chưa có dữ liệu.



def create_sheet(spreadsheet_id, sheet_name):
    """
    Tạo một sheet mới trong Google Sheets
    """
    sheet_name = sanitize_sheet_name(sheet_name)
    requests = [{
        "addSheet": {
            "properties": {
                "title": sheet_name
            }
        }
    }]
    
    body = {
        'requests': requests
    }
    
    response = app.sheets_service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body=body
    ).execute()
    return response


def get_spreadsheet_sheets(spreadsheet_id):
    """
    Lấy danh sách các trang tính trong Google Sheets
    """
    result = app.sheets_service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    sheets = result.get('sheets', [])
    return [sheet['properties']['title'] for sheet in sheets]

'''' Xử lý data (Process Data)'''

def process_excel(file):
    """
    Đọc file Excel và trả về DataFrame
    """
    df = pd.read_excel(file, header=1, engine='openpyxl')
    print("Cột của DataFrame sau khi đọc Excel:", df.columns) # In ra danh sách cột
    df['Shop'] = df['Shop'].fillna(method='ffill').astype(str)
    # Xóa khoảng trắng thừa ở đầu/cuối
    df['Shop'] = df['Shop'].str.strip()

    # Loại bỏ các ký tự lạ (nếu có)
    df['Shop'] = df['Shop'].str.replace(r'[^\w\sÀÁÂÃÈÉÊÌÍÒÓÔÕÙÚÝàáâãèéêìíòóôõùúýăđĩũơưĂĐĨŨƠƯĂắằẳẵặĐđịọọ̃ụỵý]', '', regex=True)

    df.columns = df.columns.str.strip()
    print("Cột của DataFrame sau khi đọc Excel:", df.columns) # In ra danh sách cột
    return df

def get_shop_list(df):
    if 'Shop' not in df.columns:
        raise ValueError("'Shop' column is missing in the provided Excel file.")
    
    # Danh sách shop
    shop_list = df['Shop'].drop_duplicates().tolist()
    return shop_list

def dt_by_shop_nonVAT(df):
    if 'Doanh thu thực' not in df.columns:
        raise ValueError("'Doanh thu thực' column is missing.")
    if 'Shop' not in df.columns:
        raise ValueError("'Shop' column is missing.")
    
    df['Shop'] = df['Shop'].astype(str)

    # Tính doanh thu từng shop (Không VAT)
    dtt_shop = df.groupby('Shop')['Doanh thu thực'].sum().to_dict()
    return dtt_shop

def dt_by_shop_VAT(df):
    if 'DT' not in df.columns:
        raise ValueError("'DT' column is missing.")
    if 'Shop' not in df.columns:
        raise ValueError("'Shop' column is missing.")
    
    # Tính doanh thu từng shop (VAT)
    dt_VAT = df.groupby('Shop')['DT'].sum().to_dict()
    return dt_VAT

def clean_data(data):
    cleaned_data = []
    for row in data:
        cleaned_row = []
        for cell in row:
            if isinstance(cell, float) and math.isnan(cell):  # Thay thế NaN
                cleaned_row.append(0)  # Hoặc giá trị mặc định
            elif isinstance(cell, str):
                cleaned_row.append(cell.encode('unicode_escape').decode('utf-8'))  # Mã hóa ký tự đặc biệt
            else:
                cleaned_row.append(cell)  # Dữ liệu hợp lệ, giữ nguyên
        cleaned_data.append(cleaned_row)
    return cleaned_data
