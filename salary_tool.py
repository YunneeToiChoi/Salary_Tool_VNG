import os
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
import re

def load_data(file_path, sheet_name):
    # Đọc file excel và lấy dữ liệu của sheet
    df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
    
    # Lấy thông tin tháng và năm từ các cột F, G, H, I tại dòng 1
    month_info = df.iloc[0, 5:9].values  # Cột F, G, H, I
    month_year = ' '.join([str(i) for i in month_info if pd.notna(i)]).strip()
    print(f"Tháng và năm: {month_year}")
    
    # Lấy toàn bộ dòng 8 để tìm bảng lương
    salary_row = df.iloc[7, :].values  # Lấy toàn bộ giá trị ở dòng 8
    salary_row_cleaned = [str(cell).strip() for cell in salary_row if pd.notna(cell) and str(cell).strip()]

    print("Các giá trị đọc được từ dòng 8:")
    print(salary_row_cleaned)  # In ra tất cả các giá trị không rỗng

    # Tìm tất cả các bảng lương có chứa chuỗi 'BẢNG LƯƠNG'
    salary_tables = [cell for cell in salary_row_cleaned if 'BẢNG LƯƠNG' in cell]

    if not salary_tables:
        print("Không tìm thấy bảng lương nào theo cú pháp 'BẢNG LƯƠNG'.")
    else:
        print(f"Các bảng lương tìm thấy: {salary_tables}")
    
    # Trả về dữ liệu từ dòng 8 trở đi với dòng 7 làm tiêu đề
    df_with_header = pd.read_excel(file_path, sheet_name=sheet_name, header=7)  # Dòng 7 là header
    return df_with_header, month_year, salary_tables

def clean_data(df):
    # Loại bỏ các cột không cần thiết
    df = df.dropna(how='all', axis=1)  # Loại bỏ các cột toàn NaN
    
    return df

def search_employee(data, keyword):
    # Chuyển cột 'HỌ TÊN' về kiểu chuỗi
    data['Họ tên NV'] = data['Họ tên NV'].astype(str)
    
    # Tìm kiếm nhân viên theo tên hoặc ID
    result = data[
        data['Họ tên NV'].str.contains(keyword, case=False, na=False) | 
        data['STT'].astype(str).str.contains(keyword)
    ]
    return result
def save_employee_data(employee_data, employee_name, month_year, salary_table, output_folder):
    # Tạo tên file xuất
    file_name = f'{employee_name}_{month_year}_{salary_table}.xlsx'
    output_path = os.path.join(output_folder, file_name)

    # Tạo workbook mới và viết dữ liệu vào file
    wb = Workbook()
    ws = wb.active
    ws.title = "Thông Tin Nhân Viên"
    
    # Ghi tiêu đề cột
    ws.append(list(employee_data.columns))

    # Ghi dữ liệu nhân viên
    for row in employee_data.itertuples(index=False, name=None):
        ws.append(row)

    # Định dạng tiêu đề
    for cell in ws[1]:
        cell.font = Font(bold=True, size=12, color="FFFFFF")
        cell.fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
        cell.alignment = Alignment(horizontal="center")

    # Định dạng các dòng dữ liệu
    for row in ws.iter_rows(min_row=2, min_col=1, max_col=len(employee_data.columns), max_row=len(employee_data)+1):
        for cell in row:
            cell.alignment = Alignment(horizontal="left", vertical="center")
            if isinstance(cell.value, (int, float)):
                cell.number_format = '#,##0'

    # Điều chỉnh chiều rộng cột
    min_width = 15
    for column in ws.columns:
        max_length = 0
        column_letter = column[0].column_letter
        for cell in column:
            try:
                if cell.value is not None:
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass
        adjusted_width = max(max_length + 2, min_width)
        ws.column_dimensions[column_letter].width = adjusted_width

    # Lưu file Excel
    wb.save(output_path)
    print(f"Thông tin nhân viên {employee_name} đã được lưu tại: {output_path}")

def main():
    excel_file = r'D:\Salary_Data\2024-_09_BangLuong_v01.xlsx'
    base_folder = r'D:\Salary_Data'
    # Đọc danh sách các sheet trong file Excel
    xls = pd.ExcelFile(excel_file)
    sheets = xls.sheet_names
    print("Danh sách bảng lương có sẵn:")
    for i, sheet in enumerate(sheets, 1):
        print(f"{i}. {sheet}")

    sheet_input = input("Chọn số thứ tự bảng lương: ").strip()
    if sheet_input.isdigit() and 1 <= int(sheet_input) <= len(sheets):
        sheet_name = sheets[int(sheet_input) - 1]
    else:
        print("Vui lòng chọn số hợp lệ.")
        return

    try:
        # Load dữ liệu và lấy thông tin tháng/năm và các bảng lương
        data, month_year, salary_tables = load_data(excel_file, sheet_name)
        cleaned_data = clean_data(data)

        if not salary_tables:
            print("Không tìm thấy bảng lương nào theo cú pháp 'BẢNG LƯƠNG'.")
            return
        
        print("Các bảng lương có sẵn:")
        for i, table in enumerate(salary_tables, 1):
            print(f"{i}. {table}")
        
        table_input = input("Chọn số thứ tự bảng lương: ").strip()
        if table_input.isdigit() and 1 <= int(table_input) <= len(salary_tables):
            salary_table = salary_tables[int(table_input) - 1]
        else:
            print("Vui lòng chọn số hợp lệ.")
            return

        # Nhập tên hoặc ID nhân viên
        keyword = input("Nhập tên hoặc ID nhân viên: ").strip()
        employee_data = search_employee(cleaned_data, keyword)

        if employee_data.empty:
            print(f"Không tìm thấy nhân viên với từ khóa: {keyword}")
        else:
            employee_name = employee_data.iloc[0]['HỌ TÊN']
            output_folder = os.path.join(base_folder, f'{month_year}_Ky_{salary_table}')
            if not os.path.exists(output_folder):
                os.makedirs(output_folder)

            save_employee_data(employee_data, employee_name, month_year, salary_table, output_folder)
    except Exception as e:
        print(f"Lỗi: {e}")

if __name__ == "__main__":
    main()
