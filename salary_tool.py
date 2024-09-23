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
    
    # Lấy toàn bộ dòng 9 làm tiêu đề (header) cho dataframe
    df.columns = df.iloc[8]
    
    # Loại bỏ các dòng đầu tiên trước dữ liệu thực (sau dòng 9)
    df = df.drop(index=range(9))

    # In ra các tên cột đã đọc được
    print("Tên cột đã nhận dạng:")
    print(df.columns)
    
    return df, month_year

def clean_data(df):
    # Loại bỏ các cột không cần thiết
    df = df.dropna(how='all', axis=1)  # Loại bỏ các cột toàn NaN
    
    # In ra dữ liệu đã làm sạch
    print("Dữ liệu sau khi làm sạch:")
    print(df.head(10))  # In ra 10 dòng đầu tiên để kiểm tra

    return df

def search_employee(data, keyword):
    # Chúng ta có thể cần làm sạch dữ liệu trước khi tìm kiếm
    name_column = 'Họ tên NV'
    
    try:
        # Tìm kiếm nhân viên theo tên hoặc ID
        result = data[data[name_column].str.contains(keyword, case=False, na=False)]
        
        # Kiểm tra nếu tìm thấy nhân viên
        if not result.empty:
            print("Thông tin nhân viên tìm thấy:")
            print(result)
        else:
            print("Không tìm thấy nhân viên nào.")
        
    except KeyError as e:
        print(f"Lỗi: Không tìm thấy cột '{name_column}' trong dữ liệu.")
        return None
    
    return result

def save_employee_data(employee_data, employee_name, month_year, output_folder):
    # Tạo tên file xuất
    sanitized_employee_name = re.sub(r'[\\/:"*?<>|]+', "_", employee_name)
    file_name = f'{sanitized_employee_name}_{month_year}.xlsx'
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
        # Load dữ liệu và lấy thông tin tháng/năm
        data, month_year = load_data(excel_file, sheet_name)
        cleaned_data = clean_data(data)

        # Nhập tên hoặc ID nhân viên
        keyword = input("Nhập tên hoặc ID nhân viên: ").strip()
        employee_data = search_employee(cleaned_data, keyword)

        if employee_data is None or employee_data.empty:
            print(f"Không tìm thấy nhân viên với từ khóa: {keyword}")
        else:
            employee_name = employee_data.iloc[0]['Họ tên NV']
            output_folder = os.path.join(base_folder, month_year)
            if not os.path.exists(output_folder):
                os.makedirs(output_folder)

            save_employee_data(employee_data, employee_name, month_year, output_folder)
    except Exception as e:
        print(f"Lỗi: {e}")

if __name__ == "__main__":
    main()
