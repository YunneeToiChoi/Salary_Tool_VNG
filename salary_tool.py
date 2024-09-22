import os
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill

def load_data(file_path):
    df = pd.read_excel(file_path, header=None)
    headers = df.iloc[4:6].fillna('')
    multi_headers = [(str(h1) + ' ' + str(h2)).strip() for h1, h2 in zip(headers.iloc[0], headers.iloc[1])]
    df.columns = multi_headers
    df = df[6:].reset_index(drop=True)
    return df

def clean_data(df):
    print("Dữ liệu trước khi hợp nhất:")
    print(df.head(10))

    df.columns = df.columns.map(lambda x: x if 'Unnamed' not in x else None)
    df = df.loc[:, ~df.columns.isnull()]
    df = df.loc[:, ~df.columns.duplicated(keep='first')]

    cộng_cols = df.filter(like='CÁC KHOẢN CỘNG').columns
    trừ_cols = df.filter(like='CÁC KHOẢN TRỪ').columns

    df['CÁC KHOẢN CỘNG'] = df[ cộng_cols ].sum(axis=1, numeric_only=True)
    df['CÁC KHOẢN TRỪ'] = df[ trừ_cols ].sum(axis=1, numeric_only=True)

    print("\nDữ liệu sau khi hợp nhất:")
    print(df.head(10))

    return df

def search_employee(data, keyword):
    result = data[
        data['HỌ TÊN'].str.contains(keyword, case=False, na=False) | 
        data['STT'].astype(str).str.contains(keyword)
    ]
    return result
def save_employee_data(employee_data, employee_name, month_folder):
    file_name = f'{employee_name}.xlsx'
    output_path = os.path.join(month_folder, file_name)

    # Tạo Workbook và Worksheet
    wb = Workbook()
    ws = wb.active
    ws.title = "Thông Tin Nhân Viên"

    # Ghi tiêu đề
    ws.append(list(employee_data.columns))

    # Ghi dữ liệu
    for row in employee_data.itertuples(index=False, name=None):
        ws.append(row)

    # Định dạng tiêu đề
    for cell in ws[1]:
        cell.font = Font(bold=True, size=12, color="FFFFFF")
        cell.fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
        cell.alignment = Alignment(horizontal="center")

    # Định dạng dữ liệu
    for row in ws.iter_rows(min_row=2, min_col=1, max_col=len(employee_data.columns), max_row=len(employee_data)+1):
        for cell in row:
            cell.alignment = Alignment(horizontal="left", vertical="center")
            if isinstance(cell.value, (int, float)):
                cell.number_format = '#,##0'  # Định dạng số với dấu phẩy

    # Tự động giãn cột theo nội dung và độ dài tối thiểu
    min_width = 15  # Độ dài tối thiểu của cột
    for column in ws.columns:
        max_length = 0
        column_letter = column[0].column_letter  # Lấy ký tự cột (A, B, C, ...)
        for cell in column:
            try:
                if cell.value is not None:  # Chỉ tính chiều dài cho các ô có giá trị
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass
        adjusted_width = max(max_length + 2, min_width)  # Giãn cột với chiều dài tối thiểu
        ws.column_dimensions[column_letter].width = adjusted_width

    # Lưu file
    wb.save(output_path)
    print(f'Đã lưu thông tin nhân viên {employee_name} tại: {output_path}')
    
    print("Dữ liệu nhân viên đã lưu:")
    print(employee_data)

    file_name = f'{employee_name}.xlsx'
    output_path = os.path.join(month_folder, file_name)

    # Tạo Workbook và Worksheet
    wb = Workbook()
    ws = wb.active
    ws.title = "Thông Tin Nhân Viên"

    # Ghi tiêu đề
    ws.append(list(employee_data.columns))

    # Ghi dữ liệu
    for row in employee_data.itertuples(index=False, name=None):
        ws.append(row)

    # Định dạng tiêu đề
    for cell in ws[1]:
        cell.font = Font(bold=True, size=12, color="FFFFFF")
        cell.fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
        cell.alignment = Alignment(horizontal="center")

    # Định dạng dữ liệu
    for row in ws.iter_rows(min_row=2, min_col=1, max_col=len(employee_data.columns), max_row=len(employee_data)+1):
        for cell in row:
            cell.alignment = Alignment(horizontal="left", vertical="center")
            if isinstance(cell.value, (int, float)):
                cell.number_format = '#,##0'  # Định dạng số với dấu phẩy

    # Tự động giãn cột theo nội dung
    for column in ws.columns:
        max_length = 0
        column_letter = column[0].column_letter  # Lấy ký tự cột (A, B, C, ...)
        for cell in column:
            try:
                if cell.value is not None:  # Chỉ tính chiều dài cho các ô có giá trị
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass
        adjusted_width = (max_length + 2)  # Cộng thêm không gian
        ws.column_dimensions[column_letter].width = adjusted_width

    # Lưu file
    wb.save(output_path)
    print(f'Đã lưu thông tin nhân viên {employee_name} tại: {output_path}')
    
    print("Dữ liệu nhân viên đã lưu:")
    print(employee_data)

def main():
    excel_file = r'D:\Salary_Data\Bản sao của TỔNG QUẢN CHI LƯƠNG - 2024.xlsx'
    base_folder = r'D:\Salary_Data'

    data = load_data(excel_file)
    cleaned_data = clean_data(data)

    print("Các cột trong DataFrame:")
    print(cleaned_data.columns)

    keyword = input("Nhập tên hoặc ID nhân viên: ").strip()
    employee_data = search_employee(cleaned_data, keyword)

    if employee_data.empty:
        print(f"Không tìm thấy nhân viên với từ khóa: {keyword}")
    else:
        employee_name = employee_data.iloc[0]['HỌ TÊN']
        month_folder = os.path.join(base_folder, 'Tháng_Lương_X')
        if not os.path.exists(month_folder):
            os.makedirs(month_folder)

        save_employee_data(employee_data, employee_name, month_folder)

if __name__ == "__main__":
    main()
