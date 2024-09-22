import os
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill

def load_data(file_path, sheet_name):
    df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
    headers = df.iloc[4:6].fillna('')
    multi_headers = [(str(h1) + ' ' + str(h2)).strip() for h1, h2 in zip(headers.iloc[0], headers.iloc[1])]
    df.columns = multi_headers
    df = df[6:].reset_index(drop=True)
    return df

def clean_data(df):
    print("Dữ liệu trước khi hợp nhất:")
    print(df.head(10))

    # Loại bỏ các cột không có tên và cột "Unnamed"
    df.columns = df.columns.str.strip()  # Loại bỏ khoảng trắng
    df = df.loc[:, ~df.columns.str.contains('^Unnamed') & (df.columns != '')]

    # Kiểm tra cột bị trùng
    duplicated_columns = df.columns[df.columns.duplicated(keep=False)].unique()
    if duplicated_columns.size > 0:
        print("Các cột bị trùng:", duplicated_columns)

        # Tạo DataFrame cho các cột trùng
        df_duplicates = df[duplicated_columns]
        duplicate_file_path = 'duplicates.xlsx'
        df_duplicates.to_excel(duplicate_file_path, index=False)
        print(f"Dữ liệu cột trùng đã được lưu tại: {duplicate_file_path}")

        # Bỏ các cột trùng trong dataframe chính
        df = df.loc[:, ~df.columns.duplicated(keep='first')]

    cộng_cols = df.filter(like='CÁC KHOẢN CỘNG').columns
    trừ_cols = df.filter(like='CÁC KHOẢN TRỪ').columns

    df['CÁC KHOẢN CỘNG'] = df[ cộng_cols ].sum(axis=1, numeric_only=True)
    df['CÁC KHOẢN TRỪ'] = df[ trừ_cols ].sum(axis=1, numeric_only=True)

    print("\nDữ liệu sau khi hợp nhất:")
    print(df.head(10))

    return df

def search_employee(data, keyword):
    print(data['HỌ TÊN'].dtypes)  # In ra kiểu dữ liệu của cột 'HỌ TÊN'

    data['HỌ TÊN'] = data['HỌ TÊN'].astype(str)

    result = data[
        data['HỌ TÊN'].str.contains(keyword, case=False, na=False) | 
        data['STT'].astype(str).str.contains(keyword)
    ]
    return result

def save_employee_data(employee_data, employee_name, month_folder):
    file_name = f'{employee_name}.xlsx'
    output_path = os.path.join(month_folder, file_name)

    wb = Workbook()
    ws = wb.active
    ws.title = "Thông Tin Nhân Viên"

    ws.append(list(employee_data.columns))

    for row in employee_data.itertuples(index=False, name=None):
        ws.append(row)

    for cell in ws[1]:
        cell.font = Font(bold=True, size=12, color="FFFFFF")
        cell.fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
        cell.alignment = Alignment(horizontal="center")

    for row in ws.iter_rows(min_row=2, min_col=1, max_col=len(employee_data.columns), max_row=len(employee_data)+1):
        for cell in row:
            cell.alignment = Alignment(horizontal="left", vertical="center")
            if isinstance(cell.value, (int, float)):
                cell.number_format = '#,##0'

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

    wb.save(output_path)
    print(employee_data)

def main():
    excel_file = r'D:\Salary_Data\TỔNG QUẢN CHI LƯƠNG - 2024.xlsx'
    base_folder = r'D:\Salary_Data'
    employee_name = None

    sheet_input = input("Nhập số từ 1 đến 12 để tìm sheet T{input}-2024: ").strip()
    
    if sheet_input.isdigit() and 1 <= int(sheet_input) <= 12:
        sheet_name = f"T{sheet_input}-2024"
    else:
        print("Vui lòng nhập số từ 1 đến 12.")
        return

    try:
        data = load_data(excel_file, sheet_name)
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
    except ValueError as e:
        print(f"Lỗi: {e}")

    if employee_name is not None:
        print(f'Đã lưu thông tin nhân viên {employee_name} tại: {month_folder}')
    else:
        print("Không có thông tin nhân viên để lưu.")

if __name__ == "__main__":
    main()
