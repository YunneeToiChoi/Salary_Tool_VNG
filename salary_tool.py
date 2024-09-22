import os
import pandas as pd

def load_data(file_path):
    # Đọc dữ liệu từ file Excel mà không chỉ định tiêu đề
    df = pd.read_excel(file_path, header=None)

    # Lấy tiêu đề từ hai dòng đầu tiên (dòng 5 và 6 trong Excel)
    headers = df.iloc[4:6].fillna('')
    multi_headers = [(str(h1) + ' ' + str(h2)).strip() for h1, h2 in zip(headers.iloc[0], headers.iloc[1])]
    df.columns = multi_headers

    # Bỏ hai dòng tiêu đề đã dùng
    df = df[6:].reset_index(drop=True)
    
    return df

def clean_data(df):
    # In ra dữ liệu trước khi hợp nhất
    print("Dữ liệu trước khi hợp nhất:")
    print(df.head(10))  # In ra 10 dòng đầu

    # Hợp nhất các cột bị trùng và loại bỏ cột Unnamed
    df.columns = df.columns.map(lambda x: x if 'Unnamed' not in x else None)  # Giữ lại cột không phải Unnamed
    df = df.loc[:, ~df.columns.isnull()]  # Loại bỏ các cột Unnamed
    df = df.loc[:, ~df.columns.duplicated(keep='first')]  # Giữ cột đầu tiên trong các cột trùng

    # Hợp nhất cột "CÁC KHOẢN CỘNG" và "CÁC KHOẢN TRỪ"
    cộng_cols = df.filter(like='CÁC KHOẢN CỘNG').columns
    trừ_cols = df.filter(like='CÁC KHOẢN TRỪ').columns

    # Tổng hợp các khoản cộng và các khoản trừ
    df['CÁC KHOẢN CỘNG'] = df[ cộng_cols ].sum(axis=1, numeric_only=True)
    df['CÁC KHOẢN TRỪ'] = df[ trừ_cols ].sum(axis=1, numeric_only=True)

    # In ra dữ liệu sau khi hợp nhất
    print("\nDữ liệu sau khi hợp nhất:")
    print(df.head(10))  # In ra 10 dòng đầu

    return df

def search_employee(data, keyword):
    # Tìm kiếm nhân viên theo tên hoặc ID
    result = data[
        data['HỌ TÊN'].str.contains(keyword, case=False, na=False) | 
        data['STT'].astype(str).str.contains(keyword)
    ]
    return result

def save_employee_data(employee_data, employee_name, month_folder):
    # Tạo file Excel riêng cho nhân viên
    file_name = f'{employee_name}.xlsx'
    output_path = os.path.join(month_folder, file_name)
    employee_data.to_excel(output_path, index=False)
    print(f'Đã lưu thông tin nhân viên {employee_name} tại: {output_path}')
    
    # In ra dữ liệu đã lưu
    print("Dữ liệu nhân viên đã lưu:")
    print(employee_data)

def main():
    # Đường dẫn file Excel và folder để lưu kết quả
    excel_file = r'D:\Salary_Data\Bản sao của TỔNG QUẢN CHI LƯƠNG - 2024.xlsx'  # Thay bằng đường dẫn thực tế của bạn
    base_folder = r'D:\Salary_Data'  # Folder gốc lưu các folder tháng lương

    # Tải dữ liệu
    data = load_data(excel_file)

    # Làm sạch dữ liệu
    cleaned_data = clean_data(data)

    # In ra các cột để kiểm tra
    print("Các cột trong DataFrame:")
    print(cleaned_data.columns)

    # Nhập từ khóa tìm kiếm
    keyword = input("Nhập tên hoặc ID nhân viên: ").strip()

    # Tìm kiếm nhân viên
    employee_data = search_employee(cleaned_data, keyword)

    if employee_data.empty:
        print(f"Không tìm thấy nhân viên với từ khóa: {keyword}")
    else:
        # Lấy tên nhân viên để đặt tên file
        employee_name = employee_data.iloc[0]['HỌ TÊN']

        # Tạo folder tháng lương nếu chưa tồn tại
        month_folder = os.path.join(base_folder, 'Tháng_Lương_X')
        if not os.path.exists(month_folder):
            os.makedirs(month_folder)

        # Lưu dữ liệu nhân viên
        save_employee_data(employee_data, employee_name, month_folder)

if __name__ == "__main__":
    main()
