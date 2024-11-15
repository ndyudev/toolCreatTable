import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill

def get_month_data():
    """
    Yêu cầu người dùng nhập thông tin cho từng tháng.
    Returns:
        dict: Thông tin tháng bao gồm tên tháng, doanh số, màu sắc.
    """
    month = input("Nhập tên tháng (vd: January, February,...): ")
    revenue = float(input(f"Nhập doanh số cho {month}: "))
    color = input(f"Nhập mã màu HEX cho {month} (vd: FF5733 cho màu đỏ): ")
    return {"Month": month, "Revenue": revenue, "Color": color}

def add_data():
    """
    Thêm nhiều tháng vào bảng cho đến khi người dùng không muốn nhập thêm.
    Returns:
        list: Danh sách dữ liệu các tháng.
    """
    data = []
    while True:
        month_data = get_month_data()
        data.append(month_data)
        cont = input("Bạn có muốn thêm tháng khác không? (y/n): ").strip().lower()
        if cont != 'y':
            break
    return data

def save_to_excel(data, file_path):
    """
    Lưu dữ liệu vào file Excel và áp dụng màu nền cho từng tháng.
    Args:
        data (list): Danh sách dữ liệu các tháng.
        file_path (str): Đường dẫn file Excel để lưu.
    """
    # Tạo DataFrame từ dữ liệu
    df = pd.DataFrame(data)

    # Tạo một Workbook và thêm DataFrame vào Excel
    with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Monthly Data')
        workbook = writer.book
        worksheet = writer.sheets['Monthly Data']

        # Áp dụng màu nền cho từng dòng dựa trên mã màu
        for index, row in df.iterrows():
            fill_color = PatternFill(start_color=row["Color"], end_color=row["Color"], fill_type="solid")
            worksheet[f'A{index + 2}'].fill = fill_color  # Cột tên tháng
            worksheet[f'B{index + 2}'].fill = fill_color  # Cột doanh số

    print(f"Dữ liệu đã được lưu vào {file_path} thành công!")

def main(): 
    print("Chào mừng bạn đến với tool tạo bảng!")
    file_path = "C:/Users/Chau Nhat Duy/Documents/DataAnalyst/toolCreatTable/output2.xlsx"
    data = add_data()
    save_to_excel(data, file_path)

if __name__ == "__main__":
    main()
