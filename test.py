import openpyxl
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time

# Đọc file Excel
file_path = 'C:/Users/Admin/Desktop/Book1.xlsx'  # Thay bằng đường dẫn tới file Excel của bạn
workbook = openpyxl.load_workbook(file_path)
sheet = workbook.active

# Khởi tạo trình điều khiển cho trình duyệt (ví dụ sử dụng Chrome)
driver = webdriver.Chrome(executable_path='C:/Users/Admin/Downloads/chromedriver-win64/chromedriver-win64/chromedriver.exe')  # Thay bằng đường dẫn tới chromedriver của bạn

# Mở trang web
driver.get("https://www.digikey.com/")

# Lặp qua từng ID và thực hiện tìm kiếm
for row in sheet.iter_rows(min_row=2, min_col=1, max_col=1, values_only=True):
    id = row[0]
    
    # Tìm hộp tìm kiếm và nhập ID
    search_box = driver.find_element_by_name("keywords")  # Tìm theo name
    search_box.clear()
    search_box.send_keys(id)
    search_box.send_keys(Keys.RETURN)

    # Đợi một chút để trang web tải kết quả tìm kiếm
    time.sleep(5)

    # Ở đây bạn có thể thêm mã để trích xuất thông tin từ trang kết quả nếu cần
    # Lấy thông tin sản phẩm từ kết quả tìm kiếm
    products = driver.find_elements_by_css_selector('.searchResultsTable tbody tr')
    
    for product in products:
        try:
            product_name = product.find_element_by_css_selector('.productDetail a').text
            product_price = product.find_element_by_css_selector('.productPrice').text
            
            print(f"Product: {product_name}, Price: {product_price}")
        except:
            continue

# Đóng trình duyệt
driver.quit()
