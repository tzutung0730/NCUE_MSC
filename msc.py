from selenium import webdriver
from selenium.webdriver.common.by import By
import keyboard
import openpyxl
import win32api
import win32con

# 初始化 Chrome WebDriver
driver = webdriver.Chrome()

# 設定隱性等待（最長等待 10 秒）
driver.implicitly_wait(10)

# 讀取 Excel 檔案
workbook = openpyxl.load_workbook('test.xlsx')
sheet = workbook.active
# 獲取螢幕寬度和高度
screen_width = win32api.GetSystemMetrics(win32con.SM_CXSCREEN)
screen_height = win32api.GetSystemMetrics(win32con.SM_CYSCREEN)

# 設定瀏覽器視窗大小為螢幕寬度的一半，並將其置於左半邊
window_width = screen_width // 2
window_height = screen_height
window_position_x = 0
window_position_y = 0
driver.set_window_position(window_position_x, window_position_y)
driver.set_window_size(window_width, window_height)

# 獲取帳號和密碼資訊
for row_index, row in enumerate(sheet.iter_rows(values_only=True), start=1):
    username = row[0]
    password = row[1]

    # 前往指定的網頁
    driver.get('https://youthdream.phdf.org.tw/member/login')
    driver.find_element(By.NAME, 'email').send_keys(username)
    driver.find_element(By.NAME, 'password').send_keys(password)
    # 提示目前正在輸入驗證碼的行數
    print(f"正在處理第 {row_index} 行的驗證碼")
    
    # 輸入驗證碼
    driver.find_element(By.NAME, 'captcha').click()
    keyboard.wait('enter')  

    driver.get('https://youthdream.phdf.org.tw/project/show?page=5')


    driver.get('https://youthdream.phdf.org.tw/member/profile')


# 關閉 WebDriver
driver.quit()
