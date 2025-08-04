# -*- coding: utf-8 -*-
import subprocess
import pyautogui
import pyperclip
import time
import random
import openpyxl

def paste(text):
    pyperclip.copy(text)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(0.5)

def select_score(select_img='select.png', down_presses=1, confidence=0.9):
    location = pyautogui.locateCenterOnScreen(select_img, confidence=confidence)
    if location:
        pyautogui.click(location)
        time.sleep(0.3)
        for _ in range(random.randint(3, 6)):
            pyautogui.press('down', presses=down_presses)
        pyautogui.press('enter')
        time.sleep(0.3)
    else:
        print("⚠️ 找不到下拉選單圖片")
        return False
    return True

# === 讀取 datas.xlsx ===
wb = openpyxl.load_workbook('datas.xlsx')
ws = wb.active

data_list = []
for row in ws.iter_rows(min_row=2, values_only=True):  # 跳過第一列標題
    email, name, phone = row
    data_list.append({
        'email': email,
        'name': name,
        'phone': str(phone)
    })

# ✅ 根據 Excel 內容重複填寫表單
for i, data in enumerate(data_list, start=1):
    print(f"\n🚀 開始第 {i} 筆填寫：{data['name']}")

    # === 開啟 Google 表單（無痕模式）===
    url = 'https://docs.google.com/forms/d/e/1FAIpQLSfI9iIXkKzK06LP2oFdel2O7L1FTDuLv9G1BG1yJqNUuXHZBg/viewform?usp=pp_url&pli=1'
    subprocess.Popen([
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        "--incognito", url
    ])
    
    print("🌐 表單已開啟，等待頁面載入...")
    time.sleep(6)

    # === 基本資料填寫 ===
    pyautogui.click(100, 200)  # 點進 Email 輸入框區域
    pyautogui.press('tab', presses=3)

    paste(data['email'])  # Email
    pyautogui.press('tab')
    paste(data['name'])   # 姓名
    pyautogui.press('tab')
    paste(data['phone'])  # 手機
    pyautogui.press('tab')

    time.sleep(0.5)
    pyautogui.press('pagedown', presses=2, interval=0.3)
    time.sleep(0.5)

    # === 下拉選單填答 ===
    for _ in range(4):
        if not select_score('select.png', down_presses=1, confidence=0.9):
            break

    # === 點擊提交 ===
    location = pyautogui.locateCenterOnScreen('submit.png', confidence=0.8)
    if location:
        pyautogui.click(location)
        print(f"✅ 第 {i} 筆資料已成功送出")
    else:
        print("⚠️ 找不到提交按鈕")

    time.sleep(5)  # 等待送出完成畫面穩定再繼續
