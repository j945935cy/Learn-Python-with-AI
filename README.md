# 自動填寫 Google 表單工具（Auto Google Form Filler）

這是一個使用 Python 製作的自動化工具，能夠讀取 Excel 檔案中的使用者資料，開啟 Google 表單，並自動模擬使用者填寫與提交的動作。非常適合用於表單測試、批次填答、回饋收集自動化等場景。

---

## 🔧 功能特色

- 自動從 Excel 檔案讀取使用者資料（Email、姓名、手機）
- 開啟 Google 表單並模擬填寫與提交
- 自動點選下拉選單滿意度評分（亂數給分）
- 支援無痕瀏覽模式，避免重複填寫影響
- 避免中文輸入法干擾，使用貼上方式輸入
- 圖像辨識下拉選單與提交按鈕（依照圖片位置操作）

---

## 📁 專案結構

```

form-autofill-bot/
├── main.py                 # 主程式
├── datas.xlsx              # Excel 資料檔（Email, 姓名, 手機）
├── select.png              # 下拉選單圖示（供辨識）
├── submit.png              # 提交按鈕圖示（供辨識）
├── requirements.txt        # 套件需求清單
└── README.md               # 專案說明文件

````

---

## ▶️ 執行方式

### 1. 安裝套件

建議使用虛擬環境後執行：

```bash
pip install -r requirements.txt
````

套件包含：

* `pyautogui`
* `pyperclip`
* `openpyxl`

> 若使用圖片辨識功能需額外安裝：

```bash
pip install opencv-python
```

---

### 2. 執行程式

請確認：

* Chrome 安裝路徑正確（`main.py` 中可自訂）
* `datas.xlsx` 檔案存在，欄位為 Email、姓名、手機
* `select.png`, `submit.png` 為你表單的截圖圖檔（解析度需相符）

```bash
python main.py
```

程式將依序自動打開表單、貼上資料、點選分數、並提交結果。

---

## 🧪 表單示意圖（可自行補上）

* 表單網址建議使用「允許多次填寫」模式
* 題目包含：Email、姓名、手機、四題下拉式滿意度評分題

---

## 🧠 延伸挑戰

* 自動記錄每筆填寫是否成功，回寫至 Excel
* 將整個流程封裝成 GUI 工具，供非技術者使用
* 使用 Selenium 改寫表單操作流程，更穩定不依賴圖片
* 設計可選「範圍給分」或「指定填答模式」選項

---

## 🙋‍♂️ 作者與貢獻

由 \[ j945935cy ] 製作。
如有建議與改進，歡迎提出 Pull Request 或 issue。

---

## 📄 授權 License



