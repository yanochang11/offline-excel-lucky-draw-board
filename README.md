# 🏆 尾牙春酒抽獎名單自動輪播看板 (Offline Excel Lucky Draw Board)

> **專為台灣福委會、行政人員設計的「零門檻」抽獎公告系統。不需連網、資料不外流，用 Excel 就能控制大螢幕！**

## 💡 為什麼開發這個工具？ (Why)
看到每年尾牙，行政同仁為了「如何在大螢幕上公平、漂亮地呈現得獎名單」而頭痛。市面上的線上抽獎程式需要連網（擔心個資外洩），PPT 手動播放又太累。

因此，我開發了這款 **「讀取 Excel 就能跑」** 的視窗程式。它結合了 Python 的強大與網頁的精美介面，只需一台筆電，就能搞定千人尾牙的投影需求。

這套軟體完全開源 (Open Source)，如果您對 .exe 檔有資安疑慮，歡迎直接到 GitHub 下載原始碼 (Source Code)，自行使用 Python 執行，保證乾淨無毒。

**由 [PedaleOn 騎乘不止](https://pedaleon.com/?utm_source=luckdraw&utm_medium=readme) 開發與維護。**

---

## ✨ 核心特色 (Features)
* **🔒 絕對資安**：完全**離線執行 (Offline)**，不需要連接網際網路，員工個資不出公司電腦。
* **📊 Excel 管理**：不需懂程式碼！所有名單、標題、速度設定，全部在 Excel 裡面修改。
* **⚡ 即時更新**：後台 Excel 只要輸入名字並存檔 (Ctrl+S)，前台投影幕 **5 秒內自動更新**，無縫接軌。
* **📜 智慧滾動**：支援「黏性標題 (Sticky Header)」，名單再長都能清楚知道是哪個獎項。
* **🎨 新春風格**：內建喜氣紅金配色 UI，適合尾牙、春酒、各類摸彩活動。

---

## 🚀 快速開始 (給一般使用者)
如果您不想看程式碼，只想直接使用軟體：

1.  前往 [**Releases 頁面**](https://github.com/yanochang11/offline-excel-lucky-draw-board/releases/tag/V1.0.0) 下載最新的懶人包 (ZIP)。
2.  解壓縮後，會看到 `LuckyDraw.exe` 和 `抽獎名單與設定.xlsx`。
3.  打開 Excel，填入您的獎項與名單，並設定標題。
4.  雙擊 `LuckyDraw.exe` 即可開始投影！
5.  **請注意是否被防毒軟體阻擋**，這套軟體完全開源 (Open Source)，如果您對 .exe 檔有資安疑慮，歡迎直接到 GitHub 下載原始碼 (Source Code)，自行使用 Python 執行，保證乾淨無毒。

---

## 🛠️ 開發者指南 (給工程師)
如果您想自行修改原始碼或貢獻功能，請參考以下步驟：

### 環境需求
* Python 3.8+
* 推薦使用虛擬環境 (venv)

### 安裝依賴
```bash
pip install pywebview pandas openpyxl pyinstaller
