import webview
import pandas as pd
import json
import os
import sys

# --- 預設設定 ---
DEFAULT_CONFIG = {
    "title": "幸運大抽獎",
    "subtitle": "得獎名單", 
    "refresh_rate": 5000,
    "scroll_speed": 1.5,
    "col_award": "獎項",
    "col_name": "姓名",
    "col_dept": "單位",
    "col_id": "工號"
}

EXCEL_FILENAME = '抽獎名單與設定.xlsx'

class Api:
    def get_data(self):
        if getattr(sys, 'frozen', False):
            app_path = os.path.dirname(sys.executable)
        else:
            app_path = os.path.dirname(os.path.abspath(__file__))
            
        file_path = os.path.join(app_path, EXCEL_FILENAME)

        if not os.path.exists(file_path):
            return json.dumps({"error": "找不到 Excel 檔案", "path": EXCEL_FILENAME})

        try:
            # 1. 讀取設定
            config = DEFAULT_CONFIG.copy()
            try:
                df_conf = pd.read_excel(file_path, sheet_name='系統設定')
                for _, row in df_conf.dropna().iterrows():
                    key = str(row[0]).strip()
                    val = row[1]
                    
                    if key == "活動標題": config["title"] = str(val)
                    elif key == "活動副標題": config["subtitle"] = str(val)
                    elif key == "滾動速度": config["scroll_speed"] = float(val)
                    elif key == "更新頻率": config["refresh_rate"] = int(val) * 1000
                    elif key == "欄位-獎項": config["col_award"] = str(val)
                    elif key == "欄位-姓名": config["col_name"] = str(val)
                    elif key == "欄位-單位": config["col_dept"] = str(val)
                    elif key == "欄位-工號": config["col_id"] = str(val)
            except:
                pass 

            col_award = config["col_award"]
            col_name = config["col_name"]
            col_dept = config["col_dept"]
            col_id = config["col_id"]

            # 2. 讀取名單 (強制工號轉字串)
            try:
                converters = {col_id: str}
                try:
                    df = pd.read_excel(file_path, sheet_name='得獎名單', converters=converters)
                except:
                    df = pd.read_excel(file_path, sheet_name=0, converters=converters)
            except Exception as e:
                return json.dumps({"error": f"讀取名單失敗: {str(e)}"})

            df.columns = df.columns.str.strip()

            if col_name not in df.columns or col_award not in df.columns:
                return json.dumps({"error": f"Excel 找不到欄位：[{col_name}] 或 [{col_award}]"})

            if col_id in df.columns:
                df = df.drop_duplicates(subset=[col_id], keep='first')
            
            df = df.dropna(subset=[col_name])
            df = df[df[col_name].astype(str).str.strip() != '']

            result = {}
            for _, row in df.iterrows():
                award = str(row[col_award]).strip()
                name = str(row[col_name]).strip()
                dept = str(row[col_dept]).strip() if col_dept in df.columns else ""
                
                raw_emp_id = row[col_id] if col_id in df.columns else ""
                emp_id = str(raw_emp_id).strip()
                if emp_id.lower() == 'nan': emp_id = ''
                if emp_id.endswith('.0'): emp_id = emp_id[:-2]

                if award not in result:
                    result[award] = []
                
                result[award].append({"name": name, "dept": dept, "empId": emp_id})

            return json.dumps({
                "status": "success", 
                "data": result, 
                "meta": {
                    "title": config["title"],
                    "subtitle": config["subtitle"],
                    "scroll_speed": config["scroll_speed"],
                    "refresh_rate": config["refresh_rate"]
                }
            })

        except Exception as e:
            return json.dumps({"error": f"讀取錯誤: {str(e)}"})
    
    def toggle_fullscreen(self):
        window = webview.windows[0]
        window.toggle_fullscreen()

# --- HTML/CSS/JS (全新 UI：可拖拉側邊欄 + Sticky Header + 滿版優化) ---
html_content = """
<!DOCTYPE html>
<html lang="zh-TW">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Lucky Draw System</title>
    <style>
        /* ==================== 
           全域變數與重置 
           ==================== */
        :root {
            --sidebar-bg-start: #4a0000;
            --sidebar-bg-end: #1a0505;
            --main-bg: #800000;
            --card-bg: #fffbf0;
            --gold-accent: #FFD700;
            --text-dark: #333;
            --resizer-width: 6px;
        }

        * { box-sizing: border-box; }

        body, html {
            margin: 0;
            padding: 0;
            height: 100%;
            overflow: hidden; /* 防止整頁捲動，只讓內容區捲動 */
            font-family: "Microsoft JhengHei", "Heiti TC", sans-serif;
            background-color: var(--main-bg);
            user-select: none; /* 避免拖拉時選取文字 */
        }

        /* ==================== 
           主要佈局容器 (Flexbox) 
           ==================== */
        .app-container {
            display: flex;
            width: 100vw;
            height: 100vh;
        }

        /* ==================== 
           左側側邊欄 
           ==================== */
        .sidebar {
            width: 300px; /* 初始寬度 */
            min-width: 200px;
            max-width: 600px;
            background: linear-gradient(180deg, var(--sidebar-bg-start) 0%, var(--sidebar-bg-end) 100%);
            color: var(--gold-accent);
            display: flex;
            flex-direction: column;
            padding: 20px;
            border-right: 1px solid #5a0000;
            position: relative;
            flex-shrink: 0; /* 防止被擠壓 */
            z-index: 200;
        }

        .sidebar-title {
            font-size: 2rem;
            font-weight: bold;
            margin-top: 20px;
            margin-bottom: 10px;
            line-height: 1.3;
            text-shadow: 0 2px 4px rgba(0,0,0,0.5);
        }

        .sidebar-subtitle {
            font-size: 1.2rem;
            color: rgba(255, 215, 0, 0.8);
            margin-bottom: 40px;
            padding-bottom: 20px;
            border-bottom: 1px solid rgba(255, 215, 0, 0.3);
        }

        .sidebar-info {
            margin-top: auto;
            font-size: 0.9rem;
            color: rgba(255,255,255,0.4);
            text-align: center;
        }

        /* ==================== 
           可拖拉分隔線 (Resizer) 
           ==================== */
        .resizer {
            width: var(--resizer-width);
            background: #2b0000;
            cursor: col-resize;
            flex-shrink: 0;
            transition: background 0.2s;
            position: relative;
            z-index: 300;
            box-shadow: 1px 0 0 rgba(255,255,255,0.1) inset;
        }

        /* 增加感應區域 */
        .resizer::after {
            content: ""; position: absolute; left: -5px; right: -5px; top: 0; bottom: 0; z-index: 1;
        }

        .resizer:hover, .resizer.resizing {
            background: var(--gold-accent);
            box-shadow: 0 0 10px var(--gold-accent);
        }

        /* ==================== 
           右側主要內容區 
           ==================== */
        .main-content {
            flex-grow: 1; /* 關鍵：填滿剩餘空間 */
            overflow-y: auto; /* 內容過長時捲動 */
            position: relative;
            background-color: #8B0000;
            background-image: repeating-linear-gradient(
                45deg,
                rgba(0,0,0,0.1),
                rgba(0,0,0,0.1) 10px,
                transparent 10px,
                transparent 20px
            );
            /* 隱藏捲軸但保留功能 (Chrome/Safari) */
            scrollbar-width: none; 
            -ms-overflow-style: none;
        }
        .main-content::-webkit-scrollbar { display: none; }

        /* 內容容器 */
        #content-wrapper {
            padding-bottom: 100px; /* 底部留白 */
        }

        /* ==================== 
           獎項區塊 & Sticky Header 
           ==================== */
        .prize-section {
            margin-bottom: 0px; 
            padding-bottom: 40px; /* 區塊間距 */
        }

        /* 標題吸附效果 */
        .prize-header {
            position: sticky;
            top: 0;
            z-index: 100; /* 確保在卡片上方 */
            background: linear-gradient(90deg, #a00000 0%, #600000 100%);
            border-bottom: 3px solid var(--gold-accent);
            box-shadow: 0 4px 12px rgba(0,0,0,0.4);
            padding: 15px 40px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            backdrop-filter: blur(5px);
        }

        .prize-header h2 {
            margin: 0;
            color: var(--gold-accent);
            font-size: 2rem;
            text-shadow: 2px 2px 4px rgba(0,0,0,0.6);
            letter-spacing: 1px;
        }

        .prize-count {
            background: #c0392b;
            color: white;
            padding: 5px 15px;
            border-radius: 20px;
            font-size: 1rem;
            border: 1px solid rgba(255,255,255,0.3);
        }

        /* ==================== 
           卡片網格系統 
           ==================== */
        .winner-grid {
            padding: 30px 40px;
            display: grid;
            /* 自動適應寬度，每張卡片最小 280px */
            grid-template-columns: repeat(auto-fill, minmax(280px, 1fr)); 
            gap: 20px;
        }

        .winner-card {
            background: var(--card-bg);
            border-radius: 12px;
            padding: 15px 20px;
            display: flex;
            align-items: center;
            box-shadow: 0 4px 6px rgba(0,0,0,0.2);
            border-left: 6px solid #c0392b;
            transition: transform 0.2s, box-shadow 0.2s;
            position: relative;
            overflow: hidden;
        }

        .winner-card:hover {
            transform: translateY(-3px);
            box-shadow: 0 8px 15px rgba(0,0,0,0.3);
        }

        .winner-info { flex-grow: 1; }
        
        .winner-id {
            font-size: 1.4rem;
            font-weight: bold;
            color: #333;
            margin-bottom: 4px;
        }

        .winner-dept { font-size: 0.9rem; color: #666; }

        .winner-number {
            font-size: 1.1rem;
            font-weight: bold;
            color: #fff;
            background: #333;
            padding: 4px 10px;
            border-radius: 6px;
            min-width: 50px;
            text-align: center;
        }

        /* 浮水印裝飾 */
        .winner-card::after {
            content: "LUCKY";
            position: absolute; right: -10px; bottom: -15px;
            font-size: 4rem; font-weight: bold;
            color: rgba(0,0,0,0.03);
            pointer-events: none;
            transform: rotate(-15deg);
        }

        /* ==================== 
           功能按鈕區 
           ==================== */
        #controls-area {
            position: fixed; bottom: 20px; right: 20px;
            z-index: 1000; text-align: right;
            display: flex; flex-direction: column; align-items: flex-end; gap: 5px;
        }
        
        .btn-fullscreen {
            background: #c0392b; border: 1px solid white; color: white;
            padding: 8px 15px; border-radius: 30px; cursor: pointer;
            font-size: 0.9rem; box-shadow: 0 4px 10px rgba(0,0,0,0.2);
            font-weight: bold; transition: all 0.2s;
        }
        .btn-fullscreen:hover { transform: scale(1.05); background: #a30000; }
        
        #status-bar { font-size: 11px; color: rgba(255,255,255,0.7); margin-bottom: 5px; text-shadow: 0 1px 2px #000;}

        /*錯誤訊息*/
        .error-msg { 
            color: #fff; font-size: 1.5rem; text-align: center; margin-top: 100px; 
            background: rgba(0,0,0,0.5); padding: 20px; border-radius: 10px;
        }

    </style>
</head>
<body>

    <div class="app-container">
        
        <aside class="sidebar" id="sidebar">
            <div id="main-title" class="sidebar-title">載入中...</div>
            <div id="sub-title" class="sidebar-subtitle"></div>
            
            <div class="sidebar-info">
                Designed by PedaleOn<br>
                <span style="font-size:0.8em; opacity:0.7">拖拉邊界可調整寬度 ></span>
            </div>
        </aside>

        <div class="resizer" id="resizer"></div>

        <main class="main-content" id="main-scroll-area">
            <div id="content-wrapper">
                </div>
        </main>
        
        <div id="controls-area">
            <div id="status-bar"></div>
            <button class="btn-fullscreen" onclick="callFullScreen()">⛶ 全螢幕 (F)</button>
        </div>

    </div>

    <script>
        // 設定變數
        let refreshRate = 5000;
        let scrollSpeed = 1.5;
        let isScrolling = true;
        let scrollDirection = 1;
        let currentScrollPos = 0;
        let lastDataHash = "";
        let timer = null;
        let scrollFrame = null;

        // 監聽 PyWebview 準備就緒
        window.addEventListener('pywebviewready', function() {
            updateData();
            initResizer(); // 啟動拖拉功能
        });

        // --- 資料更新邏輯 ---
        function updateData() {
            pywebview.api.get_data().then(function(response) {
                const res = JSON.parse(response);
                
                if (res.error) {
                    document.getElementById('content-wrapper').innerHTML = `<div class='error-msg'>${res.error}</div>`;
                    setTimeout(updateData, 3000);
                    return;
                }

                if (res.status === 'success') {
                    if(res.meta) {
                        document.getElementById('main-title').innerText = res.meta.title;
                        document.getElementById('sub-title').innerText = res.meta.subtitle || "";
                        scrollSpeed = res.meta.scroll_speed;
                        let newRate = res.meta.refresh_rate || 5000;
                        if (newRate !== refreshRate) refreshRate = newRate;
                    }
                    renderUI(res.data);
                    
                    const now = new Date();
                    document.getElementById('status-bar').innerText = "最後更新: " + now.getHours().toString().padStart(2,'0') + ":" + now.getMinutes().toString().padStart(2,'0') + ":" + now.getSeconds().toString().padStart(2,'0');
                    
                    if (timer) clearTimeout(timer);
                    timer = setTimeout(updateData, refreshRate);
                }
            });
        }

        // --- 渲染介面 (核心 UI 生成) ---
        function renderUI(groupedData) {
            const currentHash = JSON.stringify(groupedData);
            if (currentHash === lastDataHash) return; // 資料沒變就不重繪
            lastDataHash = currentHash;
            
            const wrapper = document.getElementById('content-wrapper');
            let html = "";
            const awards = Object.keys(groupedData);

            if (awards.length === 0) {
                wrapper.innerHTML = "<div class='error-msg'>目前沒有得獎名單，請確認 Excel。</div>";
                return;
            }

            // 遍歷每個獎項，生成 Section
            awards.forEach(award => {
                const list = groupedData[award];
                
                // 1. 獎項標題 (Sticky Header)
                html += `
                <section class="prize-section">
                    <div class="prize-header">
                        <h2>${award}</h2>
                        <span class="prize-count">共 ${list.length} 位</span>
                    </div>
                    <div class="winner-grid">`;
                
                // 2. 得獎者卡片
                list.forEach(p => {
                    html += `
                        <div class="winner-card">
                            <div class="winner-info">
                                <div class="winner-id">${p.name}</div>
                                <div class="winner-dept">${p.dept}</div>
                            </div>
                            <div class="winner-number">${p.empId}</div>
                        </div>`;
                });

                html += `</div></section>`;
            });
            
            wrapper.innerHTML = html;
            
            // 資料更新後重置捲動狀態（如果需要）
            setTimeout(() => { 
                checkAndStartScroll(); 
            }, 100);
        }

        // --- 捲動邏輯 (針對 .main-content) ---
        function checkAndStartScroll() {
            const container = document.getElementById('main-scroll-area'); // 注意這裡換成新的 ID
            if (scrollFrame) cancelAnimationFrame(scrollFrame);
            
            // 如果內容比視窗高，才需要捲動
            if (container.scrollHeight > container.clientHeight) {
                isScrolling = true;
                currentScrollPos = container.scrollTop;
                scrollLoop();
            }
        }

        function scrollLoop() {
            const container = document.getElementById('main-scroll-area');
            
            if (isScrolling) {
                currentScrollPos += (scrollSpeed * scrollDirection);
                container.scrollTop = currentScrollPos;
                
                // 到底部
                if (scrollDirection === 1 && (currentScrollPos + container.clientHeight >= container.scrollHeight - 2)) {
                    isScrolling = false;
                    setTimeout(() => {
                        scrollDirection = -1; // 往回捲
                        isScrolling = true;
                        scrollFrame = requestAnimationFrame(scrollLoop);
                    }, 3000); // 到底停留 3 秒
                    return;
                }

                // 到頂部
                if (scrollDirection === -1 && currentScrollPos <= 0) {
                    currentScrollPos = 0;
                    container.scrollTop = 0;
                    isScrolling = false;
                    setTimeout(() => {
                        scrollDirection = 1; // 往下捲
                        isScrolling = true;
                        scrollFrame = requestAnimationFrame(scrollLoop);
                    }, 3000); // 到頂停留 3 秒
                    return;
                }
            }
            scrollFrame = requestAnimationFrame(scrollLoop);
        }

        // --- 側邊欄拖拉邏輯 ---
        function initResizer() {
            const sidebar = document.getElementById('sidebar');
            const resizer = document.getElementById('resizer');
            let isResizing = false;

            resizer.addEventListener('mousedown', (e) => {
                isResizing = true;
                resizer.classList.add('resizing');
                document.body.style.cursor = 'col-resize';
            });

            document.addEventListener('mousemove', (e) => {
                if (!isResizing) return;
                let newWidth = e.clientX;
                // 設定最小與最大寬度
                if (newWidth >= 200 && newWidth <= 600) {
                    sidebar.style.width = `${newWidth}px`;
                }
            });

            document.addEventListener('mouseup', () => {
                if (isResizing) {
                    isResizing = false;
                    resizer.classList.remove('resizing');
                    document.body.style.cursor = 'default';
                }
            });
        }

        function callFullScreen() {
            pywebview.api.toggle_fullscreen();
        }

        // 鍵盤控制
        document.addEventListener('keydown', (e) => { 
            if(e.key === ' ') { // 空白鍵暫停/開始
                isScrolling = !isScrolling; 
                if(isScrolling) scrollLoop();
            }
            if(e.key === 'f' || e.key === 'F') {
                callFullScreen();
            }
        });
    </script>
</body>
</html>
"""

if __name__ == '__main__':
    api = Api()
    # 建議使用較大的初始視窗
    window = webview.create_window(
        'Lucky Draw Board', 
        html=html_content, 
        js_api=api,
        width=1280, 
        height=800,
        background_color='#800000'
    )
    webview.start(debug=False)