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

# --- HTML/CSS/JS (優化版：強化獎項一體性) ---
html_content = """
<!DOCTYPE html>
<html lang="zh-TW">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Lucky Draw System</title>
    <style>
        /* ==================== 
           全域變數 
           ==================== */
        :root {
            --sidebar-bg: #8B0000;
            --main-bg: #800000;
            --section-bg: rgba(0, 0, 0, 0.2); /* 獎項區塊背景 */
            --card-bg: #fffbf0;
            --gold-accent: #FFD700;
            --border-gold: #b8860b;
            --text-dark: #333;
            --resizer-width: 5px;
        }

        * { box-sizing: border-box; }

        body, html {
            margin: 0; padding: 0;
            height: 100%;
            overflow: hidden; 
            font-family: "Microsoft JhengHei", "Heiti TC", sans-serif;
            background-color: var(--main-bg);
            user-select: none;
        }

        .app-container {
            display: flex;
            width: 100vw;
            height: 100vh;
        }

        /* ==================== 
           左側側邊欄 (維持舊版設計) 
           ==================== */
        .sidebar {
            width: 320px;
            min-width: 200px;
            max-width: 600px;
            background: linear-gradient(160deg, #a30000 0%, #600000 100%);
            color: white;
            display: flex;
            flex-direction: column;
            justify-content: center;
            align-items: center;
            text-align: center;
            padding: 20px;
            position: relative;
            flex-shrink: 0;
            z-index: 200;
            box-shadow: 5px 0 20px rgba(0,0,0,0.5);
        }

        .sidebar::before {
            content: "";
            position: absolute; top: 0; left: 0; width: 100%; height: 100%;
            background-image: url('data:image/svg+xml;utf8,<svg width="40" height="40" viewBox="0 0 40 40" xmlns="http://www.w3.org/2000/svg"><g fill="%23ffd700" fill-opacity="0.05"><path d="M20 0l20 20-20 20L0 20z"/></g></svg>');
            pointer-events: none;
            z-index: 0;
        }
        
        .sidebar > * { z-index: 1; position: relative; }

        .main-title {
            font-size: 2.5rem;
            font-weight: bold;
            color: var(--gold-accent);
            text-shadow: 2px 2px 4px rgba(0,0,0,0.6);
            line-height: 1.3;
            margin-bottom: 10px;
            letter-spacing: 2px;
        }

        .sub-title {
            font-size: 1.2rem;
            color: rgba(255,255,255,0.9);
            letter-spacing: 2px;
            margin-bottom: 40px; 
            border-top: 1px solid rgba(255,215,0,0.3);
            border-bottom: 1px solid rgba(255,215,0,0.3);
            padding: 10px 0;
            width: 100%;
        }

        .footer-info {
            position: absolute; bottom: 20px;
            font-size: 0.8rem; color: rgba(255,255,255,0.4);
        }
        .footer-info a { color: var(--gold-accent); text-decoration: none; }

        /* ==================== 
           可拖拉分隔線 
           ==================== */
        .resizer {
            width: var(--resizer-width);
            background: var(--gold-accent);
            cursor: col-resize;
            flex-shrink: 0;
            position: relative;
            z-index: 300;
            box-shadow: 1px 0 5px rgba(0,0,0,0.3);
        }

        .resizer:hover, .resizer.resizing {
            background: #fff;
            box-shadow: 0 0 10px var(--gold-accent);
        }

        /* ==================== 
           右側主要內容區 
           ==================== */
        .main-content {
            flex-grow: 1;
            overflow-y: auto;
            position: relative;
            background-color: #8B0000;
            background-image: repeating-linear-gradient(
                45deg,
                rgba(0,0,0,0.1),
                rgba(0,0,0,0.1) 10px,
                transparent 10px,
                transparent 20px
            );
            scrollbar-width: none; 
            -ms-overflow-style: none;
        }
        .main-content::-webkit-scrollbar { display: none; }

        #content-wrapper { padding: 20px 40px 100px 40px; }

        /* ----------------------------------------------------
           關鍵修改：強化一體性的獎項區塊 (Box Design)
           ---------------------------------------------------- */
        .prize-section {
            /* 讓每個獎項像是一個獨立的「盒子」 */
            margin-bottom: 50px; /* 拉大獎項間距，避免誤會 */
            background-color: var(--section-bg); /* 整個區塊有個深色底 */
            border: 1px solid rgba(255, 215, 0, 0.3); /* 金色細邊框 */
            border-radius: 16px; /* 圓角 */
            box-shadow: 0 4px 20px rgba(0,0,0,0.2);
            position: relative; /* 為了 sticky header 定位 */
        }

        .prize-header {
            position: sticky; 
            top: 0; 
            z-index: 100;
            
            /* 設計成「盒蓋」的樣子，與下方內容無縫接軌 */
            background: linear-gradient(90deg, #b30000 0%, #800000 100%);
            border-bottom: 2px solid var(--gold-accent);
            border-radius: 15px 15px 0 0; /* 只圓上方，下方切齊 */
            
            padding: 15px 30px;
            display: flex; 
            justify-content: space-between; 
            align-items: center;
            box-shadow: 0 2px 10px rgba(0,0,0,0.3);
        }

        .prize-header h2 {
            margin: 0; color: var(--gold-accent); font-size: 2rem;
            text-shadow: 2px 2px 4px rgba(0,0,0,0.6); letter-spacing: 1px;
        }

        .prize-count {
            background: rgba(0,0,0,0.3);
            color: var(--gold-accent); 
            padding: 5px 15px;
            border-radius: 20px; font-size: 1rem;
            border: 1px solid var(--gold-accent);
        }

        /* 卡片容器：在盒子內部 */
        .winner-grid {
            padding: 30px; /* 內容與邊框的距離 */
            display: grid;
            grid-template-columns: repeat(auto-fill, minmax(280px, 1fr)); 
            gap: 20px;
        }

        .winner-card {
            background: var(--card-bg);
            border-radius: 10px; 
            padding: 15px 20px;
            display: flex; align-items: center;
            
            /* 視覺微調：讓卡片有點立體感，但不會搶過標題 */
            box-shadow: 0 2px 5px rgba(0,0,0,0.1);
            border-left: 5px solid #c0392b;
            
            transition: transform 0.2s;
            position: relative; overflow: hidden;
        }

        .winner-card:hover {
            transform: translateY(-3px); 
            box-shadow: 0 5px 15px rgba(0,0,0,0.3);
        }

        .winner-info { flex-grow: 1; }
        .winner-id { font-size: 1.4rem; font-weight: bold; color: #333; margin-bottom: 4px; }
        .winner-dept { font-size: 0.9rem; color: #666; }
        .winner-number {
            font-size: 1.1rem; font-weight: bold; color: #fff;
            background: #333; padding: 4px 10px; border-radius: 6px;
            min-width: 50px; text-align: center;
        }
        
        /* 浮水印 */
        .winner-card::after {
            content: "LUCKY"; position: absolute; right: -10px; bottom: -15px;
            font-size: 4rem; font-weight: bold; color: rgba(0,0,0,0.03);
            pointer-events: none; transform: rotate(-15deg);
        }

        /* ----------------------------------------------------
           控制區
           ---------------------------------------------------- */
        #controls-area {
            position: fixed; bottom: 20px; right: 20px; z-index: 1000;
            text-align: right; display: flex; flex-direction: column; align-items: flex-end; gap: 5px;
        }
        
        .btn-fullscreen {
            background: #c0392b; border: 1px solid white; color: white;
            padding: 8px 15px; border-radius: 30px; cursor: pointer;
            font-size: 0.9rem; box-shadow: 0 4px 10px rgba(0,0,0,0.2);
            font-weight: bold; transition: all 0.2s;
        }
        .btn-fullscreen:hover { transform: scale(1.05); background: #a30000; }
        
        #status-bar { font-size: 11px; color: rgba(255,255,255,0.7); margin-bottom: 5px; text-shadow: 0 1px 2px #000;}
        .error-msg { color: #fff; font-size: 1.5rem; text-align: center; margin-top: 100px; }
    </style>
</head>
<body>

    <div class="app-container">
        
        <aside class="sidebar" id="sidebar">
            <div id="main-title" class="main-title">載入中...</div>
            <div id="sub-title" class="sub-title"></div>
            
            <div class="footer-info">
                Designed by <a href="https://pedaleon.com" target="_blank">PedaleOn</a><br>
                <span style="font-size:0.8em; opacity:0.7">拖拉金線可調整寬度 ></span>
            </div>
        </aside>

        <div class="resizer" id="resizer"></div>

        <main class="main-content" id="main-scroll-area">
            <div id="content-wrapper"></div>
        </main>
        
        <div id="controls-area">
            <div id="status-bar"></div>
            <button class="btn-fullscreen" onclick="callFullScreen()">⛶ 全螢幕 (F)</button>
        </div>

    </div>

    <script>
        let refreshRate = 5000;
        let scrollSpeed = 1.5;
        let isScrolling = true;
        let scrollDirection = 1;
        let currentScrollPos = 0;
        let lastDataHash = "";
        let timer = null;
        let scrollFrame = null;

        window.addEventListener('pywebviewready', function() {
            updateData();
            initResizer();
        });

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

        function renderUI(groupedData) {
            const currentHash = JSON.stringify(groupedData);
            if (currentHash === lastDataHash) return;
            lastDataHash = currentHash;
            
            const wrapper = document.getElementById('content-wrapper');
            let html = "";
            const awards = Object.keys(groupedData);

            if (awards.length === 0) {
                wrapper.innerHTML = "<div class='error-msg'>目前沒有得獎名單，請確認 Excel。</div>";
                return;
            }

            awards.forEach(award => {
                const list = groupedData[award];
                html += `
                <section class="prize-section">
                    <div class="prize-header">
                        <h2>${award}</h2>
                        <span class="prize-count">共 ${list.length} 位</span>
                    </div>
                    <div class="winner-grid">`;
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
            setTimeout(() => { checkAndStartScroll(); }, 100);
        }

        function checkAndStartScroll() {
            const container = document.getElementById('main-scroll-area');
            if (scrollFrame) cancelAnimationFrame(scrollFrame);
            
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
                
                if (scrollDirection === 1 && (currentScrollPos + container.clientHeight >= container.scrollHeight - 2)) {
                    isScrolling = false;
                    setTimeout(() => { scrollDirection = -1; isScrolling = true; scrollFrame = requestAnimationFrame(scrollLoop); }, 3000);
                    return;
                }
                if (scrollDirection === -1 && currentScrollPos <= 0) {
                    currentScrollPos = 0; container.scrollTop = 0;
                    isScrolling = false;
                    setTimeout(() => { scrollDirection = 1; isScrolling = true; scrollFrame = requestAnimationFrame(scrollLoop); }, 3000);
                    return;
                }
            }
            scrollFrame = requestAnimationFrame(scrollLoop);
        }

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

        document.addEventListener('keydown', (e) => { 
            if(e.key === ' ') { 
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
    window = webview.create_window(
        'Lucky Draw Board', 
        html=html_content, 
        js_api=api,
        width=1280, 
        height=800,
        background_color='#800000'
    )
    webview.start(debug=False)