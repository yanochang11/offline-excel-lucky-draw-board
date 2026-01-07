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

# --- HTML/CSS/JS (全新側邊欄佈局) ---
html_content = """
<!DOCTYPE html>
<html lang="zh-TW">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Lucky Draw</title>
    <style>
        :root { --sidebar-bg: #8B0000; --main-bg: #fffbf0; --accent-gold: #FFD700; --text-red: #c41e3a; }
        
        body, html { 
            margin: 0; padding: 0; 
            height: 100%; 
            font-family: "Microsoft JhengHei", sans-serif; 
            overflow: hidden; 
            user-select: none; 
            display: flex; /* 改用 Flex 佈局 */
        }
        
        /* --- 左側側邊欄 (Sidebar) --- */
        .sidebar {
            width: 320px; /* 固定寬度 */
            height: 100vh;
            background: linear-gradient(160deg, #a30000 0%, #600000 100%);
            border-right: 5px solid var(--accent-gold);
            box-shadow: 5px 0 20px rgba(0,0,0,0.5);
            z-index: 100;
            display: flex;
            flex-direction: column;
            justify-content: center;
            align-items: center;
            padding: 20px;
            box-sizing: border-box;
            color: white;
            text-align: center;
            position: relative;
        }

        /* 裝飾圖案 */
        .sidebar::before {
            content: "";
            position: absolute; top: 0; left: 0; width: 100%; height: 100%;
            background-image: url('data:image/svg+xml;utf8,<svg width="40" height="40" viewBox="0 0 40 40" xmlns="http://www.w3.org/2000/svg"><g fill="%23ffd700" fill-opacity="0.05"><path d="M20 0l20 20-20 20L0 20z"/></g></svg>');
            pointer-events: none;
        }

        .main-title {
            font-size: 2.5rem;
            font-weight: bold;
            color: var(--accent-gold);
            text-shadow: 2px 2px 4px rgba(0,0,0,0.6);
            line-height: 1.3;
            margin-bottom: 10px;
            letter-spacing: 2px;
        }

        .sub-title {
            font-size: 1.2rem;
            color: rgba(255,255,255,0.9);
            letter-spacing: 2px;
            margin-bottom: 40px; /* 與下方拉開距離 */
            border-top: 1px solid rgba(255,215,0,0.3);
            border-bottom: 1px solid rgba(255,215,0,0.3);
            padding: 10px 0;
            width: 100%;
        }

        .footer-info {
            position: absolute;
            bottom: 20px;
            font-size: 0.8rem;
            color: rgba(255,255,255,0.4);
        }
        .footer-info a { color: var(--accent-gold); text-decoration: none; }

        /* --- 右側主要滾動區 (Main Content) --- */
        .main-content {
            flex: 1; /* 佔滿剩餘寬度 */
            height: 100vh;
            background-color: var(--main-bg);
            position: relative;
            background-image: radial-gradient(circle at 50% 50%, rgba(200, 0, 0, 0.02) 0%, transparent 60%);
        }

        #scroll-container { 
            width: 100%;
            height: 100vh; /* 滿版高度 */
            overflow-y: hidden; /* 隱藏原生捲軸，用 JS 控制 */
            padding: 40px; /* 內縮一點比較好看 */
            box-sizing: border-box;
        }
        
        #content-wrapper { max-width: 1200px; margin: 0 auto; padding-bottom: 100px; }
        
        /* 獎項卡片 */
        .award-group { 
            background: #fff; 
            border: 2px solid #e6cfa3; 
            border-radius: 15px; 
            margin-bottom: 40px; 
            padding: 0 25px 25px 25px; 
            box-shadow: 0 5px 15px rgba(0,0,0,0.05); 
            position: relative; 
            overflow: clip; 
        }
        
        .award-header { 
            text-align: left; /* 改為靠左 */
            position: sticky; top: 0; z-index: 10; 
            background-color: #fff; 
            padding: 20px 0 15px 0; 
            margin-bottom: 20px; 
            border-bottom: 3px double var(--text-red); 
            display: flex; justify-content: space-between; align-items: baseline;
            box-shadow: 0 5px 10px -5px rgba(0,0,0,0.1);
        }
        
        .award-name { font-size: 2.2rem; color: var(--text-red); font-weight: bold; }
        .award-count { font-size: 1rem; color: #888; }
        
        .winner-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(220px, 1fr)); gap: 15px; }
        
        .winner-item { 
            background: #fffbf0; 
            border-left: 6px solid var(--text-red); 
            border-radius: 6px; 
            padding: 12px 15px; 
            display: flex; justify-content: space-between; align-items: center; 
            box-shadow: 0 2px 4px rgba(0,0,0,0.05); 
        }
        .w-info h3 { margin: 0; font-size: 1.4rem; color: #333; }
        .w-info p { margin: 2px 0 0; font-size: 0.9rem; color: #666; }
        .w-id { background: #e0e0e0; color:#555; padding: 3px 8px; border-radius: 4px; font-size: 0.9rem; font-weight: bold; }

        /* 右下角控制區 */
        #controls-area { position: fixed; bottom: 20px; right: 20px; z-index: 1000; text-align: right; display: flex; flex-direction: column; align-items: flex-end; gap: 5px; pointer-events: auto;}
        #status-bar { font-size: 11px; color: #999; margin-bottom: 5px; background: rgba(255,255,255,0.8); padding: 2px 5px; border-radius: 3px;}
        
        .btn-fullscreen {
            background: var(--text-red);
            border: 1px solid white;
            color: white;
            padding: 8px 15px;
            border-radius: 30px;
            cursor: pointer;
            font-size: 0.9rem;
            box-shadow: 0 4px 10px rgba(0,0,0,0.2);
            transition: all 0.2s;
            display: flex; align-items: center; gap: 5px; font-weight: bold;
        }
        .btn-fullscreen:hover { transform: scale(1.05); background: #a30000; }
        
        .error-msg { color: #555; font-weight: bold; padding: 50px; text-align: center; font-size: 1.5rem;}
    </style>
</head>
<body>

    <aside class="sidebar">
        <div id="main-title" class="main-title">載入中...</div>
        <div id="sub-title" class="sub-title"></div>
        
        <div class="footer-info">
            Designed by <a href="https://pedaleon.com/?utm_source=luckdraw&utm_medium=app" target="_blank">PedaleOn</a>
        </div>
    </aside>

    <main class="main-content">
        <div id="scroll-container">
            <div id="content-wrapper"></div>
        </div>
        
        <div id="controls-area">
            <div id="status-bar"></div>
            <button class="btn-fullscreen" onclick="callFullScreen()">
                ⛶ 全螢幕 (F)
            </button>
        </div>
    </main>

    <script>
        let refreshRate = 5000;
        let scrollSpeed = 1.5;
        let isScrolling = true;
        let scrollDirection = 1;
        let currentScrollPos = 0; // 高精度位置
        let lastDataHash = "";
        let timer = null;
        let scrollFrame = null;
        
        window.addEventListener('pywebviewready', function() {
            updateData();
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
                    document.getElementById('status-bar').innerText = "更新時間: " + new Date().toLocaleTimeString();
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
                wrapper.innerHTML = "<div style='text-align:center;color:#999;margin-top:100px;font-size:1.5rem'>等待開獎中...</div>";
                return;
            }
            awards.forEach(award => {
                const list = groupedData[award];
                html += `<div class="award-group"><div class="award-header"><span class="award-name">${award}</span><span class="award-count">共${list.length}位</span></div><div class="winner-grid">`;
                list.forEach(p => {
                    html += `<div class="winner-item"><div class="w-info"><h3>${p.name}</h3><p>${p.dept}</p></div><div class="w-id">${p.empId}</div></div>`;
                });
                html += `</div></div>`;
            });
            wrapper.innerHTML = html;
            
            setTimeout(() => { 
                scrollDirection = 1; 
                checkAndStartScroll(); 
            }, 100);
        }

        function checkAndStartScroll() {
            const container = document.getElementById('scroll-container');
            if (scrollFrame) cancelAnimationFrame(scrollFrame);
            
            if (container.scrollHeight > container.clientHeight) {
                isScrolling = true;
                currentScrollPos = container.scrollTop;
                scrollLoop();
            } else {
                container.scrollTop = 0;
            }
        }

        function scrollLoop() {
            const container = document.getElementById('scroll-container');
            
            if (isScrolling) {
                currentScrollPos += (scrollSpeed * scrollDirection);
                container.scrollTop = currentScrollPos;
                
                // 到底部
                if (scrollDirection === 1 && (currentScrollPos + container.clientHeight >= container.scrollHeight - 2)) {
                    isScrolling = false;
                    setTimeout(() => {
                        scrollDirection = -1; 
                        isScrolling = true;
                        scrollFrame = requestAnimationFrame(scrollLoop);
                    }, 3000); 
                    return;
                }

                // 到頂部
                if (scrollDirection === -1 && currentScrollPos <= 0) {
                    currentScrollPos = 0;
                    container.scrollTop = 0;
                    isScrolling = false;
                    setTimeout(() => {
                        scrollDirection = 1; 
                        isScrolling = true;
                        scrollFrame = requestAnimationFrame(scrollLoop);
                    }, 3000); 
                    return;
                }
            }
            scrollFrame = requestAnimationFrame(scrollLoop);
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
    # 注意：這裡的 width 建議設寬一點，模擬寬螢幕效果
    window = webview.create_window(
        'Lucky Draw Board', 
        html=html_content, 
        js_api=api,
        width=1280, 
        height=720,
        background_color='#8B0000'
    )
    webview.start(debug=False)