import webview
import pandas as pd
import json
import os
import sys

# --- 預設設定 ---
DEFAULT_CONFIG = {
    "title": "幸運大抽獎",
    "subtitle": "得獎名單自動輪播系統", # 預設副標題
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
                    elif key == "活動副標題": config["subtitle"] = str(val) # 新增讀取副標題
                    elif key == "滾動速度": config["scroll_speed"] = float(val)
                    elif key == "更新頻率": config["refresh_rate"] = int(val) * 1000
                    elif key == "欄位-獎項": config["col_award"] = str(val)
                    elif key == "欄位-姓名": config["col_name"] = str(val)
                    elif key == "欄位-單位": config["col_dept"] = str(val)
                    elif key == "欄位-工號": config["col_id"] = str(val)
            except:
                pass # 讀取失敗則用預設值

            # 2. 讀取名單
            try:
                df = pd.read_excel(file_path, sheet_name='得獎名單')
            except:
                df = pd.read_excel(file_path, sheet_name=0)

            df.columns = df.columns.str.strip()
            col_award = config["col_award"]
            col_name = config["col_name"]
            col_dept = config["col_dept"]
            col_id = config["col_id"]

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
                emp_id = str(row[col_id]).strip() if col_id in df.columns else ""
                
                if dept == 'nan': dept = ''
                if emp_id == 'nan': emp_id = ''

                if award not in result:
                    result[award] = []
                
                result[award].append({"name": name, "dept": dept, "empId": emp_id})

            return json.dumps({
                "status": "success", 
                "data": result, 
                "meta": {
                    "title": config["title"],
                    "subtitle": config["subtitle"], # 回傳副標題
                    "scroll_speed": config["scroll_speed"],
                    "refresh_rate": config["refresh_rate"]
                }
            })

        except Exception as e:
            return json.dumps({"error": f"讀取錯誤: {str(e)}"})

# --- HTML/CSS/JS ---
html_content = """
<!DOCTYPE html>
<html lang="zh-TW">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Lucky Draw</title>
    <style>
        :root { --bg-color: #8B0000; --card-bg: #fffbf0; --border-color: #D4AF37; --accent-color: #c41e3a; }
        body, html { margin: 0; padding: 0; height: 100%; font-family: "Microsoft JhengHei", sans-serif; background-color: var(--bg-color); color: #333; overflow: hidden; user-select: none; }
        
        /* 背景紋理 */
        body::before { content: ""; position: absolute; top: 0; left: 0; width: 100%; height: 100%; background-image: repeating-linear-gradient(45deg, transparent 0, transparent 20px, rgba(255, 215, 0, 0.05) 20px, rgba(255, 215, 0, 0.05) 22px); z-index: -1; }
        
        /* 頂部 Header */
        header { text-align: center; padding: 15px; background: linear-gradient(180deg, #b30000 0%, #800000 100%); border-bottom: 4px solid var(--border-color); position: fixed; top: 0; width: 100%; z-index: 200; height: 110px; display: flex; flex-direction: column; justify-content: center; box-shadow: 0 5px 15px rgba(0,0,0,0.5); }
        h1 { margin: 0; color: var(--border-color); text-shadow: 2px 2px 0px #333; font-size: 2.5rem; letter-spacing: 5px; }
        .sub-title { color: #ffd700; font-size: 1.1rem; margin-top: 5px; letter-spacing: 2px; }
        
        /* 滾動區 */
        #scroll-container { margin-top: 115px; height: calc(100vh - 115px); overflow-y: auto; padding: 20px; box-sizing: border-box; scrollbar-width: none; }
        #scroll-container::-webkit-scrollbar { display: none; }
        #content-wrapper { max-width: 1000px; margin: 0 auto; padding-bottom: 150px; }
        
        /* 獎項卡片 */
        .award-group { background: var(--card-bg); border: 2px solid var(--border-color); border-radius: 15px; margin-bottom: 30px; padding: 0 20px 20px 20px; box-shadow: 0 8px 20px rgba(0,0,0,0.3); animation: fadeIn 0.5s ease; position: relative; overflow: clip; }
        @keyframes fadeIn { from { opacity: 0; transform: translateY(20px); } to { opacity: 1; transform: translateY(0); } }
        .award-header { text-align: center; position: sticky; top: 0; z-index: 10; background-color: var(--card-bg); padding: 20px 0 15px 0; margin-bottom: 15px; border-bottom: 3px double #e0c080; box-shadow: 0 4px 6px -4px rgba(0,0,0,0.2); }
        .award-name { font-size: 2rem; color: var(--accent-color); font-weight: bold; }
        .winner-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(200px, 1fr)); gap: 10px; }
        .winner-item { background: #fff; border-left: 5px solid var(--accent-color); border-radius: 5px; padding: 10px; display: flex; justify-content: space-between; align-items: center; box-shadow: 0 2px 5px rgba(0,0,0,0.05); }
        .w-info h3 { margin: 0; font-size: 1.3rem; }
        .w-info p { margin: 0; font-size: 0.85rem; color: #666; }
        .w-id { background: #eee; padding: 3px 6px; border-radius: 4px; font-size: 0.8rem; font-weight: bold; }
        
        /* --- Footer 版權宣告 --- */
        .footer {
            position: fixed;
            bottom: 5px;
            left: 50%;
            transform: translateX(-50%);
            color: rgba(255, 255, 255, 0.4); /* 半透明白 */
            font-size: 0.8rem;
            z-index: 1000;
            text-shadow: 1px 1px 2px rgba(0,0,0,0.8);
            pointer-events: auto;
        }
        .footer a {
            color: #FFD700; /* 金色連結 */
            text-decoration: none;
            font-weight: bold;
            transition: color 0.3s;
        }
        .footer a:hover {
            color: #fff;
            text-decoration: underline;
        }

        /* 狀態列移到右下角，避免跟 footer 撞到 */
        #status-bar { position: fixed; bottom: 5px; right: 10px; font-size: 10px; color: rgba(255,255,255,0.3); z-index: 1000; text-align: right; }
        .error-msg { color: yellow; font-weight: bold; padding: 20px; text-align: center; font-size: 1.5rem;}
    </style>
</head>
<body>
    <header>
        <h1 id="main-title">載入中...</h1>
        <div id="sub-title" class="sub-title"></div>
    </header>

    <div id="scroll-container">
        <div id="content-wrapper"></div>
    </div>

    <div class="footer">
        由 <a href="https://pedaleon.com/?utm_source=luckdraw&utm_medium=app" target="_blank">PedaleOn</a> 開發
    </div>

    <div id="status-bar"></div>

    <script>
        let refreshRate = 5000;
        let scrollSpeed = 1.5; // 可以改大一點測試，例如 3.0
        let isScrolling = true;
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
                        scrollSpeed = res.meta.scroll_speed || 1.5;
                        let newRate = res.meta.refresh_rate || 5000;
                        if (newRate !== refreshRate) refreshRate = newRate;
                    }

                    renderUI(res.data);
                    document.getElementById('status-bar').innerText = "Last Update: " + new Date().toLocaleTimeString();
                    
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
                wrapper.innerHTML = "<div style='text-align:center;color:white;margin-top:50px;font-size:1.5rem'>等待開獎中...</div>";
                return;
            }

            awards.forEach(award => {
                const list = groupedData[award];
                html += `<div class="award-group"><div class="award-header"><span class="award-name">${award}</span><span style="font-size:0.9rem; color:#666; margin-left:10px;">(共${list.length}位)</span></div><div class="winner-grid">`;
                list.forEach(p => {
                    html += `<div class="winner-item"><div class="w-info"><h3>${p.name}</h3><p>${p.dept}</p></div><div class="w-id">${p.empId}</div></div>`;
                });
                html += `</div></div>`;
            });
            wrapper.innerHTML = html;

            // --- 修正重點：資料渲染完後，立刻檢查並啟動滾動 ---
            // 稍微延遲 100ms 讓瀏覽器重新計算高度
            setTimeout(() => {
                checkAndStartScroll();
            }, 100);
        }

        function checkAndStartScroll() {
            const container = document.getElementById('scroll-container');
            const content = document.getElementById('content-wrapper');

            // 如果已經在跑，先停掉避免重複
            if (scrollFrame) cancelAnimationFrame(scrollFrame);

            // 判斷是否需要滾動 (內容高度 > 容器高度)
            if (container.scrollHeight > container.clientHeight) {
                isScrolling = true;
                scrollLoop();
            } else {
                // 內容太少，不用滾，直接歸零
                container.scrollTop = 0;
            }
        }

        function scrollLoop() {
            const container = document.getElementById('scroll-container');
            
            if (isScrolling) {
                container.scrollTop += scrollSpeed;
                
                // 判斷到底部：scrollTop + clientHeight 與 scrollHeight 的誤差在 2px 內
                // 如果到底了，或者因為浮點數誤差導致卡住
                if (container.scrollTop + container.clientHeight >= container.scrollHeight - 2) {
                    
                    // 暫停滾動
                    isScrolling = false;
                    
                    // 3秒後跳回頂部並繼續
                    setTimeout(() => {
                        container.scrollTop = 0;
                        isScrolling = true;
                        // 重新啟動迴圈
                        scrollFrame = requestAnimationFrame(scrollLoop); 
                    }, 3000);
                    
                    return; // 結束這一次的迴圈，等待 setTimeout 喚醒
                }
            }
            scrollFrame = requestAnimationFrame(scrollLoop);
        }

        // 空白鍵暫停/繼續
        document.addEventListener('keydown', (e) => { 
            if(e.key === ' ') { 
                isScrolling = !isScrolling; 
                // 如果是從暫停恢復，要確保迴圈有在跑
                if(isScrolling) scrollLoop();
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
        width=1200, 
        height=800,
        background_color='#8B0000'
    )
    webview.start(debug=False)