================================================================
  BOM Matcher v0.07  —  實機測試包
  ADATA Technology  /  2026-09-01
================================================================

Altium Text BOM + Excel 零件資料庫  →  ADATA E-BOM 格式 Excel

本包內含兩套介面（C 案 / D 案）與一套 CLI，共用同一份核心邏輯，
產出的 Excel 完全相同。附有可直接跑的範例資料。


----------------------------------------------------------------
  一、環境需求
----------------------------------------------------------------

  作業系統   Windows 10 / 11
  Python     3.10 以上（開發與驗證環境為 3.14.4）
             安裝時務必勾選「Add python.exe to PATH」
             下載：https://www.python.org/downloads/

  D 案另需   Microsoft Edge WebView2 Runtime
             Windows 11 內建，Windows 10 通常已隨 Edge 安裝。
             若 D 案開不起來才需要手動裝：
             https://developer.microsoft.com/microsoft-edge/webview2/


----------------------------------------------------------------
  二、安裝（只需做一次）
----------------------------------------------------------------

  把整包解壓到任意資料夾（路徑可含中文，但建議避免），例如：

      D:\BOM_Matcher_v0.07\

  然後雙擊：

      install_requirements.bat

  它會自動安裝 pandas / openpyxl / pillow / streamlit /
  streamlit-option-menu / pywebview，最後印出檢查結果。

  ★ 程式一律以「自己所在的資料夾」為工作目錄，
    解壓到哪裡都可以，不必固定路徑。


----------------------------------------------------------------
  三、執行（三選一）
----------------------------------------------------------------

  ▌ 先跑這個確認核心正常（不需要 streamlit / pywebview）

      run_CLI.bat

      應該會看到自動偵測到 TXT / DB / Knowledge / PCB 圖片，
      並在 output\ 產生一個 xlsx。

  ▌ C 案 — Streamlit 網頁版

      run_C_streamlit.bat
      → 瀏覽器開 http://localhost:8601
      → 關閉時請把那個黑色主控台視窗一起關掉

    ★ 一律用這支 .bat 啟動。直接雙擊 bom_matcher_C_0.07.py 不會開網頁
      （Streamlit 必須用 streamlit run 啟動），程式會印出提示並停住。

    ★ 首次啟動時 .bat 會在 %USERPROFILE%\.streamlit\ 建立一個
      credentials.toml（內容是空 email），用來跳過 Streamlit 首次執行
      的 email 詢問。不建立的話畫面會卡在 "Email:" 等你輸入。

  ▌ D 案 — pywebview 原生視窗版

      run_D_webview.bat          正常使用（無主控台）
      run_D_webview_debug.bat    出問題時用這支，會保留錯誤訊息


----------------------------------------------------------------
  四、操作流程（C / D 兩案相同）
----------------------------------------------------------------

  1. 檔案設定      確認自動偵測的四個檔案；不對就自己選
  2. 執行比對
  3. 比對結果      檢查表格，可直接改 P/N / RC/EC / Description /
                   Qty / Location
  4. Knowledge     逐筆確認，不適用的按「忽略此筆」
  5. 表頭資訊      填 PCB板號 / 供應商 / IC Type 等（可留空）
  6. 產出 E-BOM    輸出到 output\


----------------------------------------------------------------
  五、資料夾規則
----------------------------------------------------------------

  程式會掃描自己所在的資料夾與下列子資料夾：

      TXT\           Altium Text BOM（*.txt 或 *.BOM）
      ExcelDB\       零件資料庫 + Knowledge（*.xlsx）
      screenshot\    PCB 圖片（*.jpg / *.png，pcb.* 優先）
      output\        產出的 E-BOM（不存在則輸出到程式資料夾）

  ★ 自動偵測是「看內容」不是「看檔名」：
      Text BOM     必須有一列是 TAB 分隔的表頭，且欄位含
                   Item / Quantity / Reference
      零件資料庫   第 1 列須有 P/N 與 Description
      Knowledge    第 1 列須有 品名 與 包裝
    所以把自己產出的 E-BOM、或這份 README 丟回這些資料夾，
    都不會被誤認成輸入檔。


----------------------------------------------------------------
  六、本包附的範例資料
----------------------------------------------------------------

  TXT\S-M2P42280SMH086-20260812_1.BOM               44 項（預設選這支）
  TXT\S-SM2268XT_154BALL_20260515_FLASHX4_BOM.txt   54 項
  ExcelDB\BOM_20260525.xlsx                         零件資料庫 49 筆
  ExcelDB\Layout and BOM Knowledge_20260624.xlsx    Knowledge 6 筆
  pcb.jpg                                           PCB 圖片

  ▌ 預期結果 A：直接跑 run_CLI.bat

      解析 Text BOM: 44 項
      讀取零件資料庫: 49 筆
      載入 Knowledge: 6 筆
      比對完成: 命中 21 / 未命中 24 / 替代料群組 5 / Knowledge 命中 1 列
      45 項（含PCB）→ 51 列（替代料拆出 6 列）  |  Total Qty = 101

  ▌ 預期結果 B：改用另一支 TXT
      run_CLI.bat "TXT\S-SM2268XT_154BALL_20260515_FLASHX4_BOM.txt" "ExcelDB\BOM_20260525.xlsx"

      解析 Text BOM: 54 項
      比對完成: 命中 37 / 未命中 18 / 替代料群組 8 / Knowledge 命中 6 列
      55 項（含PCB）→ 66 列（替代料拆出 11 列）  |  Total Qty = 174

  ▌ 預期結果 C：驗證拆列不會重複計算 Qty
      run_CLI.bat --no-split
      → 45 項 → 45 列，Total Qty 一樣是 101

  跟這些數字對得上就代表環境沒問題。
  在 C / D 兩個介面上跑同一份資料，數字也應該完全一樣。


----------------------------------------------------------------
  七、v0.07 的主要改動
----------------------------------------------------------------

  ▌ 替代料拆列 —— 一列只有一個 P/N

      舊：一列   P/N = 61041115;61041238
      新：主列   Item=8  Qty=3  Location=C2,C3,C6  P/N=61041115
          續列   Item=空 Qty=空 Location=空        P/N=61041238

      續列的 Description 是該替代料自己的 DB 描述。
      Total Qty 只加總主列，不會因為拆列而重複計算。
      要關掉：介面上的「替代料拆列」開關，或 CLI 加 --no-split

  ▌ 產出前可人工修正表格（過去只能事後手改 Excel）
  ▌ 表頭資訊（rows 3-8）可填
  ▌ 工作目錄改為程式所在目錄
  ▌ 自動偵測改為驗內容，不再誤抓產出檔／changelog

  完整清單見 changelog_BOM_matcher_v0.07.txt


----------------------------------------------------------------
  八、常見問題
----------------------------------------------------------------

  Q. 雙擊 .bat 一閃就關
     A. 沒安裝 Python 或沒加入 PATH。先跑 install_requirements.bat
        看錯誤訊息。

  Q. C 案瀏覽器打不開 / 顯示無法連線
     A. 8601 埠被占用。改執行：
        python -m streamlit run bom_matcher_C_0.07.py --server.port 8602

  Q. D 案視窗開不起來或一閃而過
     A. 用 run_D_webview_debug.bat 看訊息。
        多半是缺 Microsoft Edge WebView2 Runtime（見第一節）。

  Q. 說找不到零件資料庫
     A. 檢查 xlsx 第 1 列是不是 P/N 與 Description 這兩個字，
        必須完全一樣（大小寫不拘）。

  Q. 產出的 Excel 開啟時 Excel 說要修復
     A. 請回報，並附上使用的 TXT 與 DB 檔。


----------------------------------------------------------------
  九、檔案清單
----------------------------------------------------------------

  bom_core_0.07.py               核心邏輯 + CLI
  bom_matcher_C_0.07.py          C 案（Streamlit）
  bom_matcher_D_0.07.py          D 案（pywebview）
  .streamlit\config.toml         C 案主題設定
  install_requirements.bat       安裝依賴
  run_CLI.bat                    CLI
  run_C_streamlit.bat            C 案
  run_D_webview.bat              D 案
  run_D_webview_debug.bat        D 案（除錯）
  requirements.txt               套件清單
  README_v0.07.txt               本文件
  changelog_BOM_matcher_v0.07.txt  完整版本歷史 v0.01-v0.07
  HANDOVER_bom_matcher.md        技術交接文件（給後續維護者）
  TXT\ ExcelDB\ output\          範例資料與輸出目錄
  pcb.jpg                        範例 PCB 圖片

================================================================
