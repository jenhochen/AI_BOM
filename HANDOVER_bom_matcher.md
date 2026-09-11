# BOM Matcher — 新對話交接事項

## 背景
Stanley（ADATA Technology）開發的 Python Windows 工具，讀取 Altium 格式 Text BOM 與 Excel 零件資料庫，輸出 ADATA E-BOM 格式 Excel。

---

## 目前版本：v0.07

v0.07 起拆成「一個核心 + 兩套前端」。v0.06 的 tkinter GUI 已由 C / D 兩案取代。

| 檔案 | 角色 | 行數 |
|---|---|---|
| `bom_core_0.07.py` | 核心邏輯（解析／比對／Knowledge／Excel 輸出）＋ CLI | ~700 |
| `bom_matcher_C_0.07.py` | **C 案**：Streamlit 網頁版 | ~330 |
| `bom_matcher_D_0.07.py` | **D 案**：pywebview 原生視窗版（HTML 內嵌） | ~640 |
| `run_C_streamlit.bat` | C 案啟動捷徑（port 8601） | — |
| `run_D_webview.bat` | D 案啟動捷徑 | — |
| `.streamlit\config.toml` | C 案主題（primaryColor `#4F46E5`） | — |

**版本規則：** 每次進版 +0.01，程式檔名含版號。

### 執行方式
```
雙擊 run_C_streamlit.bat        # C 案：瀏覽器 http://localhost:8601
雙擊 run_D_webview.bat          # D 案：原生視窗

python bom_core_0.07.py                    # CLI 全自動
python bom_core_0.07.py -o out.xlsx
python bom_core_0.07.py BOM.txt db.xlsx -k Knowledge.xlsx -p pcb.jpg
python bom_core_0.07.py --no-split         # 關閉替代料拆列

python bom_matcher_D_0.07.py --dump-html   # 輸出 D 案 UI 的 HTML
python bom_matcher_D_0.07.py --selftest    # 自動跑完流程後關閉（驗證用）
```

### 依賴套件
```
pip install pandas openpyxl pillow                # 核心
pip install streamlit streamlit-option-menu       # C 案
pip install pywebview                             # D 案
```
實測環境：Python 3.14.4 / pandas 3.0.3 / openpyxl 3.1.5 / streamlit 1.62.0 / pywebview 6.2.1

> 前端檔名含小數點，無法用一般 import，兩案皆以 `importlib` 明確載入 `bom_core_0.07.py`。改版號時記得同步兩個前端裡的路徑字串。

---

## 工作目錄與自動偵測（v0.07 重點）

**工作目錄 = 程式所在目錄**（`APP_DIR = Path(__file__).resolve().parent`），不再用 `Path.cwd()`。

| 偵測目標 | 搜尋位置 | 判定方式 |
|---|---|---|
| Text BOM | 程式目錄 + `TXT\` | 內容含 `Bill Of Materials` 或 Item/Quantity/Reference 表頭 |
| 零件資料庫 | 程式目錄 + `ExcelDB\` | 第 1 列有 `P/N` + `Description`，且 A1 不是 `ADATA Technology Co., Ltd.` |
| Knowledge | 程式目錄 + `ExcelDB\` | 第 1 列有 `品名` + `包裝`，同樣排除產出檔 |
| PCB 圖片 | 程式目錄 + `screenshot\` | `pcb.*` 優先，其次任何 jpg/png |
| 輸出目錄 | `output\`（不存在則程式目錄） | — |

⚠️ **不能只看檔名。** v0.06 就是這樣踩雷：`PREVIEW_BOM_knowledge_v0.06.xlsx` 檔名含 knowledge 但其實是 E-BOM 產出檔，會載入 57 筆垃圾並產生 21 個假命中；`changelog_*.txt` 也會被當成 BOM 檔。

---

## 核心資料流

```
parse_text_bom(txt)  →  items[]        （item / qty / ref / part / fp）
read_db(xlsx)        →  DataFrame      （P/N / Description）
load_knowledge(xlsx) →  entries[]      （pn / 品名 / 包裝 / note）
        ↓
build_rows(items, db, knowledge, split_alt=True)  →  rows[]   ★ 替代料拆列在這
        ↓  （C/D 介面在這裡讓人工修改 rows）
write_excel(rows, out, pcb_img, header_info)      →  xlsx
```

`write_excel()` 只吃 `rows`，與比對邏輯完全解耦 —— 這是 GUI 能「先改再產出」的關鍵。

### rows[] 欄位
| 欄位 | 說明 |
|---|---|
| `item` | 序號；**替代料續列為 None** |
| `seq` | 群組編號（拆列共用，不寫進 Excel） |
| `rcec` / `pn` / `desc` | RC/EC、單一 P/N、描述 |
| `qty` / `ref` | **續列為 None / ''** |
| `nc` | 原本為 NC |
| `alt` | 這列是替代料續列 |
| `group` | 此列屬於有替代料的群組（決定綠底） |
| `fp` / `part` | 原始 footprint / part 字串 |
| `kn` | Knowledge note（決定橘底） |

---

## 替代料拆列規則（v0.07 主要需求）

```
v0.06:  一列  P/N = 61041115;61041238
v0.07:  主列  Item=8   Qty=3   Location=C2,C3,C6   P/N=61041115
        續列  Item=空  Qty=空  Location=空         P/N=61041238
```

- 續列的 Description 是**該替代料自己的 DB 描述**，不是複製主料的
- **Total Qty 只加總主列**（`total_qty()` 跳過 `qty is None`），拆列不會重複計算
- 三料以上同樣支援（實測 33Ω BEAD 拆出 3 列）
- 關閉方式：CLI `--no-split`，或介面上的「替代料拆列」開關

---

## 輸出 Excel 結構

⚠️ **沒有範本檔。版面是寫死在 Python 裡的。**
視覺參考來源是 `demo document\BOM_SM2P41D8QB_20260526_M2P42280SMH086.xlsx`（工作表 `B68S 2CE`），程式是 1:1 抄的：A2:L2 合併、K7:L8 / M7:O8 跨列、**Row 36 = BOM 表頭**、表尾 A85:B86 等六格兩列合併。

```
Row  1     : ADATA Technology Co., Ltd.
Row  2     : M.2 SSD E-BOM
Row  3-8   : PCB資訊表格（A:C=標籤 D:J=值 K:L=右標籤 M:O=右值）
             ← v0.07 起可由 header_info 填入 9 個欄位
Row  9-35  : PCB圖片區（rows 高 28.2pt，A/O 欄補 medium 外框）
Row 36     : BOM表頭（BOM_HDR 常數）
Row 37+    : BOM資料列（含替代料續列）
最後一個BOM列 : PCB（EC，Description='PCB'，Qty=1）
Total列 / 表尾（Remark / 圖例 / Approval）
```

`BOM_HDR = 36`、圖片區 rows 9–35、表尾合併座標都是硬編碼的魔術數字。ADATA 官方表格若改版就得手改程式。

### 填色規則（優先序）
| 顏色 | 代碼 | 條件 |
|---|---|---|
| 橘底 | `FFE0B2` | Knowledge 命中（最高優先） |
| 黃底 | `FFF2CC` | 原本為 NC |
| 綠底 | `E2EFDA` | 屬於有替代料的群組（主列＋續列都上） |
| 藍底 | `D9E1F2` | Total 列 |

---

## RC/EC 分類邏輯

```
P/N 前四碼 = 6103 或 6104  → RC
P/N 為空時（無DB命中）：
  Description 開頭為容值（10pF, 220nF, 2.2uF...）  → RC
  Description 含「RES」                              → RC
  Description 含「Ω」或「ohm」                      → RC
  Footprint 首字元為 'C' 或 'R'                     → RC
  其他                                               → EC
注意：6900*（錫膏）= EC（v0.06 起移出 RC）
```
群組內所有拆列共用同一個 RC/EC（以主料 P/N 判定）。

---

## Knowledge 比對邏輯

Knowledge 檔欄位：`項次 | PN | 品名 | 包裝 | note`

三種命中條件（任一符合即觸發）：
1. **PN**：BOM P/N = Knowledge PN（精確）
2. **品名**：Knowledge 品名 ⊆ BOM Description（大小寫不分）
3. **包裝**：Knowledge 包裝 ⊆ BOM Description（dash-insensitive）

v0.07 起 **逐列比對**（每個拆出來的 P/N 各自比對自己的描述），命中後 note 寫入 G:O、橘底標示。兩案介面都可以逐筆「忽略」。

---

## 已知邊界情況

| 情況 | 行為 |
|---|---|
| Item 17 `BAT54KFILM` | DB 描述為 "SCHOTTY" 非型號，自動比對必然落空，P/N 空、RC/EC=EC。**v0.07 起可在介面上直接補** |
| Crystal 25MHz/6pF | 含 "6pF" 但 Description 非以容值開頭 → 正確歸 EC |
| TVS 0201 | 品名 "TVS" 觸發 Knowledge 命中（內容講 SOD-523），需人工判斷。**v0.07 起可在介面上「忽略此筆」** |
| `NC(0R_5%)` | `is_nc()` 只認完全等於 NC 的分段，這種寫法不算 NC |
| 6900*(錫膏) | 現為 EC；要改回 RC 就在 `_RC_PREFIXES` 加 `'6900'` |

---

## 兩案的坑（實作時踩過，別再踩）

### C 案 Streamlit
- **`st.data_editor` 不套用 pandas Styler 底色** → 「可編輯」與「三色底」無法並存。目前分兩個頁籤：`✎ 編輯`（emoji 狀態欄）／`🎨 色底檢視`（唯讀 Styler）。要兩者兼得得換 `streamlit-aggrid`。
- **`st.data_editor` 把 `None` / `pd.NA` / `Int64` 空值渲染成字串 `"None"`** → Item / Qty 兩欄改用字串表示，存回時再轉 int。
- 版面只能單欄由上而下，沒有原生檔案選擇器（用下拉選單＋手動輸入路徑替代）。

### D 案 pywebview
- **不能把 window 物件掛成 `self.window`** → pywebview 會遞迴序列化 js_api 的公開屬性，走進 `window.native.AccessibilityObject.Bounds.Empty...` 造成 `maximum recursion depth exceeded`，視窗當場關閉。window 一律放模組層變數 `WIN`。
- **`window.pywebview` 是載入後才注入的**，腳本執行當下判定必為 undefined。要用即時檢查 `LIVE()`，別在載入時就決定「預覽模式」。
- 表格儲存格裡放 `<input>` 時，色點要包一層 flex，否則 input 是 block 會把色點擠到上一行。
- 從非互動 shell（stdin 導向 null）直接啟動時視窗會秒關；用 `Start-Process` 並重導向 stdout/stderr 才穩。這是測試環境的問題，使用者雙擊 .bat 不受影響。

### .bat 啟動檔（v0.07 首次交付就是死在這裡）

三條硬規則，**用 `tools\gen_bats.py` 產生，不要手寫**：

1. **必須 CRLF 換行。** cmd.exe 是用位元組位移讀批次檔的，LF-only 會讓它定位錯亂、**從一行的中間開始執行** —— 實際症狀是 `echo` 被切成 `ho`、`BOM Matcher` 被切成 `atcher`。
2. **內容必須純 ASCII。** .bat 裡放中文，cmd 用哪個 codepage 讀檔無法可靠控制；Big5 解碼時中文的位元組還會**把換行吃掉**，讓下一行的指令被併進 `REM` 註解（`cd /d "%~dp0"` 就這樣消失）。中文說明一律放 README。
3. **用絕對路徑 `"%~dp0檔名"`**，不要靠 `cd` 之後的相對檔名。

另外兩個必備項目：

- 開頭加 `chcp 65001 >nul 2>&1` + `set PYTHONIOENCODING=utf-8`，這樣 Python 印出的中文與 `✓` 在 cmd 裡才不會亂碼或觸發 UnicodeEncodeError
- **必須預先建立 `%USERPROFILE%\.streamlit\credentials.toml`**（`[general]` / `email = ""`）。否則 Streamlit 首次執行會停在互動式的 `Email:` 詢問等輸入，雙擊 .bat 的人只會看到視窗卡住。credentials 只認全域路徑，放在專案的 `.streamlit\` 沒用。

### 自動偵測的自我誤判

`is_bom_txt()` 原本用「檔案內容包含 Bill Of Materials 或 Item/Quantity/Reference」判定，結果**本專案的 README 與 changelog 裡就寫著這句規則說明，說明文件自己命中了偵測規則**，被列為 BOM 候選。已改成要求 **TAB 分隔的表頭列**（與 `parse_text_bom` 實際依據一致）。

教訓：偵測規則的文件如果跟被偵測的檔案放在同一個資料夾，規則必須嚴格到不會命中自己的說明文字。改完一定要印 `discover()` 的完整候選清單確認。

---

## 工作慣例

- 程式碼不主動產出，等 Stanley 明確說「請產出程式」才產出
- 回覆用**繁體中文**
- 版本 changelog 完整保留所有歷史版本說明，不省略改寫舊版內容
- 輸出的 xlsx 樣本僅供驗證用，非程式固定產物
- 「繼續」指令 = 「從最後出現的位置繼續產出，不要重複已完成的部分」

---

## ★ 交付前的實機測試清單（2026-09-03 起強制）

> **交付任何新版本或壓縮包之前，必須逐項實際執行過並貼出結果。**
> 「程式碼看起來對」「元件建得出來」都不算驗證過。
> 這條規則是因為 v0.07 首次交付的壓縮包五個 .bat 全部無法執行 ——
> 我當時只驗證了 .py，沒有真的去雙擊 .bat。

| # | 項目 | 怎麼算過關 |
|---|---|---|
| 1 | `run_CLI.bat` | 用**真實 cmd** 執行（不是 PowerShell 管線），中文與 ✓ 不亂碼、`pause` 有停住、產出 xlsx |
| 2 | `run_C_streamlit.bat` | 真實 cmd 啟動 → 8601 埠 LISTENING → **真實瀏覽器**開得起來 → 跑完偵測/比對/編輯/產出 |
| 3 | `run_D_webview.bat` | 視窗真的開起來並截圖；或用 `--selftest` 驅動整條流程並貼出結果 |
| 4 | `install_requirements.bat` | 真實 cmd 執行到最後的檢查區塊 |
| 5 | 每個 .bat 的位元組 | `CRLF > 0`、`lone-LF == 0`、`全部 ASCII` |
| 6 | 壓縮包 | **解壓到全新路徑**再跑一次 1～4（驗證「工作目錄＝程式所在目錄」） |
| 7 | 壓縮包內容 | 檔案數對、`.streamlit\config.toml` 有進去（PowerShell 壓縮常漏掉 dot 開頭資料夾） |
| 8 | 自動偵測 | 印出 `discover()` 的完整候選清單，確認**沒有把說明文件或產出檔誤認**成輸入 |

產生 .bat 的腳本放在 `tools\gen_bats.py`，**不要手寫 .bat**（原因見下一節）。

---

## 目錄現況

```
F:\AI_BOM\
  bom_core_0.07.py                 核心 + CLI
  bom_matcher_C_0.07.py            C 案 Streamlit
  bom_matcher_D_0.07.py            D 案 pywebview
  run_C_streamlit.bat              C 案捷徑
  run_D_webview.bat                D 案捷徑
  .streamlit\config.toml           C 案主題
  changelog_BOM_matcher_v0.07.txt  完整版本歷史（v0.01–v0.07）
  HANDOVER_bom_matcher.md          本文件
  bom_matcher_0.06.py              上一版（tkinter，保留備援）
  changelog_BOM_matcher_v0.06.txt
  pcb.jpg                          PCB 圖片
  TXT\                             Altium Text BOM
  ExcelDB\                         零件資料庫（BOM_20260525.xlsx）
  screenshot\                      PCB 圖片備援位置
  output\                          產出的 E-BOM
  demo document\                   ★ E-BOM 版面的視覺參考來源
  old\                             v0.01–v0.05
```
