# BOM Matcher — 新對話交接事項

## 背景
Stanley（ADATA Technology）開發的 Python Windows 工具，讀取 Altium 格式 Text BOM 與 Excel 零件資料庫，輸出 ADATA E-BOM 格式 Excel。

---

## 目前版本：v0.10

v0.07 起拆成「一個核心 + 兩套前端」。v0.06 的 tkinter GUI 已由 C / D 兩案取代。

| 檔案 | 角色 | 行數 |
|---|---|---|
| `bom_core_0.07.py` | 核心邏輯（解析／比對／Knowledge／Excel 輸出）＋ CLI | ~700 |
| `bom_matcher_C_0.07.py` | **S 案**：Streamlit 網頁版 | ~330 |
| `bom_matcher_D_0.07.py` | **W 案**：pywebview 原生視窗版（HTML 內嵌） | ~640 |
| `run_S_streamlit.bat` | S 案啟動捷徑（port 8601） | — |
| `run_W_webview.bat` | W 案啟動捷徑 | — |
| `.streamlit\config.toml` | S 案主題（primaryColor `#4F46E5`） | — |

**版本規則：** 每次進版 +0.01，程式檔名含版號。

### 執行方式
```
雙擊 run_S_streamlit.bat        # S 案：瀏覽器 http://localhost:8601
雙擊 run_W_webview.bat          # W 案：原生視窗

python bom_core_0.07.py                    # CLI 全自動
python bom_core_0.07.py -o out.xlsx
python bom_core_0.07.py BOM.txt db.xlsx -k Knowledge.xlsx -p pcb.jpg
python bom_core_0.07.py --no-split         # 關閉替代料拆列

python bom_matcher_D_0.07.py --dump-html   # 輸出 W 案 UI 的 HTML
python bom_matcher_D_0.07.py --selftest    # 自動跑完流程後關閉（驗證用）
```

### 依賴套件
```
pip install pandas openpyxl pillow                # 核心
pip install streamlit streamlit-option-menu       # S 案
pip install pywebview                             # W 案
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

### S 案 Streamlit
- **`st.data_editor` 不套用 pandas Styler 底色** → 「可編輯」與「三色底」無法並存。目前分兩個頁籤：`✎ 編輯`（emoji 狀態欄）／`🎨 色底檢視`（唯讀 Styler）。要兩者兼得得換 `streamlit-aggrid`。
- **`st.data_editor` 把 `None` / `pd.NA` / `Int64` 空值渲染成字串 `"None"`** → Item / Qty 兩欄改用字串表示，存回時再轉 int。
- 版面只能單欄由上而下，沒有原生檔案選擇器（用下拉選單＋手動輸入路徑替代）。

### W 案 pywebview
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

## v0.10 已完成（2026-10-21）

待辦 T1–T5 全部做完，詳情見 `changelog_BOM_matcher_v0.10.txt`。重點：

| 項目 | 結果 |
|---|---|
| T1　Location 加 (T)/(B) | `tag_side()` 冪等，手打過前綴不會疊成 `(T)(T)` |
| T2　R8 電容降壓選料 | SM8008 命中 45 → **55**；Total Qty 384 **不變** |
| T3　用語統一 | 位號→Location、面別→Top/Bot、chip→篩選鈕，三版掃過 |
| T4　Top/Bot 同圖提示 | 不再靜默照貼；檔名規則只用於自動偵測 |
| T5　SOP 標註式版面 | 31 頁，PPTX + HTML，`issues:[] deadLinks:0` |
| 同名檔覆蓋確認 | core 警告 + C 勾選框 + D confirm |
| 回歸 | 三份 BOM **零退步**，Total Qty 全部相同 |

### R8 為什麼只放寬電容耐壓

耐壓不足會縮短壽命但電路仍會動，看得到、可人工判斷。誤差與電流不符
是功能性錯誤 —— 放寬等於讓程式替工程師做決定，所以 R6 的這兩條不動。

### Streamlit 的 widget 值不是刪 session_state 就能重置

v0.09 以為 `st.session_state.pop('sel_xxx')` 就能讓 `index` 重新生效，
v0.10 實測發現**不夠**：widget 值綁在「widget 身分」上，同一個 key
再建出來還是沿用舊值（目錄瀏覽器選了 pcb.jpg，下拉選單仍停在
「手動輸入路徑」）。正解是在 key 後面接一個 `gen_xxx` 計數器，
選檔時 +1，讓它變成新的 widget。

另外，選完檔案**不可以**直接寫 `st.session_state['sel_xxx']` ——
那個 selectbox 在同一次 run 裡早就建好了，會噴
`StreamlitAPIException: cannot be modified after the widget ... is instantiated`。
要先存 `pend_xxx` 再 `st.rerun()`，在 selectbox 建立之前套用。

### v0.10 追加（2026-10-21 下午）

| 項目 | 結果 |
|---|---|
| Location 前綴上畫面 | C/D 表格的 Location-Top / Bot 也帶 (T)/(B)（Q1，已確認保留）|
| 藏掉 Deploy | `.streamlit\config.toml` `toolbarMode = "viewer"`（Q3）|
| ADATA LOGO | 比照 Firmware Tool，C/D 左上角換成 `adata.png`（蜂鳥）；找不到檔案時回退成「B」徽章 |
| 彩蛋 | 按住 **Shift** 連點三次 LOGO → 小蜜蜂遊戲。沒有開關、介面上沒有任何提示（比照 Firmware Tool）|
| SOP 對照檢查 | C/D 介面逐項比對 SOP，補齊／改正十幾處；S 案截圖全部重拍，新增兩頁 → 33 頁 |
| HTML 版 SOP 截圖全空白 | `preview.py` 修正（見下）|

### 彩蛋（小蜜蜂）的結構

- 遊戲本體與 LOGO 讀取都在 `bom_core` 第 14b 節：`BEE_GAME_JS`（定義
  `window.bmBeeGame(canvas, onExit)`，回傳 `stop()`）、`app_logo_uri()`、
  `EGG_CLICKS = 3`、`EGG_WINDOW_MS = 3000`。C/D 兩案共用同一份。
- W 案：`page_html()` 把 HTML 裡的 `__LOGO_TAG__` / `__VER__` /
  `/*__BEE_GAME__*/` 等佔位字串換掉，遊戲視窗是 `#egg` 覆蓋層。
- S 案：用 **Custom Components v2**（`st.components.v2.component('bm_brand')`）
  畫品牌區。連點計數全在瀏覽器端，不回 Python；遊戲覆蓋層掛在
  `document.body`，不受側邊欄版面與 Shadow DOM 樣式限制。
  ⚠️ 不要改用 `st.components.v1.html` —— v1 已棄用，而且是 iframe，
  覆蓋層蓋不到整個頁面。
- ⚠️ 照抄 Firmware Tool 的兩個坑：**監聽 pointerdown 不是 click**
  （`<img>` 原生拖曳會吃掉 click）、圖片 `draggable=false`。
  遊戲視窗一定要有 ✕，Esc 也要能關；關閉時呼叫 `stop()` 收掉
  requestAnimationFrame 與鍵盤監聽，否則方向鍵會繼續被吃。

### HTML 版 SOP 的截圖一直是空白的

`build.js` 在每張截圖上疊一個「白色、透明度 100%」的矩形當外框。
`preview.py` 的 `fill_css()` 只取顏色、沒取透明度 → 變成實心白塊，
**把底下的截圖整張蓋掉**。QA 的 `imgFail` 只檢查圖片有沒有載入，
所以一直回報 0，沒人發現。已改成讀 `a:alpha`，透明度 100% 就不輸出底色。

教訓：版面 QA 通過 ≠ 看得到。HTML 版至少要實際截一頁有截圖的來看。

### 截圖怎麼重拍（不必開真的視窗）

- **W 案**：載入 `bom_matcher_W_0.10.py`，用 `Api().run_match()` 拿實際狀態，
  `page_html()` 加一段假的 `window.pywebview.api`（回傳同一份 state），
  用 headless Chrome 截圖。舊截圖含 31px 的 Windows 標題列，
  SOP 的 marks 座標是照那個算的 —— 重拍時把舊圖的標題列貼回去，marks 就不用改。
- **S 案**：`streamlit run` 起一個實例，用 Chrome DevTools Protocol 點按鈕後截圖。
  寬度用 1440（1280 時「瀏覽…」按鈕會折成兩行），SOP 的 `ar` 改成 `1440 / 942`。
- ⚠️ 用 Chrome，不要用 Edge headless —— Edge 會在畫面右緣疊一個浮動圖示，
  不在 DOM 裡，CSS 藏不掉。

---

## v0.09 已完成（2026-10-19）

上一版的待辦全部做完，詳情見 `changelog_BOM_matcher_v0.09.txt`。重點：

| 項目 | 結果 |
|---|---|
| 比對規則 R1–R7 | SM8008 命中 26 → **45**；三份 BOM 回歸**零退步** |
| 單面板相容 | 與 v0.08 逐格比對 **0 差異**（合併數同為 107） |
| NC(...) 剝除 | SM8008 的 nc 由 0 → **26** |
| 雙面板 | F/G 兩欄、位號對照表、三層警告、圖片切兩半 |

### 比對的三條鐵則（之後改比對邏輯前先看這個）

1. **不要用整串 Part 去 Description 找子字串。** 格式一變就全找不到，
   而且是靜默失敗。要先 `parse_part()` 拆成屬性再比。
2. **比對順序是「屬性式 → 子字串」，不可對調。** 子字串是 v0.08 的舊路徑，
   留著當退路，保證舊 BOM 不退步。改動時務必跑回歸（見下）。
3. **未命中一定要歸因（R7）。** 靜默留白是 v0.08 最大的問題 ——
   32 項電容全軍覆沒卻沒人發現。

### R6 方向性條件容易誤判為 bug

耐壓／電流 **DB ≥ BOM**、誤差 **DB ≤ BOM**。所以會出現
「DB 明明有這個值卻沒命中」的情況，那多半是**正確的**：

- `NC(100R_1%)` 不採用 `61031353`（5% 誤差超過要求的 1%）
- `FB/120R/55mR/2A` 不採用 120Ω BEAD（DB 只有 1.2~1.4A，不足 2A）

診斷會把這類歸到「規格不符（耐壓/電流/誤差）」，與「DB 沒有這個值」分開。

### 回歸測試（改比對邏輯後必跑）

```bash
python tools/regress.py        # v0.08 vs v0.09，三份 BOM，任一退步就算失敗
```

判定標準：**命中數只能增不能減，Total Qty 必須完全相同。**

---

## 雙面板

| 項目 | 說明 |
|---|---|
| 觸發條件 | 有載入位號對照表就是雙面板；沒有就退回單面板，輸出與 v0.08 相同 |
| 對照表格式 | **第一欄＝位號、第二欄＝面別**。xlsx 為主、txt 為輔 |
| 不可用表頭名判斷 | 實際檔案 TXT 是 `T/B`、XLSX 是 `T_B`，兩者不一致 |
| 面別寫法 | TOP/T/TOP SIDE/F/FRONT、BOT/B/BOTTOM/BACK 都接受 |
| 放哪裡 | 程式目錄或 `ExcelDB\`（`find_side_tables` 掃這兩處） |
| PCB 圖片 | 檔名含 top / bot 才會切兩半，例 `pcb_top.jpg` / `pcb_bot.jpg`；只有一張就不切 |
| 位號對不上 | 日誌＋Excel 淡紅底＋紅字；反向檢查也提示 |

⚠️ **對照表與 BOM 必須是同一個改版。** SM8008 實測兩份差了 6+4 個位號，
這通常代表版本不同步，不是單純「少幾個位號」。

---

## 交付規格（v0.08 起，每次更新都做）

> **★ 2026-10-21 起：每次產出新版本，一定要連 `dist\` 底下的
> 「完整離線安裝版」與「精簡版」一起重建。**
> 只更新 .py 不重建交付包，使用者拿到的還是舊版 —— 他們是從
> `dist\` 解壓縮來用的，不是跑你專案目錄裡的檔案。
> 指令：`python tools/build_dist.py`

規格對齊 `F:\AI_ini_allin1` 的 ADATA SSD Firmware Tool。

| 交付包 | 大小 | 內容 | 用在哪 |
|---|---|---|---|
| `dist\BOM_Matcher_v<VER>_完整離線安裝版` | ~113 MB | 精簡版 ＋ `python\`（Python 安裝檔）＋ `wheels\`（47 個離線套件） | 產線機器、連不上外網 |
| `dist\BOM_Matcher_v<VER>_精簡版` | ~0.6 MB | 程式 ＋ 範例資料 ＋ SOP | 辦公室機器，安裝時上網抓套件 |

兩包都含 `安裝.bat`、`安裝說明.txt`、`docs\*_操作SOP.pptx`。
程式本體與範例資料完全相同，差別只在「套件怎麼來」。

```bash
python tools/build_dist.py                   # 產出兩包
python tools/build_dist.py --refresh-wheels  # 先重抓 wheels 再打包
python tools/build_dist.py --slim-only       # 只要精簡版
```

⚠️ `proxy_tools`（pywebview 的相依）**只有 sdist 沒有 wheel**，所以
`pip download` 不能加 `--only-binary=:all:`，而且 wheels 裡必須一併放
`setuptools` 與 `wheel`，否則目標機器離線時建不起來。

---

## 操作 SOP

產出物：`docs\BOM_Matcher_v<VER>_操作SOP.pptx`（24 頁）

| 檔案 | 作用 |
|---|---|
| `docs/sop/content.js` | **全部文字內容**。要改字只動這一支 |
| `docs/sop/layout.js` | 投影片尺寸與邊界常數 |
| `docs/sop/build.js` | 產出 .pptx（pptxgenjs） |
| `docs/sop/preview.py` | **從產出的 .pptx 反向解析實際座標**畫成 preview.html |
| `docs/sop/shots/` | 9 張實機截圖（W 案用 PrintWindow 擷取視窗，非螢幕截圖） |

```bash
cd docs/sop && npm install pptxgenjs     # 第一次
node build.js
node <pptx-skill>/scripts/apply_theme.js ../BOM_Matcher_v0.08_操作SOP.pptx theme.tmp.json
python preview.py                        # 再用瀏覽器開 preview.html
```

`preview.py` 刻意**不是**重畫一份版面 —— 重畫的預覽只能證明「我以為的座標」
沒爆框，證明不了 pptx 裡真正的座標。它直接讀 pptx 的 EMU 座標畫成 HTML，
載入後的 JS 會自動掃描「超出投影片邊界」與「文字爆框」並標紅。

這台機器沒有 LibreOffice，所以不能用 pptx skill 文件裡的
`soffice --convert-to pdf` + `pdftoppm` 做視覺 QA，preview.html 是替代方案。

---

### SOP 版面固定規格（2026-10-21 指定，不要再改掉）

| 位置 | 內容 |
|---|---|
| 頁尾 | **只有「☰ 目錄」，置中**。版號、日期、頁碼都不要 |
| 截圖頁 | 方案 A 編號標記法：截圖上疊紅色編號圓點，右側說明用 ①②③ 同號對應 |

`build.js` 的 CONTENT 版面 `objects: []`、沒有 `slideNumber`；
「☰ 目錄」在 `title()` 裡用 `x: L.M, w: L.W - L.M*2, align:'center'`
畫在**投影片**上（不能放版面，pptxgenjs 在 slideLayout 寫不出
`hlinkClick` 的關聯，會變成 `rIdundefined`）。

⚠️ 這條曾經被弄丟過一次：Stanley 要求「刪掉版號與日期、目錄置中」之後，
下一輪又說「排版改回上一個方案 A」，我把頁尾一起還原了 ——
但「改回排版」指的只是截圖與說明的配置，頁尾那兩項是獨立要求、沒有被收回。
以後收到「改回上一版」時，要先確認**範圍到哪裡**，不要擴大解釋。

### 兩個對話的合併（2026-10-21）

`F:\AI_BOM` 一度有兩個對話同時在改：

| 對話 | 內容 |
|---|---|
| **BOM matcher**（主線，08-11 起） | v0.01→v0.10 全部歷史與規則，git 分支由這邊管 |
| Location tags v0.10 進行中工作（10-21 開） | v0.10 的 Location 前綴上畫面、Deploy 關閉、LOGO／彩蛋、SOP 對照介面重做 |

後者的 21 個檔已驗過並合進 commit `84f5263`，知識寫進上面的
「v0.10 追加」一節。**之後只用主線那個對話**，另一個封存。

教訓：同一個資料夾不要開兩個對話同時改，誰先存檔誰就蓋掉對方。

---

## 待辦（下一版，等 Stanley 指定版號才動工）

v0.10 的 T1–T5 全部完成，目前沒有待辦。

### 待決定項目

| # | 項目 | 提出日 | 狀態 |
|---|---|---|---|
| Q1 | C/D 介面 Location 欄也顯示 (T)/(B) | 2026-10-21 | **Stanley 確認保留**（2026-10-21）|
| Q2 | 真正的 TOP / BOT 兩張 PCB 圖 | 2026-10-21 | 等提供檔案 |
| Q3 | Streamlit 右上角 Deploy 與 ⋮ 選單 | 2026-10-21 | **已處理**（toolbarMode = viewer，SOP 已補說明）|

**Q1　Location 欄顯示前綴 —— 已確認保留**
S 案 `_rows_to_df()`、W 案 `_rows_js()` 都套上 `core.tag_side()`，
畫面與輸出 Excel 一致。改過的值寫回 `ref_top` / `ref_bot` 會帶著前綴，
寫 Excel 時 `tag_side()` 冪等，不會疊成 `(T)(T)`。

修改前後對照（SM8008 實際資料）：
`screenshot\v0.10_D_Location前綴_修改前後.png`、
`screenshot\v0.10_C_Location前綴_修改前後.png`

★ 代價：每個 Location 多 3 個字元，欄寬不變的話一格看得到的會變少
（例：Item 3 的 Location-Top 完整看得到的從約 7 個變 4 個）。
Stanley 看過預覽圖後決定保留。要改回只要拿掉那兩處 `core.tag_side()`。

**Q2　TOP / BOT 兩張 PCB 圖 —— 等檔案**
目前 TOP / BOT 都指到同一張 `pcb.jpg`，雙面板圖片區上下兩半是同一張圖
（程式會提醒「Top 與 Bot 指到同一張圖片」）。
拿到後放進 `screenshot\`，檔名含 top / bot（例 `pcb_top.jpg` / `pcb_bot.jpg`）
就會自動配對；之後要重拍 SOP 雙面板相關截圖並重建交付包。

**Q3　Streamlit 右上角 Deploy 與 ⋮ 選單 —— 已處理**
`.streamlit\config.toml` 加了 `[client] toolbarMode = "viewer"`，
把 Deploy 按鈕與開發者選項（Rerun / Clear cache）藏起來。
Deploy 是把程式發佈到 Streamlit 雲端，對內部工具沒有用途，
而且有外流原始碼與資料的風險。
實測 Streamlit 1.62：⋮ 選單只剩 **Print** 與 **Record screen**，
SOP 第 14 頁「右上角的選單（三個點）」照這兩項寫。

★ 這個設定只有在交付包裡的 `.streamlit\config.toml` 一起帶出去才有效 ——
打包時務必確認這個資料夾有進去（PowerShell 壓縮常漏掉 dot 開頭資料夾，
見交付前檢查清單第 7 項）。

---

## 用語規範（2026-10-20 起，Stanley 指定）

**只准用輸出檔裡已經存在的欄位名稱，不要另外發明中文名詞。**

| 要講的東西 | ✅ 用這個 | ❌ 不要用 |
|---|---|---|
| P/N 欄位的值（料件編號） | **料號** 或 `P/N` | 零件號、品號 |
| `R48` `U13` 這種標示（Location 欄的值） | **Location** | 位號、參考編號、reference designator |
| 正反面 | **Top / Bot**，或欄名全寫 `Location-Top Layer` / `Location-Bottom Layer` | 面別、面號、正反面別 |

適用範圍：對話回覆、程式內的訊息與警告文字、Excel 欄位與圖例、SOP、changelog，全部一致。

例：
- ❌ `⚠ 4 個位號查不到面別：R48,R54,R298,R299`
- ✅ `⚠ 4 個 Location 查不到 Top/Bot：R48,R54,R298,R299`

---

## ★ 三個版本必須同步（2026-10-20 起強制）

> **任何需求、修正、用語改動，一律同時套用到三個版本，並三個都實測過才算完成。**

| 版本 | 檔案 | 怎麼算過關 |
|---|---|---|
| CLI | `run_CLI.bat` → `bom_core_0.XX.py` 的 `main()` | 真實 cmd 跑完，訊息與產出都正確 |
| Streamlit（S 案） | `bom_matcher_C_0.XX.py` | 真實瀏覽器操作過該功能 |
| pywebview（W 案） | `bom_matcher_D_0.XX.py` | 視窗開起來截圖，或 `--selftest` 跑完並貼結果 |

共用邏輯一律放 `bom_core`，前端只做 UI；但**前端的文字、欄位、提示也要一起改**，
不能只改 core 就宣稱三版都好了。回報時三個版本的執行結果分開貼。

---

### 程式進版必須同步更新 SOP（2026-10-20 起強制）

版號一動，SOP 就要跟著走完這條流程，**不可以只改程式不改 SOP**：

1. 重新截圖：介面有變的頁面重拍，放進 `docs\sop\shots\`
2. 改文字：只動 `docs\sop\content.js`（`VER` / `DATE` 也要改）
3. 重建：`node docs/sop/build.js` → 產出 `docs\BOM_Matcher_vX.XX_操作SOP.pptx`
4. 驗證：pptx validator 必須 All validations PASSED
5. 交付包只放 **PPTX**，目錄超連結要能跳

### ★ HTML 版 SOP 已停止維護（2026-10-21 起）

`docs\sop\preview.py` 留在版控裡當參考，但**不再執行、不再進交付包**。
現有的 `BOM_Matcher_v0.10_操作SOP.html` 已經搬到 `old\`。

原因：同一份內容維護兩種格式，等於每次都要驗兩次版面。
而且 HTML 版是用 python-pptx 反解 pptx 再重畫的，pptx 的細節
（形狀填色、透明度、段落對齊…）要一項一項補 —— 已經為此修過三次 bug
（挑錯版本的 deck、沒讀透明度導致截圖全白、沒讀段落對齊導致置中失效）。
PPTX 本身就能看、能投影、能列印，不值得為第二種格式付這個成本。

SOP 的截圖以 **W 案（pywebview）** 為主，S 案只放差異頁。
所有截圖頁都要有編號標記（方案 A），`build.js` 會在建置時檢查
「標記數 == 說明條目數」，不相等直接中斷。

---

### ★ 程式目錄只留最新版（2026-10-21 起）

根目錄與 `docs\` 只保留**目前版本**的檔案，其餘全部搬到 `old\`：

```
根目錄   bom_core_0.10.py  bom_matcher_S_0.10.py  bom_matcher_W_0.10.py
         changelog_BOM_matcher_v0.10.txt  HANDOVER_bom_matcher.md
         requirements.txt  五個 .bat  pcb.jpg  adata.png
docs\    BOM_Matcher_v0.10_操作SOP.pptx
```

搬動後記得檢查有沒有程式還指著舊路徑 ——
`tools
egress.py` 會載入前一版的 core 做回歸，路徑要跟著改成 `old/`。

---

### ★ 給方案選擇時一定要附預覽圖（2026-10-21 起）

只要是「要你選 a / b / c」或「甲案 / 乙案」這種，**不可以只用文字描述**，
一定要先做出預覽圖讓 Stanley 看到長什麼樣子再決定。

適用範圍：介面版面、SOP 版面、輸出檔格式、互動方式 —— 只要看得見的都算。

做法（不需要動到正式程式）：
1. 用 HTML 模擬各方案的畫面，放在暫存目錄
2. `python -m http.server <port>` 起一個本機伺服器
3. headless Chrome 截圖：
   `chrome.exe --headless=new --window-size=W,H --screenshot=out.png http://127.0.0.1:<port>/x.html`
4. 存到 `screenshot\`，檔名帶版號與用途

W 案的真實截圖用 `tools\shots_S.py` 的 CDP 作法；
S 案（Streamlit）一定要用 CDP，因為畫面要互動才出得來。

---

### ★ 檔案被佔用時要停下來問（2026-10-21 起）

重建 SOP 或交付包時如果遇到 `EBUSY` / `Permission denied`
（多半是 OnlyOffice / PowerPoint / Excel 開著），**不要自己想辦法繞過，
也不要刪掉 `.~lock` 檔**：停下來問 Stanley。
**流程是：他關檔 → 我存檔。** 不要自己覆蓋。

理由：那是他正在看的檔案，直接覆蓋等於把他眼前的內容換掉；
刪 lock 檔更糟 —— 他一存檔就把剛建好的覆蓋回去。

`docs\sop\build.js` 有 `SOP_OUT` 環境變數，可以先建到別的路徑驗證版面，
等他關檔再寫正式檔。

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
| 2 | `run_S_streamlit.bat` | 真實 cmd 啟動 → 8601 埠 LISTENING → **真實瀏覽器**開得起來 → 跑完偵測/比對/編輯/產出 |
| 3 | `run_W_webview.bat` | 視窗真的開起來並截圖；或用 `--selftest` 驅動整條流程並貼出結果 |
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
  bom_matcher_C_0.07.py            S 案 Streamlit
  bom_matcher_D_0.07.py            W 案 pywebview
  run_S_streamlit.bat              S 案捷徑
  run_W_webview.bat                W 案捷徑
  .streamlit\config.toml           S 案主題
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
