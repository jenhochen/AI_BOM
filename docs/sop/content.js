// BOM Matcher 操作 SOP —— 全部文字內容。要改字只動這一支。
// 與 F:\AI_ini_allin1\docs\sop 同樣的作法：content / build / preview 三分。

const VER = '0.10';
const DATE = '2026-10-20';

const THEME = {
  name: 'BOM Matcher',
  headFontFace: 'Microsoft JhengHei',
  bodyFontFace: 'Microsoft JhengHei',
  colors: {
    dk1: '1A1C2E', lt1: 'FFFFFF',
    dk2: '3C4252', lt2: 'F7F8FA',
    accent1: '4F46E5',   // 程式本身的主色
    accent2: '0F9D58',   // 綠：含替代料
    accent3: 'E8863A',   // 橘：Knowledge 命中
    accent4: 'E8B931',   // 黃：原本為 NC
    accent5: '6B7180',   // 次要文字
    accent6: 'E8EAED',   // 邊框
    mark: 'D92D20',      // 截圖上的編號標記（方案 A）
    accent7: '7C6CF0',   // 紫：R8 降壓選料
    hlink: '4F46E5', folHlink: '6B7180',
  },
};

const SHOTS = 'shots/';

const SLIDES = [
  // ── 封面 ────────────────────────────────────────────────────────────
  { kind: 'cover', section: '封面',
    title: 'BOM Matcher',
    sub: `操作 SOP　v${VER}`,
    meta: [`ADATA Technology　${DATE}`,
           'Altium Text BOM ＋ 零件資料庫 → ADATA E-BOM Excel'] },

  // ── 一、認識工具 ────────────────────────────────────────────────────
  { kind: 'divider', section: '認識工具', n: '01', title: '這個工具在做什麼' },

  { kind: 'flow', section: '認識工具', title: '五個輸入，一個輸出',
    inputs: [
      { k: 'Text BOM', v: 'Altium 匯出\n*.txt 或 *.BOM', c: 'accent1' },
      { k: '零件資料庫', v: 'P/N + Description\n*.xlsx', c: 'accent1' },
      { k: 'Knowledge', v: '品名 / 包裝 / 注意事項\n*.xlsx', c: 'accent3' },
      { k: 'PCB 圖片', v: 'pcb.jpg / png\n嵌入 rows 9-35', c: 'accent5' },
    ],
    out: 'ADATA E-BOM 格式 Excel',
    outNote: '表頭、PCB 圖、BOM 明細、Total、表尾一次做好',
    note: '三種介面（CLI / C 案 / D 案）共用同一份核心邏輯，\n' +
          '同一份輸入產出的 Excel 完全相同。' },

  { kind: 'cards', section: '認識工具', title: '三種介面怎麼選',
    cards: [
      { t: 'run_CLI.bat', s: '命令列',
        b: ['不需要 streamlit / pywebview',
            '最快、最穩',
            '不能人工修正，直接出檔',
            '第一次安裝完建議先跑這支確認環境'] },
      { t: 'run_C_streamlit.bat', s: 'C 案．瀏覽器',
        b: ['在瀏覽器操作，localhost:8601',
            '可編輯表格後再產出',
            '色底需切到「色底檢視」頁籤',
            '使用期間不能關掉黑色主控台'] },
      { t: 'run_D_webview.bat', s: 'D 案．原生視窗',
        b: ['雙擊即開，不需要瀏覽器',
            '三欄式，右側即時顯示明細',
            '色底與編輯可以並存',
            '篩選鈕：未命中 / NC / 替代料 / 需確認'] },
    ] },

  // ── 二、安裝 ────────────────────────────────────────────────────────
  { kind: 'divider', section: '安裝', n: '02', title: '安裝' },

  { kind: 'compare', section: '安裝', title: '兩種交付包怎麼選',
    left: { t: '完整離線安裝版', s: '約 112 MB',
            b: ['自帶 Python 安裝檔與全部套件',
                '安裝過程完全不需要網路',
                '產線機器、連不上外網的環境用這包',
                '多了 python\\ 與 wheels\\ 兩個資料夾'] },
    right: { t: '精簡版', s: '約 0.2 MB',
             b: ['安裝時要連得上 python.org 與 PyPI',
                 '檔案小，好傳',
                 '辦公室機器用這包就夠',
                 '裝不起來時再改用完整離線版'] },
    note: '兩包的程式本體與範例資料完全相同，差別只在「套件怎麼來」。' },

  { kind: 'steps', section: '安裝', title: '安裝步驟',
    steps: [
      { n: '1', t: '解壓縮到任意資料夾',
        d: '路徑不限。程式一律以「自己所在的資料夾」為工作目錄，\n不必固定放哪裡。' },
      { n: '2', t: '執行「安裝.bat」',
        d: '已經有 Python 會自動跳過，直接裝套件。' },
      { n: '3', t: '看到「請重開視窗」就再跑一次',
        d: 'Python 剛裝好時 PATH 還沒生效，關掉視窗再執行一次\n「安裝.bat」把套件裝完。' },
      { n: '4', t: '確認三行都是 [OK]',
        d: '[OK] 核心 / [OK] C 案 / [OK] D 案。\n出現 [--] 表示該介面不能用，其餘仍可運作。' },
    ] },

  // ── 三、C 案 ────────────────────────────────────────────────────────
  { kind: 'divider', section: 'C 案操作', n: '03', title: 'C 案　瀏覽器介面' },

  { kind: 'shot', section: 'C 案操作', title: '步驟 1　檔案設定',
    img: SHOTS + 'C_01_files.jpg', ar: 800 / 523,
    bullets: [
      '四個檔案由程式自動偵測，候選依**最後修改時間**排序，最新的排第一個',
      '偵測是看「內容」不是看檔名 —— 把產出的 E-BOM 丟回資料夾不會被誤認',
      '不對就按右邊的「瀏覽…」自己選',
      '輸出路徑預設在 output\\，副檔名一律強制 .xlsx',
    ],
    warn: '確認 Text BOM 是你這次要跑的那一支。\n選錯檔案是最常見的失誤。' },

  { kind: 'shot', section: 'C 案操作', title: '檔案不在預設清單裡',
    img: SHOTS + 'C_02_browser.jpg', ar: 800 / 523,
    bullets: [
      '展開「📂 目錄瀏覽器」就能自己翻資料夾',
      '左邊是子資料夾，右邊是符合副檔名的檔案，後面附最後修改日期',
      '「⬆ 上一層」往上、「🏠 程式目錄」跳回程式所在位置',
      '點檔名即選取，上方的下拉選單會跟著換',
    ],
    warn: '「瀏覽…」開的是 Windows 原生對話框，偶爾會躲在瀏覽器視窗後面。\n找不到就改用目錄瀏覽器，功能一樣。' },

  { kind: 'shot', section: 'C 案操作', title: '步驟 2　比對結果',
    img: SHOTS + 'C_03_rows.jpg', ar: 800 / 523,
    bullets: [
      '上方六個數字：總項次 / 輸出列數 / 命中 / 未命中 / Knowledge / Total Qty',
      '表格可直接改 P/N、RC/EC、Description、Qty、Location',
      '改完按「儲存修改」才會帶進產出的 Excel',
      '狀態欄的 🟠🟡🟢⚪ 對應輸出 Excel 的橘／黃／綠底',
    ],
    warn: 'C 案的「✎ 編輯」頁籤不顯示色底 —— 這是 Streamlit 的限制。\n要看色底請切到「🎨 色底檢視」頁籤（唯讀）。' },

  { kind: 'shot', section: 'C 案操作', title: '步驟 3　產出 E-BOM',
    img: SHOTS + 'C_04_export.jpg', ar: 800 / 523,
    bullets: [
      '按「▶ 產出 Excel」，檔案直接寫到輸出路徑',
      '下方日誌會列出 PCB 圖片尺寸、實際輸出路徑、項數與 Total Qty',
      '「⤓ 下載 xlsx」可另存一份到瀏覽器的下載目錄',
    ],
    warn: '最後一行的「113 項（含PCB）→ 138 列」要對得上預期。\n列數比項數多，是因為替代料被拆成獨立列。' },

  // ── 四、D 案 ────────────────────────────────────────────────────────
  { kind: 'divider', section: 'D 案操作', n: '04', title: 'D 案　原生視窗' },

  { kind: 'shot', section: 'D 案操作', title: '步驟 1　檔案設定',
    img: SHOTS + 'D_01_files.png', ar: 1280 / 820,
    bullets: [
      '欄位與 C 案相同，每個都有「瀏覽…」',
      '**Location TOP/BOT 對照表**：選填，給了就切換成雙面板模式',
      '**PCB 圖片（BOT 面）**：雙面板時提供兩張，圖片區自動切成上下兩半',
      '「替代料拆列」打勾時一列只有一個 P/N；取消勾選會回到 a;b 單列格式',
      '下方黑色區塊是執行日誌',
      '按「▶ 執行比對」開始',
    ],
    warn: '標題列會標出啟動用的 .bat：\nBOM Matcher v0.08　[run_D_webview.bat]' },

  { kind: 'shot', section: 'D 案操作', title: '步驟 2　比對結果與人工修正',
    img: SHOTS + 'D_02_rows.png', ar: 1280 / 820,
    marks: [[1, 300, 329], [2, 390, 457], [3, 300, 259], [4, 300, 791]],
    bullets: [
      '點任何一列，右側會顯示該料件的 Footprint、RC/EC、P/N、比對狀態',
      '未命中的列顯示「— 無命中 —」，直接在格子裡補料號',
      '表格上方的**篩選鈕**：全部 / 未命中 / NC / 替代料 / Knowledge / 需確認 / Location 未知',
      '底色與輸出 Excel 一致，編輯與色底可以並存',
    ],
    warn: '改了 Qty，右上角的 Total Qty 會立刻跟著變，\n可以用來確認改對了。' },

  { kind: 'shot', section: 'D 案操作', title: '步驟 3　Knowledge 命中確認',
    img: SHOTS + 'D_03_knowledge.png', ar: 1280 / 820,
    bullets: [
      '列出所有命中 Knowledge 的料件與對應注意事項',
      '命中條件三選一：P/N 相同、品名是描述的子字串、包裝是描述的子字串',
      '判斷不適用就按「✕ 忽略此筆」',
      '忽略後該列不會標橘底，Note 也不會寫進 Excel',
    ],
    warn: '包裝比對是子字串，可能誤命中。\n例如品名「TVS」會命中講 SOD-523 的條目，要人工判斷。' },

  { kind: 'shot', section: 'D 案操作', title: '步驟 4　表頭資訊',
    img: SHOTS + 'D_04_header.png', ar: 1280 / 820,
    bullets: [
      '對應輸出 Excel 的 rows 3–8',
      'PCB板號 / PCB料號 / RC/EC NO. / PCB供應商 / PCB 規格',
      '發行序號 / 發行(改版)日期 / 目前版次 / Support IC Type',
      '留空就是空白，不影響其他功能',
    ] },

  { kind: 'shot', section: 'D 案操作', title: '步驟 5　產出 E-BOM',
    img: SHOTS + 'D_05_export.png', ar: 1280 / 820,
    bullets: [
      '確認輸出路徑後按「▶ 產出 Excel」',
      '「開啟輸出資料夾」會直接叫出檔案總管',
      '底部動作列隨時可以產出，不必回到這一頁',
    ] },

  // ── 五、雙面板 ──────────────────────────────────────────────────────
  { kind: 'divider', section: '雙面板', n: '05', title: '雙面板　TOP / BOT' },

  { kind: 'shot', section: '雙面板', title: 'Location 查不到 Top/Bot 時',
    img: SHOTS + 'D_06_unknown.png', ar: 1280 / 820,
    marks: [[1, 722, 259], [2, 300, 344], [3, 862, 356], [4, 1080, 530]],
    bullets: [
      '表格上方的「Location 未知」篩選鈕可以只看有問題的列',
      '該列**淡紅底**，並在 Location-Bot 欄寫出是哪幾個 Location',
      'Excel 的 H:O 欄也會寫同一段紅字警告',
      '執行日誌會列出受影響的 Item、Qty 與 Location',
    ],
    warn: 'Item 77 的 4 個 Location 全部查不到 —— 若靜默略過，\n'
        + 'TOP 和 BOT 都會是 0，4 顆電阻等於從兩面同時消失。' },

  { kind: 'cards', section: '雙面板', title: '雙面板的三個規則',
    cards: [
      { t: 'Location 對照表', s: 'xlsx 為主．txt 為輔',
        b: ['第一欄＝Location，第二欄＝Top/Bot',
            '不看表頭名稱（實際檔案 TXT 是 T/B、XLSX 是 T_B）',
            'TOP/T/BOT/B/BOTTOM 都接受',
            '放程式目錄或 ExcelDB\ 底下'] },
      { t: 'Excel 欄位', s: '維持 15 欄（A:O）',
        b: ['F＝Location-Top Layer',
            'G＝Location-Bottom Layer',
            'H:O＝Knowledge Note',
            'Qty 不變，TOP+BOT 各自計數'] },
      { t: 'PCB 圖片', s: '檔名含 top / bot',
        b: ['pcb_top.jpg ＋ pcb_bot.jpg',
            '圖片區切成上下兩半（rows 9-21 / 22-35）',
            '只有一張圖就不切，維持整塊',
            '沒有對照表＝單面板，與 v0.08 相同'] },
    ] },

  // ── 六、電容降壓選料 R8（v0.10 新增）──────────────────────────────
  { kind: 'divider', section: '降壓選料', n: '06', title: '電容降壓選料　R8' },

  { kind: 'shot', section: '降壓選料', title: '耐壓低於原始設計的替代料',
    img: SHOTS + 'D_07_r8.png', ar: 1280 / 820,
    marks: [[1, 683, 254], [2, 300, 344], [3, 530, 344], [4, 520, 791]],
    bullets: [
      '表格上方的「需確認」篩選鈕只看這一類，數字是筆數',
      '容值與封裝一致、**DB 耐壓低於設計值**的電容才會進這一類',
      'Description 自動加註 **(原始設計為10V, 請再確認)**，原始耐壓寫在裡面',
      '這些列標**紫底**，Excel 的圖例也會多一列說明',
    ],
    warn: '只放寬電容的耐壓。電流與誤差不放寬 ——\n'
        + '耐壓不足看得到、可人工判斷；誤差或電流不符是功能性錯誤。' },

  { kind: 'steps', section: '降壓選料', title: 'R8 怎麼判、怎麼確認',
    steps: [
      { n: '1', t: '先找規格完全符合的',
        d: '耐壓 ≥ 設計值、誤差 ≤ 設計值、電流 ≥ 設計值。找得到就用，不會走 R8。' },
      { n: '2', t: '一個都沒有才考慮降壓品',
        d: '同容值、同封裝，但耐壓低於設計值。耐壓最接近設計值的排第一個。' },
      { n: '3', t: '註記寫進 Description',
        d: '例：A/CAP 4.7uF 0402 6.3V M(+-20%) X5R (原始設計為10V, 請再確認)' },
      { n: '4', t: '人工確認',
        d: '執行日誌會把每一筆列出來。確認可用就留著，不可用就在格子裡清掉料號。' },
    ] },

  // ── 六、看懂輸出 ────────────────────────────────────────────────────
  { kind: 'divider', section: '看懂輸出', n: '07', title: '看懂產出的 Excel' },

  { kind: 'legend', section: '看懂輸出', title: '色底規則與替代料拆列',
    legend: [
      { c: 'E8863A', bg: 'FFE0B2', t: '橘底', d: 'Knowledge 命中（優先序最高）' },
      { c: 'E8B931', bg: 'FFF2CC', t: '黃底', d: '原本為 NC 的零件' },
      { c: '0F9D58', bg: 'E2EFDA', t: '綠底', d: '含替代料（已拆列）' },
      { c: '4F46E5', bg: 'D9E1F2', t: '藍底', d: 'Total 合計列' },
      { c: '7C6CF0', bg: 'EDE9FE', t: '紫底', d: 'R8 降壓選料，Description 已註記，請人工確認' },
      { c: 'C00000', bg: 'FDE1E1', t: '紅底', d: '有 Location 查不到 Top/Bot（雙面板專用）' },
    ],
    split: {
      title: '替代料拆列（v0.07 起）',
      before: 'v0.06　一列　P/N = 61041115;61041238',
      after: ['v0.08　主列　Item=8  Qty=3  Location=C2,C3,C6  P/N=61041115',
              '　　　　續列　Item=空 Qty=空 Location=空　　　 P/N=61041238'],
      note: '續列的 Description 是該替代料自己的 DB 描述。\n' +
            'Total Qty 只加總主列，不會因為拆列而重複計算。',
    } },

  { kind: 'table', section: '看懂輸出', title: '資料夾規則與自動偵測',
    head: ['資料夾', '放什麼', '怎麼認'],
    rows: [
      ['TXT\\', 'Altium Text BOM（*.txt / *.BOM）',
       '要有一列是 TAB 分隔的表頭，欄位含 Item / Quantity / Reference'],
      ['ExcelDB\\', '零件資料庫 + Knowledge（*.xlsx）',
       '零件庫第 1 列要有 P/N 與 Description；Knowledge 要有 品名 與 包裝'],
      ['screenshot\\', 'PCB 圖片（*.jpg / *.png）', 'pcb.* 優先，其餘依最後修改時間'],
      ['output\\', '產出的 E-BOM', '不存在時輸出到程式所在資料夾'],
    ],
    note: '四類候選都依「最後修改時間」排序，最新的排第一個並預設選中。\n' +
          '這些子資料夾都是選用的，找不到就退回程式所在資料夾。' },

  // ── 六、疑難排解 ────────────────────────────────────────────────────
  { kind: 'divider', section: '疑難排解', n: '08', title: '疑難排解' },

  { kind: 'qa', section: '疑難排解', title: '常見問題',
    qa: [
      { q: '雙擊 .bat 一閃就關', a: '沒裝 Python 或沒加入 PATH。先跑「安裝.bat」看訊息。' },
      { q: '雙擊 .py 沒反應 / 閃退',
        a: 'C 案的 .py 必須用 streamlit 啟動，一律走 run_C_streamlit.bat。' },
      { q: 'C 案瀏覽器打不開',
        a: '8601 埠被占用。改用 --server.port 8602 另開一個埠。' },
      { q: 'D 案視窗開不起來',
        a: '用 run_D_webview_debug.bat 看訊息，多半是缺 Edge WebView2 Runtime。' },
      { q: '說找不到零件資料庫',
        a: '檢查 xlsx 第 1 列是不是「P/N」與「Description」這兩個字。' },
      { q: '產出的檔案別的工具打不開',
        a: 'v0.08 起副檔名強制 .xlsx，不會再有 .xls 的問題。' },
    ] },

  { kind: 'final', section: '疑難排解', title: `v${VER} 的重點`,
    points: [
      ['屬性式比對', 'Part 拆成值／耐壓／誤差／電流再比，不再整串找子字串'],
      ['NC(...) 認得出來', 'SM8008 有 26 項括號式 NC，v0.08 全部漏判'],
      ['未命中會說原因', '不再靜默留白 —— 格式變了當場就看得到'],
      ['雙面板 Top / Bot', 'F、G 兩欄＋兩張 PCB 圖＋Location 對不上三層警告'],
      ['Location 加 (T)/(B)', 'Location-Top / Location-Bottom 每個值都冠上 (T) 或 (B)'],
      ['R8 電容降壓選料', '耐壓不足可選，Description 註記「請再確認」並標紫底'],
      ['SM8008 命中 26 → 45', '三份 BOM 回歸測試零退步'],
    ],
    tail: '完整清單見 changelog_BOM_matcher_v0.10.txt' },
];

module.exports = { VER, DATE, THEME, SLIDES };
