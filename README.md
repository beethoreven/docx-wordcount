# 中文

## docx-wordcount — .docx 字數統計工具

算 `.docx` 的字數／字元數／段落數，數字跟 Microsoft Word「校閱 → 字數統計」對話框**完全一樣**（唯一刻意的例外是手動換行與選擇性連字號，工具會把它們列出來，見[已知限制](#已知限制)），但不需要開 Word。有零依賴的命令列版，也有單一 `.exe`、不需安裝不需網路的 Windows GUI 版。這份文件分成兩個獨立的部分，請依需求閱讀:

- **[專案報告](#專案報告)**：為什麼要自己重新算、Word 的計數規則是怎麼量出來的、踩過哪些坑 —— 給想了解「這是什麼、為什麼長這樣」的人看。
- **[架設 SOP](#架設-sop)**：怎麼跑命令列版、怎麼打包 GUI 版、怎麼交叉驗證 —— 給想「動手把它跑起來」的人看。

這兩部分刻意分開，不要交叉閱讀；報告是背景知識，SOP 是操作手冊。

---

## 專案報告

### 這是什麼

劇本殺文本常按文本量報價，需要一次查好幾個檔案的字數。從 Word UI 一個一個開太慢；把 `.docx` 丟給 LLM 直接問字數則永遠是估的，不會對——LLM 沒有辦法真的逐字元跑 Word 內部那套分類邏輯，只能用猜的。這個工具把 Word 自己的字數統計演算法逆向工程出來，寫成一支零依賴的 Python 腳本，一份檔案約 0.02 秒，數字跟 Word 本人完全一致；文件裡如果藏有手動換行或選擇性連字號這兩種肉眼不易察覺的排版標記，則會把每一處標出來，讓使用者回文件修掉。

### 系統架構

這個 repo 是 `zh-cn-to-tw`（劇本殺繁化助手 meta-repo）底下的一個 git submodule，但跟那個產品本身**完全無關**——純粹因為開發時就近方便，才掛在同一個 meta-repo 底下管理：

```
zh-cn-to-tw/                     ← meta-repo，本機開發統一入口，本身不部署
├── zh-cn-to-tw-backend/         ← 產品的後端 API（跟本 repo 無關）
├── zh-cn-to-tw-web/             ← 產品的前端（跟本 repo 無關）
├── zh-cn-to-tw-mac/             ← 產品的 macOS 桌面殼（跟本 repo 無關）
├── zh-cn-to-tw-ocr-service/     ← 產品的本機 OCR 服務（跟本 repo 無關）
├── zh-cn-to-tw-windows/         ← 產品的 Windows 桌面殼（跟本 repo 無關）
└── docx-wordcount/              ← 本 repo，獨立的個人工具
    ├── docx_wordcount.py        ← 核心計數引擎，零依賴
    ├── verify_against_word.ps1  ← 透過 COM 拿 Word 本人的數字交叉驗證
    └── gui/wordcount_gui.py     ← Windows GUI，重用同一套計數邏輯
```

### Word 的字數規則是「量」出來的，不是查文件查出來的

Microsoft 沒有公開 `ComputeStatistics()` 的實作細節，網路上能找到的說明也大多含糊或彼此矛盾。這個工具的規則全部來自對 Word 14.0（Office 2010）的直接量測：用 COM automation 靜默開啟 Word，寫一批「每份文件只隔離一個變數」的受控測試檔（例如 `A–B`、`A—B`、`A•B` 這種單一符號夾在兩個字母中間），逐一讀 `Range.ComputeStatistics()` 回傳的字數／字元數，再從結果反推規則——不是從任何文件推的。

最花時間的一條規則是「哪些符號算字元但不算字」。長得像分隔符號的候選字元有一大票：`–—‒―‑─·●○▪‣⁃…‥§¶` 等等，逐一實測 **50 個候選符號**之後才確認，Word 真正會斷開連續字串、算 1 個字元卻算 0 個字的**只有三個**：en dash（`–` U+2013）、em dash（`—` U+2014）、項目符號（`•` U+2022）。長得幾乎一樣的中點 `·`、水平線 `─`、實心圓 `●` 全部只是普通的連續字串成員，用 Unicode 類別去推會全部推錯。

另一條反直覺的規則是「歧義寬度字元」（Unicode ambiguous width，例如 `·`、`“”`、`①`、`Ⅰ`）:這類字元沒有固定的中英文歸屬，取決於它所在的 `<w:r>` 有沒有帶 `w:rFonts w:hint="eastAsia"`——同一個字元在不同 run 裡可能一次算中文字、一次算英文字。這條是拿一份真實的樂譜檔驗出來的：檔案裡有 60 個 `·`，Word 只把其中 20 個（帶了 hint 的那些）算成中文字元，其餘 40 個算英文，而且**只看該 run 自己的屬性，段落標記跟 docDefaults 的 hint 都不會繼承下來**。

### 為什麼交叉驗證要用真正的劇本文本，不能只挑手邊隨便找到的 docx

第一版的驗證流程犯了一個典型的錯:用電腦上隨手找得到的 10 份 `.docx`（履歷、樂譜、報價單、對照表）當「真實檔案」樣本，五項統計全部跟 Word 完全吻合，看起來驗證做得很扎實。使用者一拿真正在用的劇本文本跑，`2A.docx`／`2B.docx` 兩個檔案立刻算錯。

問題出在那批「真實檔案」樣本本身不具代表性，缺了劇本文本才有的兩種內容:

1. **零寬空格（U+200B）**。`2A.docx` 裡「過了……好一會兒」中間藏了兩個零寬空格，肉眼完全看不到，是從別處複製貼上文字時常見的殘留。Word 對這種字元完全視而不見——不算字元、不算字，甚至**不會斷開連續字串**（`A` + 零寬空格 + `B` 算 1 個字，不是 2 個）。這條規則額外測了 `U+200C`（ZWNJ）、`U+200D`（ZWJ）、`U+FEFF` 三個類似字元，確認 Word 對這四個是同樣的「完全隱形」待遇；但長得很像的軟連字號 `U+00AD`、LRM/RLM（`U+200E`/`U+200F`）、WORD JOINER（`U+2060`）Word 全部照算成普通字元，不能混在一起處理。
2. **只含空白的段落不算一段**。劇本文本裡常見整段只有兩個全形空格（`　　`）當作視覺留白，Word 的段落數統計會直接跳過這種段落。真正的判定條件是「扣掉空白之後還有字元」，不是「含空白的字元數大於 0」——這個差異只有在文本裡真的出現這種空白段落時才會顯現，履歷、樂譜這類樣本裡沒有。

這件事後來寫進了 `zh-cn-to-tw` meta-repo 的 `known-issue-check`（item 27）：驗證樣本要對應工具實際服務的語料，不能只看樣本「是不是真實檔案」，還要看它「像不像真正會遇到的檔案」。

### 為什麼輸出編碼要自己釘死，不能交給執行環境決定

這支工具的 CLI 會印出檔名，而 Windows 主控台預設碼頁是 cp950（Big5）。開發機恰好設有機器層級的 `PYTHONIOENCODING=utf-8`，所以這個問題在開發過程中完全沒有暴露——直到刻意清掉這個環境變數重新測試，才在處理檔名含简体字「众」（U+4F17，Big5 沒有這個字）的檔案時重現 `UnicodeEncodeError`，整個程式當場中止。

這個工具的存在前提是「簡轉繁專案的配套工具，简体檔名是常態」，不能假設每台機器都剛好有那個環境變數。修法是在程式進入點明確 `sys.stdout.reconfigure(encoding="utf-8", errors="replace")`，不繼承任何環境設定。這件事也記進了 `known-issue-check`（item 26）：任何依賴執行環境（locale、codepage、env var）決定行為的程式，在「自己這台機器測得過」和「別人的機器測得過」之間沒有必然關係，兩邊都要各自驗證，不能只驗其中一邊。

### 手動換行與選擇性連字號：刻意不跟著 Word 算，改成標出來

一個做過微排版（換字體、加頁首頁碼、插圖片）的資料夾裡，12 個檔案有 3 個跟 Word 差 1：白未染 A／B 的「含空白」各少 1，慕歸 A 的字數、字元、含空白、中文字元各多 1。字體、頁首頁碼、圖片都不是原因（頁首頁尾本來就不算，圖片沒有文字）。把整個資料夾掃一次，偏差數量剛好等於兩種標記的出現次數；用一份「照 Word 的算法處理這兩種標記」的實驗版重算，12 個檔案全部跟 Word 一致。

- **手動換行**（Shift+Enter，Word 開編輯標記時顯示成 ↵，XML 是不帶 type 的 `<w:br/>` 或 `<w:cr/>`）：Word 算成 **1 個空白字元**，只進「含空白」，而且會斷字。程式原本的規則是「換行完全不算」，那條其實是從**分頁符號**量出來的，後來誤推廣到所有換行。分頁跟分欄符號 Word 確實不算，這部分是對的。
- **選擇性連字號**（Ctrl+-，顯示成 ¬，XML 是 `<w:softHyphen/>`）：Word **完全不算**，也不斷字。程式把它當成 U+00AD 字元，而 U+00AD 是歧義寬度字元，在中文格式的 run 裡被算成 1 個中文字，所以四項統計各多 1。要注意的是，直接打在文字裡的 U+00AD 字元 Word 是會算的，跟這個元素不一樣。

**決定：計算不改，改成標出來。** 在這批劇本裡，這兩種標記都是排版疏失——段落一律該用 Enter，選擇性連字號平常肉眼完全看不到——該修的是文件本身，不是讓字數去遷就它。所以 GUI 多一個區塊，逐一列出每一處的檔案、類型、第幾段、前後文（用 Word 自己的編輯標記符號 ↵、¬、¶ 標示位置），讓使用者回 Word 修掉。代價是：含這兩種標記的檔案，工具的數字會跟 Word 差那幾個字元，這是刻意的。偵測邏輯是獨立的 `find_hidden_marks()`，不碰計算函式，12 個檔案的計數結果在加上這個功能前後逐項相同。

**什麼情況下要改回來對齊 Word：** 如果之後這個工具會用在「Shift+Enter／Ctrl+- 是刻意使用」的文件上（例如別人的排版習慣），就該照 Word 的算法調整 `_paragraph_segments()`（換行算 1 個空白、選擇性連字號不算），標出功能保留成選項。

量測過程還順帶解開一個矛盾：自製測試檔裡的分頁符號被 Word 算成 1 個字元加 1 個字，但真實檔案裡不算。原因是測試檔沒有 `settings.xml`，Word 以 2007 相容模式開啟；真實檔案是 2010 相容模式。**自製測試檔的相容模式要跟真實檔案一致**，否則量到的規則可能只適用於測試檔。

### GUI 版經過使用者兩輪回饋才定案的三個設計

GUI 版最早的版本，為了讓「總計 X 個檔案」那一列的粗體視覺上跟旁邊的大字體總字數一致，把 Consolas 套用到整個表格跟報價文字框。使用者的回饋很直接：Consolas 本身就不好辨識，而且越用越多範圍讓問題更明顯。最終版拿掉所有自訂字體家族（包含最初就存在、從未被指名批評、但顯然也是問題一部分的 `Consolas`／`Microsoft JhengHei UI`），全部改回系統預設字體；唯一保留的「總計」列粗體，改成用 `tkinter.font` 動態取當下的預設字型再疊加 `bold`，而不是另外指定一個字體名稱——確保之後不會再不小心引入第二種字體。

同一輪回饋也拿掉了表格下方「提示：表格內容可以直接選取範圍」那行說明文字。這行字不只是「使用者說這東西只有他自己用、不需要任何提示」這麼單純的取捨——它被塞進 `ttk.Treeview` 所在的 pane 裡跟固定欄寬的表格搶版面，才是真正造成「看起來斷掉」以及「上半部寬度跟其他區塊不一致」的直接原因。拿掉這個標籤同時解決了視覺原因（不必要的說明文字）和技術原因（跟 Treeview 搶寬度）兩個問題。

報價文字格式的方括號（`【…】`）則是一次單純的溝通落差：使用者給的範例文字前後包了方括號，用意是「框住這段示意文字」，不是要求輸出本身要帶方括號，但第一版原封不動把方括號也寫進了 `build_quote_text()` 的輸出格式裡。

### 為什麼驗證 GUI 邏輯後來完全不再操作使用者的滑鼠鍵盤

第一版 GUI 打包成 `.exe` 後，曾經嘗試用 Win32 API（`SetCursorPos`、`mouse_event`、`SendKeys`）直接操作編譯後的執行檔做端到端測試——這類操作是在使用者當下正在使用的真實桌面上進行，不是隔離的測試環境。過程中至少踩到兩個坑：中文路徑透過 `SendKeys` 送出時被輸入法攔截組字，欄位裡的文字被打亂；更嚴重的是，「截圖 → 用工具顯示圖片 → 下一步再點擊」這個流程裡，顯示截圖那一步會把 Claude 桌面 App 帶回前景，等下一次點擊真正送出時，焦點已經跑到別的地方——當時螢幕上剛好開著 LINE 登入視窗跟其他應用程式，點擊因此打到了不相關的視窗。

這件事之後，GUI 邏輯的驗證方式整個換了做法：把「產生報價文字」（`build_quote_text`）、「解析單價欄位」（`parse_rate`）這些真正需要驗證正確性的邏輯，寫成不依賴任何 Tk 元件的模組層級純函式，直接用一般 Python 呼叫測試，逐字元核對輸出——完全不需要建立視窗，自然也不會有搶焦點或誤觸的風險。至於「編譯後的 `.exe` 有沒有正確包進最新程式碼」這件事，後來也放棄了對編譯後二進位檔做位元組字串搜尋（在壓縮過的封存資料裡，任何短字串都有機率巧合出現，曾經因此得到一次誤導性的「殘留」訊號），改成用 `marshal.loads()` 把打包進 PyInstaller 封存檔的模組直接解回 code object，遞迴列出所有巢狀函式的字串常數表，逐一核對——這是唯一真正對應「這支 exe 裡實際跑的是什麼程式碼」的證據，而且全程只是讀取封存檔內容，不會跳出任何視窗。

需要視覺或互動驗證的部分（字體好不好看、版面寬不寬），改成直接把打包好的 `.exe` 交給使用者本人操作確認——這正是這類工具最終應該由誰驗證的問題。

### 檔案結構

```
docx_wordcount.py          核心計數引擎（zipfile + xml.etree，零依賴），也是 CLI 入口
verify_against_word.ps1    透過 COM 靜默開啟 Word，逐檔跟 Python 的結果比對
gui/
  wordcount_gui.py          tkinter GUI，import docx_wordcount 重用同一套計數邏輯
  dist/                     PyInstaller 打包輸出（不進版控，見 .gitignore）
```

### 已知限制

- **東亞字元的 Unicode 區段大多是外推的。** 實際逐字元實測過的只有 CJK 統一表意、常見 CJK 標點、平假名、片假名、諺文音節、全形區與半寬區的代表字元；其餘同類區段（注音、康熙部首、CJK 相容字等）沒有逐一驗證過，是「同區段應該同類」的推論。
- **編號格式只驗證過兩種。** `numFmt` 支援的十幾種格式裡，真實檔案只出現過 `decimal` 跟 `bullet`，中文數字、羅馬數字、圈碼數字等分支從未跟 Word 對照過；`lvlRestart` 也沒有實作，多層清單的重新編號時機可能跟 Word 不同。
- **追蹤修訂的 `--include-deleted` 完全沒有測試覆蓋**，也沒有跟 Word「所有標記」檢視模式下的數字比對過。
- **`--include-notes` / `--include-boxes` 只驗證過一份合成測試檔**，沒有真實檔案覆蓋。
- 隱藏文字只認直接格式的 `<w:vanish/>`，不解析樣式繼承來的隱藏。
- 只支援 `.docx`／`.dotx` 這類 OOXML；舊的 `.doc` 二進位格式不支援。
- **含手動換行或選擇性連字號的檔案，數字會跟 Word 不同（刻意的）**：每個手動換行讓「含空白」少 1；每個選擇性連字號讓字元多 1，落在中文格式文字裡時字數、中文字元也各多 1。GUI 會把每一處列出來，見上方「手動換行與選擇性連字號」。
- GUI 版目前只有一次使用者親自點擊完整流程（選資料夾 → 算出結果 → 複製）並確認畫面正確的紀錄；後續每一輪修改都只驗證到純函式邏輯層級與封存檔內容層級，沒有再重複一次完整的互動流程。

---

# 架設 SOP / Setup Guide

## Part A. 命令列版

零依賴，只要標準函式庫，不需要額外 `pip install` 任何套件。

### 1. 執行

```bash
# 單檔／多檔／整個資料夾
python docx_wordcount.py "劇本.docx"
python docx_wordcount.py -r "D:\劇本"

# 貼進試算表（--tsv 印完整路徑，不會因同名檔案分不出來）
python docx_wordcount.py --tsv -r "D:\劇本"

# 對齊 Word 對話框「包含文字方塊、註腳及章節附註」那個勾選項
python docx_wordcount.py --include-notes --include-boxes a.docx
```

> PowerShell 與 cmd 不會替外部程式展開通配字元，`*.docx` 會原樣傳進來變成一個找不到的路徑。要指定一批檔案就直接給資料夾，或自己展開：`python docx_wordcount.py (Get-ChildItem *.docx).FullName`

## Part B. GUI 版

### 1. 直接用原始碼跑（開發/除錯用）

```bash
cd gui
python wordcount_gui.py
```

需要系統內建的 `tkinter`（Windows 版 Python 官方安裝檔預設就有，不需另外安裝）。

### 2. 打包成單一 exe

```bash
pip install pyinstaller
cd gui
python -m PyInstaller --onefile --windowed --clean --name ScriptWordCount --paths .. wordcount_gui.py
```

輸出在 `gui/dist/ScriptWordCount.exe`，單一檔案、不需安裝、不需網路，可以直接複製到別台機器上用。build 產物（`build/`、`dist/`、`*.spec`）不進版控。

### 3. 畫面

選資料夾（算第一層的所有 `.docx`，不進子資料夾）或單一檔案之後，下方由上到下三個區塊，中間的分隔線可以拖曳調整高度：

1. **字數表格**：每個檔案的字數、字元（不計／含空白）、中文字元、段落，最後一列是總計。可選取多列按 Ctrl+C 複製。
2. **手動換行／選擇性連字號**：每一處的檔案、類型、第幾段、前後文，標題列顯示共幾處。前後文裡 ↵ 是手動換行、¬ 是選擇性連字號、¶ 表示該處在段落結尾；可以複製前後文到 Word 用 Ctrl+F 定位。在 Word 裡要一次找出所有選擇性連字號，可用「尋找及取代 → 更多 → 指定方式 → 選擇性連字號」（`^-`）。
3. **報價文字**：依每個檔案的「字元（含空白）」自動產生，單價欄位改了會即時重算。

## Part C. 交叉驗證

改動 `docx_wordcount.py` 的計數邏輯之後，一定要跑這個，不要只靠讀程式碼判斷對不對：

```powershell
powershell -ExecutionPolicy Bypass -File verify_against_word.ps1 -Path "D:\劇本"
```

透過 COM 靜默開啟每個檔案、讀 Word 自己的 `ComputeStatistics()`，逐項比對五個統計數字。需要本機裝有 Microsoft Word，而且**驗證用的檔案要是真正會遇到的語料**（劇本文本，不是隨手找到的履歷或樂譜）——見上方專案報告「為什麼交叉驗證要用真正的劇本文本」。這支腳本比 Python 版慢得多（每檔數秒），這正是 Python 版存在的理由。

## Part D. 需求總覽

| 項目 | 版本/需求 | 用途 |
|---|---|---|
| Python | 3.x（開發時用 3.12） | 執行 `docx_wordcount.py`、GUI 版 |
| Microsoft Word | 任一版本（開發時用 Word 14.0／Office 2010） | 僅 `verify_against_word.ps1` 交叉驗證需要，一般使用不需要 |
| PyInstaller | 任一版本（開發時用 6.22） | 僅打包 GUI 版成 `.exe` 需要 |

---

# English

## docx-wordcount — .docx Word Count Tool

Counts the words, characters, and paragraphs in a `.docx` exactly the way Microsoft Word's "Review → Word Count" dialog reports them (the one deliberate exception being manual line breaks and optional hyphens, which the tool lists instead — see [Known Limitations](#known-limitations)), without opening Word. Ships as a zero-dependency CLI and as a single-file, install-free, offline Windows GUI (`.exe`). This document is split into two independent parts:

- **[Project Report](#project-report)**: why this had to be built from scratch, how Word's counting rules were reverse-engineered, and what went wrong along the way — for anyone who wants to understand what this is and why it's built this way.
- **[Setup Guide](#setup-guide)**: how to run the CLI, build the GUI, and cross-validate — for anyone who wants to get it running.

These two parts are kept deliberately separate; the report is background, the guide is a manual.

---

## Project Report

### What This Is

Script-murder-mystery scripts are commonly priced by text volume, which routinely means counting several files at once. Opening each one in Word is too slow; handing a `.docx` to an LLM and asking for its word count is always a guess, and never correct — an LLM has no way to actually run Word's internal character-by-character classification, only to estimate. This tool reverse-engineers Word's own word-count algorithm into a zero-dependency Python script — about 0.02 seconds per file, numbers that match Word exactly; if a document hides either of two hard-to-spot formatting marks — manual line breaks or optional hyphens — every occurrence is pointed out so the document itself can be fixed.

### System Architecture

This repo is a git submodule of `zh-cn-to-tw` (the Script Murder Mystery Traditionalization Assistant meta-repo), but is otherwise **entirely unrelated** to that product — it's submoduled here purely for local-dev convenience:

```
zh-cn-to-tw/                     ← meta-repo, unified local-dev entry, never deployed itself
├── zh-cn-to-tw-backend/         ← the product's backend API (unrelated to this repo)
├── zh-cn-to-tw-web/             ← the product's frontend (unrelated to this repo)
├── zh-cn-to-tw-mac/             ← the product's macOS desktop shell (unrelated to this repo)
├── zh-cn-to-tw-ocr-service/     ← the product's local OCR service (unrelated to this repo)
├── zh-cn-to-tw-windows/         ← the product's Windows desktop shell (unrelated to this repo)
└── docx-wordcount/              ← this repo, a standalone personal tool
    ├── docx_wordcount.py        ← the counting engine, zero dependencies
    ├── verify_against_word.ps1  ← cross-checks against Word's own numbers via COM
    └── gui/wordcount_gui.py     ← Windows GUI, reuses the same counting logic
```

### Word's Counting Rules Were Measured, Not Read From Any Document

Microsoft never published `ComputeStatistics()`'s implementation details, and what's out there online is mostly vague or contradicts itself. Every rule this tool follows comes from directly measuring Word 14.0 (Office 2010): driving Word silently through COM automation, writing a battery of controlled test documents that each isolate exactly one variable (e.g. a single symbol like `A–B`, `A—B`, or `A•B` sandwiched between two letters), reading `Range.ComputeStatistics()`'s word/character counts for each one, and reverse-engineering the rule from the result — never guessed from documentation.

The rule that took the longest to pin down was "which characters count as a character but not a word." There's a large family of symbols that look like plausible separators: `–—‒―‑─·●○▪‣⁃…‥§¶` and more. Only after testing **50 candidate symbols one by one** did it become clear that Word breaks a word run and charges 1 character / 0 words for **exactly three**: en dash (`–` U+2013), em dash (`—` U+2014), and the bullet (`•` U+2022). Near-identical-looking characters — the middle dot `·`, the horizontal bar `─`, the solid circle `●` — are all ordinary word-run members; guessing from Unicode category alone gets every one of them wrong.

Another counter-intuitive rule concerns "ambiguous width" characters (Unicode's own term, e.g. `·`, `“”`, `①`, `Ⅰ`): these have no fixed Chinese/Latin classification — it depends on whether the `<w:r>` they sit in carries `w:rFonts w:hint="eastAsia"`. The same character can count as Chinese in one run and Latin in another within the same document. This was measured on a real sheet-music file: it contained 60 instances of `·`, and Word counted only 20 of them (the ones in a hinted run) as CJK characters, the other 40 as Latin — and **only the run's own attribute matters; neither the paragraph mark's `rPr` nor `docDefaults`' hint is inherited**.

### Why Cross-Validation Has to Use Real Scripts, Not Whatever docx Happens to Be Lying Around

The first validation pass made a classic mistake: it used 10 `.docx` files that happened to be sitting around on the development machine (résumés, a music score, a price quote, a comparison table) as its "real-world" sample. All five statistics matched Word exactly on all ten, which looked like solid validation. The moment the user ran it against the actual scripts in use, `2A.docx` and `2B.docx` were both wrong immediately.

The problem was that the "real-world" sample wasn't representative — it was missing two things specific to script text:

1. **Zero-width spaces (U+200B).** `2A.docx` had two of them hidden in the middle of a sentence, invisible to the naked eye — a typical leftover from pasting text in from somewhere else. Word treats this character as if it doesn't exist at all: no character, no word, and it **doesn't even break a word run** (`A` + a zero-width space + `B` counts as one word, not two). Three related characters — `U+200C` (ZWNJ), `U+200D` (ZWJ), `U+FEFF` — were tested alongside it and confirmed to get the exact same "fully invisible" treatment; but visually similar characters — the soft hyphen (`U+00AD`), LRM/RLM (`U+200E`/`U+200F`), WORD JOINER (`U+2060`) — are all counted by Word as ordinary characters and cannot be lumped in with the zero-width group.
2. **A paragraph containing only whitespace doesn't count as a paragraph.** Script text often uses a paragraph holding just two full-width spaces (`　　`) as visual spacing, and Word's paragraph-count statistic skips such a paragraph entirely. The real test is "any non-whitespace character present," not "any character including whitespace present" — a distinction that only surfaces when text actually contains this kind of whitespace-only paragraph, which none of the convenience-sample files did.

This later became `known-issue-check` item 27 in the `zh-cn-to-tw` meta-repo: a validation sample needs to match the corpus the tool actually serves, not merely qualify as "real files" in the abstract — it also has to resemble what will actually be thrown at it.

### Why Output Encoding Is Pinned Explicitly Instead of Trusting the Environment

This tool's CLI prints filenames, and Windows consoles default to codepage cp950 (Big5). The development machine happened to carry a machine-level `PYTHONIOENCODING=utf-8`, so this never surfaced during development — until that variable was deliberately cleared for a test, which reproduced a `UnicodeEncodeError` (and an immediate abort) on a file whose name contained a Simplified-only character, "众" (U+4F17, absent from Big5).

Given this tool exists specifically to support a Simplified-to-Traditional project, Simplified filenames are the normal case, not an edge case — nothing can assume every machine happens to carry that same environment variable. The fix is to explicitly `sys.stdout.reconfigure(encoding="utf-8", errors="replace")` at the entry point, inheriting nothing from the ambient environment. This also became `known-issue-check` item 26: any program whose behavior depends on the ambient locale/codepage/env var carries no guarantee that "it works on my machine" says anything about whether it works on someone else's — both need to be verified independently, not just one.

### Manual Line Breaks and Optional Hyphens: Deliberately Not Counted Word's Way, Reported Instead

In a folder of scripts that had been lightly typeset (font changes, a header and page numbers, inserted images), 3 of 12 files were off from Word by one: 白未染 A/B were each one short on "characters with spaces", and 慕歸 A was one over on words, characters, characters with spaces, and CJK characters. Fonts, headers, page numbers and images were all ruled out (headers and footers are never counted; images carry no text). Scanning the whole folder showed the deviations equal, exactly, the number of two kinds of marks; recounting with an experimental copy that handles those two marks Word's way matched Word on all 12 files.

- **Manual line break** (Shift+Enter, shown as ↵ with Word's editing marks on; in XML a `<w:br/>` with no type, or `<w:cr/>`): Word charges it **one whitespace character**, counted only in "characters with spaces", and it breaks a word run. The tool's original rule, "a line break isn't a character at all", had actually been measured on **page breaks** and then wrongly generalized to every kind of break. Page and column breaks really are not counted by Word; that part stands.
- **Optional hyphen** (Ctrl+-, shown as ¬; in XML `<w:softHyphen/>`): Word **ignores it entirely** and it doesn't break a word run. The tool treated it as the character U+00AD, which is ambiguous width, so inside an East-Asian-formatted run it became one CJK character — hence all four statistics one over. Note that a literal U+00AD typed into the text itself *is* counted by Word; the element and the character behave differently.

**Decision: leave the counting alone and point these out instead.** In these scripts both marks are typesetting mistakes — paragraphs should always use Enter, and an optional hyphen is normally invisible to the eye — so the thing to fix is the document, not the count. The GUI gains a section listing every occurrence: file, kind, paragraph number, and surrounding text (marked with Word's own editing-mark glyphs ↵, ¬, ¶), so they can be fixed back in Word. The cost: for a file containing these marks, the tool's numbers differ from Word's by exactly those characters, on purpose. Detection is a separate `find_hidden_marks()` that doesn't touch the counting functions; every count across the 12 files is identical before and after the feature was added.

**When to switch back to matching Word:** if the tool starts being used on documents where Shift+Enter / Ctrl+- are intentional (someone else's formatting habits), adjust `_paragraph_segments()` to count them Word's way (a line break as one whitespace character, an optional hyphen as nothing), and keep the report as an option.

The measurements also cleared up a contradiction along the way: in a hand-built test document, Word counted a page break as one character and one word, yet in real files it doesn't. The test document had no `settings.xml`, so Word opened it in Word 2007 compatibility mode; the real files are in Word 2010 mode. **Hand-built test documents need the same compatibility mode as the real files**, or a measured rule may only hold for the test document.

### Three GUI Design Choices That Only Settled After Two Rounds of User Feedback

The GUI's earliest version applied Consolas across the whole table and quote box, originally just to make the bold "total row" visually consistent with the large Consolas total-word-count label next to it. The user's feedback was direct: Consolas itself is hard to read, and spreading it further only made the problem more visible. The final version drops every custom font family entirely — including `Consolas`/`Microsoft JhengHei UI`, which had been there from the very first version and were never individually called out, but were evidently part of the same problem. The one thing that stays bold — the "total" row — now derives its font from whatever the actual default font is via `tkinter.font`, layering on `bold`, rather than naming a typeface of its own; nothing can silently reintroduce a second font family this way.

The same round of feedback also removed the "hint: table contents can be selected and copied" label beneath the table. This wasn't purely a matter of "the user said this is a single-person tool that needs no hints" — that label had been packed into the same pane as the fixed-width `ttk.Treeview`, competing with it for space, which was the actual, direct cause of both the label looking "broken" and the upper half's width no longer matching the rest of the window. Removing it fixed the visual symptom and the underlying layout conflict at the same time.

The brackets (`【…】`) around the quote text were a simpler communication gap: the user's example text was wrapped in brackets to mark "this is illustrative text," not a request that the output itself carry brackets — but the first version copied the brackets verbatim into `build_quote_text()`'s output format.

### Why Verifying the GUI Stopped Touching the User's Mouse and Keyboard Entirely

After the first GUI build was packaged into a `.exe`, an attempt was made to drive the compiled executable end-to-end using Win32 APIs (`SetCursorPos`, `mouse_event`, `SendKeys`) — directly on the user's live desktop, not an isolated test environment. This hit at least two problems: a Chinese file path sent via `SendKeys` was intercepted and garbled by the IME mid-composition; more seriously, in the "screenshot → display it via a tool → click next" loop, the step that displayed the screenshot brought the Claude desktop app back to the foreground, so by the time the next click actually fired, focus had moved elsewhere — with a LINE login window and other apps open on screen at the time, the click landed on an unrelated window.

After that, verifying the GUI's logic changed approach entirely: the pieces that actually needed correctness checking — generating the quote text (`build_quote_text`) and parsing the rate field (`parse_rate`) — were pulled out as module-level pure functions with no Tk dependency at all, callable and testable from plain Python, checked character-for-character against expected output. No window ever needs to exist, so there's no focus-stealing or misclick risk to begin with. As for "did the compiled `.exe` actually bundle the latest code" — raw byte-string searching on the compiled binary was abandoned too (in a large blob of compressed archive data, any short string has a real chance of appearing by coincidence, which produced one genuinely misleading "still present" signal during testing); the replacement was `marshal.loads()` to decode the module bundled inside the PyInstaller archive straight back into its code object, then recursively walking every nested function's string-constant table — the only evidence that actually corresponds to "what code is really running inside this exe," and reading archive contents this way never opens a single window.

Anything that genuinely needs visual or interactive confirmation (does the font look right, is the layout the right width) is instead handed directly to the user to click through themselves — which is really the underlying question of who should be verifying this kind of thing in the first place.

### File Layout

```
docx_wordcount.py          the counting engine (zipfile + xml.etree, zero deps), also the CLI entry point
verify_against_word.ps1    drives Word silently via COM, diffs against the Python results file by file
gui/
  wordcount_gui.py          tkinter GUI, imports docx_wordcount to reuse the same counting logic
  dist/                     PyInstaller build output (gitignored)
```

### Known Limitations

- **Most East-Asian Unicode ranges are extrapolated.** Character-by-character measurement only covers CJK Unified Ideographs, common CJK punctuation, hiragana, katakana, hangul syllables, and representative characters from the fullwidth and halfwidth blocks; other same-family ranges (bopomofo, Kangxi radicals, CJK compatibility ideographs, etc.) were never individually verified — it's an inference that "the same block should behave the same," not a measurement.
- **Only two numbering formats have been verified.** Of the dozen-plus `numFmt` values, only `decimal` and `bullet` have ever shown up in real files; Chinese numerals, Roman numerals, circled digits, and the rest have never been checked against Word. `lvlRestart` also isn't implemented, so multi-level list renumbering may not match Word's timing.
- **`--include-deleted` (tracked changes) has zero test coverage**, and has never been compared against Word's "All Markup" view numbers.
- **`--include-notes` / `--include-boxes` have only been verified against one synthetic test document**, with no real-file coverage.
- Hidden text only recognizes direct formatting (`<w:vanish/>`), not hiding inherited through a style.
- Only `.docx`/`.dotx` (OOXML) is supported; the legacy binary `.doc` format is not.
- **Files containing manual line breaks or optional hyphens give numbers that differ from Word's (on purpose)**: each manual line break makes "characters with spaces" one short; each optional hyphen adds one character, and inside East-Asian-formatted text also one word and one CJK character. The GUI lists every occurrence — see "Manual Line Breaks and Optional Hyphens" above.
- The GUI has only one recorded end-to-end run where the user personally clicked through the whole flow (pick folder → see results → copy) and confirmed the screen was correct; every round of changes since has only been verified at the pure-function and packaged-archive-contents level, without repeating a full interactive pass.

---

# Setup Guide

## Part A. Command-Line Version

Zero dependencies — standard library only, no `pip install` needed.

### 1. Run it

```bash
# a single file / several files / a whole folder
python docx_wordcount.py "script.docx"
python docx_wordcount.py -r "D:\Scripts"

# paste-ready for a spreadsheet (--tsv prints full paths, so same-named files never collide)
python docx_wordcount.py --tsv -r "D:\Scripts"

# match Word's dialog checkbox "include text boxes, footnotes and endnotes"
python docx_wordcount.py --include-notes --include-boxes a.docx
```

> PowerShell and cmd don't expand wildcards for external programs — `*.docx` arrives as a literal, nonexistent path. To target a batch of files, either point at the folder directly, or expand it yourself: `python docx_wordcount.py (Get-ChildItem *.docx).FullName`

## Part B. GUI Version

### 1. Run from source (for development/debugging)

```bash
cd gui
python wordcount_gui.py
```

Needs the system's built-in `tkinter` (bundled by default with the official Windows Python installer, nothing extra to install).

### 2. Build a single exe

```bash
pip install pyinstaller
cd gui
python -m PyInstaller --onefile --windowed --clean --name ScriptWordCount --paths .. wordcount_gui.py
```

Output lands at `gui/dist/ScriptWordCount.exe` — a single file, no install, no network, safe to copy straight to another machine. Build artifacts (`build/`, `dist/`, `*.spec`) are gitignored.

### 3. The Screen

After picking a folder (counts every `.docx` at its first level, not subfolders) or a single file, three sections stack top to bottom; the dividers between them can be dragged to resize:

1. **Count table**: each file's words, characters (without/with spaces), CJK characters and paragraphs, with a total row last. Select rows and press Ctrl+C to copy.
2. **Manual line breaks / optional hyphens**: every occurrence's file, kind, paragraph number and surrounding text, with the total count in the section title. In the surrounding text, ↵ is a manual line break, ¬ an optional hyphen, and ¶ means that spot ends the paragraph; copy the text into Word's Ctrl+F to jump to it. To find every optional hyphen in Word at once, use Find and Replace → More → Special → Optional Hyphen (`^-`).
3. **Quote text**: generated from each file's "characters with spaces"; editing the rate field recalculates it immediately.

## Part C. Cross-Validation

After changing anything in `docx_wordcount.py`'s counting logic, always run this — don't trust a code review alone:

```powershell
powershell -ExecutionPolicy Bypass -File verify_against_word.ps1 -Path "D:\Scripts"
```

Drives Word silently through COM, reads its own `ComputeStatistics()` per file, and diffs all five statistics. Requires Microsoft Word installed locally, and **the files used for validation need to be the real corpus this serves** (actual scripts, not whatever résumé or sheet-music file happens to be handy) — see "Why Cross-Validation Has to Use Real Scripts" in the Project Report above. This script is much slower than the Python version (seconds per file), which is exactly why the Python version exists.

## Part D. Requirements

| Item | Version/Requirement | Used For |
|---|---|---|
| Python | 3.x (developed against 3.12) | Running `docx_wordcount.py` and the GUI |
| Microsoft Word | Any version (developed against Word 14.0 / Office 2010) | Only `verify_against_word.ps1`'s cross-validation; not needed for normal use |
| PyInstaller | Any version (developed against 6.22) | Only for building the GUI into a `.exe` |
