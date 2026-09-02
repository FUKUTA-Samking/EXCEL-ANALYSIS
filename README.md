Excel 數據分析工具（桌面版）
一個以 Python + tkinter 打造的桌面小工具，讓你匯入 Excel 資料後，
自由選擇欄位/資料列範圍、產生圖表（可疊加規格上下限 USL/LSL）、
執行統計分析（Cpk、相關性、敘述統計），最後把所有結果匯出成一份
完整的 Excel 報表。
安裝
需要 Python 3.9 以上版本。
```bash
pip install -r requirements.txt
```
> tkinter 通常隨 Python 內建；若 Linux 上出現 `No module named tkinter`，
> 請另外安裝，例如 Ubuntu/Debian：`sudo apt-get install python3-tk`
執行
```bash
python main.py
```
使用流程
① 資料選擇與規格設定
點「📂 匯入 Excel」選擇檔案，若有多個分頁可從上方下拉選單切換
在左側清單勾選要分析的欄位（可多選），並可設定「起始列／結束列」限定資料範圍
按「套用範圍 →」即可在右側預覽篩選結果
按「規格上下限設定」可為每個數值欄位輸入 LSL（下限）／USL（上限），
這組數字會同時用在圖表的參考線與 Cpk 分析
② 圖表產生
選擇圖表類型：長條圖、折線圖、圓餅圖、直方圖、散佈圖、盒鬚圖
選擇要繪製的欄位（散佈圖需另外指定 X／Y 軸）
按「產生圖表」即時預覽；若滿意可按「＋ 加入此圖表到報表」收進報表清單
可重複調整類型/欄位多產生幾張圖，通通加入報表
③ 統計分析
選擇分析方法：
Cpk 製程能力分析：依你在①設定的 LSL/USL 計算 Cp、Cpu、Cpl、Cpk，並標示製程能力評估
相關性分析：計算所選欄位間的 Pearson 相關係數矩陣
敘述統計：平均值、標準差、四分位數、偏態、峰度等
按「執行分析」在右側看結果，滿意後按「＋ 加入此結果到報表」
匯出
按右上角「💾 匯出完整 Excel 報表」，指定存檔位置即可產出包含：
報表資訊頁
篩選後的原始資料
規格上下限設定表
所有已加入的圖表（圖片內嵌）
所有已加入的分析結果（各自一個分頁）
檔案結構
```
excel_analyzer/
├── main.py            # GUI 主程式（tkinter）
├── analysis_core.py   # 核心運算邏輯（讀檔/篩選/繪圖/統計/匯出），可獨立測試、重複使用
└── requirements.txt
```
打包成 Windows 執行檔 (.exe)，不需另外安裝 Python
PyInstaller 只能在「目標系統」上打包（無法在 Linux 直接產出 Windows 的 .exe），
所以需要在一台 Windows 電腦上執行一次以下步驟：
確認該電腦有安裝 Python 3.9+（安裝時勾選「Add python.exe to PATH」）。
下載: https://www.python.org/downloads/
把 `main.py`、`analysis_core.py`、`requirements.txt`、`excel_analyzer.spec`、
`build_windows.bat` 這 5 個檔案放在同一個資料夾。
直接雙擊執行 `build_windows.bat`（會自動建立虛擬環境、安裝套件、打包）。
完成後，執行檔會出現在 `dist\ExcelAnalyzer.exe`。
之後就能把這個 `.exe` 複製到任何 Windows 電腦，直接雙擊使用，
不需要在對方電腦上安裝 Python 或任何套件。
> 只需要打包這一次；之後散發時只要給對方 `ExcelAnalyzer.exe` 這一個檔案即可。
> 若之後修改了 `main.py` 或 `analysis_core.py`，重新執行一次 `build_windows.bat` 即可更新 exe。
手動打包（不想用 .bat 腳本的話）
```cmd
pip install -r requirements.txt pyinstaller
pyinstaller excel_analyzer.spec
```
或不用 spec 檔，最簡單的一行指令：
```cmd
pyinstaller --onefile --windowed --name ExcelAnalyzer main.py
```
（此指令已在 Linux 環境下驗證過打包流程本身可正常運作，Windows 上的行為應一致；
若你在 Windows 上打包時遇到「缺少模組」的錯誤，把缺的模組名稱加進
`excel_analyzer.spec` 的 `hiddenimports` 清單即可。）
提示
Cpk 評估標準：Cpk ≥ 1.33 能力充分／1.0–1.33 能力尚可／< 1.0 能力不足（業界常見經驗值，可依公司標準調整判讀）。
若圖表中文字顯示為方框，代表系統缺少中文字型，`analysis_core.py` 會自動偵測常見中文字型
（微軟正黑體/雅黑、蘋方、Noto Sans CJK 等），建議在 Windows/macOS 上執行即可正常顯示。
