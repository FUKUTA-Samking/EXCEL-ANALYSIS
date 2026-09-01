# -*- coding: utf-8 -*-
"""
main.py - Excel 數據分析工具（桌面版）
======================================
功能：
1. 匯入待分析 Excel 檔案（支援多分頁）
2. 依欄位(columns)與資料列範圍(rows)選擇分析範圍
3. 產生圖表（長條圖/折線圖/圓餅圖/直方圖/散佈圖/盒鬚圖），可疊加 USL/LSL 上下限
4. 統計分析（Cpk 製程能力分析 / 相關性分析 / 敘述統計）
5. 匯出完整 Excel 報表（原始資料 + 規格 + 圖表 + 分析結果）

執行方式：
    python main.py
需求套件見 requirements.txt (pip install -r requirements.txt)
"""
import os
import traceback
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import pandas as pd

import matplotlib
matplotlib.use("TkAgg")
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

import analysis_core as core


class ExcelAnalyzerApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Excel 數據分析工具")
        self.geometry("1300x820")
        self.minsize(1100, 700)

        # -- 狀態資料 --
        self.filepath = None
        self.raw_df = None          # 選定分頁的完整原始資料
        self.filtered_df = None     # 依欄/列範圍篩選後的資料
        self.spec_limits = {}       # {欄位: {"LSL":..,"USL":..}}
        self.spec_entry_widgets = {}  # {欄位: (lsl_entry, usl_entry)}
        self.report_items_charts = []     # 已加入報表的圖表 [{"title","png_bytes"}]
        self.report_items_analyses = []   # 已加入報表的分析 [{"title","df"}]
        self.current_fig = None

        self._build_toolbar()
        self._build_notebook()
        self._build_statusbar()

    # ------------------------------------------------------------------
    # UI 建構
    # ------------------------------------------------------------------
    def _build_toolbar(self):
        bar = ttk.Frame(self, padding=8)
        bar.pack(side="top", fill="x")

        ttk.Button(bar, text="📂 匯入 Excel", command=self.on_import).pack(side="left")
        self.lbl_file = ttk.Label(bar, text="尚未匯入檔案", foreground="#555")
        self.lbl_file.pack(side="left", padx=10)

        ttk.Label(bar, text="分頁：").pack(side="left", padx=(20, 2))
        self.sheet_var = tk.StringVar()
        self.sheet_combo = ttk.Combobox(bar, textvariable=self.sheet_var, state="readonly", width=25)
        self.sheet_combo.pack(side="left")
        self.sheet_combo.bind("<<ComboboxSelected>>", self.on_sheet_selected)

        ttk.Button(bar, text="💾 匯出完整 Excel 報表", command=self.on_export).pack(side="right")

    def _build_statusbar(self):
        self.status_var = tk.StringVar(value="請先匯入 Excel 檔案")
        bar = ttk.Label(self, textvariable=self.status_var, anchor="w", relief="sunken", padding=4)
        bar.pack(side="bottom", fill="x")

    def _build_notebook(self):
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(side="top", fill="both", expand=True, padx=8, pady=4)

        self.tab_data = ttk.Frame(self.notebook)
        self.tab_chart = ttk.Frame(self.notebook)
        self.tab_analysis = ttk.Frame(self.notebook)

        self.notebook.add(self.tab_data, text="① 資料選擇與規格設定")
        self.notebook.add(self.tab_chart, text="② 圖表產生")
        self.notebook.add(self.tab_analysis, text="③ 統計分析")

        self._build_tab_data()
        self._build_tab_chart()
        self._build_tab_analysis()

    # -- Tab 1：資料選擇與規格 -------------------------------------------------
    def _build_tab_data(self):
        left = ttk.Frame(self.tab_data, padding=8)
        left.pack(side="left", fill="y")

        ttk.Label(left, text="欄位選擇 (可多選)", font=("", 10, "bold")).pack(anchor="w")
        self.col_listbox = tk.Listbox(left, selectmode="extended", width=26, height=18, exportselection=False)
        self.col_listbox.pack(pady=4)

        row_frame = ttk.Frame(left)
        row_frame.pack(fill="x", pady=8)
        ttk.Label(row_frame, text="資料列範圍：").grid(row=0, column=0, columnspan=2, sticky="w")
        ttk.Label(row_frame, text="起始列").grid(row=1, column=0, sticky="w")
        self.row_start_var = tk.StringVar(value="1")
        ttk.Entry(row_frame, textvariable=self.row_start_var, width=8).grid(row=1, column=1)
        ttk.Label(row_frame, text="結束列").grid(row=2, column=0, sticky="w")
        self.row_end_var = tk.StringVar(value="")
        ttk.Entry(row_frame, textvariable=self.row_end_var, width=8).grid(row=2, column=1)
        ttk.Label(row_frame, text="(結束列留空 = 至最後一列)", foreground="#777").grid(
            row=3, column=0, columnspan=2, sticky="w")

        ttk.Button(left, text="套用範圍 →", command=self.on_apply_range).pack(fill="x", pady=6)
        ttk.Button(left, text="規格上下限設定 (LSL / USL)", command=self.on_open_spec_dialog).pack(fill="x", pady=(16, 0))

        # 右側：資料預覽
        right = ttk.Frame(self.tab_data, padding=8)
        right.pack(side="left", fill="both", expand=True)
        ttk.Label(right, text="資料預覽 (前20筆)", font=("", 10, "bold")).pack(anchor="w")
        self.preview_tree = ttk.Treeview(right, show="headings")
        self.preview_tree.pack(fill="both", expand=True, pady=4)
        yscroll = ttk.Scrollbar(right, orient="vertical", command=self.preview_tree.yview)
        self.preview_tree.configure(yscroll=yscroll.set)

    # -- Tab 2：圖表產生 --------------------------------------------------------
    def _build_tab_chart(self):
        left = ttk.Frame(self.tab_chart, padding=8)
        left.pack(side="left", fill="y")

        ttk.Label(left, text="圖表類型", font=("", 10, "bold")).pack(anchor="w")
        self.chart_type_var = tk.StringVar(value=core.CHART_TYPES[0])
        ttk.Combobox(left, textvariable=self.chart_type_var, values=core.CHART_TYPES,
                     state="readonly", width=20).pack(pady=4)

        ttk.Label(left, text="欲繪製欄位 (可多選)", font=("", 10, "bold")).pack(anchor="w", pady=(12, 0))
        self.chart_col_listbox = tk.Listbox(left, selectmode="extended", width=24, height=10, exportselection=False)
        self.chart_col_listbox.pack(pady=4)

        scatter_frame = ttk.LabelFrame(left, text="散佈圖專用 (X / Y 軸)")
        scatter_frame.pack(fill="x", pady=8)
        ttk.Label(scatter_frame, text="X 軸欄位：").grid(row=0, column=0, sticky="w")
        self.scatter_x_var = tk.StringVar()
        self.scatter_x_combo = ttk.Combobox(scatter_frame, textvariable=self.scatter_x_var, width=16, state="readonly")
        self.scatter_x_combo.grid(row=0, column=1)
        ttk.Label(scatter_frame, text="Y 軸欄位：").grid(row=1, column=0, sticky="w")
        self.scatter_y_var = tk.StringVar()
        self.scatter_y_combo = ttk.Combobox(scatter_frame, textvariable=self.scatter_y_var, width=16, state="readonly")
        self.scatter_y_combo.grid(row=1, column=1)

        ttk.Button(left, text="產生圖表", command=self.on_generate_chart).pack(fill="x", pady=(10, 4))
        ttk.Button(left, text="＋ 加入此圖表到報表", command=self.on_add_chart_to_report).pack(fill="x")
        self.chart_report_count_var = tk.StringVar(value="目前報表中已有 0 張圖表")
        ttk.Label(left, textvariable=self.chart_report_count_var, foreground="#555").pack(anchor="w", pady=(8, 0))

        # 右側：圖表畫布
        right = ttk.Frame(self.tab_chart, padding=8)
        right.pack(side="left", fill="both", expand=True)
        self.chart_figure = Figure(figsize=(7, 5))
        self.chart_canvas = FigureCanvasTkAgg(self.chart_figure, master=right)
        self.chart_canvas.get_tk_widget().pack(fill="both", expand=True)

    # -- Tab 3：統計分析 --------------------------------------------------------
    def _build_tab_analysis(self):
        left = ttk.Frame(self.tab_analysis, padding=8)
        left.pack(side="left", fill="y")

        ttk.Label(left, text="分析方法", font=("", 10, "bold")).pack(anchor="w")
        self.analysis_method_var = tk.StringVar(value=core.ANALYSIS_METHODS[0])
        ttk.Combobox(left, textvariable=self.analysis_method_var, values=core.ANALYSIS_METHODS,
                     state="readonly", width=24).pack(pady=4)

        ttk.Label(left, text="分析欄位 (可多選)", font=("", 10, "bold")).pack(anchor="w", pady=(12, 0))
        self.analysis_col_listbox = tk.Listbox(left, selectmode="extended", width=24, height=12, exportselection=False)
        self.analysis_col_listbox.pack(pady=4)

        ttk.Button(left, text="執行分析", command=self.on_run_analysis).pack(fill="x", pady=(10, 4))
        ttk.Button(left, text="＋ 加入此結果到報表", command=self.on_add_analysis_to_report).pack(fill="x")
        self.analysis_report_count_var = tk.StringVar(value="目前報表中已有 0 項分析")
        ttk.Label(left, textvariable=self.analysis_report_count_var, foreground="#555").pack(anchor="w", pady=(8, 0))

        ttk.Label(left, text="※ Cpk 分析會使用「規格上下限設定」\n   中輸入的 LSL / USL", foreground="#777",
                  justify="left").pack(anchor="w", pady=(16, 0))

        # 右側：結果表格
        right = ttk.Frame(self.tab_analysis, padding=8)
        right.pack(side="left", fill="both", expand=True)
        ttk.Label(right, text="分析結果", font=("", 10, "bold")).pack(anchor="w")
        self.result_tree = ttk.Treeview(right, show="headings")
        self.result_tree.pack(fill="both", expand=True, pady=4)
        self.last_analysis_title = None
        self.last_analysis_df = None

    # ------------------------------------------------------------------
    # 事件處理：匯入 / 分頁 / 範圍
    # ------------------------------------------------------------------
    def on_import(self):
        path = filedialog.askopenfilename(
            title="選擇 Excel 檔案", filetypes=[("Excel 檔案", "*.xlsx *.xls *.xlsm")])
        if not path:
            return
        try:
            sheets = core.list_sheet_names(path)
        except Exception as e:
            messagebox.showerror("匯入失敗", f"無法讀取檔案：\n{e}")
            return
        self.filepath = path
        self.lbl_file.config(text=os.path.basename(path))
        self.sheet_combo["values"] = sheets
        self.sheet_combo.current(0)
        self._reset_report_state()
        self.on_sheet_selected()

    def on_sheet_selected(self, event=None):
        if not self.filepath or not self.sheet_var.get():
            return
        try:
            self.raw_df = core.load_sheet(self.filepath, self.sheet_var.get())
        except Exception as e:
            messagebox.showerror("讀取分頁失敗", str(e))
            return
        self.filtered_df = self.raw_df.copy()
        self.spec_limits = {}
        self.spec_entry_widgets = {}

        self.col_listbox.delete(0, "end")
        for c in self.raw_df.columns:
            self.col_listbox.insert("end", c)
        self.col_listbox.select_set(0, "end")  # 預設全選

        self.row_start_var.set("1")
        self.row_end_var.set(str(len(self.raw_df)))

        self._refresh_preview(self.filtered_df)
        self.status_var.set(
            f"已載入分頁「{self.sheet_var.get()}」，共 {len(self.raw_df)} 列、{len(self.raw_df.columns)} 欄")

    def on_apply_range(self):
        if self.raw_df is None:
            messagebox.showwarning("提醒", "請先匯入 Excel 檔案")
            return
        sel_idx = self.col_listbox.curselection()
        if not sel_idx:
            messagebox.showwarning("提醒", "請至少選擇一個欄位")
            return
        columns = [self.col_listbox.get(i) for i in sel_idx]

        try:
            row_start = int(self.row_start_var.get()) if self.row_start_var.get().strip() else None
        except ValueError:
            messagebox.showerror("錯誤", "起始列必須是整數")
            return
        try:
            row_end = int(self.row_end_var.get()) if self.row_end_var.get().strip() else None
        except ValueError:
            messagebox.showerror("錯誤", "結束列必須是整數")
            return

        self.filtered_df = core.filter_range(self.raw_df, columns=columns, row_start=row_start, row_end=row_end)
        self._refresh_preview(self.filtered_df)
        self._sync_column_choices()
        self.status_var.set(f"已套用範圍：{len(self.filtered_df)} 列 × {len(self.filtered_df.columns)} 欄")

    def _sync_column_choices(self):
        """把篩選後的欄位同步到 圖表 / 分析 tab 的清單，以及散佈圖 X/Y 下拉選單"""
        cols = list(self.filtered_df.columns)
        num_cols = core.numeric_columns(self.filtered_df)

        self.chart_col_listbox.delete(0, "end")
        for c in cols:
            self.chart_col_listbox.insert("end", c)
        self.chart_col_listbox.select_set(0, "end")

        self.analysis_col_listbox.delete(0, "end")
        for c in num_cols:
            self.analysis_col_listbox.insert("end", c)
        self.analysis_col_listbox.select_set(0, "end")

        self.scatter_x_combo["values"] = num_cols
        self.scatter_y_combo["values"] = num_cols
        if num_cols:
            self.scatter_x_var.set(num_cols[0])
            self.scatter_y_var.set(num_cols[min(1, len(num_cols) - 1)])

    def _refresh_preview(self, df):
        self.preview_tree.delete(*self.preview_tree.get_children())
        cols = list(df.columns)
        self.preview_tree["columns"] = cols
        for c in cols:
            self.preview_tree.heading(c, text=c)
            self.preview_tree.column(c, width=100, anchor="center")
        for _, row in df.head(20).iterrows():
            self.preview_tree.insert("", "end", values=list(row))
        self._sync_column_choices()

    def _reset_report_state(self):
        self.report_items_charts = []
        self.report_items_analyses = []
        self.chart_report_count_var.set("目前報表中已有 0 張圖表")
        self.analysis_report_count_var.set("目前報表中已有 0 項分析")

    # ------------------------------------------------------------------
    # 規格上下限 (LSL / USL) 對話框
    # ------------------------------------------------------------------
    def on_open_spec_dialog(self):
        if self.filtered_df is None:
            messagebox.showwarning("提醒", "請先匯入並套用範圍")
            return
        num_cols = core.numeric_columns(self.filtered_df)
        if not num_cols:
            messagebox.showinfo("提示", "目前選擇範圍內沒有數值型欄位")
            return

        win = tk.Toplevel(self)
        win.title("規格上下限設定 (LSL / USL)")
        win.geometry("420x480")

        canvas = tk.Canvas(win, borderwidth=0)
        frame = ttk.Frame(canvas)
        vsb = ttk.Scrollbar(win, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        canvas.create_window((0, 0), window=frame, anchor="nw")
        frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

        ttk.Label(frame, text="欄位", font=("", 9, "bold")).grid(row=0, column=0, padx=6, pady=6)
        ttk.Label(frame, text="LSL (下限)", font=("", 9, "bold")).grid(row=0, column=1, padx=6, pady=6)
        ttk.Label(frame, text="USL (上限)", font=("", 9, "bold")).grid(row=0, column=2, padx=6, pady=6)

        entry_map = {}
        for i, col in enumerate(num_cols, start=1):
            ttk.Label(frame, text=col).grid(row=i, column=0, sticky="w", padx=6, pady=3)
            lsl_var = tk.StringVar(value=str(self.spec_limits.get(col, {}).get("LSL", "")) if self.spec_limits.get(col, {}).get("LSL") is not None else "")
            usl_var = tk.StringVar(value=str(self.spec_limits.get(col, {}).get("USL", "")) if self.spec_limits.get(col, {}).get("USL") is not None else "")
            e_lsl = ttk.Entry(frame, textvariable=lsl_var, width=10)
            e_lsl.grid(row=i, column=1, padx=4)
            e_usl = ttk.Entry(frame, textvariable=usl_var, width=10)
            e_usl.grid(row=i, column=2, padx=4)
            entry_map[col] = (lsl_var, usl_var)

        def save_and_close():
            new_limits = {}
            for col, (lsl_var, usl_var) in entry_map.items():
                lsl_text, usl_text = lsl_var.get().strip(), usl_var.get().strip()
                try:
                    lsl = float(lsl_text) if lsl_text else None
                    usl = float(usl_text) if usl_text else None
                except ValueError:
                    messagebox.showerror("錯誤", f"「{col}」的規格值必須是數字")
                    return
                if lsl is not None or usl is not None:
                    new_limits[col] = {"LSL": lsl, "USL": usl}
            self.spec_limits = new_limits
            self.status_var.set(f"已儲存 {len(new_limits)} 個欄位的規格上下限")
            win.destroy()

        btn_frame = ttk.Frame(win)
        btn_frame.pack(side="bottom", fill="x", pady=6)
        ttk.Button(btn_frame, text="儲存並關閉", command=save_and_close).pack(side="right", padx=8)

    # ------------------------------------------------------------------
    # 圖表產生
    # ------------------------------------------------------------------
    def on_generate_chart(self):
        if self.filtered_df is None:
            messagebox.showwarning("提醒", "請先匯入並套用範圍")
            return
        sel_idx = self.chart_col_listbox.curselection()
        if not sel_idx:
            messagebox.showwarning("提醒", "請至少選擇一個欄位")
            return
        columns = [self.chart_col_listbox.get(i) for i in sel_idx]
        chart_type = self.chart_type_var.get()

        try:
            if chart_type == "散佈圖":
                x_col, y_col = self.scatter_x_var.get(), self.scatter_y_var.get()
                if not x_col or not y_col:
                    messagebox.showwarning("提醒", "散佈圖需要選擇 X 軸與 Y 軸欄位")
                    return
                core.draw_chart(self.chart_figure, self.filtered_df, columns, chart_type,
                                 self.spec_limits, x_col=x_col, y_col=y_col,
                                 title=f"{y_col} vs {x_col}")
            else:
                num_cols = [c for c in columns if pd.api.types.is_numeric_dtype(self.filtered_df[c])] \
                    if chart_type != "圓餅圖" else columns
                use_cols = num_cols if num_cols else columns
                core.draw_chart(self.chart_figure, self.filtered_df, use_cols, chart_type,
                                 self.spec_limits, title=f"{chart_type}")
            self.chart_canvas.draw()
            self.status_var.set(f"已產生「{chart_type}」")
        except Exception as e:
            traceback.print_exc()
            messagebox.showerror("繪圖失敗", str(e))

    def on_add_chart_to_report(self):
        if self.chart_figure is None or not self.chart_figure.axes:
            messagebox.showwarning("提醒", "請先產生圖表")
            return
        png = core.fig_to_png_bytes(self.chart_figure)
        title = f"{self.chart_type_var.get()}_{len(self.report_items_charts) + 1}"
        self.report_items_charts.append({"title": title, "png_bytes": png})
        self.chart_report_count_var.set(f"目前報表中已有 {len(self.report_items_charts)} 張圖表")
        self.status_var.set(f"已將圖表「{title}」加入報表")

    # ------------------------------------------------------------------
    # 統計分析
    # ------------------------------------------------------------------
    def on_run_analysis(self):
        if self.filtered_df is None:
            messagebox.showwarning("提醒", "請先匯入並套用範圍")
            return
        sel_idx = self.analysis_col_listbox.curselection()
        if not sel_idx:
            messagebox.showwarning("提醒", "請至少選擇一個數值欄位")
            return
        columns = [self.analysis_col_listbox.get(i) for i in sel_idx]
        method = self.analysis_method_var.get()

        try:
            if method == "Cpk 製程能力分析":
                missing = [c for c in columns if c not in self.spec_limits
                           or (self.spec_limits[c].get("LSL") is None and self.spec_limits[c].get("USL") is None)]
                if missing:
                    proceed = messagebox.askyesno(
                        "尚未設定規格",
                        f"以下欄位尚未設定 LSL/USL，將以「未設定」處理：\n{', '.join(missing)}\n\n是否繼續？")
                    if not proceed:
                        return
                result_df = core.calculate_cpk_batch(self.filtered_df, columns, self.spec_limits)
                title = "Cpk製程能力分析"
            elif method == "相關性分析":
                if len(columns) < 2:
                    messagebox.showwarning("提醒", "相關性分析至少需要選擇2個欄位")
                    return
                corr = core.calculate_correlation(self.filtered_df, columns)
                result_df = corr.reset_index().rename(columns={"index": "欄位"})
                title = "相關性分析"
            elif method == "敘述統計":
                result_df = core.calculate_descriptive(self.filtered_df, columns)
                title = "敘述統計"
            else:
                return

            self._show_result_table(result_df)
            self.last_analysis_title = title
            self.last_analysis_df = result_df
            self.status_var.set(f"已完成「{method}」")
        except Exception as e:
            traceback.print_exc()
            messagebox.showerror("分析失敗", str(e))

    def _show_result_table(self, df):
        self.result_tree.delete(*self.result_tree.get_children())
        cols = list(df.columns)
        self.result_tree["columns"] = cols
        for c in cols:
            self.result_tree.heading(c, text=c)
            self.result_tree.column(c, width=110, anchor="center")
        for _, row in df.iterrows():
            self.result_tree.insert("", "end", values=list(row))

    def on_add_analysis_to_report(self):
        if self.last_analysis_df is None:
            messagebox.showwarning("提醒", "請先執行分析")
            return
        title = f"{self.last_analysis_title}_{len(self.report_items_analyses) + 1}"
        self.report_items_analyses.append({"title": title, "df": self.last_analysis_df.copy()})
        self.analysis_report_count_var.set(f"目前報表中已有 {len(self.report_items_analyses)} 項分析")
        self.status_var.set(f"已將分析「{title}」加入報表")

    # ------------------------------------------------------------------
    # 匯出
    # ------------------------------------------------------------------
    def on_export(self):
        if self.filtered_df is None:
            messagebox.showwarning("提醒", "請先匯入 Excel 並選擇範圍")
            return
        if not self.report_items_charts and not self.report_items_analyses:
            proceed = messagebox.askyesno(
                "提醒", "目前尚未加入任何圖表或分析結果到報表，僅會匯出資料本身。是否繼續？")
            if not proceed:
                return

        default_name = "分析報表.xlsx"
        path = filedialog.asksaveasfilename(
            title="匯出完整 Excel 報表", defaultextension=".xlsx",
            initialfile=default_name, filetypes=[("Excel 檔案", "*.xlsx")])
        if not path:
            return

        try:
            core.export_report(
                path, self.filtered_df,
                charts=self.report_items_charts,
                analyses=self.report_items_analyses,
                spec_limits=self.spec_limits,
                sheet_prefix=self.sheet_var.get() or "資料",
            )
            self.status_var.set(f"報表已匯出至：{path}")
            messagebox.showinfo("完成", f"報表已成功匯出至：\n{path}")
        except Exception as e:
            traceback.print_exc()
            messagebox.showerror("匯出失敗", str(e))


if __name__ == "__main__":
    app = ExcelAnalyzerApp()
    app.mainloop()
