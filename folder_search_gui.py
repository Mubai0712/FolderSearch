import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import csv
import os
import threading
from pathlib import Path

class FolderSearchApp:
    def __init__(self, root):
        self.root = root
        self.root.title("FolderSearch - 表格内容搜索工具")
        self.root.geometry("1100x700")
        self.root.minsize(800, 500)
        self.root.configure(bg="#f0f2f5")
        self.setup_ui()
        self.search_results = []

    def setup_ui(self):
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("TButton", padding=6, relief="flat", background="#4a90d9", foreground="white")
        style.map("TButton", background=[("active", "#357abd")])
        style.configure("TLabel", background="#f0f2f5")
        style.configure("Treeview", rowheight=25)

        top_frame = tk.Frame(self.root, bg="#f0f2f5", padx=15, pady=10)
        top_frame.pack(fill="x")

        tk.Label(top_frame, text="选择文件夹:", bg="#f0f2f5", font=("微软雅黑", 10)).grid(row=0, column=0, sticky="w")
        self.folder_var = tk.StringVar()
        self.folder_entry = tk.Entry(top_frame, textvariable=self.folder_var, width=50, font=("微软雅黑", 9))
        self.folder_entry.grid(row=0, column=1, padx=5, pady=5)
        self.browse_btn = ttk.Button(top_frame, text="浏览...", command=self.browse_folder)
        self.browse_btn.grid(row=0, column=2, padx=5)

        tk.Label(top_frame, text="关键词:", bg="#f0f2f5", font=("微软雅黑", 10)).grid(row=1, column=0, sticky="w", pady=10)
        self.keyword_var = tk.StringVar()
        self.keyword_entry = tk.Entry(top_frame, textvariable=self.keyword_var, width=40, font=("微软雅黑", 9))
        self.keyword_entry.grid(row=1, column=1, padx=5, sticky="w")
        self.keyword_entry.bind("<Return>", lambda e: self.start_search())

        self.search_mode = tk.StringVar(value="fuzzy")
        tk.Radiobutton(top_frame, text="模糊查询", variable=self.search_mode, value="fuzzy", bg="#f0f2f5").grid(row=1, column=2, padx=2, sticky="w")
        tk.Radiobutton(top_frame, text="精确查询", variable=self.search_mode, value="exact", bg="#f0f2f5").grid(row=1, column=3, padx=2, sticky="w")

        self.search_btn = ttk.Button(top_frame, text="🔍 搜索", command=self.start_search)
        self.search_btn.grid(row=1, column=4, padx=15)

        self.progress = ttk.Progressbar(top_frame, mode="indeterminate", length=150)
        self.progress.grid(row=2, column=0, columnspan=5, pady=5, sticky="ew")

        self.status_label = tk.Label(self.root, text="就绪", bg="#f0f2f5", font=("微软雅黑", 9), anchor="w", padx=15)
        self.status_label.pack(fill="x")

        result_frame = tk.Frame(self.root, bg="white", relief="solid", bd=1)
        result_frame.pack(fill="both", expand=True, padx=15, pady=(0,10))

        columns = ("#", "文件路径", "工作表", "行号", "列名", "单元格内容")
        self.tree = ttk.Treeview(result_frame, columns=columns, show="headings", height=15, selectmode="browse")
        self.tree.heading("#", text="序号")
        self.tree.heading("文件路径", text="文件路径")
        self.tree.heading("工作表", text="工作表/CSV")
        self.tree.heading("行号", text="行号")
        self.tree.heading("列名", text="列名")
        self.tree.heading("单元格内容", text="单元格内容（匹配部分高亮）")
        self.tree.column("#", width=50, anchor="center")
        self.tree.column("文件路径", width=300)
        self.tree.column("工作表", width=100)
        self.tree.column("行号", width=60, anchor="center")
        self.tree.column("列名", width=80)
        self.tree.column("单元格内容", width=400)
        self.tree.bind("<Double-1>", self.on_item_double_click)
        self.tree.bind("<Button-3>", self.show_context_menu)

        vsb = ttk.Scrollbar(result_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        self.context_menu = tk.Menu(self.root, tearoff=0)
        self.context_menu.add_command(label="打开文件", command=self.open_selected_file)
        self.context_menu.add_command(label="打开所在文件夹", command=self.open_containing_folder)

        btn_frame = tk.Frame(self.root, bg="#f0f2f5")
        btn_frame.pack(fill="x", padx=15, pady=5)
        ttk.Button(btn_frame, text="清除结果", command=self.clear_results).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="导出结果CSV", command=self.export_results).pack(side="left", padx=5)

    def browse_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.folder_var.set(folder)

    def start_search(self):
        folder = self.folder_var.get().strip()
        keyword = self.keyword_var.get().strip()
        if not folder:
            messagebox.showwarning("提示", "请先选择文件夹")
            return
        if not keyword:
            messagebox.showwarning("提示", "请输入搜索关键词")
            return
        if not os.path.isdir(folder):
            messagebox.showerror("错误", "文件夹路径无效")
            return
        self.clear_results()
        self.status_label.config(text="正在搜索中，请稍候...")
        self.progress.start()
        self.search_btn.config(state="disabled")
        threading.Thread(target=self.search_files, args=(folder, keyword), daemon=True).start()

    def search_files(self, folder, keyword):
        self.search_results = []
        supported_ext = (".xlsx", ".xls", ".csv")
        try:
            for root_dir, dirs, files in os.walk(folder):
                for file in files:
                    ext = os.path.splitext(file)[1].lower()
                    if ext in supported_ext:
                        filepath = os.path.join(root_dir, file)
                        try:
                            self.process_file(filepath, keyword)
                        except Exception as e:
                            print(f"处理文件出错: {filepath}, 错误: {e}")
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("搜索异常", str(e)))
        finally:
            self.root.after(0, self.search_finished)

    def process_file(self, filepath, keyword):
        ext = os.path.splitext(filepath)[1].lower()
        try:
            if ext == ".csv":
                self.process_csv(filepath, keyword)
            else:
                self.process_excel(filepath, keyword)
        except Exception as e:
            raise e

    def process_csv(self, filepath, keyword):
        with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
            reader = csv.reader(f)
            for row_idx, row in enumerate(reader, start=1):
                for col_idx, cell in enumerate(row, start=1):
                    if self.cell_match(str(cell), keyword):
                        self.search_results.append((
                            filepath,
                            "CSV",
                            row_idx,
                            col_idx,
                            str(cell)
                        ))

    def process_excel(self, filepath, keyword):
        xls = pd.ExcelFile(filepath)
        for sheet_name in xls.sheet_names:
            df = xls.parse(sheet_name, dtype=str).fillna("")
            for row_idx in range(len(df)):
                for col_idx, col_name in enumerate(df.columns, start=1):
                    cell_value = str(df.iloc[row_idx, col_idx - 1])
                    if self.cell_match(cell_value, keyword):
                        self.search_results.append((
                            filepath,
                            sheet_name,
                            row_idx + 2,
                            col_name if col_name else f"列{col_idx}",
                            cell_value
                        ))

    def cell_match(self, cell_value, keyword):
        if self.search_mode.get() == "exact":
            return cell_value == keyword
        else:
            return keyword.lower() in cell_value.lower()

    def search_finished(self):
        self.progress.stop()
        self.search_btn.config(state="normal")
        total = len(self.search_results)
        if total == 0:
            self.status_label.config(text="未找到匹配内容")
            messagebox.showinfo("结果", "未找到任何匹配的单元格")
        else:
            self.status_label.config(text=f"找到 {total} 个匹配项")
            for idx, result in enumerate(self.search_results, start=1):
                filepath, sheet, row, col, value = result
                display_value = value if len(value) <= 100 else value[:100] + "..."
                self.tree.insert("", "end", values=(idx, filepath, sheet, row, col, display_value))
            messagebox.showinfo("完成", f"搜索完成，共找到 {total} 个匹配项")

    def on_item_double_click(self, event):
        self.open_selected_file()

    def show_context_menu(self, event):
        selected = self.tree.selection()
        if selected:
            self.context_menu.post(event.x_root, event.y_root)

    def open_selected_file(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showinfo("提示", "请先选择一条结果")
            return
        item = self.tree.item(selected[0])
        filepath = item
    
      
        @keyframes spin {
          from { transform: rotate(0deg); }
          to { transform: rotate(360deg); }
        }
      
      
    ['values'][1]
        try:
            os.startfile(filepath)
        except Exception as e:
            messagebox.showerror("错误", f"无法打开文件: {e}")

    def open_containing_folder(self):
        selected = self.tree.selection()
        if not selected:
            return
        item = self.tree.item(selected[0])
        filepath = item['values'][1]
        try:
            os.startfile(os.path.dirname(filepath))
        except:
            pass

    def clear_results(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.search_results.clear()
        self.status_label.config(text="就绪")

    def export_results(self):
        if not self.search_results:
            messagebox.showinfo("提示", "没有结果可导出")
            return
        file_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV文件", "*.csv")],
            title="导出结果"
        )
        if file_path:
            try:
                with open(file_path, 'w', newline='', encoding='utf-8-sig') as f:
                    writer = csv.writer(f)
                    writer.writerow(["序号", "文件路径", "工作表", "行号", "列名", "单元格内容"])
                    for idx, res in enumerate(self.search_results, 1):
                        writer.writerow([idx, res[0], res[1], res[2], res[3], res[4]])
                messagebox.showinfo("成功", f"已导出到 {file_path}")
            except Exception as e:
                messagebox.showerror("导出失败", str(e))

if __name__ == "__main__":
    root = tk.Tk()
    app = FolderSearchApp(root)
    root.mainloop()
