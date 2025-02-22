import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import pandas as pd
import pdfplumber


class TextComparator:
    def __init__(self, master):
        self.master = master
        self.master.title("文本与Excel比较工具")

        self.input_file_path = tk.StringVar()
        self.excel_files_paths = []

        self.create_widgets()

    def create_widgets(self):
        tk.Label(self.master, text="输入文件 (TXT 或 PDF):").grid(row=0, column=0, sticky='e', padx=5, pady=5)
        tk.Entry(self.master, textvariable=self.input_file_path, width=50).grid(row=0, column=1, padx=5, pady=5)
        tk.Button(self.master, text="浏览", command=self.browse_input_file).grid(row=0, column=2, padx=5, pady=5)

        tk.Label(self.master, text="Excel文件:").grid(row=1, column=0, sticky='e', padx=5, pady=5)
        self.excel_listbox = tk.Listbox(self.master, selectmode=tk.MULTIPLE, height=5, width=50, exportselection=0)
        self.excel_listbox.grid(row=1, column=1, padx=5, pady=5)
        tk.Button(self.master, text="添加Excel文件", command=self.add_excel_files).grid(row=1, column=2, padx=5, pady=5)

        tk.Button(self.master, text="比较", command=self.compare_files).grid(row=2, column=1, pady=5)
        tk.Button(self.master, text="清除", command=self.clear_all).grid(row=2, column=2, pady=5)

        self.result_text = scrolledtext.ScrolledText(self.master, wrap=tk.WORD, width=70, height=15)
        self.result_text.grid(row=3, column=0, columnspan=3, padx=5, pady=5)

        self.result_text.tag_configure("red", foreground="red")
        self.result_text.tag_configure("green", foreground="green")

    def browse_input_file(self):
        self.input_file_path.set(filedialog.askopenfilename(filetypes=[("文本或PDF文件", "*.txt *.pdf")]))

    def add_excel_files(self):
        files = filedialog.askopenfilenames(filetypes=[("Excel文件", "*.xlsx")])
        for file in files:
            if file not in self.excel_files_paths:
                self.excel_files_paths.append(file)
                self.excel_listbox.insert(tk.END, file)

    def clear_all(self):
        self.input_file_path.set("")
        self.excel_files_paths.clear()
        self.excel_listbox.delete(0, tk.END)
        self.result_text.delete(1.0, tk.END)

    def compare_files(self):
        if not self.input_file_path.get() or not self.excel_files_paths:
            messagebox.showerror("错误", "请同时选择输入文件和至少一个Excel文件。")
            return

        try:
            # 读取输入文件（支持txt和pdf）
            input_file = self.input_file_path.get()
            if input_file.endswith('.txt'):
                with open(input_file, 'r', encoding='utf-8') as f:
                    input_lines = f.readlines()
                input_data_raw = [line.rstrip() for line in input_lines if line.strip()]
            elif input_file.endswith('.pdf'):
                with pdfplumber.open(input_file) as pdf:
                    input_data_raw = []
                    for page in pdf.pages:
                        text = page.extract_text()
                        if text:
                            input_data_raw.extend([line.rstrip() for line in text.split('\n') if line.strip()])
            else:
                messagebox.showerror("错误", "不支持的文件格式，仅支持 .txt 和 .pdf")
                return

            # 读取所有Excel文件并清理数据
            all_excel_data = set()
            for excel_file in self.excel_files_paths:
                df = pd.read_excel(excel_file, header=None, dtype=str)
                cleaned_data = df.apply(lambda x: x.str.strip() if x.dtype == 'object' else x).fillna('')
                all_excel_data.update(cleaned_data.values.flatten())

            # 显示调试信息
            self.result_text.delete(1.0, tk.END)
            self.result_text.insert(tk.END, "输入文件原始数据样本（前5项）：\n")
            self.result_text.insert(tk.END, f"{input_data_raw[:5]}\n\n")
            self.result_text.insert(tk.END, "Excel数据样本（前5项）：\n")
            self.result_text.insert(tk.END, f"{list(all_excel_data)[:5]}\n\n")

            # 跟踪已见过和重复的项
            seen_items = set()  # 用于去重单个项
            duplicate_items = set()
            match_status = {}
            seen_lines = set()  # 用于去重整行

            # 处理所有项，识别重复项并比对唯一项
            for line in input_data_raw:
                items = line.split()
                for item in items:
                    if item in seen_items:
                        duplicate_items.add(item)
                    else:
                        seen_items.add(item)
                        match_status[item] = item in all_excel_data

            # 显示比较结果（去重显示）
            if not input_data_raw:
                self.result_text.insert(tk.END, "输入文件为空。")
                return

            mismatches = 0
            self.result_text.insert(tk.END, "比较结果：\n")
            for line in input_data_raw:
                if line in seen_lines:
                    continue  # 跳过重复行
                seen_lines.add(line)

                items = line.split()
                display_line = line
                self.result_text.insert(tk.END, display_line + "\n")
                start_pos = self.result_text.index("end-2l")

                for i, item in enumerate(items):
                    item_start = display_line.index(item,
                                                    i > 0 and display_line.index(items[i - 1]) + len(items[i - 1]) or 0)
                    item_end = item_start + len(item)
                    if item in duplicate_items:
                        self.result_text.tag_add("green", f"{start_pos}+{item_start}c", f"{start_pos}+{item_end}c")
                    elif not match_status[item]:
                        self.result_text.tag_add("red", f"{start_pos}+{item_start}c", f"{start_pos}+{item_end}c")
                        mismatches += 1

            # 显示统计信息
            if mismatches > 0:
                self.result_text.insert(tk.END, f"\n找到 {mismatches} 个未匹配项", "red")
            else:
                self.result_text.insert(tk.END, "\n未找到任何未匹配项。")

        except Exception as e:
            messagebox.showerror("错误", f"发生错误：{str(e)}")


if __name__ == "__main__":
    root = tk.Tk()
    app = TextComparator(root)
    root.mainloop()