import tkinter as tk
from tkinter import filedialog, scrolledtext, messagebox, ttk
import pdfplumber
import pandas as pd
import re
import threading
import sys

def run_command_line(main_text_path, compare_file_paths):
    # 保持原命令行模式不变
    def clean_main_text(text):
        return re.sub(r'\s+', ' ', text).strip().split()

    def clean_compare_text(text):
        return re.sub(r'\s+', '', text).strip()

    def read_main_file(file_path):
        if file_path.endswith('.txt'):
            with open(file_path, 'r', encoding='utf-8') as f:
                return clean_main_text(f.read())
        elif file_path.endswith('.pdf'):
            text = []
            with pdfplumber.open(file_path) as pdf:
                for page in pdf.pages:
                    text.append(page.extract_text(layout=False) or "")
            return clean_main_text(" ".join(text))
        return []

    def read_compare_file(file_path):
        if file_path.endswith(('.xlsx', '.xls')):
            df = pd.read_excel(file_path)
            words = []
            for val in df.values.flatten():
                if pd.notna(val):
                    cleaned = clean_compare_text(str(val))
                    words.append(cleaned)
            return "".join(words)
        return ""

    main_text = read_main_file(main_text_path)
    compare_text = ""
    for path in compare_file_paths:
        compare_text += read_compare_file(path)

    unique_words = sorted(set(main_text))
    found = [word for word in unique_words if word in compare_text]
    not_found = [word for word in unique_words if word not in compare_text]

    with open("C:/Temp/results.txt", "w", encoding='utf-8') as f:
        f.write("校对结果：\n\n")
        if found:
            f.write(f"匹配成功（{len(found)}个）：\n")
            for i, word in enumerate(found, 1):
                f.write(f"{i}. {word}\n")
        if not_found:
            f.write(f"\n未匹配（{len(not_found)}个）：\n")
            for i, word in enumerate(not_found, 1):
                f.write(f"{i}. {word}\n")

if len(sys.argv) > 2:
    main_file = sys.argv[1]
    compare_files = sys.argv[2:]
    run_command_line(main_file, compare_files)
    sys.exit(0)

class TextCheckerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("文本校对工具")
        self.root.geometry("800x600")

        self.main_text = []
        self.compare_text = ""
        self.not_found = []
        self.found = []
        self.current_index = -1
        self.main_file_path = ""
        self.compare_file_paths = []

        self.create_widgets()
        self.setup_bindings()

    def create_widgets(self):
        file_frame = ttk.LabelFrame(self.root, text="文件管理")
        file_frame.pack(pady=5, padx=10, fill=tk.X)

        self.main_file_btn = ttk.Button(file_frame, text="上传主文件 (TXT/PDF)", command=self.upload_main_file)
        self.main_file_btn.grid(row=0, column=0, padx=5)
        self.main_file_label = ttk.Label(file_frame, text="主文件: 未选择", anchor="w")
        self.main_file_label.grid(row=0, column=1, sticky=tk.EW)

        self.compare_file_btn = ttk.Button(file_frame, text="上传校对文件 (Excel)", command=self.upload_compare_files)
        self.compare_file_btn.grid(row=1, column=0, padx=5, pady=2)
        self.compare_file_label = ttk.Label(file_frame, text="校对文件: 未选择", anchor="w")
        self.compare_file_label.grid(row=1, column=1, sticky=tk.EW)

        control_frame = ttk.Frame(self.root)
        control_frame.pack(pady=5)
        ttk.Button(control_frame, text="清空", command=self.reset).pack(side=tk.LEFT, padx=2)
        ttk.Button(control_frame, text="查看清洗内容", command=self.show_cleaned_content).pack(side=tk.LEFT, padx=2)

        result_frame = ttk.LabelFrame(self.root, text="校对结果")
        result_frame.pack(pady=5, padx=10, fill=tk.BOTH, expand=True)

        self.result_text = scrolledtext.ScrolledText(result_frame, width=90, height=20)
        self.result_text.pack(fill=tk.BOTH, expand=True)

        nav_frame = ttk.Frame(result_frame)
        nav_frame.pack(pady=5)
        ttk.Button(nav_frame, text="上一个", command=self.prev_not_found).pack(side=tk.LEFT, padx=5)
        ttk.Button(nav_frame, text="下一个", command=self.next_not_found).pack(side=tk.LEFT, padx=5)
        ttk.Button(nav_frame, text="定位未找到词", command=self.locate_not_found).pack(side=tk.LEFT, padx=5)

        self.status_var = tk.StringVar(value="就绪")
        self.status_bar = ttk.Label(self.root, textvariable=self.status_var, anchor=tk.W)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)

    def setup_bindings(self):
        self.root.bind("<Control-f>", lambda e: self.locate_not_found())

    def clean_main_text(self, text):
        return re.sub(r'\s+', ' ', text).strip()

    def clean_compare_text(self, text):
        return re.sub(r'\s+', '', text).strip()

    def read_file(self, file_path, is_compare=False):
        try:
            if is_compare:
                if file_path.endswith(('.xlsx', '.xls')):
                    df = pd.read_excel(file_path)
                    words = []
                    for val in df.values.flatten():
                        if pd.notna(val):
                            cleaned = self.clean_compare_text(str(val))
                            words.append(cleaned)
                    return "".join(words)
                raise ValueError("仅支持Excel文件")

            if file_path.endswith('.txt'):
                with open(file_path, 'r', encoding='utf-8') as f:
                    return self.clean_main_text(f.read()).split()

            if file_path.endswith('.pdf'):
                text = []
                with pdfplumber.open(file_path) as pdf:
                    for page in pdf.pages:
                        text.append(page.extract_text(layout=False) or "")
                return self.clean_main_text(" ".join(text)).split()

            raise ValueError("不支持的文件格式")
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("错误", f"文件读取失败: {str(e)}"))
            return [] if not is_compare else ""

    def upload_main_file(self):
        path = filedialog.askopenfilename(filetypes=[("Text/PDF", "*.txt *.pdf")])
        if not path:
            return

        def process():
            self.status_var.set("正在解析主文件...")
            self.main_text = self.read_file(path)
            self.main_file_path = path
            self.root.after(0, self.update_main_display)

        threading.Thread(target=process).start()

    def update_main_display(self):
        self.main_file_label.config(text=f"主文件: {self.main_file_path}")
        self.result_text.delete(1.0, tk.END)
        self.result_text.insert(tk.END,
                              f"主文件已加载，总词数：{len(self.main_text)}，唯一词数：{len(set(self.main_text))}")
        self.status_var.set("就绪")

    def upload_compare_files(self):
        paths = filedialog.askopenfilenames(filetypes=[("Excel", "*.xlsx *.xls")])
        if not paths:
            return

        def process():
            self.status_var.set("正在处理校对文件并比对...")
            compare_text = []
            for path in paths:
                compare_text.append(self.read_file(path, is_compare=True))
            self.compare_text = "".join(compare_text)
            self.compare_file_paths = paths
            self.root.after(0, self.update_compare_display)
            # 上传完成后自动开始比对
            if self.main_text:
                self.start_comparison()
            else:
                self.status_var.set("请先上传主文件")

        threading.Thread(target=process).start()

    def update_compare_display(self):
        self.compare_file_label.config(text=f"校对文件: {', '.join(self.compare_file_paths)}")
        if not self.main_text:
            self.result_text.insert(tk.END, f"\n已加载校对文件，总字符数：{len(self.compare_text)}")

    def start_comparison(self):
        if not self.main_text:
            self.root.after(0, lambda: messagebox.showwarning("警告", "请先上传主文件！"))
            return

        def process():
            self.status_var.set("正在比对...")
            unique_words = sorted(set(self.main_text))
            self.found = [word for word in unique_words if word in self.compare_text]
            self.not_found = [word for word in unique_words if word not in self.compare_text]
            self.current_index = -1
            self.root.after(0, self.show_results)
            # 比对完成后自动跳转到第一个未匹配项
            if self.not_found:
                self.root.after(0, self.locate_not_found)

        threading.Thread(target=process).start()

    def show_results(self):
        self.result_text.delete(1.0, tk.END)
        self.result_text.insert(tk.END, "校对结果：\n\n")

        if self.found:
            self.result_text.insert(tk.END, f"匹配成功（{len(self.found)}个）：\n")
            for i, word in enumerate(self.found, 1):
                self.result_text.insert(tk.END, f"{i}. {word}\n")

        if self.not_found:
            self.result_text.insert(tk.END, f"\n未匹配（{len(self.not_found)}个）：\n", "red")
            for i, word in enumerate(self.not_found, 1):
                self.result_text.insert(tk.END, f"{i}. {word}\n", "red")
            self.result_text.tag_config("red", foreground="red")

        self.status_var.set(f"完成比对，匹配率：{len(self.found) / len(set(self.main_text)):.1%}")

    def show_cleaned_content(self):
        content_win = tk.Toplevel(self.root)
        content_win.title("清洗内容预览")
        content_win.geometry("800x600")

        notebook = ttk.Notebook(content_win)
        notebook.pack(fill=tk.BOTH, expand=True)

        main_frame = ttk.Frame(notebook)
        main_text = scrolledtext.ScrolledText(main_frame, wrap=tk.WORD)
        main_text.pack(fill=tk.BOTH, expand=True)
        main_text.insert(tk.END, " ".join(self.main_text))
        notebook.add(main_frame, text="主文件内容")

        compare_frame = ttk.Frame(notebook)
        compare_text = scrolledtext.ScrolledText(compare_frame, wrap=tk.WORD)
        compare_text.pack(fill=tk.BOTH, expand=True)
        compare_text.insert(tk.END, self.compare_text)
        notebook.add(compare_frame, text="校对文件内容")

    def prev_not_found(self):
        if not self.not_found:
            return
        self.current_index = max(0, self.current_index - 1)
        self.highlight_current()

    def next_not_found(self):
        if not self.not_found:
            return
        self.current_index = min(len(self.not_found) - 1, self.current_index + 1)
        self.highlight_current()

    def locate_not_found(self):
        if self.not_found:
            self.current_index = 0
            self.highlight_current()

    def highlight_current(self):
        self.result_text.tag_remove("highlight", 1.0, tk.END)
        if self.current_index >= 0:
            line = 4 + len(self.found) + self.current_index + 1
            self.result_text.tag_add("highlight", f"{line}.0", f"{line}.end")
            self.result_text.tag_config("highlight", background="yellow")
            self.result_text.see(f"{line}.0")

    def reset(self):
        self.main_text = []
        self.compare_text = ""
        self.found = []
        self.not_found = []
        self.current_index = -1
        self.main_file_path = ""
        self.compare_file_paths = []
        self.main_file_label.config(text="主文件: 未选择")
        self.compare_file_label.config(text="校对文件: 未选择")
        self.result_text.delete(1.0, tk.END)
        self.status_var.set("已重置")

if __name__ == "__main__":
    root = tk.Tk()
    app = TextCheckerApp(root)
    root.mainloop()