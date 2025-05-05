import os
import time
import tkinter as tk
from tkinter import filedialog, messagebox
import chardet

def detect_encoding(filepath):
    with open(filepath, 'rb') as f:
        raw = f.read(4096)  # 读取前4K字节用于检测
    result = chardet.detect(raw)
    return result['encoding']

def convert_to_utf8(filepath):
    encoding = detect_encoding(filepath)
    if not encoding or encoding.lower() == 'utf-8':
        return False  # 已是UTF-8，无需修改
    try:
        with open(filepath, 'r', encoding=encoding, errors='ignore') as f:
            content = f.read()
        with open(filepath, 'w', encoding='utf-8') as f:
            f.write(content)
        return True
    except Exception as e:
        print(f"Error converting {filepath}: {e}")
        return False

def process_directory(folder_path, log_callback):
    start_time = time.time()
    modified_count = 0
    file_count = 0

    for root, dirs, files in os.walk(folder_path):
        for file in files:
            if file.lower().endswith('.cs'):
                file_count += 1
                full_path = os.path.join(root, file)
                if convert_to_utf8(full_path):
                    modified_count += 1
                    log_callback(f"✔ Modified: {full_path}")
                else:
                    log_callback(f"⏩ Skipped: {full_path}")

    elapsed_time = time.time() - start_time
    return modified_count, file_count, elapsed_time

# -------- UI ------------
class App:
    def __init__(self, root):
        self.root = root
        self.root.title("C# 文件编码转 UTF-8 工具")
        self.root.geometry("600x400")

        self.label = tk.Label(root, text="请选择包含 .cs 文件的文件夹", font=("Arial", 12))
        self.label.pack(pady=10)

        self.select_button = tk.Button(root, text="选择文件夹", command=self.select_folder)
        self.select_button.pack(pady=5)

        self.run_button = tk.Button(root, text="开始转换", command=self.run, state='disabled')
        self.run_button.pack(pady=5)

        self.text = tk.Text(root, height=15)
        self.text.pack(pady=10, padx=10, fill=tk.BOTH, expand=True)

        self.folder_path = ""

    def log(self, msg):
        self.text.insert(tk.END, msg + "\n")
        self.text.see(tk.END)
        self.root.update()

    def select_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.folder_path = folder
            self.label.config(text=f"已选择文件夹: {folder}")
            self.run_button.config(state='normal')

    def run(self):
        self.text.delete(1.0, tk.END)
        self.log(f"开始处理文件夹: {self.folder_path}")
        modified, total, elapsed = process_directory(self.folder_path, self.log)
        self.log(f"\n完成 ✅\n总文件数: {total}\n修改数量: {modified}\n耗时: {elapsed:.2f} 秒")
        messagebox.showinfo("完成", f"修改完成！\n总文件数: {total}\n修改数量: {modified}\n耗时: {elapsed:.2f} 秒")

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
