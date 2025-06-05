import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import re
import pandas as pd
import os
from pathlib import Path

class TextLocalizationGUI:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("文本本地化Excel生成器")
        self.root.geometry("768x700")  # 增加高度以容纳新功能
        self.root.resizable(True, True)
        
        # 数据存储
        self.data = []
        self.output_path = tk.StringVar()
        self.output_path.set(os.path.join(os.getcwd(), "text_localization.xlsx"))
        
        self.setup_ui()
    
    def setup_ui(self):
        # 主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 配置权重使界面可以调整大小
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(1, weight=1)
        
        # 标题
        title_label = ttk.Label(main_frame, text="文本本地化Excel生成器", 
                               font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 10))
        
        # 输入区域标签
        input_label = ttk.Label(main_frame, text="请输入TextData代码:")
        input_label.grid(row=1, column=0, sticky=tk.NW, padx=(0, 10))
        
        # 文本输入区域 - 使用ScrolledText
        self.text_input = scrolledtext.ScrolledText(
            main_frame, 
            width=60, 
            height=15,
            font=("Consolas", 10),
            wrap=tk.WORD
        )
        self.text_input.grid(row=1, column=1, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), 
                            padx=(0, 0), pady=(0, 10))
        
        # 添加示例文本（包含注释的格式）
        sample_text = '''new TextData(TextKeyword.START_NEW_GAME, "ゲームを始める"), //开始游戏
new TextData(TextKeyword.CONTINUE_GAME, "ゲームを続ける"), //继续游戏
new TextData(TextKeyword.LOADING, "Loading..."),
new TextData(TextKeyword.TYT_JYB, "がんばろう！"), //加油吧！
new TextData(TextKeyword.TYT_JPYL, "賞品一覧"), //奖品一览
new TextData(TextKeyword.SETTINGS, "設定"),'''
        
        self.text_input.insert(tk.END, sample_text)
        
        # 操作按钮框架
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(10, 0))
        button_frame.columnconfigure(1, weight=1)
        
        # 解析按钮
        parse_button = ttk.Button(button_frame, text="解析数据", command=self.parse_data)
        parse_button.grid(row=0, column=0, padx=(0, 10))
        
        # 清空按钮
        clear_button = ttk.Button(button_frame, text="清空输入", command=self.clear_input)
        clear_button.grid(row=0, column=1, padx=(0, 10))
        
        # 预览按钮
        preview_button = ttk.Button(button_frame, text="预览数据", command=self.preview_data)
        preview_button.grid(row=0, column=2)
        
        # 文件输出设置框架
        output_frame = ttk.LabelFrame(main_frame, text="输出设置", padding="10")
        output_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(15, 0))
        output_frame.columnconfigure(1, weight=1)
        
        # 输出路径选择
        path_label = ttk.Label(output_frame, text="输出路径:")
        path_label.grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        
        path_entry = ttk.Entry(output_frame, textvariable=self.output_path, width=50)
        path_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 10))
        
        browse_button = ttk.Button(output_frame, text="浏览...", command=self.browse_output_path)
        browse_button.grid(row=0, column=2)
        
        # 生成Excel按钮
        generate_frame = ttk.Frame(main_frame)
        generate_frame.grid(row=4, column=0, columnspan=3, pady=(15, 0))
        
        self.generate_button = ttk.Button(
            generate_frame, 
            text="生成Excel文件", 
            command=self.generate_excel,
            style="Accent.TButton"
        )
        self.generate_button.pack()
        
        # 导入Excel功能框架
        import_frame = ttk.LabelFrame(main_frame, text="导入Excel生成代码", padding="10")
        import_frame.grid(row=6, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(15, 0))
        import_frame.columnconfigure(1, weight=1)
        
        # Excel文件路径选择
        excel_label = ttk.Label(import_frame, text="Excel文件:")
        excel_label.grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        
        self.excel_path = tk.StringVar()
        excel_entry = ttk.Entry(import_frame, textvariable=self.excel_path, width=50)
        excel_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 10))
        
        browse_excel_button = ttk.Button(import_frame, text="浏览...", command=self.browse_excel_file)
        browse_excel_button.grid(row=0, column=2)
        
        # 语言选择
        lang_label = ttk.Label(import_frame, text="生成语言:")
        lang_label.grid(row=1, column=0, sticky=tk.W, padx=(0, 10), pady=(10, 0))
        
        self.selected_language = tk.StringVar(value="中文")
        lang_frame = ttk.Frame(import_frame)
        lang_frame.grid(row=1, column=1, sticky=tk.W, pady=(10, 0))
        
        chinese_radio = ttk.Radiobutton(lang_frame, text="中文", variable=self.selected_language, value="中文")
        chinese_radio.pack(side=tk.LEFT, padx=(0, 20))
        
        english_radio = ttk.Radiobutton(lang_frame, text="英文", variable=self.selected_language, value="英文")
        english_radio.pack(side=tk.LEFT)
        
        # 导入和生成按钮
        import_button_frame = ttk.Frame(import_frame)
        import_button_frame.grid(row=2, column=0, columnspan=3, pady=(15, 0))
        
        import_button = ttk.Button(import_button_frame, text="导入Excel并生成代码", command=self.import_and_generate_code)
        import_button.pack(side=tk.LEFT, padx=(0, 10))
        
        copy_button = ttk.Button(import_button_frame, text="复制生成的代码", command=self.copy_generated_code)
        copy_button.pack(side=tk.LEFT)
        
        # 生成的代码显示区域
        code_label = ttk.Label(main_frame, text="生成的TextData代码:")
        code_label.grid(row=7, column=0, sticky=tk.NW, padx=(0, 10), pady=(15, 0))
        
        self.code_output = scrolledtext.ScrolledText(
            main_frame, 
            width=60, 
            height=10,
            font=("Consolas", 10),
            wrap=tk.WORD
        )
        self.code_output.grid(row=7, column=1, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), 
                             padx=(0, 0), pady=(15, 0))
        
        # 状态栏
        self.status_var = tk.StringVar()
        self.status_var.set("就绪")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, 
                              relief=tk.SUNKEN, anchor=tk.W)
        status_bar.grid(row=8, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(10, 0))
    
    def browse_excel_file(self):
        """浏览选择Excel文件"""
        filename = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[("Excel文件", "*.xlsx"), ("Excel文件", "*.xls"), ("所有文件", "*.*")]
        )
        if filename:
            self.excel_path.set(filename)
    
    def import_and_generate_code(self):
        """导入Excel并生成TextData代码"""
        excel_file = self.excel_path.get()
        if not excel_file:
            messagebox.showwarning("警告", "请先选择Excel文件！")
            return
        
        if not os.path.exists(excel_file):
            messagebox.showerror("错误", "Excel文件不存在！")
            return
        
        try:
            # 读取Excel文件
            df = pd.read_excel(excel_file, engine='openpyxl')
            
            # 检查必要的列是否存在
            required_columns = ['TextKeyword', '日文', '中文', '英文']
            missing_columns = [col for col in required_columns if col not in df.columns]
            
            if missing_columns:
                messagebox.showerror("错误", f"Excel文件缺少必要的列: {', '.join(missing_columns)}")
                return
            
            # 获取选择的语言
            selected_lang = self.selected_language.get()
            lang_column = selected_lang
            
            # 生成TextData代码
            code_lines = []
            success_count = 0
            
            for index, row in df.iterrows():
                keyword = row['TextKeyword']
                text_content = row[lang_column]
                
                # 跳过空值
                if pd.isna(keyword) or pd.isna(text_content) or not str(text_content).strip():
                    continue
                
                # 生成TextData行
                code_line = f'new TextData(TextKeyword.{keyword}, "{text_content}"),'
                code_lines.append(code_line)
                success_count += 1
            
            if code_lines:
                # 显示生成的代码
                generated_code = '\n'.join(code_lines)
                self.code_output.delete("1.0", tk.END)
                self.code_output.insert(tk.END, generated_code)
                
                self.status_var.set(f"成功生成 {success_count} 行{selected_lang}TextData代码")
                messagebox.showinfo("生成成功", f"成功从Excel生成 {success_count} 行{selected_lang}TextData代码！")
            else:
                self.status_var.set("没有找到有效的数据")
                messagebox.showwarning("警告", f"Excel文件中没有找到有效的{selected_lang}数据！")
                
        except Exception as e:
            error_msg = f"导入Excel文件时出错: {str(e)}"
            self.status_var.set("导入失败")
            messagebox.showerror("错误", error_msg)
    
    def copy_generated_code(self):
        """复制生成的代码到剪贴板"""
        code_content = self.code_output.get("1.0", tk.END).strip()
        if not code_content:
            messagebox.showinfo("提示", "没有代码可以复制！")
            return
        
        try:
            self.root.clipboard_clear()
            self.root.clipboard_append(code_content)
            self.status_var.set("代码已复制到剪贴板")
            messagebox.showinfo("成功", "代码已复制到剪贴板！")
        except Exception as e:
            messagebox.showerror("错误", f"复制失败: {str(e)}")
    
    def parse_textdata_line(self, line):
        """解析TextData行，忽略注释和额外空格"""
        # 移除行末注释 (// 注释内容)
        if '//' in line:
            comment_pos = line.find('//')
            line = line[:comment_pos].strip()
        
        # 清理行内容，去除前后空白和逗号
        line = line.strip().rstrip(',')
        
        # 增强的正则表达式，支持更多空格和字符变化
        # 支持TextKeyword中的数字、下划线、大小写字母
        pattern = r'new\s+TextData\s*\(\s*TextKeyword\.([A-Z0-9_]+)\s*,\s*"([^"]+)"\s*\)'
        match = re.search(pattern, line, re.IGNORECASE)
        
        if match:
            keyword = match.group(1).upper()  # 统一转为大写
            japanese_text = match.group(2)
            return keyword, japanese_text
        return None, None
    
    def parse_data(self):
        """解析输入的数据"""
        input_text = self.text_input.get("1.0", tk.END).strip()
        if not input_text:
            messagebox.showwarning("警告", "请先输入TextData代码！")
            return
        
        self.data = []
        lines = input_text.split('\n')
        success_count = 0
        failed_lines = []
        
        for line_num, line in enumerate(lines, 1):
            original_line = line.strip()
            if original_line and 'new TextData' in original_line:
                keyword, japanese_text = self.parse_textdata_line(original_line)
                if keyword and japanese_text:
                    self.data.append({
                        'TextKeyword': keyword,
                        '日文': japanese_text,
                        '中文': '',  # 保持空白
                        '英文': ''   # 保持空白
                    })
                    success_count += 1
                else:
                    failed_lines.append(f"第{line_num}行: {original_line}")
        
        if success_count > 0:
            status_msg = f"解析成功！共解析 {success_count} 条数据"
            if failed_lines:
                status_msg += f"，{len(failed_lines)} 条失败"
            self.status_var.set(status_msg)
            
            # 显示解析结果
            result_msg = f"成功解析 {success_count} 条TextData数据！"
            if failed_lines:
                result_msg += f"\n\n解析失败的行：\n" + "\n".join(failed_lines[:5])  # 只显示前5个失败的行
                if len(failed_lines) > 5:
                    result_msg += f"\n... 还有 {len(failed_lines) - 5} 行失败"
            
            messagebox.showinfo("解析完成", result_msg)
        else:
            self.status_var.set("解析失败，未找到有效的TextData")
            messagebox.showerror("解析失败", "未找到有效的TextData格式数据！")
    
    def clear_input(self):
        """清空输入框"""
        self.text_input.delete("1.0", tk.END)
        self.data = []
        self.status_var.set("输入已清空")
    
    def preview_data(self):
        """预览解析的数据"""
        if not self.data:
            messagebox.showinfo("提示", "请先解析数据！")
            return
        
        # 创建预览窗口
        preview_window = tk.Toplevel(self.root)
        preview_window.title("数据预览")
        preview_window.geometry("600x400")
        
        # 创建Treeview显示数据
        columns = ('序号', 'TextKeyword', '日文', '中文', '英文')
        tree = ttk.Treeview(preview_window, columns=columns, show='headings', height=15)
        
        # 设置列标题
        for col in columns:
            tree.heading(col, text=col)
            if col == '序号':
                tree.column(col, width=50)
            elif col == 'TextKeyword':
                tree.column(col, width=200)
            else:
                tree.column(col, width=120)
        
        # 添加数据
        for i, item in enumerate(self.data, 1):
            tree.insert('', tk.END, values=(
                i, 
                item['TextKeyword'], 
                item['日文'], 
                item['中文'], 
                item['英文']
            ))
        
        # 添加滚动条
        scrollbar = ttk.Scrollbar(preview_window, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(10, 0), pady=10)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y, padx=(0, 10), pady=10)
    
    def browse_output_path(self):
        """浏览输出路径"""
        filename = filedialog.asksaveasfilename(
            title="选择输出文件",
            defaultextension=".xlsx",
            filetypes=[("Excel文件", "*.xlsx"), ("所有文件", "*.*")]
        )
        if filename:
            self.output_path.set(filename)
    
    def generate_excel(self):
        """生成Excel文件"""
        if not self.data:
            messagebox.showwarning("警告", "请先解析数据！")
            return
        
        output_file = self.output_path.get()
        if not output_file:
            messagebox.showwarning("警告", "请选择输出路径！")
            return
        
        try:
            # 确保输出目录存在
            output_dir = os.path.dirname(output_file)
            if output_dir and not os.path.exists(output_dir):
                os.makedirs(output_dir)
            
            # 创建DataFrame并导出
            df = pd.DataFrame(self.data)
            df.to_excel(output_file, index=False, engine='openpyxl')
            
            self.status_var.set(f"Excel文件已生成: {output_file}")
            
            # 成功提示
            result = messagebox.askyesno(
                "生成成功", 
                f"Excel文件已成功生成！\n路径: {output_file}\n共导出 {len(self.data)} 条记录\n\n是否打开文件所在文件夹？"
            )
            
            if result:
                # 打开文件所在文件夹
                folder_path = os.path.dirname(os.path.abspath(output_file))
                if os.name == 'nt':  # Windows
                    os.startfile(folder_path)
                elif os.name == 'posix':  # macOS and Linux
                    os.system(f'open "{folder_path}"' if sys.platform == 'darwin' else f'xdg-open "{folder_path}"')
                    
        except Exception as e:
            error_msg = f"生成Excel文件时出错: {str(e)}"
            self.status_var.set("生成失败")
            messagebox.showerror("错误", error_msg)
    
    def run(self):
        """运行GUI"""
        self.root.mainloop()

def main():
    try:
        import pandas as pd
    except ImportError:
        messagebox.showerror("依赖缺失", "请先安装pandas库：\npip install pandas openpyxl")
        return
    
    app = TextLocalizationGUI()
    app.run()

if __name__ == "__main__":
    main()