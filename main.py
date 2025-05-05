
import tkinter as tk  # 用于生成用户界面
from tkinter import ttk

import excel_utils.excel_utils as excel_utils  # 导入 excel_utils 模块
# from excel_utils.excel_utils import select_execel_file, process_excel_file, save_json_file, show_messagebox


# 存储不同工具的页面（Frame）
tool_pages = {}

# 切换页面函数


def show_page(tool_name, frame):
    # 清除右侧区域
    for widget in frame.winfo_children():
        widget.destroy()

    # 如果有对应页面，显示它；否则显示默认空白页
    if tool_name in tool_pages:
        tool_pages[tool_name](frame)


# 示例：添加几个按钮和对应页面
tools = [
    ("Excel转JSON", "excel_json"),
    ("编码转换", "encoding_convert"),
    ("占位功能", "placeholder"),
]


# 添加各页面的实际 UI（后续你可以扩展每个工具的功能页面）

# Excel转JSON工具页面
def create_excel_json_page(parent):
    label = tk.Label(parent, text="这是 Excel 转 JSON 工具界面", font=("Arial", 14))
    label.pack(pady=20)
    excel_utils.create_excel_tool_ui(parent)


def create_encoding_convert_page(parent):
    label = tk.Label(parent, text="这是 编码转换 工具界面", font=("Arial", 14))
    label.pack(pady=20)


def create_placeholder_page(parent):
    label = tk.Label(parent, text="这是一个占位工具界面", font=("Arial", 14))
    label.pack(pady=20)


def create_tool_page(left_frame,right_frame):
    for tool_label, tool_key in tools:
        def make_callback(name=tool_key):
            return lambda: show_page(name,right_frame)
        btn = tk.Button(left_frame, text=tool_label, width=20, command=make_callback())
        btn.pack(pady=5)


def main():
    # 创建主窗口
    root = tk.Tk()
    root.title("工具集合")
    root.geometry("768x500")
    # 左侧工具栏框架（固定宽度）
    left_frame = tk.Frame(root, width=200, bg="#f0f0f0")
    left_frame.pack(side="left", fill="y")
    # 右侧内容区框架（可变化宽度）
    right_frame = tk.Frame(root, bg="#C7C7C7")
    right_frame.pack(side="right", expand=True, fill="both")
# 创建按钮和页面内容（这里只是演示）
    create_tool_page(left_frame,right_frame)

        # 注册页面函数到字典
    tool_pages["excel_json"] = create_excel_json_page
    tool_pages["encoding_convert"] = create_encoding_convert_page
    tool_pages["placeholder"] = create_placeholder_page
    # 运行主循环
    root.mainloop()

# 运行主程序
if __name__ == "__main__":
    main()
