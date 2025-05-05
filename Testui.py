import tkinter as tk

def login():
    username = entry_user.get()
    password = entry_pass.get()
    print(f"用户名: {username}, 密码: {password}")

# 创建主窗口
root = tk.Tk()
root.title("登录界面")
root.geometry("300x150")

# 创建控件
label_user = tk.Label(root, text="用户名:")
entry_user = tk.Entry(root)

label_pass = tk.Label(root, text="密码:")
entry_pass = tk.Entry(root, show="*")

btn_login = tk.Button(root, text="登录", command=login)

# 布局控件（使用 grid）
label_user.grid(row=0, column=0, padx=10, pady=10, sticky="e")
entry_user.grid(row=0, column=1, padx=10, pady=10)

label_pass.grid(row=1, column=0, padx=10, pady=10, sticky="e")
entry_pass.grid(row=1, column=1, padx=10, pady=10)

btn_login.grid(row=2, column=0, columnspan=2, pady=10)

# 进入主事件循环
root.mainloop()
