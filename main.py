import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, messagebox

# 初始化全局变量
df_small = None
df_class = None
result_df = pd.DataFrame()

def load_file(file_label, df_global, file_type):
    file_path = filedialog.askopenfilename(
        title=f"选择{file_type}文件",
        filetypes=[("Excel files", "*.xlsx")],
        initialdir=os.path.expanduser('~')
    )
    if file_path:
        try:
            df = pd.read_excel(file_path)
            globals()[df_global] = df
            file_label.config(text=file_path)  # 修正这里
            print(f"{file_type}文件加载成功：", file_path)
        except Exception as e:
            messagebox.showerror("错误", f"读取{file_type}文件时出错：{e}")

def synthesize_files():
    global df_small, df_class, result_df
    try:
        if df_small is None or df_class is None:
            messagebox.showerror("错误", "请先选择两个文件。")
            return

        # 假设小分文件的第二列是姓名，班级姓名文件的第一列是姓名
        names_small = df_small.iloc[:, 1].astype(str)
        names_class = df_class.iloc[:, 0].astype(str)

        common_names = names_small[names_small.isin(names_class)]
        if common_names.empty:
            messagebox.showinfo("注意", "没有找到匹配的姓名。")
            return

        result_df = df_small[names_small.isin(common_names)]
        
        # 将结果保存到result.xlsx文件中
        result_df.to_excel('result.xlsx', index=False)
        messagebox.showinfo("完成", "合成完成，结果文件已保存。")
    except Exception as e:
        messagebox.showerror("错误", f"合成文件时出错：{e}")

# 创建主窗口
root = tk.Tk()
root.title("姓名匹配与数据合成工具")

# 文件路径显示变量
small_file_path = tk.StringVar()
class_file_path = tk.StringVar()

# 创建选择小分文件的按钮和显示框
small_file_label = tk.Label(root, text="选择小分.xlsx文件：")
small_file_label.pack()
small_file_button = tk.Button(root, text="选择文件", command=lambda: load_file(small_file_label, 'df_small', '小分'))
small_file_button.pack()
small_file_entry = tk.Entry(root, textvariable=small_file_path, width=80)
small_file_entry.pack()

# 创建选择班级姓名文件的按钮和显示框
class_file_label = tk.Label(root, text="选择班级姓名.xlsx文件：")
class_file_label.pack()
class_file_button = tk.Button(root, text="选择文件", command=lambda: load_file(class_file_label, 'df_class', '班级姓名'))
class_file_button.pack()
class_file_entry = tk.Entry(root, textvariable=class_file_path, width=80)
class_file_entry.pack()

# 创建合成按钮
synthesize_button = tk.Button(root, text="合成", command=synthesize_files)
synthesize_button.pack(pady=20)

# 运行主循环
root.mainloop()
