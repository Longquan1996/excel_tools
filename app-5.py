# 增加pdf重命名功能
import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import re
import pdfplumber

def add_file():
    file = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    file = os.path.normpath(os.path.abspath(file))  # 获取标准化的绝对路径
    if file and (file not in file_list):
        file_list.append(file)
        file_display.insert(tk.END, os.path.basename(file) + "\n")
        update_sheet_options()

def select_folder():
    folder = filedialog.askdirectory()
    if folder:
        for file in os.listdir(folder):
            if file.endswith((".xlsx", ".xls")):
                full_path = os.path.join(folder, file)
                full_path = os.path.normpath(os.path.abspath(full_path))  # 获取标准化的绝对路径
                if full_path not in file_list:
                    file_list.append(full_path)
                    file_display.insert(tk.END, os.path.basename(full_path) + "\n")
        update_sheet_options()

def delete_file():
    selected_files = list(file_display.curselection()) # index  list
    if not selected_files:
        messagebox.showwarning("警告", "请选择要删除的文件！")
        return
    selected_files.sort(reverse=True) # 反向删除索引，避免顺序问题
    for selected in selected_files: # index
        file_list.pop(selected)  # 从 file_list 中移除文件
        file_display.delete(selected)  # 从显示区域中移除
    update_sheet_options()  # 更新页签选择列表

def update_sheet_options():
    file_count_label.set(f"共 {file_display.size()} 个文件")
    if not file_list:
        return
    try:
        first_file = file_list[0]
        excel_file = pd.ExcelFile(first_file)
        sheet_names = excel_file.sheet_names  # 获取所有工作表名称
        select_sheet['values'] = sheet_names
    except Exception as e:
        messagebox.showerror("错误", f"无法读取页签名称: {e}")

def clear_combined_data():
    global combined_data  # 确保在函数开始时声明 global
    if not combined_data:
        messagebox.showinfo("信息", "没有数据可清空。")
        return

    answer = messagebox.askyesno("确认", "您确定要清空汇总数据吗？")
    if answer:
        combined_data = {}  # 清空汇总结果
        update_combined_info()  # 更新已汇总页签信息
        save_button.config(state=tk.DISABLED)

def update_combined_info():
    combined_info = "\n".join(combined_data.keys()) if combined_data else "无"
    combined_info_label.set(f"已汇总的页签:\n{combined_info}")

def combine_excel_files():
    global combined_data
    if not file_list:
        messagebox.showerror("错误", "请添加至少一个 Excel 文件！")
        return

    if not select_sheet.get():
        messagebox.showerror("错误", "请选择至少一个页签！")
        return
    
    try:
        total = len(file_list)
        count = 0

        sheet_name = select_sheet.get()
        sheet_data = pd.DataFrame()

        for idx, file_path in enumerate(file_list):
            if idx == 0:  # 第一次读取，确定表头行
                df = pd.read_excel(file_path, sheet_name=sheet_name, header=None, dtype=str)
                if flag:
                    title_row = None
                    max_count = max([df.iloc[i].notnull().sum() for i in range(4)])
                    for i in range(3,-1,-1):
                        if df.iloc[i].notnull().sum() == max_count:
                            title_row = i
                            break
                    if title_row is None:
                        raise Exception(f"文件 {file_path} 中页签 {sheet_name} 没有找到表头行！请手动设置！")
                else:
                    title_row = int(entry.get())
                df = pd.read_excel(file_path, sheet_name=sheet_name, header=title_row, dtype=str)
                sheet_data = pd.concat([sheet_data, df], ignore_index=True, join='outer')

            else:  # 后续文件跳过表头行
                df = pd.read_excel(file_path, sheet_name=sheet_name, header=title_row, dtype=str)
                sheet_data = pd.concat([sheet_data, df], ignore_index=True, join='outer')
            count += 1
            progress_bar.config(value=count, maximum=total)
            root.update_idletasks() # 更新进度条显示
        combined_data[sheet_name] = sheet_data
        save_button.config(state=tk.NORMAL)

    except Exception as e:
        messagebox.showerror("未知错误", str(e))

    update_combined_info()

def save_combined_file():
    if not combined_data:
        messagebox.showerror("错误", "没有任何汇总数据！")
        return
    
    output_file = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if not output_file:
        messagebox.showinfo("信息", "您没有选择文件，保存操作已取消。")
        return
    try:
        with pd.ExcelWriter(output_file) as writer:
            for sheet_name, df in combined_data.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)
        messagebox.showinfo("完成", f"汇总文件已保存为: {output_file}")
    except Exception as e:
        messagebox.showerror("保存错误", str(e))

def show_help():
    messagebox.showinfo("帮助", "备注：\n请确保各excel的表头行内容一致。\n为了方便，所有内容都会当做文本来处理，需要调整格式的话请在合并后再处理。\n合并时请关闭需合并的excel，不然会漏掉。\n自动识别表头行，表头有合并或者表头超过第4行，可能会识别错误，识别错误的话需要自己看excel然后设置。")

def open_file(event):
    selected_index = file_display.curselection() # tuple
    if selected_index:
        file_name = file_list[selected_index[0]]
        try:
            os.startfile(file_name)
        except Exception as e:
            print(f"无法打开文件: {e}")

# 新窗口：表格拆分功能
def excel_split_window():
    split_window = tk.Toplevel(root)
    split_window.title("Excel 表格拆分工具")
    
    # 设置窗口大小
    window_width = 450
    window_height = 550

    # 获取主窗口的位置和大小
    main_x = root.winfo_x()
    main_y = root.winfo_y()
    main_width = root.winfo_width()
    main_height = root.winfo_height()

    # 计算 Toplevel 窗口的中心位置
    toplevel_x = main_x + main_width
    toplevel_y = main_y

    # 设置 Toplevel 窗口的几何参数
    split_window.geometry(f"{window_width}x{window_height}+{toplevel_x}+{toplevel_y}")

    # 选择文件
    def select_file():
        file = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
        if file:
            file_split.set(os.path.normpath(os.path.abspath(file)))
            file_name_var.set(os.path.basename(file))
            update_sheet_options()

    # 变量声明
    tk.Button(split_window, text="选择文件", command=select_file).grid(row=0, column=0, padx=5, pady=5)
    file_split = tk.StringVar()
    file_name_var = tk.StringVar()
    tk.Label(split_window, text="已选择文件:").grid(row=1, column=0, padx=5, pady=5, sticky="e")
    tk.Entry(split_window, textvariable=file_name_var, state="readonly", width=40).grid(row=1, column=1, padx=5, pady=5, sticky="w")

    tk.Label(split_window, text="请选择页签:").grid(row=2, column=0, padx=5, pady=5, sticky="e")
    select_sheet = ttk.Combobox(split_window, values=[], state="readonly")
    select_sheet.grid(row=2, column=1, padx=5, pady=5, sticky="w")

    def update_sheet_options():
        file_path = file_split.get()
        if not file_path:
            return
        try:
            excel_file = pd.ExcelFile(file_path)
            sheet_names = excel_file.sheet_names  # 获取所有工作表名称
            select_sheet['values'] = sheet_names
        except Exception as e:
            messagebox.showerror("错误", f"无法读取文件: {e}")

    tk.Label(split_window, text="请选择拆分依据:").grid(row=3, column=0, padx=5, pady=5, sticky="e")
    select_title = ttk.Combobox(split_window, values=[], state="readonly")
    select_title.grid(row=3, column=1, padx=5, pady=5, sticky="w")

    # 绑定一个事件，当选择页签时，更新拆分依据
    select_sheet.bind("<<ComboboxSelected>>", lambda event: update_title_options())
    title_row = 0
    def update_title_options():
        nonlocal title_row
        title_row = 0
        file_path = file_split.get()
        sheet_name = select_sheet.get()
        if not file_path or not sheet_name:
            return
        try:
            df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
            # title_row = 0
            max_count = max([df.iloc[i].notnull().sum() for i in range(4)])
            for i in range(3,-1,-1):
                if df.iloc[i].notnull().sum() == max_count:
                    title_row = i
                    break
            titles = df.iloc[title_row].tolist()
            select_title['values'] = titles
        except Exception as e:
            messagebox.showerror("错误", f"无法读取文件: {e}")

    # 拆分操作
    catagory_df = {}
    catagory_count = tk.IntVar(value=0)
    catagory_tk = tk.StringVar()
    catagory = []
    def split_excel():
        nonlocal catagory_df, catagory
        file_path = file_split.get()
        if not file_path:
            messagebox.showwarning("警告", "请选择要拆分的文件！")
            return

        sheet_name = select_sheet.get()
        if not sheet_name:
            messagebox.showwarning("警告", "请选择要拆分的工作表！")
            return
        
        select_title_name = select_title.get()
        if not select_title_name:
            messagebox.showwarning("警告", "请选择拆分依据！")
            return

        try:
            df = pd.read_excel(file_path, sheet_name=sheet_name, header=title_row, dtype=str)
            catagory = df[select_title_name].unique()
            catagory_tk.set(catagory)
            catagory_df = {}
            for i in catagory:
                catagory_df[i] = df[df[select_title_name] == i]
            catagory_count.set(len(catagory_df))
        except Exception as e:
            messagebox.showerror("拆分错误", str(e))
    tk.Button(split_window, text="开始拆分", command=split_excel).grid(row=4, column=1, padx=5, pady=5, sticky="w")

    # 显示拆分简要信息
    tk.Label(split_window, text="拆分数量:").grid(row=5, column=0, padx=5, pady=5, sticky="e")
    tk.Entry(split_window, textvariable=catagory_count, state="readonly", width=10).grid(row=5, column=1, padx=5, pady=5, sticky="w")
    tk.Label(split_window, text="具体:").grid(row=6, column=0, padx=5, pady=5, sticky="e")
    tk.Label(split_window, textvariable=catagory_tk, relief="sunken", anchor='w', padx=5, pady=5).grid(row=6, column=1, sticky="w")
    
    mode = tk.IntVar(value=0)
    tk.Radiobutton(split_window, text="保存为多个文件", variable=mode, value=1).grid(row=7, column=0, padx=5, pady=5, sticky="w")
    tk.Radiobutton(split_window, text="保存为多个页签", variable=mode, value=2).grid(row=7, column=1, padx=5, pady=5, sticky="w")

    # 保存拆分结果
    def save_split_file():
        if mode.get() == 0:
            messagebox.showwarning("警告", "请选择保存模式！")
            return
        if mode.get() == 1: # 保存为多个文件
            output_folder = filedialog.askdirectory()
            if not output_folder:
                messagebox.showinfo("信息", "您没有选择文件夹，保存操作已取消。")
                return
            try:
                for key, value in catagory_df.items():
                    output_file = os.path.join(output_folder, f"{key}.xlsx")
                    value.to_excel(output_file, sheet_name=key, index=False)
                messagebox.showinfo("完成", f"拆分文件已保存至: {output_folder}")
            except Exception as e:
                messagebox.showerror("保存错误", str(e))
        elif mode.get() == 2: # 保存为多个页签
            output_file = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
            if not output_file:
                messagebox.showinfo("信息", "您没有选择文件，保存操作已取消。")
                return
            try:
                with pd.ExcelWriter(output_file) as writer:
                    for key, value in catagory_df.items():
                        value.to_excel(writer, sheet_name=key, index=False)
                messagebox.showinfo("完成", f"拆分文件已保存为: {output_file}")
            except Exception as e:
                messagebox.showerror("保存错误", str(e))

    tk.Button(split_window, text="保存", command=save_split_file).grid(row=8, column=1, columnspan=2, pady=5, padx=5, sticky="w")

    split_window.mainloop()

# 新窗口：pdf重命名功能
def pdf_rename_window():
    pdf_window = tk.Toplevel(root)
    pdf_window.title("PDF 发票重命名工具")
    
    # 设置窗口大小
    window_width = 450
    window_height = 550

    # 获取主窗口的位置和大小
    main_x = root.winfo_x()
    main_y = root.winfo_y()
    main_width = root.winfo_width()
    main_height = root.winfo_height()

    # 计算 Toplevel 窗口的中心位置
    toplevel_x = main_x + main_width
    toplevel_y = main_y + main_height

    # 设置 Toplevel 窗口的几何参数
    pdf_window.geometry(f"{window_width}x{window_height}+{toplevel_x}+{toplevel_y}")

    # 选择文件夹下的所有pdf文件
    def select_folder():
        folder = filedialog.askdirectory()
        if folder:
            pdf_files.extend([os.path.join(folder, file) for file in os.listdir(folder) if file.endswith(".pdf")])
            pdf_count.set(len(pdf_files))
            pdf_list.delete(0, tk.END)
            for file in pdf_files:
                pdf_list.insert(tk.END, os.path.basename(file))
    
    # 变量声明
    pdf_files = []
    pdf_count = tk.IntVar(value=0)
    tk.Button(pdf_window, text="选择文件夹", command=select_folder).grid(row=0, column=0, padx=5, pady=5)
    pdf_list = tk.Listbox(pdf_window, height=8, width=40, font=("Times New Roman", 8), selectmode=tk.SINGLE)
    pdf_list.grid(row=1, column=0, columnspan=2, pady=5)    
    tk.Label(pdf_window, text="文件数:").grid(row=2, column=0, padx=5, pady=5, sticky="e")
    tk.Entry(pdf_window, textvariable=pdf_count, state="readonly", width=10).grid(row=2, column=1, padx=5, pady=5, sticky="w")
    

    # 提供名字的输入
    tk.Label(pdf_window, text="报销人名字:").grid(row=4, column=0, padx=5, pady=5, sticky="e")
    name_entry = tk.Entry(pdf_window)
    name_entry.grid(row=4, column=1, padx=5, pady=5, sticky="w")
    
    # 读取pdf文件内容
    def extract_text_from_pdf(pdf_path):
        with pdfplumber.open(pdf_path) as pdf:
            text = ''
            for page in pdf.pages:
                text += page.extract_text()
        return text
    
    # 重命名pdf文件
    def rename_pdf():
        if not pdf_files:
            messagebox.showwarning("警告", "没有文件！")
            return
        cout = 1
        name = name_entry.get()
        for pdf_file in pdf_files:
            text = extract_text_from_pdf(pdf_file)
            match_number = re.search(r'发票号码[:：]\s*(\d+)', text)
            if not match_number:
                continue
            number = match_number.group(1)
            company_name_pattern = r'名\s*称[:：]\s*([^\s\n]+)'
            company_names = re.findall(company_name_pattern, text)  
            if not company_names:
                continue
            company_map = {
                '广西扬翔集团股份有限公司': '扬翔股份',
                '广西扬翔农牧有限责任公司': '扬翔农牧',
                '广西扬翔猪基因科技有限公司': '猪基因',
                '贵港瑞康饲料有限公司': '贵港瑞康',
                '南宁扬翔农牧有限公司': '南宁扬翔'
            }
            company = ''
            for company_name in company_names:
                if company_name in company_map:
                    company = company_map[company_name]
            amount_pattern = r'[（(]小\s*写[）)]\s*[￥¥]([\d,]+(?:\.\d{2})?)'
            match_amount = re.search(amount_pattern, text)
            if not match_amount:
                continue
            amount = match_amount.group(1)

            # 获取原文件的目录并构建新的文件路径
            dir_name = os.path.dirname(pdf_file)
            output_name_suffix = f'{cout}、{company}_{name}_{number}_{amount}元.pdf'
            output_file = os.path.join(dir_name, output_name_suffix)  # 完整的输出路径
            try:
                os.rename(pdf_file, output_file)  # 使用完整的文件路径
                cout += 1
            except Exception as e:
                messagebox.showerror("重命名错误", str(e))
        # 更新文件列表，重新显示新的文件名
        pdf_files.clear()
        pdf_count.set(0)
        pdf_list.delete(0, tk.END)
        for file in os.listdir(dir_name):
            if file.endswith(".pdf"):
                pdf_files.append(os.path.join(dir_name, file))
                pdf_list.insert(tk.END, os.path.basename(file))
        pdf_count.set(len(pdf_files))


        messagebox.showinfo("完成", f"已重命名 {cout-1} 个文件！")

    tk.Button(pdf_window, text="开始重命名", command=rename_pdf).grid(row=5, column=1, padx=5, pady=5, sticky="w")

    # 打开文件
    def open_file(event):
        selected_index = pdf_list.curselection()
        if selected_index:
            file_name = pdf_files[selected_index[0]]
            try:
                os.startfile(file_name)
            except Exception as e:
                print(f"无法打开文件: {e}")

    # 绑定双击事件，双击打开文件
    pdf_list.bind("<Double-Button-1>", open_file)

    pdf_window.mainloop()

file_list = []
combined_data = {}
combined_title_rows = {}

# 创建主窗口
root = tk.Tk()
root.title("HR表格处理工具集合")
root.geometry('450x550')  # 设置初始窗口大小

# 创建菜单栏
menu_bar = tk.Menu(root)
root.config(menu=menu_bar)
feature_menu = tk.Menu(menu_bar, tearoff=0)
menu_bar.add_cascade(label="其他功能", menu=feature_menu)
feature_menu.add_command(label="Excel 表格拆分工具", command=excel_split_window)
feature_menu.add_command(label="PDF 发票重命名工具", command=pdf_rename_window)
feature_menu.add_separator()
feature_menu.add_command(label="退出", command=root.quit)
menu_bar.add_cascade(label="使用帮助", command=show_help)

# 文件操作 Frame
file_frame = tk.Frame(root)
file_frame.grid(row=0, column=0, padx=5, pady=5, sticky='ew')

tk.Button(file_frame, text="添加文件夹", command=select_folder).grid(row=0, column=0, padx=5, pady=5)
tk.Button(file_frame, text="添加文件", command=add_file).grid(row=0, column=1, padx=5, pady=5)
tk.Button(file_frame, text="删除文件", command=delete_file).grid(row=0, column=2, padx=5, pady=5)

tk.Label(file_frame, text="文件列表:").grid(row=1, column=0, sticky="w")
file_display = tk.Listbox(file_frame, height=8, width=70, font=("Times New Roman", 8), selectmode=tk.SINGLE)
file_display.grid(row=2, column=0, columnspan=3, pady=5)

file_count_label = tk.StringVar(value=f"共 {file_display.size()} 个文件")
tk.Label(file_frame, textvariable=file_count_label).grid(row=3, column=0, sticky="w", padx=5)

ttk.Separator(root, orient="horizontal").grid(row=1, column=0, columnspan=3, sticky="ew", pady=5)

# 页签选择 Frame
sheet_frame = tk.Frame(root)
sheet_frame.grid(row=2, column=0, padx=5, pady=5, sticky='ew')

tk.Label(sheet_frame, text="选择需合并页签:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
sheet_name_list = []
select_sheet = ttk.Combobox(sheet_frame, values=sheet_name_list, state="readonly")
select_sheet.grid(row=0, column=1, padx=5, pady=5, sticky="w")
# select_sheet.set("需合并的页签")

flag = tk.BooleanVar()
flag.set(True)  # 默认打开
tk.Checkbutton(sheet_frame, text="自动识别表头行(可能出错)", variable=flag,
               command=lambda: entry.config(state="disabled" if flag.get() else "normal")).grid(row=1, column=1, padx=5, pady=5, sticky="w")

tk.Label(sheet_frame, text="手动输入表头行:").grid(row=2, column=0, padx=5, pady=5, sticky="e")
entry = tk.Entry(sheet_frame, validate="key", 
                 validatecommand=(root.register(lambda P: P == "" or P.isdigit()), '%P'))
entry.grid(row=2, column=1, padx=5, pady=5, sticky="w")
entry.config(state="disabled")  # 根据Checkbutton状态禁用

ttk.Separator(root, orient="horizontal").grid(row=3, column=0, columnspan=3, sticky="ew")
# a = tk.Label(sheet_frame, text=f"{select_sheet.get()}:")
# a.grid(row=3, column=0, padx=5, pady=5, sticky="e")
# select_sheet.bind("<<ComboboxSelected>>", lambda event: a.config(text=f"{select_sheet.get()}:")) # 绑定选择box事件

# 保存 Frame
save_frame = tk.Frame(root)
save_frame.grid(row=4, column=0, padx=10, pady=10, sticky='ew')

progress_bar = ttk.Progressbar(save_frame, length=300, mode="determinate")
progress_bar.grid(row=0, column=0, pady=10, columnspan=2)

tk.Button(save_frame, text="开始合并", command=combine_excel_files, bg="lightblue").grid(row=1, column=0, padx=10, pady=5)
tk.Button(save_frame, text="清空汇总结果", command=clear_combined_data, bg="lightyellow").grid(row=1, column=1, padx=10, pady=5)
save_button = tk.Button(save_frame, text="另存为", command=save_combined_file, bg="lightgreen", state=tk.DISABLED)
save_button.grid(row=1, column=2, padx=10, pady=5)

combined_info_label = tk.StringVar(value=f"已汇总的页签:\n无")
tk.Label(root, textvariable=combined_info_label, relief="sunken", anchor="w", padx=5, pady=5).grid(row=5, column=0, columnspan=3, sticky="ew")

# 绑定双击事件，双击打开文件，单个文件。多个文件的话，只打开第一个（暂时选择只能选择单个文件的模式）
file_display.bind("<Double-Button-1>", open_file)

root.mainloop()
