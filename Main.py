# -*- coding: utf-8 -*-
import pandas as pd
import numpy as np
import tkinter
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox

global_data = {}

# 字体
message_font = ("黑体", 13)


def get_cost(closing_cost, goods_cost, packing_cost, freight, platform=0):
    """
    :param closing_cost:
    :param goods_cost:
    :param packing_cost:
    :param freight:
    :param platform:
    :return:
    """
    return goods_cost + packing_cost + freight + closing_cost * (1 - platform)


def import_button():
    """
    导入总表点击事件
    :return:
    """
    filename = filedialog.askopenfilename()
    if str(filename).endswith(".xlsx"):
        temp = pd.read_excel(filename)
        if temp.columns.size == 18:
            global_data["总表dataframe"] = temp
            template_toplevel()
        else:
            message = "总表应有18列，您选择的表格有" + str(temp.columns.size) + "列"
            messagebox.askokcancel("操作错误", message)
    else:
        messagebox.askokcancel("操作错误", "请选择表格文件！")


def template_toplevel():
    # 模板子窗格
    template_window = tkinter.Toplevel()
    template_window.title("编辑与选择模板")
    template_window.geometry("1200x900+350+50")
    template_window.resizable(0, 0)

    # tab
    notebook = ttk.Notebook(template_window)
    # 包装模板
    packing_frame = tkinter.Frame(notebook, width=1170, height=900)
    tkinter.Label(packing_frame, text="已有包装模板", width=130, height=3, font=message_font).place(x=0, y=0)
    # 包装表格
    packing_list_frame = tkinter.Frame(packing_frame, width=1170, height=750, bg="#BDBDBD")
    packing_form = ttk.Treeview(packing_list_frame, show="headings", height=36)
    packing_scroll = tkinter.Scrollbar(packing_list_frame)
    packing_columns = [
        "品种", "重量范围", "保温箱数量", "保温袋数量", "冰袋数量", "干冰数量", "防水袋数量", "纸箱数量", "包装总价"
    ]
    packing_form["columns"] = packing_columns
    for i in range(len(packing_columns)):
        packing_form.column(packing_columns[i], width=int(1155 / len(packing_columns)), anchor="center")
        packing_form.heading(packing_columns[i], text=packing_columns[i])
    for i in range(50):
        packing_form.insert("", i, text="line", values=("虾", "50-80", "1", i + 2, i + 3, i - 1, i * 2, i, i * 100))

    packing_form.pack(side=tkinter.LEFT, fill=tkinter.Y)
    packing_scroll.pack(side=tkinter.RIGHT, fill=tkinter.Y)
    packing_scroll.config(command=packing_form.yview)
    packing_form.config(yscrollcommand=packing_scroll.set)
    packing_list_frame.place(y=50)
    # 包装页面按钮
    tkinter.Button(packing_frame, text="  导入新包装模板  ", font=message_font).place(relx=0.05, rely=0.91)
    tkinter.Button(packing_frame, text="  确定模板开始导入  ", font=message_font).place(relx=0.75, rely=0.91)
    # 快递模板
    express_frame = tkinter.Frame(notebook, width=1170, height=800, bg="blue")
    notebook.add(packing_frame, text="       包装模板        ")
    notebook.add(express_frame, text="         运费模板       ")
    notebook.place(x=10, y=5)


# 主窗口
window = tkinter.Tk()

whole_width = window.winfo_screenwidth()
whole_height = window.winfo_screenheight()
whole_width = whole_width - 50
whole_height = whole_height - 100
window.title("利润计算系统")
window.resizable(0, 0)  # 阻止Python GUI的大小调整
size = str(whole_width) + "x" + str(whole_height) + "+0+0"
window.geometry(size)

# 菜单栏
menubar = tkinter.Menu(window)
window.config(menu=menubar)
menu1 = tkinter.Menu(menubar, tearoff=False)
for item in ['python', 'c', 'java', 'c++', 'c#', 'php', 'B', '退出']:
    if item == "退出":
        menu1.add_separator()
        menu1.add_command(label=item, command=window.quit)
    else:
        menu1.add_command(label=item)
menubar.add_cascade(label='语言', menu=menu1)


def showMenu(event):
    menubar.post(event.x_root, event.y_root)


window.bind("<Button-3>", showMenu)

# 主界面功能按钮

# 导入总表按钮
button_import_whole = tkinter.Button(window, text="    导入总表查看利润    ", font=message_font, command=import_button)
button_import_whole.place(x=10, y=5)

# 主界面搜索按钮
button_search = tkinter.Button(window, text="  按关键字搜索订单  ", font=message_font)
button_search.place(x=whole_width - 500, y=5)

# 排序
sort_title = tkinter.Label(window, text=" 按列排序 ", font=message_font)
sort_title.place(x=whole_width - 300, y=7)
sort_choices = tkinter.StringVar()
sort_com = ttk.Combobox(window, textvariable=sort_choices)
sort_com.place(x=whole_width - 200, y=7)
# 排序选项
sort_com["value"] = ("测试1", "测试2", "测试3")

# 表格容器
form_context = tkinter.Frame(window, width=whole_width - 20, height=whole_height - 200, bg="#BDBDBD")

# 表格
whole_tree = ttk.Treeview(form_context, show="headings", height=40)
# 滚动条
whole_scroll = tkinter.Scrollbar(form_context)
# 填充表数据
whole_columns = (
    "订单号", "店铺", "发货仓库", "食材", "量",
    "保温箱数量", "保温袋数量", "冰袋数量", "干冰数量", "防水袋数量",
    "纸箱数量", "包装总价", "食材总成本", "运费", "平台扣点", "纯利润"
)
whole_tree["columns"] = whole_columns
for i in range(len(whole_columns)):
    whole_tree.column(whole_columns[i], width=int((whole_width - 50) / len(whole_columns)), anchor="center")
    whole_tree.heading(whole_columns[i], text=whole_columns[i])
for i in range(50):
    whole_tree.insert("", i, text="line", values=(i, "店三大铺", "萨达", "润体乳", "123", "3543",
                                                  "1", "2", "3", "4", "5",
                                                  "6", "234234", "234", "23", "234"))
whole_tree.pack(side=tkinter.LEFT, fill=tkinter.Y)
whole_scroll.pack(side=tkinter.RIGHT, fill=tkinter.Y)
whole_scroll.config(command=whole_tree.yview)
whole_tree.config(yscrollcommand=whole_scroll.set)
form_context.place(x=10, y=40)

# 底部界面
export_excel_button = tkinter.Button(window, text="     导出表格     ", font=message_font)
export_excel_button.place(x=10, rely=0.9)

import_shunfeng = tkinter.Button(window, text="    导入顺丰账单与上总表进行匹配    ", font=message_font)
import_shunfeng.place(relx=0.15, rely=0.9)

whole_profit = tkinter.Label(window, font=("黑体", 20), fg="red")
whole_profit.place(relx=0.8, rely=0.9)
profit_str = "总利润为："
whole_profit["text"] = profit_str + "2165"

window.mainloop()
