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
    """
    选择模版
    :return:
    """
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
    tkinter.Button(packing_frame, text="  导入新包装模板  ", font=message_font, command=add_packing_template).place(relx=0.05,
                                                                                                             rely=0.91)
    tkinter.Button(packing_frame, text="  确定模板开始导入  ", font=message_font).place(relx=0.75, rely=0.91)
    # 快递模板
    express_frame = tkinter.Frame(notebook, width=1170, height=800)
    notebook.add(packing_frame, text="       包装模板        ")
    notebook.add(express_frame, text="         运费模板       ")
    notebook.place(x=10, y=5)
    # 运费tab
    tkinter.Label(express_frame, text="已有快递公司", width=130, height=3, font=message_font).place(x=0, y=0)
    # 运费表格
    express_list_frame = tkinter.Frame(express_frame, width=1170, height=750, bg="#BDBDBD")
    express_form = ttk.Treeview(express_list_frame, show="headings", height=36)
    express_scroll = tkinter.Scrollbar(express_list_frame)
    express_columns = [
        "公司名字"
    ]
    express_form["columns"] = express_columns
    for i in range(len(express_columns)):
        express_form.column(express_columns[i], width=1155, anchor="center")
        express_form.heading(express_columns[i], text=express_columns[i])
    express_form.insert("", i, text="line", values=("顺丰标快(增量式)"))
    express_form.insert("", i, text="line", values=("顺丰特惠(区间式)"))
    express_form.pack(side=tkinter.LEFT, fill=tkinter.Y)
    express_scroll.pack(side=tkinter.RIGHT, fill=tkinter.Y)
    express_scroll.config(command=express_form.yview)
    express_form.config(yscrollcommand=express_scroll.set)
    express_list_frame.place(y=50)

    def open_section_express(event):
        """
        点击运费公司名
        :param event:
        :return:
        """
        section_express(express_form.item(express_form.selection(), "values"))

    express_form.bind("<Double-Button-1>", open_section_express)

    # 运费页面按钮
    tkinter.Button(express_frame, text="  导入新运费模板  ", font=message_font, command=add_express_template).place(relx=0.05,
                                                                                                             rely=0.91)
    tkinter.Button(express_frame, text="  确定模板开始导入  ", font=message_font).place(relx=0.75, rely=0.91)


def add_packing_template():
    """
    添加包装模版
    :return:
    """
    add_pacing_window = tkinter.Toplevel()
    add_pacing_window.title("导入新的包装模板")
    add_pacing_window.geometry("350x600+500+150")
    add_pacing_window.resizable(0, 0)
    add_pacing_window.attributes("-topmost", 1)
    add_pacing_window.wm_attributes("-topmost", 1)

    # 品种
    y_local = 0.07
    tkinter.Label(add_pacing_window, text="品种：", font=message_font).place(relx=0.1, rely=y_local)
    kind_cv = tkinter.StringVar()
    kind_com = ttk.Combobox(add_pacing_window, textvariable=kind_cv)
    kind_com.place(relx=0.3, rely=y_local)
    kind_com["value"] = ("虾", "鱼", "蟹")

    # 重量范围
    y_local = y_local + 0.1
    tkinter.Label(add_pacing_window, text="重量范围：", font=message_font).place(relx=0.1, rely=y_local)
    weight_entry_low = tkinter.Entry(add_pacing_window, width=5)
    weight_entry_low.place(relx=0.33, rely=y_local)
    tkinter.Label(add_pacing_window, font=message_font, text="kg ~").place(relx=0.44, rely=y_local)
    weight_entry_high = tkinter.Entry(add_pacing_window, width=5)
    weight_entry_high.place(relx=0.55, rely=y_local)
    tkinter.Label(add_pacing_window, font=message_font, text="kg").place(relx=0.65, rely=y_local)

    # 保温箱
    y_local = y_local + 0.1
    tkinter.Label(add_pacing_window, text="保温箱数量：", font=message_font).place(relx=0.1, rely=y_local)
    box_cv = tkinter.StringVar()
    box_com = ttk.Combobox(add_pacing_window, textvariable=box_cv, width=7)
    box_com.place(relx=0.4, rely=y_local)
    box_com["value"] = (1, 2, 3, 4)
    tkinter.Label(add_pacing_window, font=message_font, text="个").place(relx=0.6, rely=y_local)

    # 保温袋
    y_local = y_local + 0.1
    tkinter.Label(add_pacing_window, text="保温袋数量：", font=message_font).place(relx=0.1, rely=y_local)
    bag_cv = tkinter.StringVar()
    bag_com = ttk.Combobox(add_pacing_window, textvariable=bag_cv, width=7)
    bag_com.place(relx=0.4, rely=y_local)
    bag_com["value"] = (1, 2, 3, 4)
    tkinter.Label(add_pacing_window, font=message_font, text="个").place(relx=0.6, rely=y_local)

    # 干冰
    y_local = y_local + 0.1
    tkinter.Label(add_pacing_window, text="干冰数量：", font=message_font).place(relx=0.1, rely=y_local)
    ice_count = tkinter.Entry(add_pacing_window, width=5)
    ice_count.place(relx=0.37, rely=y_local)
    tkinter.Label(add_pacing_window, font=message_font, text="kg").place(relx=0.57, rely=y_local)

    # 防水袋
    y_local = y_local + 0.1
    tkinter.Label(add_pacing_window, text="防水袋数量：", font=message_font).place(relx=0.1, rely=y_local)
    waterproof_cv = tkinter.StringVar()
    waterproof_com = ttk.Combobox(add_pacing_window, textvariable=waterproof_cv, width=7)
    waterproof_com.place(relx=0.4, rely=y_local)
    waterproof_com["value"] = (1, 2, 3, 4)
    tkinter.Label(add_pacing_window, font=message_font, text="个").place(relx=0.6, rely=y_local)

    # 纸箱数量
    y_local = y_local + 0.1
    tkinter.Label(add_pacing_window, text="纸箱数量：", font=message_font).place(relx=0.1, rely=y_local)
    paper_box_cv = tkinter.StringVar()
    paper_box_com = ttk.Combobox(add_pacing_window, textvariable=paper_box_cv, width=7)
    paper_box_com.place(relx=0.4, rely=y_local)
    paper_box_com["value"] = (1, 2, 3, 4)
    tkinter.Label(add_pacing_window, font=message_font, text="个").place(relx=0.6, rely=y_local)

    # 总价显示
    y_local = y_local + 0.1
    str = "包装总价：" + "123"
    tkinter.Label(add_pacing_window, text=str, font=("黑体", 18), fg="red").place(relx=0.1, rely=y_local)
    tkinter.Button(add_pacing_window, text="  确定添加  ", font=message_font).place(relx=0.5, rely=y_local + 0.1)


def add_express_template():
    """
    添加运费模版
    :return:
    """
    add_express_window = tkinter.Toplevel()
    add_express_window.title("添加运费模版")
    add_express_window.geometry("300x200+800+400")
    add_express_window.resizable(0, 0)
    add_express_window.attributes("-toolwindow", 1)
    add_express_window.wm_attributes("-topmost", 1)

    tkinter.Label(add_express_window, text="请选择添加类型", font=message_font).place(relx=0.27, rely=0.1)
    express_iv = tkinter.IntVar()
    tkinter.Radiobutton(add_express_window, text="增量式", value=1, variable=express_iv).place(relx=0.15, rely=0.3)
    tkinter.Radiobutton(add_express_window, text="区间式", value=2, variable=express_iv).place(relx=0.55, rely=0.3)
    tkinter.Button(add_express_window, text=" 选择文件开始导入 ", command=get_express_xslm).place(relx=0.28, rely=0.7)


def get_express_xslm():
    """
    获取运费模版文件
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


def section_express(name):
    """
    区间式显示详细窗口
    :return:
    """
    section_window = tkinter.Toplevel()
    section_window.title(name)
    section_window.geometry("450x600+700+150")
    section_window.resizable(0, 0)
    section_window.attributes("-topmost", 1)
    section_window.wm_attributes("-topmost", 1)
    # 运费表格
    section_list_frame = tkinter.Frame(section_window, width=450, height=500, bg="#BDBDBD")
    section_form = ttk.Treeview(section_list_frame, show="headings", height=15)
    section_scroll = tkinter.Scrollbar(section_list_frame)
    section_columns = [
        "目的省份", "首1kg(元)", "1.5kg-3kg(元)", ">3kg(元)"
    ]
    section_form["columns"] = section_columns
    for i in range(len(section_columns)):
        if i == 0:
            wid = int(450 / 3)
        else:
            wid = int(450 / 3 * 2 / (len(section_columns) - 1))
        section_form.column(section_columns[i], width=wid, anchor="center")
        section_form.heading(section_columns[i], text=section_columns[i])
    section_form.insert("", 0, text="line", values=("四川、山西", 5, 6, 3))
    section_form.insert("", 0, text="line", values=("北京、上海、北京、上海、北京、上海", 7, 8, 9))
    section_form.insert("", 0, text="line", values=("北京、上海、北京、上海、北京、上海、北京、上海、北京、上海、北京、上海", 12, 23, 43))
    section_form.pack(side=tkinter.LEFT, fill=tkinter.Y)
    section_scroll.pack(side=tkinter.RIGHT, fill=tkinter.Y)
    section_scroll.config(command=section_form.yview)
    section_form.config(yscrollcommand=section_scroll.set)
    section_list_frame.pack()


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
