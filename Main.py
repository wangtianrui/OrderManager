# -*- coding: utf-8 -*-
import pandas as pd
import tkinter
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox
import pickle
import numpy as np
import os

global_data = {}
"""
总表dataframe
食材信息
地区信息
"""
packing_models = []
express_models = {}
if not os.path.exists(r"./packing_file.txt"):
    writer_packing_file = open("./packing_file.txt", "wb")
    pickle.dump(packing_models, writer_packing_file, -1)
    writer_express_file = open("./express_file.txt", "wb")
    pickle.dump(express_models, writer_express_file, -1)
else:
    reader_packing_file = open("./packing_file.txt", "rb")
    reader_express_file = open("./express_file.txt", "rb")
    packing_models = pickle.load(reader_packing_file)
    express_models = pickle.load(reader_express_file)


def update():
    """
    更新两个模型
    :return:
    """
    writer_packing_file = open("./packing_file.txt", "wb")
    pickle.dump(packing_models, writer_packing_file, -1)
    writer_express_file = open("./express_file.txt", "wb")
    pickle.dump(express_models, writer_express_file, -1)


def get_foodinfor():
    """
    获取食材信息
    :return:
    """
    data = global_data["总表dataframe"]
    drop_list = ['g', 'k', '斤', '克', '半', '一', '二', 0, 1, 2, 3, 4, 5, 6, 7, 8, 9]
    temp = data[["货品名称", "下单数量", "预估重量"]]

    def drop_char(x):
        for i in drop_list:
            x = x.replace(str(i), '')
        return x

    temp["货品名称"] = temp["货品名称"].apply(lambda x: drop_char(x))
    global_data["食材信息"] = temp
    global_data["食材名"] = tuple(temp["货品名称"].unique())
    print(global_data["食材名"])


def get_addressinfor():
    """
    获取地区信息
    :return:
    """
    data = global_data["总表dataframe"]
    temp = data["收货地区"]
    print(temp)
    global_data["地区信息"] = temp


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
            get_addressinfor()
            get_foodinfor()
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
        "品种", "重量范围", "保温箱重量", "保温袋重量", "冰袋重量", "干冰重量", "防水袋重量", "纸箱重量", "包装总价"
    ]
    packing_form["columns"] = packing_columns
    for i in range(len(packing_columns)):
        packing_form.column(packing_columns[i], width=int(1155 / len(packing_columns)), anchor="center")
        packing_form.heading(packing_columns[i], text=packing_columns[i])
    for item in packing_models:
        packing_form.insert("", "end", values=item)
    packing_form.pack(side=tkinter.LEFT, fill=tkinter.Y)
    packing_scroll.pack(side=tkinter.RIGHT, fill=tkinter.Y)
    packing_scroll.config(command=packing_form.yview)
    packing_form.config(yscrollcommand=packing_scroll.set)
    packing_list_frame.place(y=50)

    def deleteitem(event):
        """
        删除包装模板
        :param event:
        :return:
        """
        if tkinter.messagebox.askokcancel("提醒", "确定要删除该模板吗？"):
            item = packing_form.selection()

            value = packing_form.item(item)["values"]
            for index in range(len(value)):
                value[index] = str(value[index])
            value[-1] = float(value[-1])
            print(packing_models)
            packing_models.remove(value)
            packing_form.delete(item)
            update()
            tkinter.messagebox.showinfo('提醒', '删除成功')

    packing_form.bind("<Double-Button-1>", deleteitem)
    # 包装页面按钮
    tkinter.Button(packing_frame, text="  导入新包装模板  ", font=message_font,
                   command=lambda: add_packing_template(packing_form)).place(relx=0.05,
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
    for name in express_models:
        print(name)
        express_form.insert("", "end", values=(name))
    express_form.pack(side=tkinter.LEFT, fill=tkinter.Y)
    express_scroll.pack(side=tkinter.RIGHT, fill=tkinter.Y)
    express_scroll.config(command=express_form.yview)
    express_form.config(yscrollcommand=express_scroll.set)
    express_list_frame.place(y=50)

    def open_detail_express(event):
        """
        点击运费公司名
        :param event:
        :return:
        """
        name = express_form.item(express_form.selection(), "values")

        if str(name).find("区间式") != -1:
            section_express(name, express_form, express_form.selection())
        else:
            incremental_express(name, express_form, express_form.selection())

    express_form.bind("<Double-Button-1>", open_detail_express)

    # 运费页面按钮
    tkinter.Button(express_frame, text="  导入新运费模板  ", font=message_font,
                   command=lambda: add_express_template(express_form)).place(relx=0.05,
                                                                             rely=0.91)
    tkinter.Button(express_frame, text="  确定模板开始导入  ", font=message_font).place(relx=0.75, rely=0.91)


def add_packing_template(packing_form):
    inputer = []
    """
    添加包装模版
    :return:
    """

    add_pacing_window = tkinter.Toplevel()
    add_pacing_window.title("导入新的包装模板")
    add_pacing_window.geometry("350x600+700+150")
    add_pacing_window.resizable(0, 0)
    add_pacing_window.attributes("-topmost", 1)
    add_pacing_window.wm_attributes("-topmost", 1)
    input_x_1 = 0.4
    input_x_2 = 0.66

    # 品种
    y_local = 0.07
    tkinter.Label(add_pacing_window, text="品种：", font=message_font).place(relx=0.1, rely=y_local)
    kind_cv = tkinter.StringVar()
    kind_com = ttk.Combobox(add_pacing_window, textvariable=kind_cv)
    kind_com.place(relx=input_x_1, rely=y_local)
    kind_com["value"] = global_data["食材名"]

    # 重量范围
    y_local = y_local + 0.1
    tkinter.Label(add_pacing_window, text="重量范围：", font=message_font).place(relx=0.1, rely=y_local)
    weight_entry_low = tkinter.Entry(add_pacing_window, width=5)
    weight_entry_low.place(relx=input_x_1, rely=y_local)
    tkinter.Label(add_pacing_window, font=message_font, text="kg ~").place(relx=input_x_1 + 0.12, rely=y_local)
    weight_entry_high = tkinter.Entry(add_pacing_window, width=5)
    weight_entry_high.place(relx=input_x_2, rely=y_local)
    tkinter.Label(add_pacing_window, font=message_font, text="kg").place(relx=0.8, rely=y_local)

    # 保温箱
    y_local = y_local + 0.1
    tkinter.Label(add_pacing_window, text="保温箱重量：", font=message_font).place(relx=0.1, rely=y_local)
    box_count = tkinter.Entry(add_pacing_window, width=5)
    box_count.place(relx=input_x_1, rely=y_local)
    tkinter.Label(add_pacing_window, font=message_font, text="kg").place(relx=0.52, rely=y_local)

    box_cost = tkinter.Entry(add_pacing_window, width=5)
    box_cost.place(relx=input_x_2, rely=y_local)
    tkinter.Label(add_pacing_window, font=message_font, text="元").place(relx=0.8, rely=y_local)

    # 保温袋
    y_local = y_local + 0.1
    tkinter.Label(add_pacing_window, text="保温袋重量：", font=message_font).place(relx=0.1, rely=y_local)
    bag_count = tkinter.Entry(add_pacing_window, width=5)
    bag_count.place(relx=input_x_1, rely=y_local)
    tkinter.Label(add_pacing_window, font=message_font, text="kg").place(relx=0.52, rely=y_local)
    bag_cost = tkinter.Entry(add_pacing_window, width=5)
    bag_cost.place(relx=input_x_2, rely=y_local)
    tkinter.Label(add_pacing_window, font=message_font, text="元").place(relx=0.8, rely=y_local)

    # 冰袋
    y_local = y_local + 0.1
    tkinter.Label(add_pacing_window, text="冰袋重量：", font=message_font).place(relx=0.1, rely=y_local)
    icebag_count = tkinter.Entry(add_pacing_window, width=5)
    icebag_count.place(relx=input_x_1, rely=y_local)
    tkinter.Label(add_pacing_window, font=message_font, text="kg").place(relx=0.52, rely=y_local)
    icebag_cost = tkinter.Entry(add_pacing_window, width=5)
    icebag_cost.place(relx=input_x_2, rely=y_local)
    tkinter.Label(add_pacing_window, font=message_font, text="元").place(relx=0.8, rely=y_local)

    # 干冰
    y_local = y_local + 0.1
    tkinter.Label(add_pacing_window, text="干冰重量：", font=message_font).place(relx=0.1, rely=y_local)
    ice_count = tkinter.Entry(add_pacing_window, width=5)
    ice_count.place(relx=input_x_1, rely=y_local)
    tkinter.Label(add_pacing_window, font=message_font, text="kg").place(relx=0.52, rely=y_local)
    ice_cost = tkinter.Entry(add_pacing_window, width=5)
    ice_cost.place(relx=input_x_2, rely=y_local)
    tkinter.Label(add_pacing_window, font=message_font, text="元").place(relx=0.8, rely=y_local)

    # 防水袋
    y_local = y_local + 0.1
    tkinter.Label(add_pacing_window, text="防水袋重量：", font=message_font).place(relx=0.1, rely=y_local)
    waterproof_count = tkinter.Entry(add_pacing_window, width=5)
    waterproof_count.place(relx=input_x_1, rely=y_local)
    tkinter.Label(add_pacing_window, font=message_font, text="kg").place(relx=0.52, rely=y_local)
    waterproof_cost = tkinter.Entry(add_pacing_window, width=5)
    waterproof_cost.place(relx=input_x_2, rely=y_local)
    tkinter.Label(add_pacing_window, font=message_font, text="元").place(relx=0.8, rely=y_local)

    # 纸箱重量
    y_local = y_local + 0.1
    tkinter.Label(add_pacing_window, text="纸箱重量：", font=message_font).place(relx=0.1, rely=y_local)
    paper_box_count = tkinter.Entry(add_pacing_window, width=5)
    paper_box_count.place(relx=input_x_1, rely=y_local)
    tkinter.Label(add_pacing_window, font=message_font, text="kg").place(relx=0.52, rely=y_local)
    paper_box_cost = tkinter.Entry(add_pacing_window, width=5)
    paper_box_cost.place(relx=input_x_2, rely=y_local)
    tkinter.Label(add_pacing_window, font=message_font, text="元").place(relx=0.8, rely=y_local)

    # 总价显示
    y_local = y_local + 0.05
    str = tkinter.StringVar()
    str.set("包装总价:%2.f元" % (0.00))

    tkinter.Label(add_pacing_window, textvariable=str, font=("黑体", 16), fg="red").place(relx=0.1, rely=y_local)

    def get_cost():
        whole_cost = 0
        whole_cost += float(box_cost.get())
        whole_cost += float(bag_cost.get())
        whole_cost += float(icebag_cost.get())
        whole_cost += float(ice_cost.get())
        whole_cost += float(waterproof_cost.get())
        whole_cost += float(paper_box_cost.get())
        str.set("包装总价:%.2f元" % (whole_cost))

    def add():
        inputer.append(kind_com.get())
        inputer.append(weight_entry_low.get() + "~" + weight_entry_high.get() + "kg")
        inputer.append(box_count.get())
        inputer.append(bag_count.get())
        inputer.append(icebag_count.get())
        inputer.append(ice_count.get())
        inputer.append(waterproof_count.get())
        inputer.append(paper_box_count.get())
        cost = 0
        cost += float(box_cost.get())
        cost += float(bag_cost.get())
        cost += float(icebag_cost.get())
        cost += float(ice_cost.get())
        cost += float(waterproof_cost.get())
        cost += float(paper_box_cost.get())
        inputer.append(cost)
        packing_form.insert("", 0, text="end", values=inputer)
        packing_models.append(inputer)
        update()
        add_pacing_window.destroy()

    tkinter.Button(add_pacing_window, text="  确定添加  ", font=message_font, command=add).place(relx=0.6,
                                                                                             rely=y_local + 0.05)
    tkinter.Button(add_pacing_window, text="计算包装总价", font=message_font, command=get_cost).place(relx=0.1,
                                                                                                rely=y_local + 0.05)


def add_express_template(express_form):
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

    def get_express_xslm():
        """
        获取运费模版文件
        :return:
        """
        add_express_window.destroy()
        print(express_iv.get())
        filename = filedialog.askopenfilename()
        if str(filename).endswith(".xlsx"):
            temp = pd.read_excel(filename)
            if temp.columns.size == 3:
                if messagebox.askokcancel("提醒", "是否是增量式？") and express_iv.get() == 1:
                    name = str(filename).split("/")[-1].split(".")[0] + "(增量式)"
                    global_data[name] = temp
                    express_models[name] = {}
                    express_models[name]["columns"] = temp.columns
                    express_models[name]["value"] = np.array(temp).tolist()
                    express_models[name]["type"] = 1
                    express_form.insert("", 0, "end", values=(name))
                    messagebox.askokcancel("提醒", "导入成功")

            elif temp.columns.size > 3:
                if messagebox.askokcancel("提醒", "是否是区间式？") and express_iv.get() == 2:
                    name = str(filename).split("/")[-1].split(".")[0] + "(区间式)"
                    global_data[name] = temp
                    express_models[name] = {}
                    express_models[name]["columns"] = temp.columns
                    express_models[name]["value"] = np.array(temp).tolist()
                    express_models[name]["type"] = 2
                    print(name)
                    express_form.insert("", 0, "end", values=(name))

                    messagebox.askokcancel("提醒", "导入成功")

            else:
                messagebox.askokcancel("操作错误", "模板表格有问题，增量式应为3列，区间式应大于3列，请确认！")
        else:
            messagebox.askokcancel("操作错误", "请选择表格文件！")
        update()

    tkinter.Button(add_express_window, text=" 选择文件开始导入 ", command=get_express_xslm).place(relx=0.28, rely=0.7)


def section_express(name, express_form, choose_item):
    """
    区间式显示详细窗口
    :return:
    """
    name = name[0]
    section_window = tkinter.Toplevel()
    section_window.title(name)
    section_window.geometry("1000x600+300+150")
    section_window.resizable(0, 0)
    section_window.attributes("-topmost", 1)
    section_window.wm_attributes("-topmost", 1)
    # 运费表格
    section_list_frame = tkinter.Frame(section_window, width=1000, height=500, bg="#BDBDBD")
    section_form = ttk.Treeview(section_list_frame, show="headings", height=25)
    section_scroll = tkinter.Scrollbar(section_list_frame)
    section_columns = express_models[name]["columns"].tolist()
    print(section_columns)
    section_form["columns"] = section_columns
    for i in range(len(section_columns)):
        if i == 0:
            wid = int(980 / 3 * 2)
        else:
            wid = int(980 / 3 / (len(section_columns) - 1))

        section_form.column(str(section_columns[i]), width=wid, anchor="center")
        section_form.heading(section_columns[i], text=section_columns[i])
    for item in express_models[name]["value"]:
        section_form.insert("", 0, text="end", values=item)

    section_form.pack(side=tkinter.LEFT, fill=tkinter.Y)
    section_scroll.pack(side=tkinter.RIGHT, fill=tkinter.Y)
    section_scroll.config(command=section_form.yview)
    section_form.config(yscrollcommand=section_scroll.set)
    section_list_frame.pack()

    def delete_item():
        express_models.pop(name)
        express_form.delete(choose_item)
        update()
        section_window.destroy()

    tkinter.Button(section_window, text="删除该模版", font=message_font, command=delete_item).pack()


def incremental_express(name, express_form, choose_item):
    """
    增量式显示详细窗口
    :return:
    """
    name = name[0]
    incremental_window = tkinter.Toplevel()
    incremental_window.title(name)
    incremental_window.geometry("1000x600+300+150")
    incremental_window.resizable(0, 0)
    incremental_window.attributes("-topmost", 1)
    incremental_window.wm_attributes("-topmost", 1)
    # 运费表格
    incremental_list_frame = tkinter.Frame(incremental_window, width=1000, height=500, bg="#BDBDBD")
    incremental_form = ttk.Treeview(incremental_list_frame, show="headings", height=25)
    incremental_scroll = tkinter.Scrollbar(incremental_list_frame)
    incremental_columns = express_models[name]["columns"]
    incremental_form["columns"] = incremental_columns.tolist()
    for i in range(len(incremental_columns)):
        if i == 0:
            wid = int(980 / 3 * 2)
        else:
            wid = int(980 / 3 / (len(incremental_columns) - 1))
        incremental_form.column(incremental_columns[i], width=wid, anchor="center")
        incremental_form.heading(incremental_columns[i], text=incremental_columns[i])
    for item in express_models[name]["value"]:
        incremental_form.insert("", 0, text="end", values=item)

    incremental_form.pack(side=tkinter.LEFT, fill=tkinter.Y)
    incremental_scroll.pack(side=tkinter.RIGHT, fill=tkinter.Y)
    incremental_scroll.config(command=incremental_form.yview)
    incremental_form.config(yscrollcommand=incremental_scroll.set)
    incremental_list_frame.pack()

    def delete_item():
        express_models.pop(name)
        express_form.delete(choose_item)
        update()
        incremental_window.destroy()

    tkinter.Button(incremental_window, text="删除该模版", font=message_font, command=delete_item).pack()


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
    "保温箱重量", "保温袋重量", "冰袋重量", "干冰重量", "防水袋重量",
    "纸箱重量", "包装总价", "食材总成本", "运费", "平台扣点", "纯利润"
)
whole_tree["columns"] = whole_columns
for i in range(len(whole_columns)):
    whole_tree.column(whole_columns[i], width=int((whole_width - 50) / len(whole_columns)), anchor="center")
    whole_tree.heading(whole_columns[i], text=whole_columns[i])
for i in range(50):
    whole_tree.insert("", i, text="end", values=(i, "店三大铺", "萨达", "润体乳", "123", "3543",
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
