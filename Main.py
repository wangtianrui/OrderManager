# -*- coding: utf-8 -*-
import pandas as pd
import tkinter
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox
import pickle
import numpy as np
import os

global_data = {"扣点比例": 0.0}
"""
总表dataframe
食材信息
地区信息
所有信息
显示的信息
食材重量范围
"""

lack_packing = []
lack_express = []
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
    # print(global_data["食材名"])

    food_weight = data[["货品名称", "预估重量"]]
    food_weight["货品名称"] = food_weight["货品名称"].apply(lambda x: drop_char(x))
    food_weight_min = food_weight.groupby("货品名称")["预估重量"].min()
    food_weight_max = food_weight.groupby("货品名称")["预估重量"].max()
    names = np.array(food_weight_max.index)
    names = names.reshape((len(names), 1))
    food_weight_max_min = pd.concat([food_weight_min, food_weight_max], axis=1)
    food_weight_max_min = np.append(names, np.array(food_weight_max_min), axis=1)
    global_data["食材重量范围"] = food_weight_max_min

    """
    获取所有需要的信息
    :return:
    """

    data = global_data["总表dataframe"]
    whole_data = data[["订单编号", "店铺", "仓库", "货品名称", "预估重量", "订单预估成本", "订单支付金额", "收货地区", "物流公司"]]

    def drop_location_char(x):
        drop_list = ['省', '市', '自治区', '自治州', '壮族', '维吾尔', '回族', ]
        for i in drop_list:
            x = x.replace(str(i), '')
        x = x.split(" ")[:2]
        return x

    whole_data["收货地区"] = whole_data["收货地区"].apply(lambda x: drop_location_char(x))

    def drop_char(x):
        drop_list = ['g', 'k', '斤', '克', '半', '一', '二', 0, 1, 2, 3, 4, 5, 6, 7, 8, 9]
        for i in drop_list:
            x = x.replace(str(i), '')
        return x

    whole_data["货品名称"] = whole_data["货品名称"].apply(lambda x: drop_char(x))
    global_data["所有信息"] = whole_data
    # print(global_data)


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


def import_button(form):
    """
    导入总表点击事件
    :return:
    """
    filename = filedialog.askopenfilename()
    if str(filename).endswith(".xlsx"):
        temp = pd.read_excel(filename)
        if temp.columns.size == 18:
            global_data["总表dataframe"] = temp

            get_foodinfor()
            template_toplevel(form)
        else:
            message = "总表应有18列，您选择的表格有" + str(temp.columns.size) + "列"
            messagebox.askokcancel("操作错误", message)
    else:
        messagebox.askokcancel("操作错误", "请选择表格文件！")


def has_packing(name, low, high):
    """
    判断是否有name产品的包装
    :param name:
    :param low:
    :param high:
    :return:
    """
    low_ok = False
    high_ok = False
    for item in packing_models:
        model_name = item[0]
        if model_name == name:
            section = item[1]
            section = section[:-2]
            section = section.split("~")
            section_low = section[0]
            section_high = section[1]
            if float(low) >= float(section_low):
                low_ok = True
            if float(high) <= float(section_high):
                high_ok = True
    return (low_ok and high_ok)


def get_pack(name, weight):
    """
    查询包装
    :param name:
    :param weight:
    :return:
    """
    for item in packing_models:
        model_name = item[0]
        if model_name == name:
            section = item[1]
            section = section[:-2]
            section = section.split("~")
            section_low = section[0]
            section_high = section[1]
            if weight >= float(section_low) and weight <= float(section_high):
                return item[2:]
    return -1


def platform_point(form):
    """
    计算平台扣点
    :return:
    """
    input_percent_window = tkinter.Toplevel()
    input_percent_window.title("输入平台扣点比例")
    input_percent_window.geometry("300x200+800+200")
    input_percent_window.resizable(0, 0)
    tkinter.Label(input_percent_window, text="平台扣点比例：", font=message_font).place(relx=0.15, rely=0.3)
    editer = tkinter.Entry(input_percent_window, width=5)
    editer.place(relx=0.6, rely=0.3)
    tkinter.Label(input_percent_window, text="%", font=message_font).place(relx=0.7, rely=0.3)

    def get_percent():
        percent = float(editer.get())
        if messagebox.askokcancel("提醒", "确认平台扣点比例为百分之%.2f?" % (percent)):
            global_data["扣点比例"] = percent
            import_whole_data(form)
            input_percent_window.destroy()

    tkinter.Button(input_percent_window, text=" 确定 ", font=message_font, command=get_percent).place(rely=0.7, relx=0.40)


def import_whole_data(form):
    """
    导入总表的最终逻辑
    :return:
    """
    # 食材重量范围
    food_weight_min_max = global_data["食材重量范围"]
    lack_packing.clear()
    for item in food_weight_min_max:
        name = item[0]
        low = item[1]
        hight = item[2]
        if not has_packing(name, low, hight):
            lack_packing.append(item)
    if len(lack_packing) == 0:
        box = []
        bag = []
        ice_bg = []
        ice = []
        waterproof = []
        paper_box = []
        cost = []
        all_weight = []

        whole_data = global_data["所有信息"]
        packing_list = []
        name_weight = np.array(whole_data[["货品名称", "预估重量"]])
        packing_list.clear()
        for item in name_weight:
            if get_pack(item[0], item[1]) != -1:
                packing_list.append(get_pack(item[0], item[1]))
            else:
                messagebox.askokcancel("警告", "未找到%s%.2f的模板" % (item[0], item[1]))
                break
        for item in packing_list:
            box.append(item[0])
            bag.append(item[1])
            ice_bg.append(item[2])
            ice.append(item[3])
            waterproof.append(item[4])
            paper_box.append(item[5])
            cost.append(item[6])
            sum = 0
            for index in range(7):
                sum += float(item[index])
            all_weight.append(sum)
        box = np.reshape(box, (len(box), 1))
        bag = np.reshape(bag, (len(bag), 1))
        ice_bg = np.reshape(ice_bg, (len(ice_bg), 1))
        ice = np.reshape(ice, (len(ice), 1))
        waterproof = np.reshape(waterproof, (len(waterproof), 1))
        paper_box = np.reshape(paper_box, (len(paper_box), 1))
        cost = np.reshape(cost, (len(cost), 1))
        all_weight = np.reshape(all_weight, (len(all_weight), 1))
        temp = np.append(box, bag, axis=1)
        temp = np.append(temp, ice_bg, axis=1)
        temp = np.append(temp, ice, axis=1)
        temp = np.append(temp, waterproof, axis=1)
        temp = np.append(temp, paper_box, axis=1)
        temp = np.append(temp, cost, axis=1)
        temp = np.append(temp, all_weight, axis=1)

        new_df = pd.DataFrame(columns=["保温箱重量", "保温袋重量", "冰袋重量", "干冰重量", "防水袋重量", "纸箱重量", "包装总价", "包装总重"],
                              data=temp)
        whole_data = whole_data.join(new_df)
        global_data["所有信息"] = whole_data
        # print(packing_list)
    else:
        message = "%s未找到%.1f~%.1f区间的包装模板，请补充!" % (lack_packing[0][0], lack_packing[0][1], lack_packing[0][2])
        messagebox.askokcancel("错误", message)
    # 运费
    whole_data = global_data["所有信息"]
    express_data = whole_data[["收货地区", "物流公司", "包装总重", "预估重量"]]
    all_express = express_data["物流公司"].unique()
    lack_express.clear()
    for item in all_express:
        if item not in express_models.keys():
            lack_express.append(item)
    if len(lack_express) != 0:
        messagebox.askokcancel("错误", "未找到%s的运费模板，请补充!" % (lack_express[0]))
    else:

        # print(np.array(express_data))
        express_cost = []
        for item in np.array(express_data):
            province = item[0][0]
            city = item[0][1]
            company = item[1]
            weight = float(item[2]) + float(item[3])
            express_cost.append(calculate_express_cost(province, city, company, weight))
        express_cost = np.reshape(express_cost, (len(express_cost), 1))
        express_cost = pd.DataFrame(columns=["运费"], data=express_cost)
        whole_data = whole_data.join(express_cost)
        global_data["所有信息"] = whole_data
        # print(whole_data)
    # 计算平台扣点
    data = global_data["所有信息"]
    pay_data = np.array(data["订单支付金额"]).tolist()
    platform_cost = []
    percent = global_data["扣点比例"]
    for item in pay_data:
        platform_cost.append(round(item * (percent / 100.0), 2))
    platform_cost = np.reshape(platform_cost, (len(platform_cost), 1))
    platform_cost = pd.DataFrame(columns=["平台扣点"], data=platform_cost)
    data = data.join(platform_cost)
    global_data["所有信息"] = data

    profits = []
    profit_data = np.array(data[["订单支付金额", "订单预估成本", "包装总价", "运费", "平台扣点"]], dtype=float)
    for item in profit_data:
        profits.append([round(item[0] - item[1] - item[2] - item[3] - item[4], 2)])
    whole_profit["text"] = (profit_str + str(np.sum(profits)) + "元")
    profits = pd.DataFrame(columns=["利润"], data=profits)
    data = data.join(profits)
    global_data["所有信息"] = data

    show_data = global_data["所有信息"]
    temp = show_data[["订单编号", "店铺", "仓库", "货品名称", "预估重量", "保温箱重量",
                      "保温袋重量", "冰袋重量", "干冰重量", "防水袋重量", "纸箱重量",
                      "包装总价", "订单预估成本", "运费", "平台扣点", "利润"]]
    global_data["last_data"] = temp
    show_data = np.array(temp).tolist()
    for index in range(len(show_data)):
        form.insert("", index, text="end", values=show_data[index])
    sort_com["value"] = tuple(list(temp.columns.values))


def calculate_express_cost(province, city, company, weight):
    """
    计算运费
    :param province:
    :param city:
    :param company:
    :param weight:
    :return:
    """
    cost = 0
    calculate_model = express_models[company]
    section = calculate_model["columns"]
    values = calculate_model["value"]
    type = calculate_model["type"]

    if type == 2:
        # 区间式
        section = get_section(section)
        contain_index = -1
        xuzhong = 0
        for index in range(len(section)):
            if weight >= float(section[index][0]) and weight <= float(section[index][1]):
                contain_index = index
                if contain_index == 2:
                    contain_index = contain_index - 1
                    xuzhong = weight - float(section[index][0])
                    weight = float(section[index][0])
                break
        if contain_index == -1:
            messagebox.askokcancel("错误", "未在%s中找到重量%.2f的收费标准,请确认后重新导表！" % (company, weight))
        else:
            for item in values:
                item[0] = str(item[0])
                # print(item[0], city)
                if item[0].find(city) != -1:
                    cost = float(item[contain_index + 1]) * weight + xuzhong * float(item[contain_index + 2])
                    break
                elif item[0].find(province) != -1:
                    cost = float(item[contain_index + 1]) * weight + xuzhong * float(item[contain_index + 2])
                    break
                else:
                    cost = 0
    if type == 1:
        # 增量式
        xuzhong = 0
        section = get_section(section)
        print(section)
        split = float(section[0][1])
        for item in values:
            if item[0].find(city) != -1:
                if weight > split:
                    xuzhong = weight - split
                    weight = split
                cost = float(item[1]) * weight + xuzhong * float(item[2])
                break
            elif item[0].find(province) != -1:
                cost = float(item[1]) * weight + xuzhong * float(item[2])
                break
            else:
                cost = 0
    return cost


def get_section(index):
    keep = ['1', '2', '3', '4', '5', '6', '7', '8', '9', '0', '.', '-', '>', '<']
    sections = np.array(index[1:]).tolist()
    for i in range(len(sections)):
        for j in sections[i]:
            if j not in keep:
                sections[i] = str(sections[i]).replace(j, '')

    for i in range(len(sections)):
        if len(sections[i]) > 1:
            if sections[i].find("-") != -1:
                sections[i] = sections[i].split("-")
            elif sections[i].find(">") != -1:
                sections[i] = sections[i].split(">")
                sections[i][0] = sections[i][1]
                sections[i][1] = '1000'
            elif sections[i].find("<") != -1:
                sections[i] = sections[i].split("<")
                sections[i][0] = 0
        else:
            sections[i] = [0, sections[i]]
    return sections


def template_toplevel(form):
    """
    选择模版
    :return:
    """

    # 模板子窗格
    template_window = tkinter.Toplevel()
    template_window.title("编辑与选择模板")
    template_window.geometry("1200x900+350+50")
    template_window.resizable(0, 0)
    pack_lack_str = tkinter.StringVar()
    express_lack_str = tkinter.StringVar()

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
            getpacklackinfor(pack_lack_str)
            tkinter.messagebox.showinfo('提醒', '删除成功')

    packing_form.bind("<Double-Button-1>", deleteitem)

    # 包装页面按钮
    tkinter.Button(packing_frame, text="  导入新包装模板  ", font=message_font,
                   command=lambda: add_packing_template(packing_form, pack_lack_str)).place(relx=0.05,
                                                                                            rely=0.91)

    def see_pack_lack():
        """
        查看缺少包装的食材
        :return:
        """
        lack_window = tkinter.Toplevel()
        lack_window.title("缺少模板")
        lack_window.geometry("600x400+800+250")
        lack_window.resizable(0, 0)
        tkinter.Label(lack_window, text="缺少模板的食材以及对应区间", width=40, height=2, font=message_font).place(x=0, y=0)
        # 包装表格
        lack_list_frame = tkinter.Frame(lack_window, width=600, height=400, bg="#BDBDBD")
        lack_form = ttk.Treeview(lack_list_frame, show="headings", height=15)
        lack_scroll = tkinter.Scrollbar(lack_list_frame)
        lack_columns = [
            "食材名", "目前总表里已有的重量区间（最小值与最大值）"
        ]
        lack_form["columns"] = lack_columns
        for i in range(len(lack_columns)):
            lack_form.column(lack_columns[i], width=int(570 / len(lack_columns)), anchor="center")
            lack_form.heading(lack_columns[i], text=lack_columns[i])
        for item in lack_packing:
            temp = [item[0]]
            print(item)
            temp.append("%.2f~%.2fkg" % (item[1], item[2]))
            lack_form.insert("", "end", values=temp)
        lack_form.pack(side=tkinter.LEFT, fill=tkinter.Y)
        lack_scroll.pack(side=tkinter.RIGHT, fill=tkinter.Y)
        lack_scroll.config(command=lack_form.yview)
        lack_form.config(yscrollcommand=lack_scroll.set)
        lack_list_frame.place(y=10)

        # 包装页面按钮
        tkinter.Button(lack_window, text="  导入新包装模板  ", font=message_font,
                       command=lambda: add_packing_template(packing_form, pack_lack_str)).place(relx=0.05,
                                                                                                rely=0.91)

        def back():
            lack_window.destroy()

        tkinter.Button(lack_window, text="  返回  ", font=message_font,
                       command=back).place(relx=0.75,
                                           rely=0.91)

    getpacklackinfor(pack_lack_str)
    if len(lack_packing) != 0:
        tkinter.Button(packing_frame, textvariable=pack_lack_str, font=("黑体", 16), fg="red",
                       command=see_pack_lack).place(relx=0.35, rely=0.91)

    tkinter.Button(packing_frame, text="  确定模板开始导入  ", font=message_font, command=lambda: platform_point(form)).place(
        relx=0.75,
        rely=0.91)
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
        # express_models[name]["type"]
        name = name[0]
        print(name)

        if express_models[name]["type"] == 2:
            section_express(name, express_form, express_form.selection(), express_lack_str)
        else:
            incremental_express(name, express_form, express_form.selection())

    express_form.bind("<Double-Button-1>", open_detail_express)

    # 运费页面按钮
    tkinter.Button(express_frame, text="  导入新运费模板  ", font=message_font,
                   command=lambda: add_express_template(express_form, express_lack_str)).place(relx=0.05,
                                                                                               rely=0.91)

    def see_express_lack():
        """
        查看缺少的快递模板
        :return:
        """
        lack_express_window = tkinter.Toplevel()
        lack_express_window.title("缺少模板")
        lack_express_window.geometry("600x400+800+250")
        lack_express_window.resizable(0, 0)
        tkinter.Label(lack_express_window, text="缺少的货运公司模板", width=40, height=2, font=message_font).place(x=0, y=0)
        # 包装表格
        lack_express_list_frame = tkinter.Frame(lack_express_window, width=600, height=400, bg="#BDBDBD")
        lack_express_form = ttk.Treeview(lack_express_list_frame, show="headings", height=15)
        lack_express_scroll = tkinter.Scrollbar(lack_express_list_frame)
        lack_express_columns = [
            "公司名"
        ]
        lack_express_form["columns"] = lack_express_columns
        for i in range(len(lack_express_columns)):
            lack_express_form.column(lack_express_columns[i], width=int(570 / len(lack_express_columns)),
                                     anchor="center")
            lack_express_form.heading(lack_express_columns[i], text=lack_express_columns[i])
        for item in lack_express:
            lack_express_form.insert("", "end", values=(item))
        lack_express_form.pack(side=tkinter.LEFT, fill=tkinter.Y)
        lack_express_scroll.pack(side=tkinter.RIGHT, fill=tkinter.Y)
        lack_express_scroll.config(command=lack_express_form.yview)
        lack_express_form.config(yscrollcommand=lack_express_scroll.set)
        lack_express_list_frame.place(y=10)

        # 包装页面按钮
        tkinter.Button(lack_express_window, text="  导入新运费模板  ", font=message_font,
                       command=lambda: add_express_template(lack_express_form, express_lack_str)).place(relx=0.05,
                                                                                                        rely=0.91)

        def back():
            lack_express_window.destroy()

        tkinter.Button(lack_express_window, text="  返回  ", font=message_font,
                       command=back).place(relx=0.75,
                                           rely=0.91)

    getexpresslackinfor(express_lack_str)
    if len(lack_express) != 0:
        tkinter.Button(express_frame, textvariable=express_lack_str, font=("黑体", 16), fg="red",
                       command=see_express_lack).place(relx=0.35, rely=0.91)

    tkinter.Button(express_frame, text="  确定模板开始导入  ", font=message_font, command=lambda: platform_point(form)).place(
        relx=0.75,
        rely=0.91)


def getpacklackinfor(lack_str):
    """
    计算缺少多少包装模板
    :param lack_str:
    :return:
    """
    food_weight_min_max = global_data["食材重量范围"]
    lack_packing.clear()
    for item in food_weight_min_max:
        name = item[0]
        low = item[1]
        hight = item[2]
        if not has_packing(name, low, hight):
            lack_packing.append(item)

    lack_str.set("还有%d个食材没找到模板，点我查看" % (len(lack_packing)))


def getexpresslackinfor(lack_str):
    """
    计算缺少多少运费模板
    :param lack_str:
    :return:
    """
    data = global_data["所有信息"]
    all_express = data["物流公司"].unique()
    lack_express.clear()
    for item in all_express:
        if item not in express_models.keys():
            lack_express.append(item)
    print(lack_express)
    lack_str.set("还有%d个快递公司没找到模板，点我查看" % (len(lack_express)))


def add_packing_template(packing_form, lack_str):
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
        getpacklackinfor(lack_str)

    tkinter.Button(add_pacing_window, text="  确定添加  ", font=message_font, command=add).place(relx=0.6,
                                                                                             rely=y_local + 0.05)
    tkinter.Button(add_pacing_window, text="计算包装总价", font=message_font, command=get_cost).place(relx=0.1,
                                                                                                rely=y_local + 0.05)


def add_express_template(express_form, express_lack_str):
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
                    name = str(filename).split("/")[-1].split(".")[0] + ""
                    global_data[name] = temp
                    express_models[name] = {}
                    express_models[name]["columns"] = temp.columns
                    express_models[name]["value"] = np.array(temp).tolist()
                    express_models[name]["type"] = 1
                    express_form.insert("", 0, "end", values=(name))
                    messagebox.askokcancel("提醒", "导入成功")

            elif temp.columns.size > 3:
                if messagebox.askokcancel("提醒", "是否是区间式？") and express_iv.get() == 2:
                    name = str(filename).split("/")[-1].split(".")[0] + ""
                    global_data[name] = temp
                    express_models[name] = {}
                    express_models[name]["columns"] = temp.columns
                    express_models[name]["value"] = np.array(temp).tolist()
                    express_models[name]["type"] = 2
                    print(name)
                    express_form.insert("", 0, "end", values=(name))

                    messagebox.askokcancel("提醒", "导入成功")
                    getexpresslackinfor(express_lack_str)

            else:
                messagebox.askokcancel("操作错误", "模板表格有问题，增量式应为3列，区间式应大于3列，请确认！")
        else:
            messagebox.askokcancel("操作错误", "请选择表格文件！")
        update()

    tkinter.Button(add_express_window, text=" 选择文件开始导入 ", command=get_express_xslm).place(relx=0.28, rely=0.7)


def section_express(name, express_form, choose_item, lack_str):
    """
    区间式显示详细窗口
    :return:
    """

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
        getexpresslackinfor(lack_str)

    tkinter.Button(section_window, text="删除该模版", font=message_font, command=delete_item).pack()


def incremental_express(name, express_form, choose_item):
    """
    增量式显示详细窗口
    :return:
    """
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
button_import_whole = tkinter.Button(window, text="    导入总表查看利润    ", font=message_font,
                                     command=lambda: import_button(whole_tree))
button_import_whole.place(x=10, y=5)

# 主界面搜索按钮
button_search = tkinter.Button(window, text="  按关键字搜索订单  ", font=message_font)
button_search.place(x=whole_width - 500, y=5)


# 排序
def sort_data(*args):
    choose_name = sort_com.get()
    print(choose_name)
    if choose_name != "":
        x = whole_tree.get_children()
        for item in x:
            whole_tree.delete(item)
        global_data["last_data"].sort_values(choose_name, inplace=True)
        temp = global_data["last_data"]
        show_data = np.array(temp).tolist()
        for index in range(len(temp)):
            whole_tree.insert("", index, text="end", values=show_data[index])


sort_title = tkinter.Label(window, text=" 按列排序 ", font=message_font)
sort_title.place(x=whole_width - 300, y=7)
sort_choices = tkinter.StringVar()
sort_com = ttk.Combobox(window, textvariable=sort_choices)
sort_com.place(x=whole_width - 200, y=7)
sort_com.bind("<<ComboboxSelected>>", sort_data)
# 排序选项


# 表格容器
form_context = tkinter.Frame(window, width=whole_width - 20, height=whole_height - 200, bg="#BDBDBD")

# 表格
whole_tree = ttk.Treeview(form_context, show="headings", height=40)
# 滚动条
whole_scroll = tkinter.Scrollbar(form_context)
# 填充表数据
whole_columns = (
    "订单号", "店铺", "发货仓库", "食材", "食材重量",
    "保温箱重量", "保温袋重量", "冰袋重量", "干冰重量", "防水袋重量",
    "纸箱重量", "包装总价", "食材总成本", "运费", "平台扣点", "纯利润"
)
whole_tree["columns"] = whole_columns
for i in range(len(whole_columns)):
    whole_tree.column(whole_columns[i], width=int((whole_width - 50) / len(whole_columns)), anchor="center")
    whole_tree.heading(whole_columns[i], text=whole_columns[i])

whole_tree.pack(side=tkinter.LEFT, fill=tkinter.Y)
whole_scroll.pack(side=tkinter.RIGHT, fill=tkinter.Y)
whole_scroll.config(command=whole_tree.yview)
whole_tree.config(yscrollcommand=whole_scroll.set)
form_context.place(x=10, y=40)


def export():
    print(global_data.keys())
    if "last_data" in global_data.keys():
        path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                            filetypes=[('表格文件', '.xlsx'), ('text files', '.txt')],
                                            initialfile="导出结果")
        # print(path)
        global_data["last_data"].to_excel(path)
    else:
        messagebox.askokcancel("操作错误", "还未导入总表")


# 底部界面
export_excel_button = tkinter.Button(window, text="     导出表格     ", command=export, font=message_font)
export_excel_button.place(x=10, rely=0.9)

import_shunfeng = tkinter.Button(window, text="    导入顺丰账单与上总表进行匹配    ", font=message_font)
import_shunfeng.place(relx=0.15, rely=0.9)

whole_profit = tkinter.Label(window, font=("黑体", 20), fg="red")
whole_profit.place(relx=0.8, rely=0.9)
profit_str = "总利润为："
whole_profit["text"] = profit_str + "0"

window.mainloop()
