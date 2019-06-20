import pandas as pd

path = r"E:\androidprogrames\formoney\1800\总表.xlsx"

data = pd.read_excel(path)
print(data)

columns = data.columns


drop_list = ['g', 'k', '斤', '克', '半', '一', '二', 0, 1, 2, 3, 4, 5, 6, 7, 8, 9]

temp = data[["货品名称", "下单数量", "预估重量"]]


def drop_char(x):
    for i in drop_list:
        x = x.replace(str(i), '')
    return x


temp["货品名称"] = temp["货品名称"].apply(lambda x: drop_char(x))
print(temp)
