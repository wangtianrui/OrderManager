import pandas as pd

# path = r"E:\androidprogrames\formoney\1800\总表.xlsx"
#
# data = pd.read_excel(path)
# print(data)
#
# columns = data.columns
#

import numpy as np

a = [1, 2, 3]
b = [1, 2, 3]
c = [1, 2, 3, 4, 5]


def is_sub_set(a, b):
    a = set(a)
    b = set(b)
    return a.issubset(b)


