# -*- coding: utf-8 -*-
# author: Kindle Hsieh time: 2021/3/6

"""
shopee競賽，相同的使用者找出來，並將id組合起來，算出他的Contacts個數。
資料是 contacts.json
"""

import pandas as pd
import numba as nb
from numba.typed import List
from tqdm import tqdm
# 讀入大會的檔案。
df = pd.read_json('/Users/kindlemac/PycharmProjects/shopee/[For Participants] Multi-Channel Contacts Problem (1)/contacts.json')

# df.head()
# len(df)

# 資料前處理。
# 將相同的資料組合起來，並將空白的資料剔除。
email = df[['Email', 'Id']]
email = email[email['Email'] != '']
email_group = email.groupby('Email')['Id'].apply(list)

phone = df[['Phone', 'Id']]
phone = phone[phone['Phone'] != '']
phone_group = phone.groupby('Phone')['Id'].apply(list)

orderid = df[['OrderId', 'Id']]
orderid = orderid[orderid['OrderId'] != '']
orderid_group = orderid.groupby('OrderId')['Id'].apply(list)

# 最核心的操作。
# 將擁有相同資料的組合起來。 利用資料表的特性，一個並上一個，最後只要整理id的合成，就可以避免做過多的迴圈操作。
df = df.merge(email_group, left_on='Email', right_on='Email', how='left', suffixes=('', '_x'))
df = df.merge(phone_group, left_on='Phone', right_on='Phone', how='left', suffixes=('', '_y'))
df = df.merge(orderid_group, left_on='OrderId', right_on='OrderId', how='left', suffixes=('', '_z'))

# 將空值的部分放上空字串。
df.fillna('', inplace=True)


# 將id 合成，並且由小排到大，且不重複。
def id_concat_sort(array):
    i = array['Id_x']
    j = array['Id_y']
    k = array['Id_z']
    return sorted(list(set(i) | set(j) | set(k)))
df['ticket'] = df.apply(id_concat_sort, axis=1)

# 處理間接關係。
tqdm.pandas()
@nb.jit(nopython=True)
def _ggg(j, lists):
    gh = []
    for e in lists:
        if j in e:
            gh.append(e)
    return gh

def ggg(array):
    global nb_lists
    hhh = []
    for j in array:
        # print(j)
        hhh.extend(_ggg(j, nb_lists))
    return sorted(list(set([h for hh in hhh for h in hh])))


lists = df.ticket.to_list()
nb_lists = List()
for l in lists:
    nb_lists.append(l)
df['ticket_'] = df['ticket'].apply(ggg)


df['ticket_'] = df['ticket_'].apply(lambda x: '-'.join([str(xx) for xx in x]))  # 因為要是str才能合成字串。


# 因為同group的會有相同的tickets，且之前有保留了contacts個數，所以可以利用groupby sum算出總數。
Contacts_ = df[['Contacts', 'ticket_']].groupby('ticket_')['Contacts'].sum()
df = df.merge(Contacts_, how='left', left_on='ticket_', right_on="ticket_", suffixes=('', '_y'))
df['ticket_trace/contact'] = df.apply(lambda x: x['ticket_'] + ", " + str(x['Contacts_y']), axis=1)

# 組合大會要的格式。
submission = df[['Id', 'ticket_trace/contact']]
submission.columns = ['ticket_id', 'ticket_trace/contact']

submission.to_csv("#1 - 不吃蝦子倒吐蝦皮 - open.csv", index=False)




