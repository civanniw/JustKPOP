# -*- coding: utf-8 -*-
"""
Created on Sat May 29 16:51:37 2021

@author: Wangzw
"""


import requests
from bs4 import BeautifulSoup
import datetime
import json
import pandas as pd



# input id of Cbar's EVENT pages, find from the url.
indexs = ['3132063', '3132101', '3085108', '3095412', '3132677', '3084528', '3085268', '3084517', '3087545', '3084357', '3084889', '2474404']

record = []
for index in indexs:
    # access Cbar's EVENT pages
    url = 'https://www.ktown4u.cn/eventsub?eve_no='+index+'&biz_no=599'
    header={
            'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/90.0.4430.85 Safari/537.36 Edg/90.0.818.49'
            }     
    r=requests.get(url, headers=header)
    r.raise_for_status()
    
    # get data
    data = BeautifulSoup(r.content,'lxml')

    fanc_name = data.find_all(class_='fanc-name')[0].strong.text
    items = data.find_all(class_='item')
    
    for item in items:
        Name = item.find_all('span')[1].text.replace('\r','').replace('\n','').replace('\t','')
        try:
            No = item.a['href'].split('mst_fanc_goods_no=')[1]
            url = 'https://www.ktown4u.cn/selectFancGoodsTotalSalesList?shopNo=197&goodsNo=&fancGoodsNo=' + str(No)
        except:
            No1 = item.a['href'].split('&')[1].split('=')[1]
            No2 = item.a['href'].split('&')[2].split('=')[1]
            url = 'https://www.ktown4u.cn/selectFancGoodsTotalSalesList?shopNo=197&goodsNo=' + str(No2)+'&fancGoodsNo=' + str(No1)  
        r=requests.get(url, headers=header)
        sales = json.loads(r.content)
        for sale in sales:
            record.append([fanc_name, Name, sale['GOODS_NM'], sale['TOTAL_SALES']])
    
    print(fanc_name)
    

# filtering process 
record = pd.DataFrame(record)

def iden_version(version):
    if 'K版' in version:
        return 'K版'
    elif 'INSIDE Ver.' in version:
        return 'INSIDE Ver.'
    elif 'OUTSIDE Ver.' in version:
        return 'OUTSIDE Ver.'
    elif '定金' in version:
        return '定金'
    elif '不运回' in version:
        return '不运回'
    else:
        return version

# item recording process
record[1] = record[1].apply(lambda x: x.split(' N.Flying')[0])
record[2] = record[2].apply(lambda x: iden_version(x))
record.columns = ['站子','类型','版本','数量']
record_sum = record[record['版本']!= 'K版'].groupby(by = ['站子','类型']).sum()

# write to excel
d = datetime.datetime.now()
with pd.ExcelWriter('YOUR DIR PATH HERE/nflying-'+str(d.month).zfill(2) + str(d.day).zfill(2)+'.xlsx') as writer:  # doctest: +SKIP
    record_sum.to_excel(writer, sheet_name='Sum')
    record.to_excel(writer, sheet_name='All')     
