import os
from xlutils.copy import copy
import xlrd
import xlwt

def saveexcel(xlsfile,item,colname):
    try:
        if not os.path.exists(xlsfile):
            book=xlwt.Workbook(encoding='utf-8')
            sheet1=book.add_sheet('Sheet 1')
            colindex = 0
            for col in colname:
                sheet1.write(0,colindex,col)
                colindex += 1
            print('ok')
            book.save(xlsfile)
        book=xlrd.open_workbook(xlsfile)
        rsheet=book.sheet_by_index(0)
        rows=rsheet.nrows
        wbook=copy(book)
        #使用get_sheet获取副本要操作的sheet
        wsheet=wbook.get_sheet(0)
        #写入数据参数
        for col in range(len(colname)):
            wsheet.write(rows,col,item[col])
        #保存
        wbook.save(xlsfile)
        flag='T'
    except:
        flag='F'
    return flag
  
  
import requests
from bs4 import BeautifulSoup
import time
import datetime


year = 2021
startweek = 1
endweek = 30
path = 'YOUR DIR PATH HERE'
colname = ["year","week","rank","title","singer","album", "digital"]


starttime=time.time()

for i in range(startweek, endweek):   
    xlsfile= path + 'gaon_'+str(year)+'_week'+str(i)+'.xls'
    link='http://gaonchart.co.kr/main/section/chart/online.gaon?nationGbn=T&serviceGbn=ALL&targetTime='+str(i).zfill(2)+'&hitYear='+str(year)+'&termGbn=week'
    header={
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/61.0.3163.79 Safari/537.36'
    }
    request=requests.get(link,headers=header)
    request.raise_for_status()
    soup = BeautifulSoup(request.content,'lxml')

    #ranking = soup.find_all(class_='ranking')
    subjects=soup.find_all(class_='subject')
    counts=soup.find_all(class_='count')
    
    for s in range(len(subjects)):
        title=subjects[s].find_all('p')[0].text
        singer=subjects[s].find_all('p')[1].text.split('|')[0]
        album=subjects[s].find_all('p')[1].text.split('|')[1]
        count=int(counts[s].find('p').text.replace(',',''))
        song=[year,i,s+1,title,singer,album,count]
        saveexcel(xlsfile,song,colname)    
    print(str(year)+'-'+str(i))

    
print('共用时：'+str(time.time()-starttime)+'s')  
