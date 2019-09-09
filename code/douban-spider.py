import xlwt
import re
import urllib.request
import urllib.error
import time
from lxml import etree
import requests
import numpy as np

#创建excel文件，一个记录电影详情，一个记录有多个项内容的节点用于统计
workbook=xlwt.Workbook(encoding='utf-8')
worksheet=workbook.add_sheet('总览')
worksheet2=workbook.add_sheet('分支详细')
row0=[u'电影',u'年份',u'地区',u'语言',u'类型',u'导演',u'主演',u'评分',u'五星',u'四星',u'三星',u'二星',u'一星',u'评分人数',u'排名',u'链接']
row1=[u'地区详情',u'语言详情',u'类型详情',u'主演详情',u'标签详情']
for k in range(len(row0)):
    worksheet.write(0,k,row0[k])
for n in range(len(row1)):
    worksheet2.write(0,n,row1[n])

#浏览器伪装
header=('User-Agent','Mozilla/5.0 (Windows NT 6.1; Win64; x64; rv:67.0) Gecko/20100101 Firefox/67.0')
opener=urllib.request.build_opener()
opener.addheaders=[header]
urllib.request.install_opener(opener)
url='https://movie.douban.com/top250?start=0'
num=0
w_len = 0
x_len = 0
y_len = 0
z_len = 0
v_len = 0
for i in range(10):
    req=requests.get(url)
    html=etree.HTML(req.text)
    url_detail=html.xpath('//ol/li/div[@class="item"]/div[@class="info"]/div[@class="hd"]/a/@href')
    for j in range(len(url_detail)):
        #进行数据爬取，数据筛选以及清洗
        req_detail=requests.get(url_detail[j])
        html_detail=etree.HTML(req_detail.text)
        data=urllib.request.urlopen(url_detail[j]).read().decode()
        name=html_detail.xpath('//span[@property="v:itemreviewed"]/text()')
        year=html_detail.xpath('//span[@class="year"]/text()')
        pat_year='\((.*?)\)'
        year_num=re.compile(pat_year).findall(year[0])
        director=html_detail.xpath('//div[@id="info"]/span[1]/span[2]/a/text()')
        if name == ['二十二']:
            actor=['韦绍兰','罗善学']
        else:
            actor=html_detail.xpath('//a[@rel="v:starring"]/text()')
        movie_type=html_detail.xpath('//span[@property="v:genre"]/text()')
        pat_loc='制片国家/地区:</span> (.*?)<br'
        pat_lng='语言:</span> (.*?)<br'
        movie_loc=re.compile(pat_loc).findall(data)
        movie_lng=re.compile(pat_lng).findall(data)
        pat_rank='\d{1,3}'
        rank=html_detail.xpath('//span[@class="top250-no"]/text()')
        rank_num=re.compile(pat_rank).findall(rank[0])
        vote=html_detail.xpath('//strong[@property="v:average"]/text()')
        people=html_detail.xpath('//span[@property="v:votes"]/text()')
        star_5=html_detail.xpath('//div[@class="ratings-on-weight"]/div[1]/span[2]/text()')
        star_4=html_detail.xpath('//div[@class="ratings-on-weight"]/div[2]/span[2]/text()')
        star_3=html_detail.xpath('//div[@class="ratings-on-weight"]/div[3]/span[2]/text()')
        star_2=html_detail.xpath('//div[@class="ratings-on-weight"]/div[4]/span[2]/text()')
        star_1=html_detail.xpath('//div[@class="ratings-on-weight"]/div[5]/span[2]/text()')
        movie_tag=html_detail.xpath('//div[@class="tags-body"]/a/text()')

        #写第一个sheet
        worksheet.write(i*len(url_detail)+j+1,0,name)
        worksheet.write(i*len(url_detail)+j+1,1,year_num)
        worksheet.write(i*len(url_detail)+j+1,2,movie_loc)
        worksheet.write(i*len(url_detail)+j+1,3,movie_lng)
        worksheet.write(i*len(url_detail)+j+1,4,movie_type)
        worksheet.write(i*len(url_detail)+j+1,5,director)
        worksheet.write(i*len(url_detail)+j+1,6,actor[0])
        worksheet.write(i*len(url_detail)+j+1,7,vote)
        worksheet.write(i*len(url_detail)+j+1,8,star_5)
        worksheet.write(i*len(url_detail)+j+1,9,star_4)
        worksheet.write(i*len(url_detail)+j+1,10,star_3)
        worksheet.write(i*len(url_detail)+j+1,11,star_2)
        worksheet.write(i*len(url_detail)+j+1,12,star_1)
        worksheet.write(i*len(url_detail)+j+1,13,people)
        worksheet.write(i*len(url_detail)+j+1,14,rank_num)
        worksheet.write(i*len(url_detail)+j+1,15,url_detail[j])

        #写第二个sheet
        #分析数据发现地区以及语言未分开，使用spilt进行分离
        movie_loc_spilt=movie_loc[0].split(' / ')
        movie_lng_spilt=movie_lng[0].split(' / ')
        for w in range(len(movie_loc_spilt)):
            worksheet2.write(w_len+w+1,0,movie_loc_spilt[w])
        for x in range(len(movie_lng_spilt)):
            worksheet2.write(x_len+x+1,1,movie_lng_spilt[x])
        for y in range(len(movie_type)):
            worksheet2.write(y_len+y+1,2,movie_type[y])
        #取前三分之一的演员进行分析
        len_actor=0
        if len(actor) < 3:
            len_actor = 1
        else:
            len_actor = len(actor)//3
        for z in range(len_actor):
            worksheet2.write(z_len+z+1,3,actor[z])
        for v in range(len(movie_tag)):
            worksheet2.write(v_len+v+1,4,movie_tag[v])
        w_len += len(movie_loc_spilt)
        x_len += len(movie_lng_spilt)
        y_len += len(movie_type)
        z_len += len_actor
        v_len += len(movie_tag)
        #因免费的有效的IP代理现在有点难找，所以选用暂停来降低访问频率，避免被网站禁止
        time.sleep(np.random.randint(1,3))
    num += 25
    url='https://movie.douban.com/top250?start='+str(num)
    time.sleep(np.random.randint(2,4))
workbook.save('douban-full.xlsx')
