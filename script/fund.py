# coding=utf-8


import requests
import re
# 如有不明白的地方,还请访问博文: www.3code.info。内有使用简介
import sys

from xlwt.compat import unicode_type

reload(sys)
sys.setdefaultencoding('utf-8')

import xlrd
import xlwt
import time, datetime
from xlutils.copy import copy
from xlrd import open_workbook
from xlwt import easyxf

# 宏定义数据

path = '/Users/3code/fund/data.xls'
url = 'http://www.jjmmw.com/fund/sypm/?order=1year&limit=0'
# 储存的日期
timeNow = time.strftime('%Y-%m-%d %Hh', time.localtime(time.time()))
# 今天日期
timeToday = time.strftime('%Y-%m-%d', time.localtime(time.time()))
# 前一天日期
timePrevious = time.strftime('%Y-%m-%d', time.localtime(time.time() - 24 * 60 * 60))




class jSpider(object):
    def __init__(self):
        print '开始爬取内容。。。'

    # getsource用来获取网页源代码
    def getSource(self, url):
        html = requests.get(url)
        return html.text

    def getMainText(self, source):
        everyclass = re.search(
            '(<table width="950" border="0" align="center" cellpadding="0" cellspacing="0" class="table02" id="fundlist" data-lt="sypm">.*?</table>)',
            source, re.S).group(1)
        return everyclass

    def getEveryClass(self, source):
        maintext = re.findall('(<td class="t13".*?</td></tr>)', source, re.S)
        return maintext


    def getInfo(self, eachclass):
        info = {}
        info["num"] = re.search('<td width="3%" align="center">(.*?)</td>', eachclass, re.S).group(1)
        info["id"] = re.search('target="_blank">(.*?)</a></td>', eachclass, re.S).group(1)
        info['name'] = re.search('target="_blank" class="blue ellipsis" title="(.*?)">', eachclass, re.S).group(1)
        info['time'] = re.search('<td width="7%" align="center">(.*?)</td>', eachclass, re.S).group(1)
        info["price1"] = re.findall('<td width="6%" align="center">(.*?)</td>', eachclass, re.S)[1]
        info["price2"] = re.findall('<td width="6%" align="center">(.*?)</td>', eachclass, re.S)[2]

        return info


    def saveInfo(self, classinfo):

        # 写入文本:
        # r:读 w:写 a:追加 b:二进制(可添加在其他模式) +:读写(可添加在其他模式)
        f = open("/Users/3code/fund/" + timePrevious + ".txt", 'a+')
        for each in classinfo:
            f.writelines('xuhao:' + each['num'] + '\n')
            f.writelines('daima:' + each['id'] + '\n')
            f.writelines('name:' + each['name'] + '\n')
            f.writelines('time:' + each['time'] + '\n')
            f.writelines('price1:' + each['price1'] + '\n')
            f.writelines('price2:' + each['price2'] + '\n')
            f.writelines('url:http://www.jjmmw.com/fund/' + each['id'] + '/' + '\n\n')
        f.close()

        rb = open_workbook(path, formatting_info=True)
        r_sheet = rb.sheet_by_index(0)
        wb = copy(rb)
        sheet = wb.get_sheet(0)
        j = r_sheet.ncols
        i = 1
        timeShould = timePrevious
        timeQ = timePrevious
        tListB = []
        tListS = []
        print sheet

        fir = ['ID', 'Time', 'NowPrice', 'PlusPrice', 'WriteTime']


        for each in classinfo:
            sheet.write(i, j + 0, each['id'])
            sheet.write(i, j + 1, each['time'])
            sheet.write(i, j + 2, each['price1'])
            sheet.write(i, j + 3, each['price2'])
            i += 1
            timeQ = each['time']
            if timeQ > timeShould:
                tListB.append(each['id']+' '+timeQ)
            elif timeQ < timeShould:
                tListS.append(each['id']+' '+timeQ)


        for k in range(j, j + 5):
            sheet.write(0, k, fir[k - j])
            print k
            if k == j + 4:
                sheet.write(1, k, timeNow.decode('utf-8'))
                sheet.write(2, k, 'UNupdate:' + str(len(tListS)))

        print "未及时更新列表:"
        print tListS;
        print path
        print "数量:" + str(len(tListS))
        print "数据错误列表:"
        print tListB

        wb.save(path)
        print path


# 判断写入的时间是否已经存在:
def read_excel():
    rb = open_workbook(path, formatting_info=True)
    r_sheet = rb.sheet_by_index(0)
    # 遍历第2行所有列,是否有时间数据c
    print "\n已经爬取的列数:" + str(r_sheet.ncols) + "\n"
    for c in range(0, r_sheet.ncols):
        age_nov = r_sheet.cell(1, c).value
        print age_nov
        if ((c + 1) % 5 == 0) and (c != 0):
            print "\n"
        if age_nov == timeNow:
            return 1
    else:
        return 0


# 按日期更新
def getTime():
    timeT = datetime.date.today()
    weekOne = timeT.weekday()
    print weekOne
    if weekOne >= 1 and weekOne <= 5:
        print weekOne
        return timeToday
    else:
        print "今天是周日或者周一,不获取信息。周六获取周五,周二获取周一的信息"
        return 0


if __name__ == '__main__':

    classinfo = []
    # 首先判断时间是否满足:
    if getTime() == 0:
        exit()

    j3codeSpider = jSpider()
    print u'正在处理页面：' + url
    html = j3codeSpider.getSource(url)
    # html = html.decode('utf-8')
    maint = j3codeSpider.getMainText(html)
    everyclass = j3codeSpider.getEveryClass(maint)

    for each in everyclass:
        info = j3codeSpider.getInfo(each)
        classinfo.append(info)

    if read_excel():
        print "\n 已经爬取过数据!!!"
    else:
        j3codeSpider.saveInfo(classinfo)
    u = "23"

'''

'''