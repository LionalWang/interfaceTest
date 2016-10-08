# coding=utf-8
import xlrd
import xlwt
import json
import urllib2


data = xlrd.open_workbook('interface.xlsx')  # 打开文档

table = data.sheets()[0]  # 读取第一个sheet

newData = xlwt.Workbook()  # 创建一个新的文档

sheet = newData.add_sheet(u'result', cell_overwrite_ok=True)  # 创建sheet

url = table.cell(1, 0).value

for i in range(1, table.ncols):
    a = table.row_values(0)[i]
    a = str(a)
    b = table.row_values(1)[i]
    b = str(b)
    url = url+"&"+a+"="+b

response = urllib2.urlopen(url)  # 访问接口

apiContent = response.read()  # 储存返回的接口数据

jsonObject = json.loads(apiContent)  # 把返回的字符串变成字典


def processlist(result, key):
    if key.isdigit():
        return result[int(key)]
    else:
        return result[key]


for i in range(4, table.nrows):
    initWord = table.col_values(0)[i]
    dictWord = initWord.split(".")
    wordLength = len(dictWord)

    if dictWord[0] in jsonObject:
        value = jsonObject[dictWord[0]]
        if isinstance(value, list) | isinstance(value, dict):
            for j in range(1, wordLength):
                value = processlist(value, dictWord[j])

    res = False
    if table.col_values(1)[i] == 'str':
        res = isinstance(value, unicode)
    elif table.col_values(1)[i] == 'int':
        res = isinstance(value, int)
    elif table.col_values(1)[i] == 'boolean':
        res = isinstance(value, bool)

    sheet.write(i, 0, initWord)
    sheet.write(i, 1, table.col_values(1)[i])
    sheet.write(i, 2, res)

newData.save('result.xls')
