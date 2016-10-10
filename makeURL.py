# coding=utf-8
import xlrd
import xlwt
import json
import urllib2


def start(filename, number):

    try:
        data = xlrd.open_workbook(filename)  # 打开文档
    except Exception, e:
        print "Cant find file", e

    else:
        for table in data.sheets():  # 循环所有sheet
            new_data = xlwt.Workbook()  # 创建一个新的文档
            sheet = new_data.add_sheet(table.name, cell_overwrite_ok=True)  # 创建sheet
            url = table.cell(1, 0).value  # 创建url

            for i in range(1, table.ncols):
                a = table.row_values(0)[i]
                a = str(a)
                b = table.row_values(1)[i]
                b = str(b)
                url = url+"&"+a+"="+b

            try:
                response = urllib2.urlopen(url)  # 访问接口
                api_content = response.read()  # 储存返回的接口数据
                json_object = json.loads(api_content)  # 把返回的字符串变成字典
            except Exception, e:
                print "Cant open url", e

            else:
                for i in range(4, table.nrows):
                    init_word = table.col_values(0)[i]
                    dict_word = init_word.split(".")
                    word_length = len(dict_word)

                    if dict_word[0] in json_object:
                        value = json_object[dict_word[0]]
                        if isinstance(value, list) | isinstance(value, dict):
                            for j in range(1, word_length):
                                value = processlist(value, dict_word[j])

                    res_type = False
                    res_value = False
                    if table.col_values(1)[i] == 'str':
                        res_type = isinstance(value, unicode)
                    elif table.col_values(1)[i] == 'int':
                        res_type = isinstance(value, int)
                    elif table.col_values(1)[i] == 'boolean':
                        res_type = isinstance(value, bool)

                    if value == table.col_values(2)[i]:
                        res_value = True

                    sheet.write(i, 0, init_word)
                    sheet.write(i, 1, table.col_values(1)[i])
                    sheet.write(i, 2, table.col_values(2)[i])
                    sheet.write(i, 3, res_type)
                    sheet.write(i, 4, res_value)

                    print table.col_values(2)[i]
                    print value
                    print res_value
                    print " "

                new_data.save('result'+str(number)+'.xls')


def processlist(result, key):  # 处理list或object的情况
        if key.isdigit():
            return result[int(key)]
        else:
            return result[key]
