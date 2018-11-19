# -*- coding: utf-8 -*-

"""
  从Unit2_tag.xlsx提取点信息，在Redis中建立点键

  代码没有使用类
"""

import xlrd
import redis

try:
    conn = redis.Redis('localhost')
except:
    pass


def TagDefToRedisHashKey(tagdeflist):
    pipe = conn.pipeline()
    for element in tagdeflist:
        pipe.hmset(
            element['id'], {'desc': element['desc'], 'value': "-10000", 'ts': ""})
    pipe.execute()


def open_excel(file):
    try:
        data = xlrd.open_workbook(file)
        return data
    except Exception as e:
        print(e)


def tagdef_from_excel_sheet(excelfile, rowindex, colidindex, coldescindex, sheet_name):
    sheet = excelfile.sheet_by_name(sheet_name)
    nrows = sheet.nrows
    tagdeflist = []
    for rownum in range(rowindex, nrows):
        tagdef = {'id': u'', 'desc': u''}
        tagdef['id'] = sheet.cell(rownum, colidindex).value
        tagdef['desc'] = sheet.cell(rownum, coldescindex).value
        tagdeflist.append(tagdef)
    return tagdeflist


def UnitTagDefToRedisHash(Sheetname):

    tagdeflist = tagdef_from_excel_sheet(
        excelfile, rowbegindex, colidindex, coldescindex, Sheetname)
    TagDefToRedisHashKey(tagdeflist)

    TagCount = len(tagdeflist)
    print(tagdeflist[TagCount - 1])
    print('TagCount= ', TagCount, ' Redis Keys=', conn.dbsize())

if __name__ == "__main__":

    rowbegindex = 2
    colidindex = 2
    coldescindex = 1

    excelfile = open_excel('./unit2_tag.xlsx')

    UnitTagDefToRedisHash(u'DCS2AI')
