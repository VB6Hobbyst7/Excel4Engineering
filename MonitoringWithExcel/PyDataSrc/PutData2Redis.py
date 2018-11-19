
# -*- coding: utf-8 -*-
"""
 　Put Data to Redis
"""
from datetime import datetime
import redis
import xlrd

class PeriodSensorTask():

    def __init__(self, file, rowindex, colidindex, colvalueindex, sheet_name):
        
        self.conn = redis.Redis('localhost')
        
        self.excelfile = self.open_excel(file)
        self.rowindex = rowindex
        self.colidindex = colidindex
        self.colvalueindex = colvalueindex
        self.sheet = self.excelfile.sheet_by_name(sheet_name)
        self.nrows = self.sheet.nrows

        self.taglist = self.tagname_from_excel()
        self.tagvaluelist = []

        self.sheetname = sheet_name
        print(self.sheetname)

    def open_excel(self, file):
        try:
            data = xlrd.open_workbook(file)
            return data
        except Exception as e:
            print(e)

    def tagname_from_excel(self):
        taglist = []
        for rownum in range(self.rowindex, self.nrows):
            tagdef = {'id': u'', 'value': '', 'time': ''}
            tagdef['id'] = self.sheet.cell(rownum, self.colidindex).value
            taglist.append(tagdef)
        return taglist

    def tagvalue_from_excel(self):
        taglist = []
        for rownum in range(self.rowindex, self.nrows):
            tagdef = {'id': u'', 'value': '', 'ts': ''}
            tagdef['id'] = self.sheet.cell(rownum, self.colidindex).value
            tagdef['value'] = self.sheet.cell(rownum, self.colvalueindex).value
            taglist.append(tagdef)

        self.colvalueindex = self.colvalueindex + 1
        if self.colvalueindex > 30:
           self.colvalueindex = 6

        return taglist

    def SendToRedisHash(self):

        self.tagvaluelist = self.tagvalue_from_excel()

        curtime = datetime.now()

        pipe = self.conn.pipeline()
        for element in self.tagvaluelist:
            pipe.hmset(
                element['id'], {'value': element['value'], 'ts': curtime})
        pipe.execute()

        print('Point:',self.tagvaluelist[10]['id'],self.conn.hmget(self.tagvaluelist[10]['id'], 'value', 'ts'))

