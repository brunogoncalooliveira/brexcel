# coding=utf8
from openpyxl import load_workbook

class RExcel:

    def __init__(self, filename, sheet="Sheet1"):
        self.filename = filename
        self.sheet = sheet
        self.wb2 = load_workbook(filename)
        self.sheet = self.wb2[sheet]

    def getDictByField(self, name_of_field, fields=[]):
        first = True
        allfields = False
        if fields == []:
            allfields = True

        header = {}

        chave = name_of_field
        arr = {}

        for row in self.sheet.rows:
            if first:
                for i in row:
                    header[ i.value ] =  i.col_idx - 1
                    if allfields and i.value != name_of_field:
                         fields.append(i.value)
                first = False
            else:
                if row[header[ chave ]].value != '':
                    if row[header[ chave ]].value not in arr:
                        arr[row[header[ chave ]].value] = []
                    tmp = {}
                    for i in fields:
                        tmp[i] = row[header[ i ]].value
                    arr[row[header[ chave ]].value].append(tmp)

        return arr

    def getDict(self):
        first = True
        header = {}
        arr = []
        for row in self.sheet.rows:
            if first:
                for i in row:
                    header[ i.value ] =  i.col_idx - 1
                first = False
            else:
                tmp = {}
                for i in header:
                    tmp[i] = row[header[ i ]].value
                arr.append(tmp)
        return arr
