# coding=utf8
from brexcel.wexcel import WExcel


arr = [{'Name': u'root node',  'top_id': '',  'id': '1'},
       {'Name': u'First Leaf',  'top_id': '1',  'id': '2'},
       {'Name': u'Second leaf inside root node',  'top_id': '1',  'id': '3'},
       {'Name': u'another root node',  'top_id': '',  'id': '4'}]


f = WExcel(arr)
f.header_order = ['id', 'top_id', 'Name']
f.header_alias = {'id': 'id', 'top_id': 'Top Node'}
f.SaveExcelAs('myfilename.xlsx', 'Sheet1')

