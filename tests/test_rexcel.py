# coding=utf8
from brexcel.rexcel import RExcel
from pprint import pprint



# example 1

f = RExcel('ISO3166-2.xlsx')
arr = f.getDictByField('Code')  # returns all columns
pprint(arr['PT'])




# example 2

f = RExcel('ISO3166-2.xlsx')
arr = f.getDictByField('Code', ['Name']) # returns only column 'Name'
pprint(arr['PT'])



# example 3

f = RExcel('ISO3166-2.xlsx')
arr = f.getDict()
pprint(arr[0]) # prints first record




# example 4

f = RExcel('Financial Sample.xlsx')
arr = f.getDictByField('Country', ['Segment', 'Product', 'Units Sold'])
pprint(arr)

"""

prints:
{u'Canada': [{'Product': u'Carretera',
              'Segment': u'Government',
              'Units Sold': 1618.5},
             {'Product': u'Montana',
              'Segment': u'Channel Partners',
              'Units Sold': 2518L},
             ...
 u'France': [{'Product': u'Carretera',
              'Segment': u'Midmarket',
              'Units Sold': 2178L},
             {'Product': u'Montana',
              'Segment': u'Government',
              'Units Sold': 1899L},
"""