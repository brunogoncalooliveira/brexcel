brexcel
======

#### MOTIVATION

Recently, I had to parse lots of excel files that were sent by bussiness departs and some IT apps. I had to parse these excel files and enrich them with other data (metrics, documentation, etc).

This work consisted in the following flow:

[![](http://yuml.me/b9072e48)](DataExample)

This is the my first motivation.

I never packaged a python project so... my second motivation is to package a python project.



#### What brexcel does


Simply:

- reads an excel file into a dict
- writes a dict into an excel file

#### Reading excel files

```python
# coding=utf8
from brexcel.rexcel import RExcel
from pprint import pprint

f = RExcel('ISO3166-2.xlsx')
arr = f.getDictByField('Code', ['Name'])

pprint(arr)
```


### Writing excel files

```python
from brexcel.rexcel import RExcel

arr = [{'Name': u'root node',  'top_id': '',  'id': '1'},
       {'Name': u'First Leaf',  'top_id': '1',  'id': '2'},
       {'Name': u'Second leaf inside root node',  'top_id': '1',  'id': '3'},
       {'Name': u'another root node',  'top_id': '',  'id': '4'}]

f = WExcel(arr)
f.header_order = ['id', 'top_id', 'Name']
f.header_alias = {'id': 'id', 'top_id': 'Top Node'}
f.SaveExcelAs('myfilename.xlsx', 'Sheet1')
```

The above will create an excel file like this:

| id          | Top Node | Name              |
 ------------ | ---------| ------------------
| 1           |          | root node  |
| 2           | 1        | First Leaf |
| 3           | 1        | Second leaf inside root node |
| 4           |          | another root node |


> **Note:**
> - See the test files inside tests folder.
