brexcel
======

###MOTIVATION

Recently, I had to parse lots of excel files that were sent by bussiness departs and some IT apps. I had to parse these excel files and enrich them with other data (metrics, documentation, etc).

This work consisted in the following flow:

[![](http://yuml.me/b9072e48)](DataExample)

This is the my first motivation.

I never packaged a python project so... my second motivation is to package a python project.



###What brexcel does


Simply:

- reads an excel file into a dict
- writes a dict into an excel file

####Reading excel files

```
# coding=utf8
from brexcel.rexcel import RExcel
from pprint import pprint

f = RExcel('ISO3166-2.xlsx')
arr = f.getDictByField('Code', ['Name'])

pprint(arr)
```
