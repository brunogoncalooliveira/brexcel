# coding=utf8
"""

TO DO

"""
from unittest import TestCase

from pprint import pprint

from brexcel.rexcel import RExcel

class TestRExcel(TestCase):
    def test_is_string(self):

        f = RExcel('ISO3166-2.xlsx')
        print('ok')

        self.assertTrue(isinstance(f.filename, basestring))
        #arr = f.getDictByField('Grupo', ['Tabela'])
