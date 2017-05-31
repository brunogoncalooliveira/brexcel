# coding=utf8
from unittest import TestCase

from pprint import pprint

from brexcel.rexcel import RExcel
from brexcel.wexcel import WExcel

class TestRExcel(TestCase):
    def test_is_string(self):

        f = RExcel('ISO3166-2.xlsx')
        self.assertTrue(isinstance(f.filename, basestring))
        #arr = f.getDictByField('Grupo', ['Tabela'])

#f = RExcel('sample.xlsx', 'survey')
#f = RExcel('FuncionalidadesNegocio.xlsx')
#arr = f.getDictByField('ID da Funcionalidade')
#f.getDict()
#pprint(arr)
#exit()

"""


arr = [{'Alias': '581 - Consultar todas as tabelas gerais do banco',
  'NivelSup': '',
  'id': 'gt1'},
 {'Alias': u'795 Alterar c\xf3digos de diverg\xeancia',
  'NivelSup': 'gt831',
  'id': 'gt843'},
 {'Alias': u'799 Alterar descritivos motivos devolu\xe7\xe3o cheques',
  'NivelSup': 'gt831',
  'id': 'gt844'},
 {'Alias': u'804 Alterar prec\xe1rio base da telecompensa\xe7\xe3o',
  'NivelSup': 'gt831',
  'id': 'gt845'}]

#pprint(arr)



f = WExcel(arr)
f.header_order = ['NivelSup', 'id', 'Alias']
f.header_alias = {'id': 'campo1', 'NivelSup': 'Nível Superior'}
f.SaveExcelAs('myfilename.xlsx', 'Sheet1')


"""