# -*- coding: utf-8 -*-
# -----------------------------------------------------------------------------------
#
#    Copyright (C) 2017  jeo Software  (http://www.jeosoft.com.ar)
#    All Rights Reserved.
#
#    This program is free software: you can redistribute it and/or modify
#    it under the terms of the GNU Affero General Public License as
#    published by the Free Software Foundation, either version 3 of the
#    License, or (at your option) any later version.
#
#    This program is distributed in the hope that it will be useful,
#    but WITHOUT ANY WARRANTY; without even the implied warranty of
#    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#    GNU Affero General Public License for more details.
#
#    You should have received a copy of the GNU Affero General Public License
#    along with this program.  If not, see <http://www.gnu.org/licenses/>.
#
# -----------------------------------------------------------------------------------

# -*- coding: utf-8 -*-
import openpyxl


class Product(object):
    def __init__(self, data):
        self._code = data[0]
        self._desc = data[1]
        self._list = float(data[2])
        self._cost = float(data[3])

    @property
    def code(self):
        return self._code

    @property
    def name(self):
        return self._desc

    @property
    def list(self):
        return self._list

    @property
    def cost(self):
        return self._cost

    def dump(self):
        return u'[{:10}]  ${:6.2f}  ${:6.2f}  {:10}  '.format(self.code, self.list, self.cost, self.name)


class Onceworksheet(object):
    def __init__(self, filename):

        def decode(acell):
            ret = acell.value
            return ret

        self._prods = []

        wb = openpyxl.load_workbook(filename=filename, read_only=True)
        sheet = wb.get_sheet_by_name('Hoja1')

        # leo toda la planilla
        for row in sheet.iter_rows(min_row=3, min_col=1, max_col=4, max_row=600):
            rowlist = []
            # itero en toda la fila por celdas
            for cell in row:
                rowlist.append(decode(cell))
            self._prods.append(Product(rowlist))

    def list(self):
        return self._prods

    def prod(self, default_code):
        """ Search a product in once worksheet """
        ret = False
        for p in self._prods:
            if p.code == default_code:
                ret = p
        return ret

    def prods(self):
        return self._prods
