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
# ----------------------------------------------------------------------------------
import odoorpc
from lib import Onceworksheet
from secret import odoo_key

# conectar con odoo
odoo = odoorpc.ODOO(odoo_key['server'], port=odoo_key['port'])
odoo.login(odoo_key['database'], odoo_key['username'], odoo_key['password'])

prods_odoo_obj = odoo.env['product.product']

ONCE_CATEGS = [5]


def process_all_worksheet_prods():
    """ actualiza o agrega los productos en odoo """
    print 'actualizando productos en odoo'

    once_wk = Onceworksheet('farmacia-once.xlsx')
    for once_prod in once_wk.prods():
        # busco en odoo
        ids = prods_odoo_obj.search([('default_code', '=', once_prod.code)])
        if ids:
            # está en odoo
            odoo_prod = prods_odoo_obj.browse(ids)
            # asegurar que solo hay uno
            if len(ids) > 1:
                print '{} está {} veces'.format(once_prod.code, len(ids))
            assert len(ids) == 1
            for prod in odoo_prod:
                if once_prod.list - prod.lst_price > 0.01:
                    print u'{:.2f} %  [{}] {}'.format(
                            ((once_prod.list - prod.lst_price)/once_prod.list) * 100,
                            once_prod.code,
                            once_prod.name
                    )
                else:
                    print '.',
                prod.lst_price = once_prod.list
                prod.standard_price = once_prod.cost
                prod.name = once_prod.name
                prod.cost_method = 'real'
                prod.sale_ok = True
        else:
            # no está en odoo, lo agregamos
            id_prod = prods_odoo_obj.create({
                'default_code': once_prod.code,
                'standard_price': once_prod.cost,
                'lst_price': once_prod.list,
                'name': once_prod.name,
                'cost_method': 'real',
                'categ_id': ONCE_CATEGS[0],
                'sale_ok': True
            })
            print
            print 'add  ----------------------', once_prod.code, once_prod.name


def list_categ():
    categ_obj = odoo.env['product.category']
    categs = categ_obj.browse(ONCE_CATEGS)
    for cat in categs:
        print cat.id, cat.name


def list_odoo_products():
    ids = prods_odoo_obj.search([('categ_id', '=', ONCE_CATEGS[0])])
    prods = prods_odoo_obj.browse(ids)
    for prod in prods:
        print u'"{}","{}",{},{}'.format(prod.default_code, prod.name, prod.list_price, prod.standard_price)


def list_once_worksheet():
    once_wk = Onceworksheet('farmacia-once.xlsx')
    for prod in once_wk.prods():
        print prod.dump()


#list_once_worksheet()
process_all_worksheet_prods()
#list_categ()
#list_odoo_products()


