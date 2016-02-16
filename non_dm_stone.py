#!/usr/bin/env python
# coding: utf-8

import xmlrpclib
import xlrd

# database name as per changes
dbname = 'fusion_2_15'
# dbname = 'jc17'
username = 'admin'
pwd = 'admin'
# server address
server = 'http://161.202.180.220:9092'
# server = 'http://127.0.0.1:8069'

sock_common = xmlrpclib.ServerProxy(server + '/xmlrpc/common')
sock = xmlrpclib.ServerProxy(server + '/xmlrpc/object')
uid = sock_common.login(dbname, username, pwd)


workbook = xlrd.open_workbook('style stone 090216 final.xlsx')
worksheet = workbook.sheet_by_name('Stlstn')

num_rows = worksheet.nrows - 1
num_cells = worksheet.ncols - 1
curr_row = 0

cst = []
while curr_row < num_rows:
    curr_row += 1
    row = worksheet.row(curr_row)
    if row[1].value != "DM":
        # print "===> ", curr_row,
        cst_cm = row[1].value + " " + row[2].value + " " + row[3].value
        sp = sock.execute(dbname, uid, pwd, 'style.product', 'search', [('name', '=', cst_cm)])
        if not sp:
            style_product_id = {
                'name': cst_cm,
                'product_type': 'stockable',
                'category': 'stone',
                'sale_purchase_selection': 'purchase',
                'unit_of_measure': 1,
                'stone_type': 'gem_stone',
            }
        # print "========>", style_product_id
            print "=cst_stone=====", sock.execute(dbname, uid, pwd, 'style.product', 'create', style_product_id)
        print "====>", curr_row
        setting = ""
        if row[5].value == "WSET":
            setting = "wset"
        if row[5].value == "MSET":
            setting = "mset"

        master_style_id = sock.execute(dbname, uid, pwd, 'master.style', 'search', [('name', '=', row[0].value)])
        if not master_style_id:
            master_style_id = ""
        else:
            master_style_id = master_style_id[0]
        product_id = sock.execute(dbname, uid, pwd, 'style.product', 'search', [('name', '=', cst_cm), ('category', '=', 'stone')])
        if not product_id:
            product_id = ""
        else:
            product_id = product_id[0]
        stone_id_data = sock.execute(dbname, uid, pwd, 'product.stone', 'search', [('master_id', '=', master_style_id), ('product_id', '=', product_id), ('stone_seive_size', '=', row[4].value), ('stone_wt', '=', row[8].value), ('stone_ct', '=', row[7].value), ('stone_setting', '=', setting)])
        if not stone_id_data:
            stone_ids = {
                 'master_id': master_style_id,
                 'product_id': product_id,
                 'stone_seive_size': row[4].value,
                 'stone_qty': row[6].value,
                 'stone_uom': 1,
                 'stone_wt': row[8].value,
                 'stone_ct': row[7].value,
                 'stone_setting': setting,
             }
            print "=======stone ids----->", sock.execute(dbname, uid, pwd, 'product.stone', 'create', stone_ids)
        if master_style_id:
            style_product = sock.execute(dbname, uid, pwd, 'style.product', 'search', [('product_id', '=', master_style_id), ('is_master', '=', False)])
            # print "=====", style_product
            for style in style_product:
                stone_id_data = sock.execute(dbname, uid, pwd, 'product.stone', 'search', [('style_id', '=', style), ('product_id', '=', product_id), ('stone_seive_size', '=', row[4].value), ('stone_wt', '=', row[8].value), ('stone_ct', '=', row[7].value), ('stone_setting', '=', setting)])
                if not stone_id_data:
                    stone_ids = {
                         'style_id': style,
                         'product_id': product_id,
                         'stone_seive_size': row[4].value,
                         'stone_qty': row[6].value,
                         'stone_uom': 1,
                         'stone_wt': row[8].value,
                         'stone_ct': row[7].value,
                         'stone_setting': setting,
                     }
                    print "=======stone ids----->", sock.execute(dbname, uid, pwd, 'product.stone', 'create', stone_ids)
print "========cst stone added--------------"
