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

# cnt = 0
# create umo
umo = sock.execute(dbname, uid, pwd, 'product.uom.category.style', 'search', [('name', '=', 'Unit(s)')])
if not umo:
    umo = {
        'name': 'Unit(s)'
    }
    sock.execute(dbname, uid, pwd, 'product.uom.category.style', 'create', umo)
print "======umo======="
umo = sock.execute(dbname, uid, pwd, 'product.uom.category.style', 'search', [('name', '=', 'Weight')])
if not umo:
    umo = {
        'name': 'Weight'
    }
    sock.execute(dbname, uid, pwd, 'product.uom.category.style', 'create', umo)
print "======umo======="

u_m_o = sock.execute(dbname, uid, pwd, 'product.uom.style', 'search', [('name', '=', 'Unit(s)')])
if not u_m_o:
    umo = sock.execute(dbname, uid, pwd, 'product.uom.category.style', 'search', [('name', '=', 'Unit(s)')])
    u_m_o = {
        'name': 'Unit(s)',
        'category_id': umo[0]
    }
    sock.execute(dbname, uid, pwd, 'product.uom.style', 'create', u_m_o)
print "======u_m_o======="

u_m_o = sock.execute(dbname, uid, pwd, 'product.uom.style', 'search', [('name', '=', 'g')])
if not u_m_o:
    umo = sock.execute(dbname, uid, pwd, 'product.uom.category.style', 'search', [('name', '=', 'Weight')])
    u_m_o = {
        'name': 'g',
        'category_id': umo[0]
    }
    sock.execute(dbname, uid, pwd, 'product.uom.style', 'create', u_m_o)
print "======u_m_o======="

workbook = xlrd.open_workbook('stylemaster 010216.xlsx')
worksheet = workbook.sheet_by_name('Main Page data')

num_rows = worksheet.nrows - 1
num_cells = worksheet.ncols - 1
curr_row = 0

print "Importing master style..."

while curr_row < num_rows:
    curr_row += 1
    # cnt += 1
    row = worksheet.row(curr_row)
    print "=master style====curr_row======", curr_row
    partner_id = sock.execute(dbname, uid, pwd, 'res.partner', 'search', [('name', '=', row[5].value)])
    if not partner_id:
        if row[5].value or not row[5].value == "":
            res_partner = {
                'name': row[5].value,
                'customer': True,
            }
            print "partner_id======>", sock.execute(dbname, uid, pwd, 'res.partner', 'create', res_partner)
    metal_id = sock.execute(dbname, uid, pwd, 'product.type', 'search', [('name', '=', row[7].value)])
    if not metal_id:
        metal_id = {
            'name': row[7].value,
        }
        print "metal_id=====>", sock.execute(dbname, uid, pwd, 'product.type', 'create', metal_id)
    jewel_code_id = sock.execute(dbname, uid, pwd, 'product.jewel.code', 'search', [('name', '=', row[1].value), ('code', '=', row[2].value)])
    if not jewel_code_id:
        jewel_code_id = {
            'code': row[2].value,
            'name': row[1].value,
        }
        print "jewel_code_id====>", sock.execute(dbname, uid, pwd, 'product.jewel.code', 'create', jewel_code_id)

    master_style_id = sock.execute(dbname, uid, pwd, 'master.style', 'search', [('name', '=', row[0].value)])
    if not master_style_id:
        jewel_code_id = sock.execute(dbname, uid, pwd, 'product.jewel.code', 'search', [('name', '=', row[1].value), ('code', '=', row[2].value)])
        metal_id = sock.execute(dbname, uid, pwd, 'product.type', 'search', [('name', '=', row[7].value)])
        partner_id = sock.execute(dbname, uid, pwd, 'res.partner', 'search', [('name', '=', row[5].value)])
        if not partner_id:
            partner_id = ""
        else:
            partner_id = partner_id[0]
        master_style = {
            'name': row[0].value,
            'jewel_code_id': jewel_code_id[0],
            'description': row[1].value,
            'customer_id': partner_id,
            'metal_type_id': metal_id[0],
            'size': row[9].value,
            'length_mm': row[14].value,
            'width_mm': row[15].value,
            'height_mm': row[16].value,
            'length_inch': row[17].value,
            'width_inch': row[18].value,
            'height_inch': row[19].value,
            'shrank_width_mm': row[21].value,
            'shrank_width_inch': row[22].value,
            'sale_purchase_selection': 'sale',
            'unit_of_measure': 1,
        }
        print "=======master_style========", sock.execute(dbname, uid, pwd, 'master.style', 'create', master_style)
print "style master uploaded Successfully...!"


workbook = xlrd.open_workbook('Jno master 270116.xlsx')
worksheet = workbook.sheet_by_name('pstlmast')

print "========product variant create==============="
num_rows = worksheet.nrows - 1
num_cells = worksheet.ncols - 1
curr_row = 0


while curr_row < num_rows:
    curr_row += 1
    # cnt += 1
    row = worksheet.row(curr_row)
    print "=product variant====curr_row======", curr_row
    master_style_id = sock.execute(dbname, uid, pwd, 'master.style', 'search', [('name', '=', row[1].value)])
    if master_style_id:
        product_variant = sock.execute(dbname, uid, pwd, 'style.product', 'search', [('product_id', '=', master_style_id[0]), ('name', '=', row[2].value), ('j_no', '=', row[0].value)])
        if not product_variant:
            color = sock.execute(dbname, uid, pwd, 'product.color', 'search', [('name', '=', row[6].value)])
            if not color and row[6].value != '':
                color_name = {
                 'name': row[6].value,
                 }
                print "==color--created==>", sock.execute(dbname, uid, pwd, 'product.color', 'create', color_name)
            color = sock.execute(dbname, uid, pwd, 'product.color', 'search', [('name', '=', row[6].value)])
            if color:
                color = color[0]
            else:
                color = ''
            product_variant = {
                'name': row[2].value,
                'j_no': row[0].value,
                'product_id': master_style_id[0],
                'variants_metal_color_id': color,
                'variants_description': row[5].value,
                'sale_purchase_selection': 'sale',
                'is_master': False,
            }
            print "product_variant====>", sock.execute(dbname, uid, pwd, 'style.product', 'create', product_variant)


# general detail add in variants
workbook = xlrd.open_workbook('stylemaster 010216.xlsx')
worksheet = workbook.sheet_by_name('Main Page data')

num_rows = worksheet.nrows - 1
num_cells = worksheet.ncols - 1
curr_row = 0
while curr_row < num_rows:
    curr_row += 1
    row = worksheet.row(curr_row)
    #  add metal in master style line
    print "===mastre=row=====>", curr_row
    master_style = sock.execute(dbname, uid, pwd, 'master.style', 'search', [('name', '=', row[0].value)])
    # print "===========", master_style
    if master_style:
        style_product = sock.execute(dbname, uid, pwd, 'style.product', 'search', [('product_id', '=', master_style[0]), ('is_master', '=', False)])
        for style in style_product:
            product_id = sock.execute(dbname, uid, pwd, 'style.product', 'search', [('id', '=', style)])
            jewel_code_id = sock.execute(dbname, uid, pwd, 'product.jewel.code', 'search', [('name', '=', row[1].value), ('code', '=', row[2].value)])
            metal_id = sock.execute(dbname, uid, pwd, 'product.type', 'search', [('name', '=', row[7].value)])
            partner_id = sock.execute(dbname, uid, pwd, 'res.partner', 'search', [('name', '=', row[5].value)])
            if not partner_id:
                partner_id = ""
            else:
                partner_id = partner_id[0]
            variant_detail = {
                'variants_jewel_code_id': jewel_code_id[0],
                'variants_metal_type_id': metal_id[0],
                'variants_customer_id': partner_id,
                'variants_size': row[9].value,
                'variants_length_mm': row[14].value,
                'variants_width_mm': row[15].value,
                'variants_height_mm': row[16].value,
                'variants_length_inch': row[17].value,
                'variants_width_inch': row[18].value,
                'variants_height_inch': row[19].value,
                'variants_shrank_width_mm': row[21].value,
                'variants_shrank_width_inch': row[22].value,
                'unit_of_measure': 1
            }
            print "======variant_detail====>", product_id[0], sock.execute(dbname, uid, pwd, 'style.product', 'write', product_id[0], variant_detail)


workbook = xlrd.open_workbook('stylemaster 010216.xlsx')
worksheet = workbook.sheet_by_name('Metal Weight')

num_rows = worksheet.nrows - 1
num_cells = worksheet.ncols - 1
curr_row = 0
while curr_row < num_rows:
    curr_row += 1
    row = worksheet.row(curr_row)
    #  add metal in master style line
    print "===metal_ids=row=====>", curr_row
    metal_ids = sock.execute(dbname, uid, pwd, 'style.product', 'search', [('name', '=', row[1].value)])
    if not metal_ids:
        metal = {
            'name': row[1].value,
            'sale_purchase_selection': 'purchase',
            'category': 'metal',
            'unit_of_measure': 2,
        }
        print "=========metal=======>", sock.execute(dbname, uid, pwd, 'style.product', 'create', metal)
    master_style_id = sock.execute(dbname, uid, pwd, 'master.style', 'search', [('name', '=', row[0].value)])
    product_id = sock.execute(dbname, uid, pwd, 'style.product', 'search', [('name', '=', row[1].value), ('category', '=', 'metal')])
    if not master_style_id:
        master_style_id = ""
    else:
        master_style_id = master_style_id[0]
    metal_ids = sock.execute(dbname, uid, pwd, 'product.metal', 'search', [('master_id', '=', master_style_id), ('product_id', '=', product_id[0]), ('metal_actual_wt', '=', row[4].value)])
    if not metal_ids:
        metal_ids = {
            'master_id': master_style_id,
            'product_id': product_id[0],
            'metal_qty': 1,
            'metal_uom': 2,
            'metal_actual_wt': row[4].value,
        }
    print "=======metal_ids======", sock.execute(dbname, uid, pwd, 'product.metal', 'create', metal_ids)
    if master_style_id:
        style_product = sock.execute(dbname, uid, pwd, 'style.product', 'search', [('product_id', '=', master_style_id), ('is_master', '=', False)])
        product_id = sock.execute(dbname, uid, pwd, 'style.product', 'search', [('name', '=', row[1].value), ('category', '=', 'metal')])
        for style in style_product:
            metal_ids = sock.execute(dbname, uid, pwd, 'product.metal', 'search', [('style_id', '=', style), ('product_id', '=', product_id[0]), ('metal_actual_wt', '=', row[4].value)])
            if not metal_ids:
                # print "====style_product======>", style
                metal_ids = {
                    'style_id': style,
                    'product_id': product_id[0],
                    'metal_qty': 1,
                    'metal_uom': 2,
                    'metal_actual_wt': row[4].value,
                }
                print "=======metal_ids===in variant===", curr_row, sock.execute(dbname, uid, pwd, 'product.metal', 'create', metal_ids)

print "metal weight uploaded Successfully...!"

workbook = xlrd.open_workbook('stylemaster 010216.xlsx')
worksheet = workbook.sheet_by_name('total Dm and color stone Wt')

num_rows = worksheet.nrows - 1
num_cells = worksheet.ncols - 1
curr_row = 0

print "Importing master style..."

while curr_row < num_rows:
    curr_row += 1
    # cnt += 1
    row = worksheet.row(curr_row)
    print "=master style====curr_row======", curr_row
    master_style_id = sock.execute(dbname, uid, pwd, 'master.style', 'search', [('name', '=', row[0].value)])
    if master_style_id:
        master_style_id = master_style_id[0]
    if master_style_id:
        stone_total = {
            'total_diamond': row[1].value,
            'total_diamond_wt': row[2].value,
            'total_stone': row[3].value,
            'total_stone_wt': row[4].value,
            'total_pcs': row[5].value,
            'total_pcs_wt': row[6].value
        }
        print "==master ==stone_total=======", sock.execute(dbname, uid, pwd, 'master.style', 'write', master_style_id, stone_total)
    if master_style_id:
        style_product = sock.execute(dbname, uid, pwd, 'style.product', 'search', [('product_id', '=', master_style_id), ('is_master', '=', False)])
        # product_id = sock.execute(dbname, uid, pwd, 'style.product', 'search', [('name', '=', row[1].value), ('category', '=', 'metal')])
        for style in style_product:

            product_id = sock.execute(dbname, uid, pwd, 'style.product', 'search', [('id', '=', style)])
            stone_total = {
                'total_diamond': row[1].value,
                'total_diamond_wt': row[2].value,
                'total_stone': row[3].value,
                'total_stone_wt': row[4].value,
                'total_pcs': row[5].value,
                'total_pcs_wt': row[6].value
                 }
            print "======variant_detail===total diamond=>", product_id[0], sock.execute(dbname, uid, pwd, 'style.product', 'write', product_id[0], stone_total)

print "total diamond and weight uploaded Successfully...!"

# workbook = xlrd.open_workbook('stylemaster 010216.xlsx')
# worksheet = workbook.sheet_by_name('Stones breakup')

# num_rows = worksheet.nrows - 1
# num_cells = worksheet.ncols - 1
# curr_row = 0

# print "Importing Stones breakup..."

# while curr_row < num_rows:
#     curr_row += 1
#     cnt += 1
#     row = worksheet.row(curr_row)
#     print "=row number.>>>>>>>>>>>>>>>", curr_row
#     if row[3].value:
#         if row[3].value == 'WH':
#             product = row[1].value + " " + row[2].value + " " + "WHITE"
#         else:
#             product = row[1].value + " " + row[2].value + " " + row[3].value
#     else:
#         product = row[1].value + " " + row[2].value + " " + "WHITE"
#     product_id = sock.execute(dbname, uid, pwd, 'style.product', 'search', [('name', '=', product)])
#     if not product_id:
#         style_product_id = {
#             'name': product,
#             'product_type': 'stockable',
#             'category': 'stone',
#             'sale_purchase_selection': 'purchase',
#             'unit_of_measure': 1,
#         }
#         print "=======Stones breakup==================", sock.execute(dbname, uid, pwd, 'style.product', 'create', style_product_id)
#     stone_seive = sock.execute(dbname, uid, pwd, 'stone.size', 'search', [('name', '=', (row[4].value).strip())])
#     if not stone_seive:
#         stone = {
#             'name': (row[4].value).strip(),
#         }
#         print "=====stone size created==>", sock.execute(dbname, uid, pwd, 'stone.size', 'create', stone)
#     master_style_id = sock.execute(dbname, uid, pwd, 'master.style', 'search', [('name', '=', row[0].value)])
#     if row[3].value:
#         if row[3].value == 'WH':
#             product = row[1].value + " " + row[2].value + " " + "WHITE"
#         else:
#             product = row[1].value + " " + row[2].value + " " + row[3].value
#     else:
#         product = row[1].value + " " + row[2].value + " " + "WHITE"

#     product_id = sock.execute(dbname, uid, pwd, 'style.product', 'search', [('name', '=', product), ('category', '=', 'stone')])
#     setting = ""
#     if row[5].value == "WSET":
#         setting = "wset"
#     if row[5].value == "MSET":
#         setting = "mset"
#     if not master_style_id:
#         master_style_id = ""
#     else:
#         master_style_id = master_style_id[0]
#     stone_seive = sock.execute(dbname, uid, pwd, 'stone.size', 'search', [('name', '=', row[4].value)])
#     if stone_seive:
#         stone_seive = stone_seive[0]
#     else:
#         stone_seive = ""
#     stone_ids = {
#          'master_id': master_style_id,
#          'product_id': product_id[0],
#          'stone_seive_size': stone_seive,
#          'stone_qty': row[6].value,
#          'stone_uom': 2,
#          'stone_wt': row[8].value,
#          'stone_ct': row[7].value,
#          'stone_setting': setting,
#      }
#     print "=======stone ids----->", sock.execute(dbname, uid, pwd, 'product.stone', 'create', stone_ids)
# print "=====================stone_ids successfully addedd=========="
print "===========all done ==============cnt============",
