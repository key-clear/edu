#!/usr/bin/env python
# coding: utf-8

import xmlrpclib
import xlrd

# database name as per changes
dbname = 'fusion_2_15'
# dbname = 'jc16'
username = 'admin'
pwd = 'admin'
# server address
server = 'http://161.202.180.220:9092'
server = 'http://127.0.0.1:8080'

sock_common = xmlrpclib.ServerProxy(server + '/xmlrpc/common')
sock = xmlrpclib.ServerProxy(server + '/xmlrpc/object')
uid = sock_common.login(dbname, username, pwd)


workbook = xlrd.open_workbook('style stone 090216 final.xlsx')
worksheet = workbook.sheet_by_name('Stlstn')

num_rows = worksheet.nrows - 1
num_cells = worksheet.ncols - 1
curr_row = 0

print "Importing master style..."


def add_stone(row, stone):
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

    product_id = sock.execute(dbname, uid, pwd, 'style.product', 'search', [('name', '=', stone), ('category', '=', 'stone')])
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

while curr_row < num_rows:
    curr_row += 1
    row = worksheet.row(curr_row)
    if row[1].value == "DM":
        if row[4].value:
            if row[4].value in ['00000-0000', '-000000', '-00000']:
                stone = row[1].value + " " + row[2].value + " " + row[3].value + " " + "+00000-0000"
                add_stone(row, stone)
                # print "1==>", curr_row, row[4].value, stone
            elif row[4].value in ['0000-000', '-000']:
                stone = row[1].value + " " + row[2].value + " " + row[3].value + " " + "+0000-000"
                add_stone(row, stone)
                # print "2==>", curr_row, row[4].value, stone
            elif row[4].value in ['+00', '+000', '+00-0', '+000-00', '-00', '000-00']:
                stone = row[1].value + " " + row[2].value + " " + row[3].value + " " + "+000-00"
                add_stone(row, stone)
                # print "3==>", curr_row, row[4].value, stone
            elif row[4].value in ['+0', '+0-1']:
                stone = row[1].value + " " + row[2].value + " " + row[3].value + " " + "+0-1"
                add_stone(row, stone)
                # print "4==>", curr_row, row[4].value, stone
            elif row[4].value in ['1.5-2', '1-1.5']:
                stone = row[1].value + " " + row[2].value + " " + row[3].value + " " + "1.5-2"
                add_stone(row, stone)
                # print "5==>", curr_row, row[4].value, stone
            elif row[4].value in ['2.5-3', '2-2.5', '3.5-4', '3-3.4', '3-3.5']:
                stone = row[1].value + " " + row[2].value + " " + row[3].value + " " + "2.5-3"
                add_stone(row, stone)
                # print "6==>", curr_row, row[4].value, stone
            elif row[4].value in ['4.5-5', '4-4.5']:
                stone = row[1].value + " " + row[2].value + " " + row[3].value + " " + "2.5-3"
                add_stone(row, stone)
                # print "7==>", curr_row, row[4].value, stone
            elif row[4].value in ['5.5-6', '5-5.5', '6-6.5']:
                stone = row[1].value + " " + row[2].value + " " + row[3].value + " " + "5.5-6"
                add_stone(row, stone)
                # print "8==>", curr_row, row[4].value, stone
            elif row[4].value in ['6.5-7', '7.5-8', '7-7.5']:
                stone = row[1].value + " " + row[2].value + " " + row[3].value + " " + "6.5-7"
                add_stone(row, stone)
                # print "9==>", curr_row, row[4].value, stone
            elif row[4].value in ['8.5-9', '8-8.5', '9.5-10', '9-9.5']:
                stone = row[1].value + " " + row[2].value + " " + row[3].value + " " + "8.5-9"
                add_stone(row, stone)
                # print "10==>", curr_row, row[4].value, stone
            elif row[4].value in ['10.5-11', '10-10.5', '11.5-12', '11-11.5']:
                stone = row[1].value + " " + row[2].value + " " + row[3].value + " " + "10.5-11"
                add_stone(row, stone)
                # print "11==>", curr_row, row[4].value, stone
            elif row[4].value in ['12.5-13', '12-12.5', '13.5-14', '13-13.5', '14.5-15', '14-14.5']:
                stone = row[1].value + " " + row[2].value + " " + row[3].value + " " + "12.5-13"
                add_stone(row, stone)
                # print "12==>", curr_row, row[4].value, stone
            elif row[4].value in ['15.5-16', '15-15.5', '16.5-17', '16-16.5', '17.5-18', '17-17.5', '18.5-19', '18-18.5', '19.5-20', '19-19.5', '0.17 CTS', '0.20 CTS']:
                stone = row[1].value + " " + row[2].value + " " + row[3].value + " " + "15.5-16"
                add_stone(row, stone)
                # print "13==>", curr_row, row[4].value, stone
            elif row[4].value in ['0.22 CTS', '0.25 CTS']:
                stone = row[1].value + " " + row[2].value + " " + row[3].value + " " + "0.22 CTS"
                add_stone(row, stone)
                # print "14==>", curr_row, row[4].value, stone
            elif row[4].value in ['0.30 CTS']:
                stone = row[1].value + " " + row[2].value + " " + row[3].value + " " + "0.30 CTS"
                add_stone(row, stone)
                # print "15==>", curr_row, row[4].value, stone
            elif row[4].value in ['0.33 CTS', '0.35 CTS', '0.38 CTS', '0.40 CTS', '0.45 CTS']:
                stone = row[1].value + " " + row[2].value + " " + row[3].value + " " + "0.33 CTS"
                add_stone(row, stone)
                # print "16==>", curr_row, row[4].value, stone
            elif row[4].value in ['0.50 CTS', '0.55 CTS', '0.60 CTS']:
                stone = row[1].value + " " + row[2].value + " " + row[3].value + " " + "0.50 CTS"
                add_stone(row, stone)
                # print "17==>", curr_row, row[4].value, stone
            elif row[4].value in ['0.65 CTS', '0.70 CTS']:
                stone = row[1].value + " " + row[2].value + " " + row[3].value + " " + "0.65 CTS"
                add_stone(row, stone)
                # print "18==>", curr_row, row[4].value, stone
            elif row[4].value in ['0.75 CTS', '0.78 CTS', '0.80 CTS']:
                stone = row[1].value + " " + row[2].value + " " + row[3].value + " " + "0.75 CTS"
                add_stone(row, stone)
                # print "19==>", curr_row, row[4].value, stone
            elif row[4].value in ['0.85 CTS', '0.90 CTS']:
                stone = row[1].value + " " + row[2].value + " " + row[3].value + " " + "0.85 CTS"
                add_stone(row, stone)
                # print "20==>", curr_row, row[4].value, stone
            elif row[4].value in ['1.25 CTS', '1.40 CTS', '1.50 CTS', '1.75 CTS', '1.80 CTS', '1.90 CTS', '2.00 CTS', '2.50 CTS', '3.00 CTS']:
                stone = row[1].value + " " + row[2].value + " " + row[3].value + " " + "1.25 CTS"
                add_stone(row, stone)
                # print "21==>", curr_row, row[4].value, stone
            elif row[4].value in ['0.90 MM', '-1.0MM']:
                stone = row[1].value + " " + row[2].value + " " + row[3].value + " " + "0.9 MM"
                add_stone(row, stone)
                # print "22==>", curr_row, row[4].value, stone
            elif row[4].value in ['1.0MM']:
                stone = row[1].value + " " + row[2].value + " " + row[3].value + " " + "1.0 MM"
                add_stone(row, stone)
                # print "23==>", curr_row, row[4].value, stone
            elif row[4].value in ['1.10 MM', '1.10 MM-1.15 MM', '1.15 MM']:
                stone = row[1].value + " " + row[2].value + " " + row[3].value + " " + "1.1 MM"
                add_stone(row, stone)
                # print "24==>", curr_row, row[4].value, stone
            elif row[4].value in ['1.20 MM', '1.20 MM-1.25 MM', '1.25 MM-1.30 MM']:
                stone = row[1].value + " " + row[2].value + " " + row[3].value + " " + "1.2 MM"
                add_stone(row, stone)
                # print "25==>", curr_row, row[4].value, stone
            elif row[4].value in ['1.30 MM', '1.30 MM-1.35 MM', '1.35 MM-1.40 MM']:
                stone = row[1].value + " " + row[2].value + " " + row[3].value + " " + "1.3 MM"
                add_stone(row, stone)
                # print "26==>", curr_row, row[4].value, stone
            elif row[4].value in ['1.40 MM', '1.45 MM']:
                stone = row[1].value + " " + row[2].value + " " + row[3].value + " " + "1.4 MM"
                add_stone(row, stone)
                # print "27==>", curr_row, row[4].value, stone
            elif row[4].value in ['1.50 MM', '1.5 MM-1.2 MM']:
                stone = row[1].value + " " + row[2].value + " " + row[3].value + " " + "1.5 MM"
                add_stone(row, stone)
                # print "28==>", curr_row, row[4].value, stone
            elif row[4].value in ['1.60 MM', '1.65 MM']:
                stone = row[1].value + " " + row[2].value + " " + row[3].value + " " + "1.6 MM"
                add_stone(row, stone)
                # print "29==>", curr_row, row[4].value, stone
            elif row[4].value in ['1.70 MM', '1.75 MM', '1.70 MM-1.80 MM']:
                stone = row[1].value + " " + row[2].value + " " + row[3].value + " " + "1.7 MM"
                add_stone(row, stone)
                # print "30==>", curr_row, row[4].value, stone
            elif row[4].value in ['1.80 MM', '1.80 MM-1.90 MM']:
                stone = row[1].value + " " + row[2].value + " " + row[3].value + " " + "1.8 MM"
                add_stone(row, stone)
                # print "31==>", curr_row, row[4].value, stone
            elif row[4].value in ['1.90 MM', '1.90 MM-2.00 MM']:
                stone = row[1].value + " " + row[2].value + " " + row[3].value + " " + "1.9 MM"
                add_stone(row, stone)
                # print "32==>", curr_row, row[4].value, stone
            elif row[4].value in ['2.00 MM', '2*1', '2*1.5']:
                stone = row[1].value + " " + row[2].value + " " + row[3].value + " " + "2.0 MM"
                add_stone(row, stone)
                # print "33==>", curr_row, row[4].value, stone
            elif row[4].value in ['2.1 MM', '2.1*1.3*9', '2.1*1.5', '2.10 MM']:
                stone = row[1].value + " " + row[2].value + " " + row[3].value + " " + "2.1 MM"
                add_stone(row, stone)
                # print "34==>", curr_row, row[4].value, stone
            elif row[4].value in ['2.20 MM', '2.2*1.7', '2.25 MM']:
                stone = row[1].value + " " + row[2].value + " " + row[3].value + " " + "2.2 MM"
                add_stone(row, stone)
                # print "35==>", curr_row, row[4].value, stone
            elif row[4].value in ['2.30 MM', '2.3*1.25', '2.3*1.2', '2.3*1.0', '2.3*1']:
                stone = row[1].value + " " + row[2].value + " " + row[3].value + " " + "2.3 MM"
                add_stone(row, stone)
                # print "36==>", curr_row, row[4].value, stone
            elif row[4].value in ['2.4*1.3', '2.4*1.4*1', '2.4*1.5', '2.40 MM']:
                stone = row[1].value + " " + row[2].value + " " + row[3].value + " " + "2.4 MM"
                add_stone(row, stone)
                # print "37==>", curr_row, row[4].value, stone
            elif row[4].value in ['2.5*1.25', '2.5*1.5', '2.5*2', '2.50 MM']:
                stone = row[1].value + " " + row[2].value + " " + row[3].value + " " + "2.5 MM"
                add_stone(row, stone)
                # print "38==>", curr_row, row[4].value, stone
            elif row[4].value in ['2.6*1.2', '2.6*1.25', ' 2.6*1.3', '2.60 MM', '2.6*1.3*1']:
                stone = row[1].value + " " + row[2].value + " " + row[3].value + " " + "2.6 MM"
                add_stone(row, stone)
                # print "39==>", curr_row, row[4].value, stone
            elif row[4].value in ['2.70 MM', '2.75 MM', '2.75*1.5']:
                stone = row[1].value + " " + row[2].value + " " + row[3].value + " " + "2.7 MM"
                add_stone(row, stone)
                # print "40==>", curr_row, row[4].value, stone
            elif row[4].value in ['2.80 MM', '2.8 MM', '2.8*1.3']:
                stone = row[1].value + " " + row[2].value + " " + row[3].value + " " + "2.8 MM"
                add_stone(row, stone)
                # print "41==>", curr_row, row[4].value, stone
            elif row[4].value in ['2.90 MM', '2.9 MM']:
                stone = row[1].value + " " + row[2].value + " " + row[3].value + " " + "2.9 MM"
                add_stone(row, stone)
                # print "42==>", curr_row, row[4].value, stone
            elif row[4].value in ['3*1', '3*1.25', '3*1.3', '3*1.4*1', '3*1.41*', '3*2*1.5', '3.00 MM', '3.0 MM']:
                stone = row[1].value + " " + row[2].value + " " + row[3].value + " " + "3.0 MM"
                add_stone(row, stone)
                # print "43==>", curr_row, row[4].value, stone
            elif row[4].value in ['3.10 MM', '3.0 MM']:
                stone = row[1].value + " " + row[2].value + " " + row[3].value + " " + "3.1 MM"
                add_stone(row, stone)
                # print "44==>", curr_row, row[4].value, stone
            elif row[4].value in ['3.2*1.25', '3.25 MM', '3.2*1.5', '3.20 MM', '3.2 MM']:
                stone = row[1].value + " " + row[2].value + " " + row[3].value + " " + "3.2 MM"
                add_stone(row, stone)
                # print "45==>", curr_row, row[4].value, stone
            elif row[4].value in ['3.40 MM', '3.4 MM']:
                stone = row[1].value + " " + row[2].value + " " + row[3].value + " " + "3.4 MM"
                add_stone(row, stone)
                # print "46==>", curr_row, row[4].value, stone
            elif row[4].value in ['3.50 MM', '3.5 MM', '3.5*1.5', '3.5*2', '3.5*2.5']:
                stone = row[1].value + " " + row[2].value + " " + row[3].value + " " + "3.5 MM"
                add_stone(row, stone)
                # print "47==>", curr_row, row[4].value, stone
            elif row[4].value in ['3.60 MM', '3.6 MM']:
                stone = row[1].value + " " + row[2].value + " " + row[3].value + " " + "3.6 MM"
                add_stone(row, stone)
                # print "48==>", curr_row, row[4].value, stone
            elif row[4].value in ['3.7 MM', '3.75 MM']:
                stone = row[1].value + " " + row[2].value + " " + row[3].value + " " + "3.7 MM"
                add_stone(row, stone)
                # print "49==>", curr_row, row[4].value, stone
            elif row[4].value in ['4.00 MM', '4*2*1.5', '4*2']:
                stone = row[1].value + " " + row[2].value + " " + row[3].value + " " + "4.0 MM"
                add_stone(row, stone)
                # print "50==>", curr_row, row[4].value, stone
            elif row[4].value in ['4.30 MM', '4.3*2', '4.3 MM']:
                stone = row[1].value + " " + row[2].value + " " + row[3].value + " " + "4.3 MM"
                add_stone(row, stone)
                # print "51==>", curr_row, row[4].value, stone
            elif row[4].value in ['4.50 MM', '4.5*1.5', '4.5*2', '4.5 MM']:
                stone = row[1].value + " " + row[2].value + " " + row[3].value + " " + "4.5 MM"
                add_stone(row, stone)
                # print "52==>", curr_row, row[4].value, stone
            elif row[4].value in ['4.60 MM', '4.6 MM']:
                stone = row[1].value + " " + row[2].value + " " + row[3].value + " " + "4.6 MM"
                add_stone(row, stone)
                # print "53==>", curr_row, row[4].value, stone
