#! /usr/bin/env python3

import pymssql
import openpyxl
import os.path

#TO BE REPLACED MY JSON FILE IN HOME DIR
user_vars = [
    {'name':'Hopkins', 'server':'jhopkins1-win7\\rea9', 'password':'jackputt', 'user':'sa', 'database':'REA9'}
]

#CONSTANTS
REA_GROUP = "'Property Owner Scrape'"
TEST_GROUP = "'Centers - 50-125K'"


def add_rea_user():
    pass

def load_rea_users():
    pass

def get_connection(user):
    conn = pymssql.connect(host=user['server'], user=user['user'], password=user['password'], database=user['database'])
    return conn

def get_user(user_last):

    for d in user_vars:
        if user_last.lower() == d['name'].lower():
            return d

def set_rea_group(group_name):

    REA_GROUP = "\'" + group_name + "\'"

def build_sql_query(property_field):

    return 'SELECT property_info.'+property_field+ \
           ' FROM property_info ' \
           ' JOIN object_info ON property_info.property_key = object_info.property_key' \
           ' JOIN Groups ON Groups.object_key = object_info.object_key' \
           ' WHERE Groups.name = ' + TEST_GROUP

def get_data(conn, sql_field):
    cur = conn.cursor()
    stmt = build_sql_query(sql_field)
    cur.execute(stmt)
    data = []
    for d in cur:
        data.append(d[0])
    return data

def load_col(workbook, data_array, file_path, col_loc, header_name):
    print(file_path, col_loc, header_name)

def append_outfile(data_array, file_path, col_loc, header_name):
    
    if os.path.isfile(file_path):
        wb = openpyxl.load_workbook(file_path)
        load_col(wb, data_array, file_path, col_loc, header_name)
    else:
        wb = openpyxl.Workbook()
        load_col(wb, data_array, file_path, col_loc, header_name)





