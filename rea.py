#! /usr/bin/env python3

#Module is used to connect to REA server, pull information needed, and to write data to a 
#csv for the purpose of web parsing. 

import pymssql

#TO BE REPLACED MY JSON FILE IN HOME DIR
user_vars = [
    {'name':'_USER_', 'server':'_SERVER_', 'password':'_PW_', 'user':'_USER_', 'database':'_DB_'}
]

#CONSTANTS
REA_GROUP = "'Property Owner Scrape'"
TEST_GROUP = "'A-Starbucks'"


#FUNCTIONS
def add_rea_user():
    pass

def load_rea_users():
    pass

#Returns a pymssql connection specified by the list user_vars (to be updated to JSON loaded)
def get_connection(user):
    conn = pymssql.connect(host=user['server'], user=user['user'], password=user['password'], database=user['database'])
    print('\nConnected with user: ' + user['name'] + ' on ' + user['server'] + '\n')
    print('REA Data Pull...')
    return conn

#Returns the user info by their last name (the arguement to the function)
def get_user(user_last):
    for d in user_vars:
        if user_last.lower() == d['name'].lower():
            return d

#Sets the REA group to perform the SQL query on 
def set_rea_group(group_name):
    REA_GROUP = "\'" + group_name + "\'"

#Helper function to build a standard SQL function which queries the property_info table in the
#specified group name (REA_GROUP). The arguement is the property_field to be returned from the 
#property_info table. This function should only be a helper function. 
def __build_sql_query(property_field):
    '''
    return 'SELECT property_info.'+property_field+ \
           ' FROM property_info ' \
           ' JOIN object_info ON property_info.property_key = object_info.property_key' \
           ' JOIN Groups ON Groups.object_key = object_info.object_key' \
           ' WHERE Groups.name = ' + TEST_GROUP
    '''

    return 'SELECT property_info.'+property_field+ \
           ' FROM property_info ' \
           ' JOIN object_info ON property_info.property_key = object_info.property_key' \
           ' JOIN Groups ON Groups.object_key = object_info.object_key'


#This is the starter function for the above function. The property_field arguement is the arguement
#for the build_sql_querey parameter. bulid_sql_query should not be called directly, only modified as
#needed to change the implementation of the returned SQL statement verbage
def get_data(conn, property_field):
    cur = conn.cursor()
    stmt = __build_sql_query(property_field)
    cur.execute(stmt)
    data = []
    for d in cur:
        data.append(str(d[0]))
    return data

#This function updates an REA field with the data from the Excel field for a given property ID
def update_field(conn, prop_id, excel_field, rea_field):

    cur = conn.cursor()

    stmt = 'UPDATE property_info' \
           ' SET ' + rea_field + ' = \'' + str(excel_field) + '\''\
           ' WHERE property_key = \'' + str(prop_id) + '\''

    cur.execute(stmt)
    conn.commit()
    print('Updating REA field ' + rea_field + ' with ' + excel_field)









