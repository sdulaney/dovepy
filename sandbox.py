#! /usr/bin/env python3


rea_field = 'user 12'
excel_field = 'bob'
prop_id = '2438753298573289652'

stmt = 'UPDATE property_info' \
           ' SET ' + rea_field + ' = \'' + str(excel_field) + '\''\
           ' WHERE property_key = \'' + str(prop_id) + '\''

print(stmt)