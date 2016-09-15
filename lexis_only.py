#! /usr/bin/env python3

import data
import lexis_2
from Property import Property

def main():
    file_path = '/PATH_TO_PROPERTY_LIST'

    prop_list = []

    '''
    USE REA9

    SELECT contact_info.contact_key, contact_info.last_name, contact_info.first_name, address_info.address_1, address_info.city, address_info.state, address_info.zip_code
    FROM contact_info
    JOIN object_info ON contact_info.contact_key = object_info.contact_key
    JOIN Groups ON Groups.object_key = object_info.object_key
    JOIN address_info ON address_info.contact_key = contact_info.contact_key
    WHERE Groups.name = 'Test Run'
    '''

    #Load .xlsx data
    ids = data.get_column_data('contact_key', file_path)
    addresses = data.get_column_data('address_1', file_path)
    cities = data.get_column_data('city', file_path)
    states = data.get_column_data('state', file_path)
    zips = data.get_column_data('zip_code', file_path)
    last_names = data.get_column_data('last_name', file_path)
    first_names = data.get_column_data('first_name', file_path)


    lexis_2.login('USER_FIRST', 'USER_LAST', 'LEXIS_ADDITIONAL', 'LOCATION')

    for i in range(len(addresses)):
        prop = Property(ids[i])
        prop.set_owner_address(addresses[i], cities[i], states[i], zips[i])
        prop.set_owner_name(first_names[i], last_names[i])
        prop.owner_is_business = False
        prop_list.append(prop)

    row = 2

    for prop in prop_list: 
        
        
        #If there is a match, write adjacent columns with the data and get Lexis data to append to prop object
        if row:
            
            #Get Lexis data and append to prop
            try:
                lexis_dict = lexis_2.run_crawl(prop)
                prop.set_lexis_data(lexis_dict)
            except:
                lexis_2.reset()
                continue


            #Write out .xlsx file
            print('Writing ' + prop.owner_last + ', ' + prop.owner_first)
            data.write_cell(file_path, row, 8, prop.owner_last, 'owner_last')
            data.write_cell(file_path, row, 9, prop.owner_first, 'owner_first')
            data.write_cell(file_path, row, 10, prop.owner_address, 'owner_address')
            data.write_cell(file_path, row, 11, prop.owner_city, 'owner_city')
            data.write_cell(file_path, row, 12, prop.owner_state, 'owner_state')
            data.write_cell(file_path, row, 13, prop.owner_zip, 'owner_zip_code')
            data.write_cell(file_path, row, 14, prop.lexis_phones, 'lexis_phones')
            data.write_cell(file_path, row, 15, prop.lexis_emails, 'lexis_emails')

        row = row + 1


if __name__ == "__main__":
    main()