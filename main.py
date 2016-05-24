#! /usr/bin/env python3

import rea
import prospect
import data
import lexis_2
import os

from Property import Property
from apscheduler.schedulers.blocking import BlockingScheduler


def run_script():
    
    #VARS - to be updated later..only for prototyping 
    #Path to .xlsx file build 
    file_path = '/users/bryan/desktop/rea_holder.xlsx'

    #Write structure for REA to Excel mapping (column letter : field)
    property_write_fields = [
        ['A', 'property_key'],
        ['B', 'property_name'],
        ['C', 'address_1'],
        ['D', 'city'], 
        ['E', 'state'], 
        ['F', 'zip_code']
    ]


    #Add control to fork program depending on provided list or REA pull?

    #Get user and SQL server connection to REA database. Comment out these lines to run a custom sheat
    user = rea.get_user('Hopkins')
    conn = rea.get_connection(user)

    #Write initial .xlsx file structure according to var property_write_fields 
    for pair in property_write_fields:
        data_ = rea.get_data(conn, pair[1])
        data.create_initial_outfile(data_, file_path, pair[0], pair[1])

    conn.close()

    #Retrieve necessary data for prospect now scrape (format: address_1, city, state, zip_code)
    address_list = data.get_column_data('address_1', file_path)
    city_list = data.get_column_data('city', file_path)
    state_list = data.get_column_data('state', file_path)
    zip_code_list = data.get_column_data('zip_code', file_path)

    #Start prospectnow crawl and assign property data to prop_objects. prop_objects is a nested
    #dictionary that contains property data at the key 'data_packet'.  
    #address_list = data.build_address_list(address_list, city_list, state_list, zip_code_list)
    address_list = data.build_random_address_list(address_list, city_list, state_list, zip_code_list, 20)
    prop_objects = prospect.get_html('jack.hopkins@marcusmillichap.com','jackputt',address_list)
    data_packet = []
    data_packet_prop = []

    #Load property objects from prospectnow (which are dictionaries) into single lists 
    for d in prop_objects[1:len(prop_objects)]:
        data_packet = data_packet + d['data_packet']

    for d in prop_objects[1:len(prop_objects)]:
        data_packet_prop.append(d['data_packet_prop'])

    #Load property class list
    property_list = []
    for d in data_packet_prop:
        for key, value in d.items():
            prop = Property(key)
            prop.set_property_address(value['address'], value['city'], value['county_state_code'], value['zip_code'])
            
            #"Maybe" fields for property data
            try:
                prop.loan_amount = value['loanamount']
            except:
                pass

            try:
                prop.purchase_price = value['purchase_price']
            except:
                pass

            try:
                prop.purchase_date = value['date_purchased']
            except:
                pass

            for d in data_packet:
                if d['id'] == key:
                    prop.set_owner_address(d['owner_address'], d['owner_city'], d['state_code'], d['owner_zip_code'])
                    prop.set_owner_name(d['owner_firstname'], d['owner_lastname'])

                if prop.owner_last != '':
                    prop.owner_is_business = False

            property_list.append(prop)

    #Login to LexisNexis
    lexis_2.login('Bryan', 'Wheeler', 'bryan.wheeler@marcusmillichap.com', 'Newport Beach')

    #List of dictionaries storing property data
    prop_data_list = []

    #Write property data to .xlsx
    for prop in property_list: 

        #Match property address from ProspectNow with REA address (Not all REA addresses return a property) 
        row = data.search_rows(file_path, prop.property_address, 'C')
        
        #If there is a match, write adjacent columns with the data and get Lexis data to append to prop object
        if row:
            
            #Get Lexis data and append to prop
            try:
                lexis_dict = lexis_2.run_crawl(prop)
                prop.set_lexis_data(lexis_dict)
                if prop.owner_is_business is True:
                    prop.owner_entity = prop.owner_first
                    prop.owner_first = ''
            except:
                #ADDED 5/24
                property_dict = {}

                property_dict.update({'property_key': data.get_property_key(file_path, row)})
                
                if prop.owner_is_business is True:
                    print('Writing ' + prop.owner_entity)
                    data.write_cell(file_path, row, 7, prop.owner_entity, 'owner_entity')
                    property_dict.update({'owner_entity': prop.owner_entity})
                else:
                    print('Writing ' + prop.owner_last + ', ' + prop.owner_first)
                    data.write_cell(file_path, row, 7, '', 'owner_entity')
                    property_dict.update({'owner_entity': ''})

                data.write_cell(file_path, row, 8, prop.owner_last, 'owner_last')
                data.write_cell(file_path, row, 9, prop.owner_first, 'owner_first')
                property_dict.update({'owner_fullname': prop.owner_last + ', ' + prop.owner_first})
                data.write_cell(file_path, row, 10, prop.owner_address, 'owner_address')
                property_dict.update({'owner_address': prop.owner_address})
                data.write_cell(file_path, row, 11, prop.owner_city, 'owner_city')
                property_dict.update({'owner_city': prop.owner_city})
                data.write_cell(file_path, row, 12, prop.owner_state, 'owner_state')
                property_dict.update({'owner_state': prop.owner_state})
                data.write_cell(file_path, row, 13, prop.owner_zip, 'owner_zip_code')
                property_dict.update({'owner_zip': prop.owner_zip})
                data.write_cell(file_path, row, 14, '', 'lexis_phones')
                property_dict.update({'lexis_phones': ''})
                data.write_cell(file_path, row, 15, '', 'lexis_emails')
                property_dict.update({'lexis_emails': ''})
                data.write_cell(file_path, row, 16, '', 'lexis_name')
                property_dict.update({'lexis_name': ''})
                data.write_cell(file_path, row, 17, prop.loan_amount, 'loan_amount')
                property_dict.update({'loan_amount': prop.loan_amount})
                data.write_cell(file_path, row, 18, prop.purchase_price, 'purchase_price')
                property_dict.update({'purchase_price': prop.purchase_price})
                data.write_cell(file_path, row, 19, prop.purchase_date, 'date_purchased')
                property_dict.update({'purchase_date': prop.purchase_date})

                prop_data_list.append(property_dict)
                #END 5/24 ADD

                lexis_2.reset()
                continue


            #Write out .xlsx file and append_prop_data_list
            
            property_dict = {}

            property_dict.update({'property_key': data.get_property_key(file_path, row)})
            
            if prop.owner_is_business is True:
                print('Writing ' + prop.owner_entity)
                data.write_cell(file_path, row, 7, prop.owner_entity, 'owner_entity')
                property_dict.update({'owner_entity': prop.owner_entity})
            else:
                print('Writing ' + prop.owner_last + ', ' + prop.owner_first)
                data.write_cell(file_path, row, 7, '', 'owner_entity')
                property_dict.update({'owner_entity': ''})

            data.write_cell(file_path, row, 8, prop.owner_last, 'owner_last')
            data.write_cell(file_path, row, 9, prop.owner_first, 'owner_first')
            property_dict.update({'owner_fullname': prop.owner_last + ', ' + prop.owner_first})
            data.write_cell(file_path, row, 10, prop.owner_address, 'owner_address')
            property_dict.update({'owner_address': prop.owner_address})
            data.write_cell(file_path, row, 11, prop.owner_city, 'owner_city')
            property_dict.update({'owner_city': prop.owner_city})
            data.write_cell(file_path, row, 12, prop.owner_state, 'owner_state')
            property_dict.update({'owner_state': prop.owner_state})
            data.write_cell(file_path, row, 13, prop.owner_zip, 'owner_zip_code')
            property_dict.update({'owner_zip': prop.owner_zip})
            data.write_cell(file_path, row, 14, prop.lexis_phones, 'lexis_phones')
            property_dict.update({'lexis_phones': prop.lexis_phones})
            data.write_cell(file_path, row, 15, prop.lexis_emails, 'lexis_emails')
            property_dict.update({'lexis_emails': prop.lexis_emails})
            data.write_cell(file_path, row, 16, prop.lexis_name, 'lexis_name')
            property_dict.update({'lexis_name': prop.lexis_name})
            data.write_cell(file_path, row, 17, prop.loan_amount, 'loan_amount')
            property_dict.update({'loan_amount': prop.loan_amount})
            data.write_cell(file_path, row, 18, prop.purchase_price, 'purchase_price')
            property_dict.update({'purchase_price': prop.purchase_price})
            data.write_cell(file_path, row, 19, prop.purchase_date, 'date_purchased')
            property_dict.update({'purchase_date': prop.purchase_date})

            prop_data_list.append(property_dict)

    #connection for REA update
    conn = rea.get_connection(user)

    #Loop and update REA fields
    for data_dict in prop_data_list:
        rea.update_field(conn, data_dict['property_key'], data_dict['owner_entity'], 'user_1')
        rea.update_field(conn, data_dict['property_key'], data_dict['lexis_name'], 'user_2')
        rea.update_field(conn, data_dict['property_key'], data_dict['owner_fullname'], 'user_3')
        rea.update_field(conn, data_dict['property_key'], data_dict['owner_address'], 'user_4')
        rea.update_field(conn, data_dict['property_key'], data_dict['owner_city'], 'user_5')
        rea.update_field(conn, data_dict['property_key'], data_dict['owner_state'], 'user_6')
        rea.update_field(conn, data_dict['property_key'], data_dict['owner_zip'], 'user_7')
        rea.update_field(conn, data_dict['property_key'], data_dict['lexis_phones'], 'user_8')
        rea.update_field(conn, data_dict['property_key'], data_dict['loan_amount'], 'user_9')
        rea.update_field(conn, data_dict['property_key'], data_dict['purchase_price'], 'user_10')
        rea.update_field(conn, data_dict['property_key'], data_dict['purchase_date'], 'user_11')
        rea.update_field(conn, data_dict['property_key'], data_dict['lexis_emails'], 'usermulti')

    conn.close()

    lexis_2.quit()

    os.remove(file_path)





def run_sheet_only():
#######################################################################################################################
#RUN PROPERTIES ALREADY CONTAINED IN AN EXCEL FILE IN THE ORDER "ADDRESS, CITY, STATE, ZIP CODE"


    #VARS - to be updated later..only for prototyping 
    #Path to .xlsx file build 
    file_path = '/users/bryan/desktop/prospect_holder.xlsx'

#Retrieve necessary data for prospect now scrape (format: address_1, city, state, zip_code)
    address_list = data.get_column_data('address_1', file_path)
    city_list = data.get_column_data('city', file_path)
    state_list = data.get_column_data('state', file_path)
    zip_code_list = data.get_column_data('zip_code', file_path)

    #Start prospectnow crawl and assign property data to prop_objects. prop_objects is a nested
    #dictionary that contains property data at the key 'data_packet'.  
    #address_list = data.build_address_list(address_list, city_list, state_list, zip_code_list)
    address_list = data.build_address_list(address_list, city_list, state_list, zip_code_list)
    prop_objects = prospect.get_html('jack.hopkins@marcusmillichap.com','jackputt',address_list)
    data_packet = []
    data_packet_prop = []

    #Load property objects from prospectnow (which are dictionaries) into single lists 
    for d in prop_objects[1:len(prop_objects)]:
        data_packet = data_packet + d['data_packet']

    for d in prop_objects[1:len(prop_objects)]:
        data_packet_prop.append(d['data_packet_prop'])

    #Load property class list
    property_list = []
    for d in data_packet_prop:
        for key, value in d.items():
            prop = Property(key)
            prop.set_property_address(value['address'], value['city'], value['county_state_code'], value['zip_code'])
            
            #"Maybe" fields for property data
            try:
                prop.loan_amount = value['loanamount']
            except:
                pass

            try:
                prop.purchase_price = value['purchase_price']
            except:
                pass

            try:
                prop.purchase_date = value['date_purchased']
            except:
                pass

            for d in data_packet:
                if d['id'] == key:
                    prop.set_owner_address(d['owner_address'], d['owner_city'], d['state_code'], d['owner_zip_code'])
                    prop.set_owner_name(d['owner_firstname'], d['owner_lastname'])

                if prop.owner_last != '':
                    prop.owner_is_business = False

            property_list.append(prop)

    #Login to LexisNexis
    lexis_2.login('Bryan', 'Wheeler', 'bryan.wheeler@marcusmillichap.com', 'Newport Beach')

    #List of dictionaries storing property data
    prop_data_list = []

    #Write property data to .xlsx
    for prop in property_list: 

        #Match property address from ProspectNow with REA address (Not all REA addresses return a property) 
        row = data.search_rows(file_path, prop.property_address, 'A')
        
        #If there is a match, write adjacent columns with the data and get Lexis data to append to prop object
        if row:
            
            #Get Lexis data and append to prop
            try:
                lexis_dict = lexis_2.run_crawl(prop)
                prop.set_lexis_data(lexis_dict)
                if prop.owner_is_business is True:
                    prop.owner_entity = prop.owner_first
                    prop.owner_first = ''
            except:
                data.write_cell(file_path, row, 7, prop.owner_entity, 'owner_entity')
                data.write_cell(file_path, row, 8, prop.owner_last, 'owner_last')
                data.write_cell(file_path, row, 9, prop.owner_first, 'owner_first')
                data.write_cell(file_path, row, 10, prop.owner_address, 'owner_address')
                data.write_cell(file_path, row, 11, prop.owner_city, 'owner_city')
                data.write_cell(file_path, row, 12, prop.owner_state, 'owner_state')
                data.write_cell(file_path, row, 13, prop.owner_zip, 'owner_zip_code')
                data.write_cell(file_path, row, 17, prop.loan_amount, 'loan_amount')
                data.write_cell(file_path, row, 18, prop.purchase_price, 'purchase_price')
                data.write_cell(file_path, row, 19, prop.purchase_date, 'date_purchased')
                lexis_2.reset()
                continue


            #Write out .xlsx file and append_prop_data_list
            
            property_dict = {}

            property_dict.update({'property_key': data.get_property_key(file_path, row)})
            
            if prop.owner_is_business is True:
                print('Writing ' + prop.owner_entity)
                data.write_cell(file_path, row, 7, prop.owner_entity, 'owner_entity')
                property_dict.update({'owner_entity': prop.owner_entity})
            else:
                print('Writing ' + prop.owner_last + ', ' + prop.owner_first)
                data.write_cell(file_path, row, 7, '', 'owner_entity')
                property_dict.update({'owner_entity': ''})

            data.write_cell(file_path, row, 8, prop.owner_last, 'owner_last')
            data.write_cell(file_path, row, 9, prop.owner_first, 'owner_first')
            property_dict.update({'owner_fullname': prop.owner_last + ', ' + prop.owner_first})
            data.write_cell(file_path, row, 10, prop.owner_address, 'owner_address')
            property_dict.update({'owner_address': prop.owner_address})
            data.write_cell(file_path, row, 11, prop.owner_city, 'owner_city')
            property_dict.update({'owner_city': prop.owner_city})
            data.write_cell(file_path, row, 12, prop.owner_state, 'owner_state')
            property_dict.update({'owner_state': prop.owner_state})
            data.write_cell(file_path, row, 13, prop.owner_zip, 'owner_zip_code')
            property_dict.update({'owner_zip': prop.owner_zip})
            data.write_cell(file_path, row, 14, prop.lexis_phones, 'lexis_phones')
            property_dict.update({'lexis_phones': prop.lexis_phones})
            data.write_cell(file_path, row, 15, prop.lexis_emails, 'lexis_emails')
            property_dict.update({'lexis_emails': prop.lexis_emails})
            data.write_cell(file_path, row, 16, prop.lexis_name, 'lexis_name')
            property_dict.update({'lexis_name': prop.lexis_name})
            data.write_cell(file_path, row, 17, prop.loan_amount, 'loan_amount')
            property_dict.update({'loan_amount': prop.loan_amount})
            data.write_cell(file_path, row, 18, prop.purchase_price, 'purchase_price')
            property_dict.update({'purchase_price': prop.purchase_price})
            data.write_cell(file_path, row, 19, prop.purchase_date, 'date_purchased')
            property_dict.update({'purchase_date': prop.purchase_date})

            prop_data_list.append(property_dict)



def main():

    #run_script()
    #sched = BlockingScheduler()
    #sched.add_job(run_script, 'interval', minutes=30)
    #sched.start()

    run_sheet_only()



if __name__ == "__main__":
    main()