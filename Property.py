#! /usr/bin/env python3

#Owner object class to store address and other usefull property information

class Property:

    property_address = ''
    property_city = ''
    property_state = ''
    property_zip = ''

    owner_address = ''
    owner_city = ''
    owner_state = ''
    owner_zip = ''

    owner_last = ''
    owner_first = ''

    lexis_phones = ''
    lexis_emails = ''
    lexis_name = ''

    loan_amount = ''

    purchase_price = ''
    purchase_date = ''

    owner_entity = ''

    owner_is_business = True 

    #State translation
    states = {
        'AK': 'Alaska',
        'AL': 'Alabama',
        'AR': 'Arkansas',
        'AS': 'American Samoa',
        'AZ': 'Arizona',
        'CA': 'California',
        'CO': 'Colorado',
        'CT': 'Connecticut',
        'DC': 'District of Columbia',
        'DE': 'Delaware',
        'FL': 'Florida',
        'GA': 'Georgia',
        'GU': 'Guam',
        'HI': 'Hawaii',
        'IA': 'Iowa',
        'ID': 'Idaho',
        'IL': 'Illinois',
        'IN': 'Indiana',
        'KS': 'Kansas',
        'KY': 'Kentucky',
        'LA': 'Louisiana',
        'MA': 'Massachusetts',
        'MD': 'Maryland',
        'ME': 'Maine',
        'MI': 'Michigan',
        'MN': 'Minnesota',
        'MO': 'Missouri',
        'MP': 'Northern Mariana Islands',
        'MS': 'Mississippi',
        'MT': 'Montana',
        'NA': 'National',
        'NC': 'North Carolina',
        'ND': 'North Dakota',
        'NE': 'Nebraska',
        'NH': 'New Hampshire',
        'NJ': 'New Jersey',
        'NM': 'New Mexico',
        'NV': 'Nevada',
        'NY': 'New York',
        'OH': 'Ohio',
        'OK': 'Oklahoma',
        'OR': 'Oregon',
        'PA': 'Pennsylvania',
        'PR': 'Puerto Rico',
        'RI': 'Rhode Island',
        'SC': 'South Carolina',
        'SD': 'South Dakota',
        'TN': 'Tennessee',
        'TX': 'Texas',
        'UT': 'Utah',
        'VA': 'Virginia',
        'VI': 'Virgin Islands',
        'VT': 'Vermont',
        'WA': 'Washington',
        'WI': 'Wisconsin',
        'WV': 'West Virginia',
        'WY': 'Wyoming'
    }

    def __init__(self, id_):
        self.id_ = id_

    def set_property_address(self, address, city, state, zip_):
        self.property_address = address
        self.property_city = city
        self.property_state = state
        self.property_zip = zip_

    def set_owner_address(self, address, city, state, zip_):
        self.owner_address = address
        self.owner_city = city
        self.owner_state = state
        self.owner_zip = zip_

    def set_owner_name(self, first, last):
        self.owner_last = last
        self.owner_first = first
        if self.owner_last != '':
            self.owner_is_business = False

    def get_street_number(self):
        full_address = self.property_address
        splitter = full_address.find(' ')
        number = full_address[0:splitter]
        rest = full_address[splitter:]
        return (number, rest)

    def set_lexis_data(self, lexis_dict):
        self.lexis_phones = lexis_dict['phones']
        self.lexis_emails = lexis_dict['emails']
        self.lexis_name = lexis_dict['name']

    def get_state(self, is_cap):
        if is_cap is True:
            state_abbr = self.owner_state
            return self.states[state_abbr].upper()
        else:
            state_abbr = self.owner_state
            return self.states[state_abbr]





        

