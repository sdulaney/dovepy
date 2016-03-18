#! /usr/bin/env python3

import rea

def main():
    
    user = rea.get_user('Hopkins')
    conn = rea.get_connection(user)
    data = rea.get_data(conn, 'property_name')
    rea.append_outfile(data, '/users/bryan/desktop/test.xlsx', 'A', 'Property Name')

    conn.close()

if __name__ == "__main__":
    main()