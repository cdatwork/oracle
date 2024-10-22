import cx_Oracle # type: ignore
import struct
import os
import glob
import keyring # type: ignore
import pandas as pd # type: ignore
import io

class Oracle:
    def __init__(self):
        self.user_id = None
        self.password = None
        self.environment = 'prod'

    def retrieve_credentials_from_keyring(self):
        '''This retrieves your password from Windows Credential Manager.

        Here's additional information about how Credentials Manager should be populated for this to work:
        # {
        #    'Internet or network address': oracle_usr@oracle',
        #    'User name': 'oracle_usr',
        #    'Password': <your_user_id_goes_here>,
        #    'Persistence': 'Enterprise'
        # }
        # {
        #    'Internet or network address': oracle',
        #    'User name': '<your_user_id_goes_here>',
        #    'Password': <your_password_goes_here>,
        #    'Persistence': 'Enterprise'
        # }
        '''
        self.user_id = keyring.get_password('oracle_user@oracle', 'oracle_usr') # must have this stored in Windows 'Credential Manager' > 'Windows Credentials' >
        self.password = keyring.get_password('oracle', self.user_id) # must have this stored in Windows 'Credential Manager' > 'Windows Credentials' >
    
    def connect(self, user_id, password, database_server_url, database_service_name, database_port='1521', verbose=True):
        '''Connects to Oracle and returns an Oracle connection to the class instance "conn" property.
        This method will determine the bitness of the current Python environment to determine how to connect.
            
        user_id = your oracle user_id
        password = your oracle password

        Set verbose = False to suppress print statements
        '''
        bits = struct.calcsize('P')*8

        if bits == 64:
            if verbose:
                print('Connecting using the 64-bit method')
            
            instantclient_paths = glob.glob(r'C:\oracle\instantclient*') + glob.glob(r'C:\oracle\product\*\*\instantclient*')
            if len(instantclient_paths):
                instantclient_paths  = ';'.join(instantclient_paths) + ';'
            os.environ['PATH'] = instantclient_paths + os.environ['PATH']
            dsn_tns = cx_Oracle.makedsn(database_server_url, database_port, service_name=database_service_name)
            conn = cx_Oracle.connect(user=user_id, password=password, dsn=dsn_tns)

            if verbose:
                print('Connected')
        
        elif bits == 32:
            if verbose:
                print('Connecting using the 32-bit method')
            
            conn = cx_Oracle.connect(user=user_id, password=password, dsn=database_service_name)

            if verbose:
                print('Connected')
        else:
            print(f'Invalid bitness detected. 32-bit or 64-bit expected. {bits} actual.')
            # raise exception
        
        self.conn = conn