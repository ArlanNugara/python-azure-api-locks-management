import requests
import json
import pandas as pd
import ast
import csv
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl import load_workbook
from datetime import datetime
import time

# Get today's date
todays_date = datetime.today().strftime('%Y-%m-%d')

def update_locks_at_scope(main_url, scope, token_header):
    get_excel = pd.read_excel('locks.xlsx', sheet_name = scope, header=0)
    print("Scope Set to ", scope)
    # Generate create csv files
    create_target_columns = get_excel[['Scope', 'Lock Name', 'Lock Level', 'Lock Notes', 'Create Locks']]
    create_drop_nan_value = create_target_columns.dropna()
    create_drop_nan_value.to_csv('create.csv', index=False, header=True, mode='a')
    # Generate delete csv files
    delete_target_columns = get_excel[['Scope', 'Lock Name', 'Lock Level', 'Lock Notes', 'Delete Locks']]
    delete_drop_nan_value = delete_target_columns.dropna()
    delete_drop_nan_value.loc[delete_drop_nan_value['Delete Locks'] == 1]
    delete_drop_nan_value.to_csv('delete.csv', index=False, header=True, mode='a')
    # Generate update csv files
    update_target_columns = get_excel[['Scope', 'Lock Name', 'Lock Level', 'Lock Notes', 'Update Locks']]
    update_drop_nan_value = update_target_columns.dropna()
    update_drop_nan_value.loc[update_drop_nan_value['Update Locks'] == 1]
    update_drop_nan_value.to_csv('update.csv', index=False, header=True, mode='a')
    # Create the master csv file
    updated_locks_csv_header = ['Scope', 'Lock Name', 'Lock Level', 'Lock Notes', 'Operation', 'Status Response']
    with open('updated_lock_details.csv', mode='w', newline='') as updated_locks_header:
        csvwriter = csv.writer(updated_locks_header, delimiter=',')
        csvwriter.writerow(updated_locks_csv_header)
    
    # Start Create / Update Lock Process
    api_files = ('create', 'update')
    for lock_file in api_files:
        with open(''+lock_file+'.csv', mode='r', newline='') as create_update_locks:
            createupdatecsvreader = csv.reader(create_update_locks, delimiter=',')
            next(createupdatecsvreader)
            print('Starting '+lock_file+' process')
            for create_update_lock_entry in createupdatecsvreader:
                create_update_lock_scope = create_update_lock_entry[0]
                create_update_lock_name = create_update_lock_entry[1]
                create_update_lock_level = create_update_lock_entry[2]
                create_update_lock_notes = create_update_lock_entry[3]
                create_update_lock_payload = {"properties": {"level": create_update_lock_level, "notes": create_update_lock_notes}}
                print('Starting '+lock_file+' process for '+create_update_lock_scope+' with lock name '+create_update_lock_name+'')
                create_update_api_request = requests.put(url=''+main_url+''+create_update_lock_scope+'/providers/Microsoft.Authorization/locks/'+create_update_lock_name+'?api-version=2016-09-01', headers = token_header, data = json.dumps(create_update_lock_payload))
                create_update_api_request_to_json = create_update_api_request.json()
                if create_update_api_request.status_code == 200 or create_update_api_request.status_code == 201:
                    print('Successfully completed '+lock_file+' locks on scope '+create_update_lock_scope+'')
                    print('\n\n\n')
                    create_update_lock_details = [create_update_lock_scope, create_update_api_request_to_json['name'], create_update_api_request_to_json['properties']['level'], create_update_api_request_to_json['properties']['notes'], lock_file, create_update_api_request.status_code]
                else:
                    print('Failed to '+lock_file+' locks at '+create_update_lock_scope+'')
                    print('\n\n\n')
                    create_update_lock_details = [create_update_lock_scope, create_update_lock_name, create_update_lock_level, create_update_api_request_to_json['error']['message'], lock_file, create_update_api_request.status_code]
                with open('updated_lock_details.csv', mode='a', newline='') as create_update_results_csv:
                    create_update_results = csv.writer(create_update_results_csv, delimiter=',')
                    create_update_results.writerow(create_update_lock_details)

    # Start Delete locks process
    with open('delete.csv', mode='r', newline='') as delete_locks:
        deletecsvreader = csv.reader(delete_locks, delimiter=',')
        next(deletecsvreader)
        print("Starting Delete Lock Process")
        for delete_lock_entry in deletecsvreader:
            delete_lock_scope = delete_lock_entry[0]
            delete_lock_name = delete_lock_entry[1]
            delete_lock_level = delete_lock_entry[2]
            delete_lock_notes = delete_lock_entry[3]
            print('Starting delete process for '+delete_lock_scope+' with lock name '+delete_lock_name+'')
            delete_locks_request = requests.delete(url = ""+main_url+""+delete_lock_scope+"/providers/Microsoft.Authorization/locks/"+delete_lock_name+"?api-version=2016-09-01", headers = token_header)
            if delete_locks_request.status_code == 200 or delete_locks_request.status_code == 204:
                delete_lock_details = [delete_lock_scope, delete_lock_name, delete_lock_level, delete_lock_notes, 'delete', delete_locks_request.status_code]
                print('Successfully completed delete locks on scope '+delete_lock_scope+'')
                print('\n\n\n')
            else:
                delete_locks_response_to_json = delete_locks_request.json()
                delete_lock_error_message = delete_locks_response_to_json['error']['message']
                delete_lock_details = [delete_lock_scope, delete_lock_name,delete_lock_level,  delete_lock_error_message, 'delete', delete_locks_request.status_code]
                print('Failed to delete locks on scope '+delete_lock_scope+'')
                print('/n/n/n')
            with open('updated_lock_details.csv', mode='a', newline='') as delete_results_csv:
                deletecsvwriter = csv.writer(delete_results_csv, delimiter=',')
                deletecsvwriter.writerow(delete_lock_details)
    
    # Update the excel file for audit data
    audit_csv_data = pd.read_csv('updated_lock_details.csv')
    audit_csv_writer = pd.ExcelWriter('locks.xlsx', engine='openpyxl', mode='a', if_sheet_exists='overlay')
    audit_csv_data.to_excel(audit_csv_writer, sheet_name='Audit-'+todays_date+'', index=False, header=True, startrow=0)
    audit_csv_writer.save()

    audit_workbook = load_workbook(filename = 'locks.xlsx')
    audit_worksheet = audit_workbook['Audit-'+todays_date+'']
    audit_worksheet.auto_filter.ref = audit_worksheet.dimensions
    audit_workbook.save('locks.xlsx')
    audit_workbook.close()