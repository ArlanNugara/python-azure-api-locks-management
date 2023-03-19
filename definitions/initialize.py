import requests
import os
import json
import sys
from .get_locks import get_all_locks
from .update_locks import update_locks_at_scope
import configparser

# Start Lock Process
def start_lock_process(cloud_type, task_type, scope):
    # Read the Config
    config = configparser.ConfigParser()
    config.read('param.ini')
    if cloud_type == "Azure":
        URL = config['Azure']['URL']
        RESOURCE = config['Azure']['RESOURCE']
        LOGIN_URL = ""+URL+""+os.environ.get('ARM_TENANT_ID')+"/oauth2/token"
        PARAMS_MGMT = {'grant_type':'client_credentials','client_id': ''+os.environ.get('ARM_CLIENT_ID')+'','client_secret':''+os.environ.get('ARM_CLIENT_SECRET')+'','resource':RESOURCE}
    elif cloud_type == 'AzureChina':
        URL = config['AzureChina']['URL']
        RESOURCE = config['AzureChina']['RESOURCE']
        LOGIN_URL = ""+URL+""+os.environ.get('ARM_TENANT_ID')+"/oauth2/token"
        PARAMS_MGMT = {'grant_type':'client_credentials','client_id': ''+os.environ.get('ARM_CLIENT_ID')+'','client_secret':''+os.environ.get('ARM_CLIENT_SECRET')+'','resource':RESOURCE}
    else:
        print("Cloud Type Should be either 'Azure' or 'AzureChina'.")
        sys.exit(1)
    # # Login and get access token
    login = requests.post(url = LOGIN_URL, data = PARAMS_MGMT)
    jsonfy = login.json()
    token = jsonfy["access_token"]
    query_header = {"Content-Type": "application/json", "Authorization": "Bearer "+token}
    print("Login Successful")
    if task_type == "get":
        print("Starting GET operation")
        get_all_locks(RESOURCE, scope, query_header)
    elif task_type == "update":
        print("Starting UPDATE operation")
        update_locks_at_scope(RESOURCE, scope, query_header)
    else:
        print("Task type should be either 'get' or 'update'.")
        sys.exit(1)