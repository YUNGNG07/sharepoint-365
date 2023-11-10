from office365.runtime.auth.client_credential import ClientCredential
from office365.sharepoint.client_context import ClientContext
import json

with open('config.json', 'r') as f:
    credentials = json.load(f)
    client_id = credentials['client_id']
    client_secret = credentials['client_secret']
    site_url = credentials['site_url']

# client_id = '3268db4f-6258-4fae-bc12-7982c1eb1783'
# client_secret = 'HYtKhZwvJ03862cL16MiAOr5V41LAfaZDPHMTqmJh3I='
# site_url = 'https://keysighttech.sharepoint.com/sites/TestProjects-InventoryManagementSystem'

credentials = ClientCredential(client_id, client_secret)
context = ClientContext(site_url).with_credentials(credentials)
web = context.web
context.load(web)
context.execute_query()
info = web.properties

for key, value in info.items():
    print(f'{key}: {value}')
