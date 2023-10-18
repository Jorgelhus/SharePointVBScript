"""
Script Description:
-------------------
This script retrieves the names and folder IDs of folders in a specific document library or drive
using the Microsoft Graph API and Azure AD app credentials.

Script Usage:
-------------
1. Replace the placeholders for 'client_id', 'client_secret', 'tenant_id', and 'drive_id' with your Azure AD app registration details and drive ID.
2. Run the script to list the folders and their IDs in the specified drive.

Script Variables:
----------------
- client_id: Azure AD app client ID.
- client_secret: Azure AD app client secret.
- tenant_id: Azure AD tenant ID.
- drive_id: ID of the specific document library or drive.
"""

import requests
import json

# Define your Azure AD app registration details
client_id = "YOUR_CLIENT_ID"  # Replace with your Azure AD app client ID
client_secret = "YOUR_CLIENT_SECRET"  # Replace with your Azure AD app client secret
tenant_id = "YOUR_TENANT_ID"  # Replace with your Azure AD tenant ID

# Define the Drive ID for the specific document library or drive
drive_id = "YOUR_DRIVE_ID"

# Get an access token using client credentials flow
token_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
token_data = {
    "grant_type": "client_credentials",
    "scope": "https://graph.microsoft.com/.default",
    "client_id": client_id,
    "client_secret": client_secret,
}
token_response = requests.post(token_url, data=token_data)
access_token = token_response.json().get("access_token")

# Define the URL for listing the children (items) in the drive
list_children_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/root/children"

# Set up the request with the access token
headers = {
    "Authorization": f"Bearer {access_token}"
}

# Send the request
response = requests.get(list_children_url, headers=headers)

# Extract and print the name and folder ID for each item
data = response.json()
for item in data['value']:
    if 'folder' in item:
        folder_name = item['name']
        folder_id = item['id']
        print(f"Folder Name: {folder_name}, Folder ID: {folder_id}")
