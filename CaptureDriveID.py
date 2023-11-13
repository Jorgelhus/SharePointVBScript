"""
Script Description:
-------------------
This script retrieves information about drives in a SharePoint site using the Microsoft Graph API.

Script Usage:
-------------
1. Replace the placeholders for 'client_id', 'client_secret', 'tenant_id', and 'site_id' with your Azure AD app registration details and SharePoint site ID.
2. Run the script to obtain drive information from the SharePoint site.

Script Variables:
----------------
- client_id: Azure AD app client ID.
- client_secret: Azure AD app client secret.
- tenant_id: Azure AD tenant ID.
- site_id: SharePoint site ID.
"""

import requests
import json

# Your Azure AD app registration details
client_id = "YOUR_CLIENT_ID"
client_secret = "YOUR_CLIENT_SECRET"
tenant_id = "YOUR_TENANT_ID"

# Your SharePoint site ID
site_id = "YOUR_SITE_ID"

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

# Microsoft Graph API endpoint to get drives (document libraries) in the site
list_drives_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives"

# Headers for the request
headers = {
    "Authorization": f"Bearer {access_token}",
    "Content-Type": "application/json",
}

# Make the request to get the list of drives
drives_response = requests.get(list_drives_url, headers=headers)

# Check if the request was successful (status code 200)
if drives_response.status_code == 200:
    drives_data = drives_response.json()

    # Extract and print the IDs of the drives (document libraries)
    for drive in drives_data['value']:
        print(f"Drive Name: {drive['name']}, Drive ID: {drive['id']}")
else:
    print(f"Failed to retrieve drives. Status code: {drives_response.status_code}")
    print("Response content:")
    print(drives_response.text)
