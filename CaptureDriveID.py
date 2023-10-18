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

# Define your Azure AD app registration details
client_id = "YOUR_CLIENT_ID"  # Replace with your Azure AD app client ID
client_secret = "YOUR_CLIENT_SECRET"  # Replace with your Azure AD app client secret
tenant_id = "YOUR_TENANT_ID"  # Replace with your Azure AD tenant ID

# Define your SharePoint site ID
site_id = "YOUR_SITE_ID"  # Replace with your SharePoint site ID

try:
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

    # Define the URL for listing drives in the site
    list_drives_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives"

    # Set up the request with the access token
    headers = {
        "Authorization": f"Bearer {access_token}"
    }

    # Send the request
    response = requests.get(list_drives_url, headers=headers)

    if response.status_code == 200:
        # Parse the JSON response
        data = response.json()
        
        # Extract information
        drive_id = data['value'][0]['id']

        # Format and print the result
        formatted_output = f'drive-id: {drive_id}'
        print(formatted_output)
    else:
        print(f"Request failed with status code: {response.status_code}")
except Exception as e:
    print(f"An error occurred: {e}")
