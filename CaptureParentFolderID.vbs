import requests

# Define your Azure AD app registration details
client_id = "your-client-id"
client_secret = "your-client-secret"
tenant_id = "your-tenant-id"

# Define the Drive ID for the specific document library or drive
drive_id = "your-drive-id"

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

# Print the response, and look for the Parent Folder ID
print(response.json())
