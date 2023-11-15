import requests
import sys

# Define constants for HTTP methods
HTTP_GET = "GET"
HTTP_POST = "POST"
HTTP_PUT = "PUT"
HTTP_PATCH = "PATCH"

# Check if the folder name and client value are provided as command-line arguments
if len(sys.argv) < 3:
    print("Usage: python CreateFolder.py <FolderName> <ClientValue>")
    sys.exit()

# Get the folder name and client value from the command-line arguments
folder_name = sys.argv[1]
client_value = sys.argv[2]

# Define your SharePoint site ID, drive ID, list ID and Graph API URL
list_id = "list-id"
site_id = "site-id"
drive_id = "drive-id"
graph_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/root:/"

# Define your Azure AD app registration details
client_id = "client-id"
client_secret = "secret-id"
tenant_id = "tenant-id"

# Get an access token using client credentials flow
token_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
token_data = {
    "grant_type": "client_credentials",
    "scope": "https://graph.microsoft.com/.default",
    "client_id": client_id,
    "client_secret": client_secret
}

token_response = requests.post(token_url, data=token_data)
token_json = token_response.json()

# Check if the request was successful
if token_response.status_code == 200:
    # Parse the access token from the response JSON
    access_token = token_json["access_token"]

    # Create a folder with the provided name
    folder_data = {"name": folder_name, "folder": {}, "@microsoft.graph.conflictBehavior": "fail"}
    create_folder_url = graph_url + folder_name

    # Send a PUT request to create the folder
    headers = {"Authorization": "Bearer " + access_token, "Content-Type": "application/json"}
    create_folder_response = requests.put(create_folder_url, json=folder_data, headers=headers)

    # Check the response status
    if create_folder_response.status_code == 201:
        print(f"Folder '{folder_name}' created successfully.")

        # After creating the folder, run the additional GET request
        get_list_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{list_id}?expand=columns,items(expand=fields)"
        get_list_response = requests.get(get_list_url, headers=headers)

        # Check the response status for the GET request
        if get_list_response.status_code == 200:
            print("GET request for the list succeeded.")
            
            # Find the item with the name of the created folder in the "FileLeafRef" field
            items = get_list_response.json()["items"]
            folder_item = next((item for item in items if item["fields"]["FileLeafRef"] == folder_name), None)

            # Print the ID of the item
            if folder_item:
                item_id = folder_item["id"]
                print(f"ID of the item with name '{folder_name}': {item_id}")

                # Run the PATCH request to update the "Client" field
                patch_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{list_id}/items/{item_id}/fields"
                patch_data = {"Client": client_value}
                patch_response = requests.patch(patch_url, json=patch_data, headers=headers)

                # Check the response status for the PATCH request
                if patch_response.status_code == 200:
                    print("PATCH request succeeded.")
                else:
                    print(f"Failed to execute PATCH request. Status code: {patch_response.status_code}")
                    print("Response content:")
                    print(patch_response.text)
            else:
                print(f"No item found with name '{folder_name}' in the 'FileLeafRef' field.")

        else:
            print(f"Failed to execute GET request for the list. Status code: {get_list_response.status_code}")
            print("Response content:")
            print(get_list_response.text)

    elif create_folder_response.status_code == 409:
        print(f"Folder '{folder_name}' already exists.")
    else:
        print(f"Failed to create folder. Status code: {create_folder_response.status_code}")
        print("Response content:")
        print(create_folder_response.text)
else:
    print(f"Failed to obtain access token. Status code: {token_response.status_code}")
    print("Response content:")
    print(token_response.text)
