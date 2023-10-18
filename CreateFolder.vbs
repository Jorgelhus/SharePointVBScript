'------------------------------------------------------------
' CreateFolder.vbs
'
' Description:
'   This VBScript is designed to create a folder in a SharePoint site using the Microsoft Graph API.
'   It uses client credentials flow to obtain an access token and then sends a PUT request to create the folder.
'
' Usage:
'   cscript CreateFolder.vbs <FolderName>
'
' Parameters:
'   <FolderName> - The name of the folder you want to create. This should be provided as a command-line argument.
'
' Optional Parameters:
'   - The following optional parameters can be adjusted to change script behavior:
'     - site_id - Your SharePoint site ID.
'     - graph_url - Microsoft Graph API URL.
'     - client_id - Azure AD app client ID.
'     - client_secret - Azure AD app client secret.
'     - tenant_id - Your Azure AD tenant ID.
'     - @microsoft.graph.conflictBehavior - Conflict behavior when the folder name already exists.
'       Options:
'         - "replace" - Overwrite the existing folder (default).
'         - "rename" - Automatically rename the new folder to avoid conflicts.
'         - "fail" - Fail the operation if a folder with the same name exists.
'
' Note:
'   This script uses simplified JSON handling and is intended for basic folder creation.
'   More advanced JSON parsing may be needed for complex scenarios.
'
'------------------------------------------------------------



Option Explicit

' Define constants for HTTP methods
Const HTTP_GET = "GET"
Const HTTP_POST = "POST"
Const HTTP_PUT = "PUT"

' Check if the folder name is provided as a command-line argument
If WScript.Arguments.Count = 0 Then
    WScript.Echo "Usage: cscript CreateFolder.vbs <FolderName>"
    WScript.Quit
End If

' Get the folder name from the command-line argument
Dim folder_name
folder_name = WScript.Arguments(0)

' Define your SharePoint site ID and Graph API URL
Dim site_id
site_id = "siteid"  ' Use your site ID

Dim graph_url
graph_url = "https://graph.microsoft.com/v1.0/sites/" & site_id & "/drive/root:/"

' Define your Azure AD app registration details
Dim client_id, client_secret, tenant_id
client_id = "clientid"
client_secret = "clientsecret"
tenant_id = "tentantid"

' Create an HTTP request object
Dim httpRequest
Set httpRequest = CreateObject("MSXML2.ServerXMLHTTP")

' Get an access token using client credentials flow
Dim token_url, token_data, access_token
token_url = "https://login.microsoftonline.com/" & tenant_id & "/oauth2/v2.0/token"
token_data = "grant_type=client_credentials" & _
             "&scope=https://graph.microsoft.com/.default" & _
             "&client_id=" & client_id & _
             "&client_secret=" & client_secret

httpRequest.Open HTTP_POST, token_url, False
httpRequest.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
httpRequest.send token_data

' Check if the request was successful
If httpRequest.status = 200 Then
    ' Parse the access token from the response text (simplified)
    access_token = Mid(httpRequest.responseText, InStr(httpRequest.responseText, """access_token"":""") + Len("""access_token"":"""))
    access_token = Left(access_token, InStr(access_token, """") - 1)
    
    ' Create a folder with the provided name
    Dim folder_data, create_folder_url, create_folder_response
    folder_data = "{""name"":""" & folder_name & """,""folder"":{},""@microsoft.graph.conflictBehavior"":""fail""}"
    create_folder_url = graph_url & folder_name
    
    ' Send a PUT request to create the folder
    httpRequest.Open HTTP_PUT, create_folder_url, False
    httpRequest.setRequestHeader "Authorization", "Bearer " & access_token
    httpRequest.setRequestHeader "Content-Type", "application/json"
    httpRequest.send folder_data
    
    ' Check the response status
    If httpRequest.status = 201 Then
        WScript.Echo "Folder '" & folder_name & "' created successfully."
    ElseIf httpRequest.status = 409 Then
        WScript.Echo "Folder '" & folder_name & "' already exists."
    Else
        WScript.Echo "Failed to create folder. Status code: " & httpRequest.status
        WScript.Echo "Response content:"
        WScript.Echo httpRequest.responseText
    End If
Else
    WScript.Echo "Failed to obtain access token. Status code: " & httpRequest.status
    WScript.Echo "Response content:"
    WScript.Echo httpRequest.responseText
End If

' Clean up the HTTP request object
Set httpRequest = Nothing
