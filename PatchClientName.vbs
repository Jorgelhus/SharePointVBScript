Option Explicit

Class VbsJson
    'Author: Demon
    'Date: 2012/5/3
    'Website: http://demon.tw
    Private Whitespace, NumberRegex, StringChunk
    Private b, f, r, n, t

    Private Sub Class_Initialize
        Whitespace = " " & vbTab & vbCr & vbLf
        b = ChrW( 8 )
        f = vbFormFeed
        r = vbCr
        n = vbLf
        t = vbTab

        Set NumberRegex = New RegExp
        NumberRegex.Pattern = "(-?(?:0|[1-9]\d*))(\.\d+)?([eE][-+]?\d+)?"
        NumberRegex.Global = False
        NumberRegex.MultiLine = True
        NumberRegex.IgnoreCase = True

        Set StringChunk = New RegExp
        StringChunk.Pattern = "([\s\S]*?)([""\\\x00-\x1f])"
        StringChunk.Global = False
        StringChunk.MultiLine = True
        StringChunk.IgnoreCase = True
    End Sub
   
    'Return a JSON string representation of a VBScript data structure
    'Supports the following objects and types
    '+-------------------+---------------+
    '| VBScript | JSON |
    '+======================================
    '| Dictionary | object |
    '+-------------------+---------------+
    '| Array | array |
    '+-------------------+---------------+
    '| String | string |
    '+-------------------+---------------+
    '| Number | number |
    '+-------------------+---------------+
    '| True | true |
    '+-------------------+---------------+
    '| False | false |
    '+-------------------+---------------+
    '| Null | null |
    '+-------------------+---------------+
    Public Function Encode( ByRef obj)
        Dim buf, i, c, g
        Set buf = CreateObject ( "Scripting.Dictionary" )
        Select Case VarType (obj)
            Case vbNull
                buf.Add buf.Count, "null"
            Case vbBoolean
                If obj Then
                    buf.Add buf.Count, "true"
                Else
                    buf.Add buf.Count, "false"
                End If
            Case vbInteger , vbLong , vbSingle , vbDouble
                buf.Add buf.Count, obj
            Case vbString
                buf.Add buf.Count, """"
                For i = 1 To Len (obj)
                    c = Mid (obj, i, 1 )
                    Select Case c
                        Case """" buf.Add buf.Count, "\"""
                        Case "\" buf.Add buf.Count, "\\"
                        Case "/" buf.Add buf.Count, "/"
                        Case b buf.Add buf.Count, "\b"
                        Case f buf.Add buf.Count, "\f"
                        Case r buf.Add buf.Count, "\r"
                        Case n buf.Add buf.Count, "\n"
                        Case t buf.Add buf.Count, "\t"
                        Case Else
                            If AscW(c) >= 0 And AscW(c) <= 31 Then
                                c = Right ( "0" & Hex (AscW(c)), 2 )
                                buf.Add buf.Count, "\u00" & c
                            Else
                                buf.Add buf.Count, c
                            End If
                    End Select
                Next
                buf.Add buf.Count, """"
            Case vbArray + vbVariant
                g = True
                buf.Add buf.Count, "["
                For Each i In obj
                    If g Then g = False Else buf.Add buf.Count, ","
                    buf.Add buf.Count, Encode(i)
                Next
                buf.Add buf.Count, "]"
            Case vbObject
                If TypeName (obj) = "Dictionary" Then
                    g = True
                    buf.Add buf.Count, "{"
                    For Each i In obj
                        If g Then g = False Else buf.Add buf.Count, ","
                        buf.Add buf.Count, """" & i & """" & ":" & Encode(obj(i))
                    Next
                    buf.Add buf.Count, "}"
                Else
                    Err.Raise 8732,, "None dictionary object"
                End If
            Case Else
                buf.Add buf.Count, """" & CStr (obj) & """"
        End Select
        Encode = Join (buf.Items, "" )
    End Function

    'Return the VBScript representation of ``str(``
    'Performs the following translations in decoding
    '+---------------+-------------------+
    '| JSON | VBScript |
    '+======================================
    '| object | Dictionary |
    '+---------------+-------------------+
    '| array | Array |
    '+---------------+-------------------+
    '| string | String |
    '+---------------+-------------------+
    '| number | Double |
    '+---------------+-------------------+
    '| true | True |
    '+---------------+-------------------+
    '| false | False |
    '+---------------+-------------------+
    '| null | Null |
    '+---------------+-------------------+
    Public Function Decode( ByRef str)
        Dim idx
        Idx = SkipWhitespace(str, 1 )

        If Mid (str, idx, 1 ) = "{" Then
            Set Decode = ScanOnce(str, 1 )
        Else
            Decode = ScanOnce(str, 1 )
        End If
    End Function
   
    Private Function ScanOnce( ByRef str, ByRef idx)
        Dim c, ms

        Idx = SkipWhitespace(str, idx)
        c = Mid (str, idx, 1 )

        If c = "{" Then
            Idx = idx + 1
            Set ScanOnce = ParseObject(str, idx)
            Exit Function
        ElseIf c = "[" Then
            Idx = idx + 1
            ScanOnce = ParseArray(str, idx)
            Exit Function
        ElseIf c = """" Then
            Idx = idx + 1
            ScanOnce = ParseString(str, idx)
            Exit Function
        ElseIf c = "n" And StrComp ( "null" , Mid (str, idx, 4 )) = 0 Then
            Idx = idx + 4
            ScanOnce = Null
            Exit Function
        ElseIf c = "t" And StrComp ( "true" , Mid (str, idx, 4 )) = 0 Then
            Idx = idx + 4
            ScanOnce = True
            Exit Function
        ElseIf c = "f" And StrComp ( "false" , Mid (str, idx, 5 )) = 0 Then
            Idx = idx + 5
            ScanOnce = False
            Exit Function
        End If
       
        Set ms = NumberRegex.Execute( Mid (str, idx))
        If ms.Count = 1 Then
            Idx = idx + ms( 0 ).Length
            ScanOnce = CDbl (ms( 0 ))
            Exit Function
        End If
       
        Err.Raise 8732,, "No JSON object could be ScanOnced"
    End Function

    Private Function ParseObject( ByRef str, ByRef idx)
        Dim c, key, value
        Set ParseObject = CreateObject ( "Scripting.Dictionary" )
        Idx = SkipWhitespace(str, idx)
        c = Mid (str, idx, 1 )
       
        If c = "}" Then
            Exit Function
        ElseIf c <> """" Then
            Err.Raise 8732,, "Expecting property name"
        End If

        Idx = idx + 1
       
        Do
            Key = ParseString(str, idx)

            Idx = SkipWhitespace(str, idx)
            If Mid (str, idx, 1 ) <> ":" Then
                Err.Raise 8732,, "Expecting : delimiter"
            End If

            Idx = SkipWhitespace(str, idx + 1 )
            If Mid (str, idx, 1 ) = "{" Then
                Set value = ScanOnce(str, idx)
            Else
                Value = ScanOnce(str, idx)
            End If
            ParseObject.Add key, value

            Idx = SkipWhitespace(str, idx)
            c = Mid (str, idx, 1 )
            If c = "}" Then
                Exit Do
            ElseIf c <> "," Then
                Err.Raise 8732,, "Expecting , delimiter"
            End If

            Idx = SkipWhitespace(str, idx + 1 )
            c = Mid (str, idx, 1 )
            If c <> """" Then
                Err.Raise 8732,, "Expecting property name"
            End If

            Idx = idx + 1
        Loop

        Idx = idx + 1
    End Function
   
    Private Function ParseArray( ByRef str, ByRef idx)
        Dim c, values, value
        Set values = CreateObject ( "Scripting.Dictionary" )
        Idx = SkipWhitespace(str, idx)
        c = Mid (str, idx, 1 )

        If c = "]" Then
            ParseArray = values.Items
            Exit Function
        End If

        Do
            Idx = SkipWhitespace(str, idx)
            If Mid (str, idx, 1 ) = "{" Then
                Set value = ScanOnce(str, idx)
            Else
                Value = ScanOnce(str, idx)
            End If
            values.Add values.Count, value

            Idx = SkipWhitespace(str, idx)
            c = Mid (str, idx, 1 )
            If c = "]" Then
                Exit Do
            ElseIf c <> "," Then
                Err.Raise 8732,, "Expecting , delimiter"
            End If

            Idx = idx + 1
        Loop

        Idx = idx + 1
        ParseArray = values.Items
    End Function
   
    Private Function ParseString( ByRef str, ByRef idx)
        Dim chunks, content, terminator, ms, esc, char
        Set chunks = CreateObject ( "Scripting.Dictionary" )

        Do
            Set ms = StringChunk.Execute( Mid (str, idx))
            If ms.Count = 0 Then
                Err.Raise 8732,, "Unterminated string starting"
            End If
           
            Content = ms( 0 ).Submatches( 0 )
            Terminator = ms( 0 ).Submatches( 1 )
            If Len (content) > 0 Then
                chunks.Add chunks.Count, content
            End If
           
            Idx = idx + ms( 0 ).Length
           
            If terminator = """" Then
                Exit Do
            ElseIf terminator <> "\" Then
                Err.Raise 8732,, "Invalid control character"
            End If
           
            Esc = Mid (str, idx, 1 )

            If esc <> "u" Then
                Select Case esc
                    Case """" char = """"
                    Case "\" char = "\"
                    Case "/" char = "/"
                    Case "b" char = b
                    Case "f" char = f
                    Case "n" char = n
                    Case "r" char = r
                    Case "t" char = t
                    Case Else Err.Raise 8732,, "Invalid escape"
                End Select
                Idx = idx + 1
            Else
                Char = ChrW( "&H" & Mid (str, idx + 1 , 4 ))
                Idx = idx + 5
            End If

            chunks.Add chunks.Count, char
        Loop

        ParseString = Join (chunks.Items, "" )
    End Function

    Private Function SkipWhitespace( ByRef str, ByVal idx)
        Do While idx <= Len (str) And _
            InStr (Whitespace, Mid (str, idx, 1 )) > 0
            Idx = idx + 1
        Loop
        SkipWhitespace = idx
    End Function

End Class 

' Define constants for HTTP methods
Const HTTP_GET = "GET"
Const HTTP_PATCH = "PATCH"
Const HTTP_POST = "POST"

' Check if the folder name and client value are provided as command-line arguments
If WScript.Arguments.Count < 2 Then
    WScript.Echo "Usage: cscript GetAndPatchList.vbs <FolderName> <ClientValue>"
    WScript.Quit
End If

' Get the folder name and client value from the command-line arguments
Dim folder_name, client_value
folder_name = WScript.Arguments(0)
client_value = WScript.Arguments(1)

' Define your SharePoint site ID, list ID, and Graph API URL
Dim list_id, site_id
list_id = "%7Bc172bec8-0448-4116-8214-55b3885e95f4%7D"
site_id = "ea8eda3f-cc3e-49c9-9a13-2b7a1d7dad41"

' Define your Azure AD app registration details
Dim client_id, client_secret, tenant_id
client_id = "abf3c9c9-59cd-4d56-b4ba-32265a4bace3"
client_secret = "0FD8Q~hbQKixc~K93BxRY39zQmpTxtFzD9TSJaBX"
tenant_id = "a7346d8c-bf93-4fe8-ac40-640e06c90cf3"

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

    ' Run the additional GET request
    Dim get_list_url, get_list_response
    get_list_url = "https://graph.microsoft.com/v1.0/sites/" & site_id & "/lists/" & list_id & "?expand=columns,items(expand=fields)"

    ' Send a GET request to get the list
    httpRequest.Open HTTP_GET, get_list_url, False
    httpRequest.setRequestHeader "Authorization", "Bearer " & access_token
    httpRequest.setRequestHeader "Content-Type", "application/json"
    httpRequest.send

    ' Check the response status for the GET request
    If httpRequest.status = 200 Then
        WScript.Echo "GET request for the list succeeded."

        ' Use VbsJson to parse the JSON response
        Dim jsonParser
        Set jsonParser = New VbsJson
        Dim jsonObject
        Set jsonObject = jsonParser.Decode(httpRequest.responseText)

        ' Print the objects inside the jsonObject
        ' PrintJsonObject jsonObject, 0

        ' Search for the target folder_name within the JSON object
        Dim folderId
        folderId = FindFolderId(folder_name, jsonObject)

        If Not IsEmpty(folderId) Then
            ' Run the PATCH request to update the "Client" field
            Dim patch_url, patch_data, patch_response
            patch_url = "https://graph.microsoft.com/v1.0/sites/" & site_id & "/lists/" & list_id & "/items/" & folderId & "/fields"
            patch_data = "{""Client"": """ & client_value & """}"

            ' Send a PATCH request to update the "Client" field
            httpRequest.Open HTTP_PATCH, patch_url, False
            httpRequest.setRequestHeader "Content-Type", "application/json"
            httpRequest.setRequestHeader "Authorization", "Bearer " & access_token
            httpRequest.send patch_data

            ' Check the response status for the PATCH request
            If httpRequest.status = 200 Then
                WScript.Echo "PATCH request succeeded."
            Else
                WScript.Echo "Failed to execute PATCH request. Status code: " & httpRequest.status
                WScript.Echo "Response content:"
                WScript.Echo httpRequest.responseText
            End If
        Else
            WScript.Echo "No item found with name '" & folder_name & "' in the 'FileLeafRef' field."
        End If
    Else
        WScript.Echo "Failed to execute GET request for the list. Status code: " & httpRequest.status
        WScript.Echo "Response content:"
        WScript.Echo httpRequest.responseText
    End If

Else
    WScript.Echo "Failed to obtain access token. Status code: " & httpRequest.status
    WScript.Echo "Response content:"
    WScript.Echo httpRequest.responseText
End If

Function FindFolderId(folderName, jsonObject)
    ' Check if the 'items' property exists in the JSON object
    If IsObject(jsonObject) Then
        'WScript.Echo "Phase 1"
        ' Check if 'items' property is an array
        If IsArray(jsonObject("items")) Then
            'WScript.Echo "Phase 2 - Array found"
            ' Iterate through each item in the 'items' array
            Dim item
            For Each item In jsonObject("items")
                'WScript.Echo "Phase 3"
                ' Output the entire 'item' object for inspection
                ' PrintJsonObject item, 0 ' This function should handle nested objects and arrays
                ' Check if 'Fields' property exists in the item
                If IsObject(item("fields")) Then
                    'WScript.Echo "Phase 4"
                    ' Check if 'FileLeafRef' field matches the folder name
                    Dim fileLeafRef
                    fileLeafRef = LCase(item("fields")("FileLeafRef"))
                    'WScript.Echo "Comparing: " & fileLeafRef & " with " & LCase(folderName)
                    If fileLeafRef = LCase(folderName) Then
                        ' Return the ID if the folder is found
                        FindFolderId = item("fields")("id")
                        'WScript.Echo "Phase 5"
                        'WScript.Echo "Folder ID: " & FindFolderId
                        Exit Function
                    End If
                End If
            Next
        End If
    End If

    ' Return 0 if the folder is not found
    FindFolderId = 0
    WScript.Echo "No Phase"
    WScript.Echo "Folder not found"
End Function





Sub PrintJsonObject(obj, indent)
    Dim key, value
    For Each key In obj
        If IsObject(obj(key)) Then
            WScript.Echo Space(indent) & key & ":"
            If IsArray(obj(key)) Then
                ' Handle array of objects
                Dim i
                For i = 0 To UBound(obj(key))
                    WScript.Echo Space(indent + 2) & "Item " & i + 1 & ":"
                    PrintJsonObject obj(key)(i), indent + 4
                Next
            Else
                PrintJsonObject obj(key), indent + 2
            End If
        Else
            ' Handle non-string, non-numeric values gracefully
            On Error Resume Next
            value = obj(key)
            On Error GoTo 0

            If IsNull(value) Then
                WScript.Echo Space(indent) & key & ": Null"
            ElseIf IsNumeric(value) Then
                WScript.Echo Space(indent) & key & ": " & value
            ElseIf VarType(value) = vbDate Then
                WScript.Echo Space(indent) & key & ": #" & FormatDateTime(value, vbShortDate) & "#"
            ElseIf VarType(value) = vbBoolean Then
                WScript.Echo Space(indent) & key & ": " & CStr(value)
            ElseIf VarType(value) = vbString Then
                WScript.Echo Space(indent) & key & ": """ & Replace(CStr(value), """", """""") & """"
            Else
                ' Additional information about the type
                WScript.Echo Space(indent) & key & ": (Unsupported type - " & TypeName(value) & ")"
            End If

            ' Additional check for items to print more details
            If key = "items" And IsArray(value) Then
                Dim itemIndex, itemName, itemId, fieldKey
                For itemIndex = LBound(value) To UBound(value)
                    WScript.Echo Space(indent + 2) & "Item " & itemIndex + 1 & ":"
                    If IsObject(value(itemIndex)) Then
                        If IsObject(value(itemIndex)("fields")) Then
                            WScript.Echo Space(indent + 4) & "Fields:"
                            For Each fieldKey In value(itemIndex)("fields")
                                WScript.Echo Space(indent + 6) & fieldKey & ": " & value(itemIndex)("fields")(fieldKey)
                            Next
                        End If
                        If IsObject(value(itemIndex)("id")) Then
                            WScript.Echo Space(indent + 4) & "ID: " & value(itemIndex)("id")
                        End If
                    End If
                Next
            End If
        End If
    Next
End Sub
