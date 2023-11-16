# SharePointVBScript
Script Process to create a folder on the root of a Sharepoint Portal

### CreateFolder.vbs
Generates the folder on the root of the selected site (using **site-id** and **drive-id** as reference).

### CreateFolderStructure.vbs
Generates folder inside another folder, following a pre-defined structure (requires **drive-id** and **parent-folder-id** as reference).

### CaptureDriveID.py
Script in Python to be ran for the first time on setup so the **drive-id** can be captured (output on the console).

### Capture ParentFolderID.py
Script in Python to be ran for the first time on setup so the **parent-folder-id** can be captured (output on the console).

### FolderWMeta.py
Script in Python that creates a folder on the sharepoint, get the list of the files, detect the id of the file in regards to the list and updates the "Client" column.

### PatchClientName.vbs
Script in VBS that pulls the list of items on the sharepoint list selected, and then compares to the folder Name given, pulling its ID in the list so the "Client" column can be updated. Does not create a folder. Just update the list.
