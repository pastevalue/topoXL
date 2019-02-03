Attribute VB_Name = "UDF_FileInfo"
''=======================================================
''Called by:
''    Modules: None
''    Classes: None
''Calls:
''    Modules: None
''    Classes: None
''=======================================================
Option Explicit

'Verfify if a file exists
Public Function fileInfoFileFolderExist(fullPath As String) As Boolean
    On Error GoTo EarlyExit
    If fullPath = "" Then
        fileInfoFileFolderExist = False
    Else
        If Not Dir(fullPath, vbDirectory) = vbNullString Then fileInfoFileFolderExist = True
    End If
EarlyExit:
    On Error GoTo 0
End Function


