Attribute VB_Name = "UDF_FileInfo"
''' TopoXL: Excel UDF library for land surveyors
''' Copyright (C) 2019 Bogdan Morosanu and Cristian Buse
''' This program is free software: you can redistribute it and/or modify
''' it under the terms of the GNU General Public License as published by
''' the Free Software Foundation, either version 3 of the License, or
''' (at your option) any later version.
'''
''' This program is distributed in the hope that it will be useful,
''' but WITHOUT ANY WARRANTY; without even the implied warranty of
''' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
''' GNU General Public License for more details.
'''
''' You should have received a copy of the GNU General Public License
''' along with this program.  If not, see <https://www.gnu.org/licenses/>.

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


