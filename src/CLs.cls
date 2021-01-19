VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CLs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

''========================================================================
'' Description
'' Represents a collection of CL objects.
'' This is to be used for storage of multiple CL objects. Each object is
'' identified by its name (key used by the collection)
''========================================================================

'@Folder("TopoXL.CL")
Option Explicit

Private m_CLs As Collection

' Class constructor with default values
Private Sub Class_Initialize()
    Set m_CLs = New Collection
End Sub

' Adds a new CL object to the CLs collection
' Returns True if the CL was added successfully
Public Function addCL(clObj As CL) As Boolean
    On Error GoTo ErrHandler
    m_CLs.Add clObj, clObj.name
    addCL = True
    Exit Function
ErrHandler:
    addCL = False
End Function

' Returns a CL object by its name
Public Function getCL(ByVal name As String) As CL
    On Error GoTo ErrHandler
    Set getCL = m_CLs.item(name)
    Exit Function
ErrHandler:
    Set getCL = Nothing
End Function

