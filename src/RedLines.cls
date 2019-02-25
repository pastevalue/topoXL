VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RedLines"
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

''=======================================================
''Called by:
''    Modules: RedLinesInit
''    Classes: None
''Calls:
''    Modules: None
''    Classes: RedLine
''=======================================================
Option Explicit
Private m_redLines As Collection

'Constructor
Private Sub Class_Initialize()
    Set m_redLines = New Collection
End Sub

'add new Red Line
Public Function addRedLine(rl As RedLine) As Boolean
    On Error GoTo ErrHandler
    m_redLines.Add rl, rl.name
    addRedLine = True
    Exit Function
    
ErrHandler:
    addRedLine = False
End Function

'gets a Rl object by it's name
Public Function getRedLine(name As String) As RedLine
    On Error GoTo ErrHandler
    Set getRedLine = m_redLines.item(name)
    Exit Function
ErrHandler:
    Set getRedLine = Nothing
End Function

