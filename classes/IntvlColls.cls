VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IntvlColls"
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
''    Modules: IntervalCollectionsInit
''    Classes: None
''Calls:
''    Modules: None
''    Classes: IntvlColl
''=======================================================
Option Explicit
Private m_intvlColls As Collection

'Constructor
Private Sub Class_Initialize()
    Set m_intvlColls = New Collection
End Sub

'add new Red Line
Public Function addIntvlColl(ic As IntvlColl) As Boolean
    On Error GoTo ErrHandler
    m_intvlColls.Add ic, ic.name
    addIntvlColl = True
    Exit Function
    
ErrHandler:
    addIntvlColl = False
End Function

'gets a IC object by it's name
Public Function getIntvlColl(name As String) As IntvlColl
    On Error GoTo ErrHandler
    Set getIntvlColl = m_intvlColls.item(name)
    Exit Function
ErrHandler:
    Set getIntvlColl = Nothing
End Function


