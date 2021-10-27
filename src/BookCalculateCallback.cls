VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BookCalculateCallback"
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
'' Description:
''    - Allows for a method (callback) to be called (by name) when a sheet
''      calculate event is triggered in a workbook
'' Note:
''    - The callback must expect a single Sheet argument:
''      Sub/Function Callback(Sh as Object)
''========================================================================

'@Folder("TopoXL.Init")

Option Explicit

Private WithEvents m_book As Workbook
Attribute m_book.VB_VarHelpID = -1
Private m_callbackName As String
Private m_ignoreErrors As Boolean

Public Sub init(book As Workbook _
    , ByVal callbackName_ As String _
    , ByVal ignoreErrors_ As Boolean _
)
    Set m_book = book
    m_callbackName = callbackName_
    m_ignoreErrors = ignoreErrors_
End Sub

' Event triggered when a Sheet is calculating
Private Sub m_book_SheetCalculate(ByVal sh As Object)
    If m_callbackName = vbNullString Then Exit Sub
    '
    If m_ignoreErrors Then On Error Resume Next
    Run "'" & m_book.name & "'!" & m_callbackName, sh 'Application.Run not allowed by Bitdefender
    If m_ignoreErrors Then On Error GoTo 0
End Sub

Private Property Get Workbook() As Workbook
    Set Workbook = m_book
End Property
Private Property Get CallbackName() As String
    CallbackName = m_callbackName
End Property
Private Property Get IgnoreErrors() As Boolean
    IgnoreErrors = m_ignoreErrors
End Property
