Attribute VB_Name = "RedLinesInit"
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
''    Modules: RLenums
''    Classes: RedLine, RedLines, RLarcCircle, RLarcClothoid, RLelement, RLline
''=======================================================
Option Private Module
Option Explicit
Public RED_LINES As New RedLines

'table name prefix
Private Const TBL_NAME_PREFIX As String = "tblRL"
Public Const TBL_ELEM_TYPE_COL As String = "Element Type"
Public Const TBL_ELEM_INPUT_TYPE_COL As String = "Input Type"


Public Sub initRedLines()
    Dim WS As Worksheet
    Dim tbl As ListObject
    Dim tempRL As RedLine
    
    'loop all worksheets
    For Each WS In ThisWorkbook.Worksheets
        'loop all tables
        For Each tbl In WS.ListObjects
            Set tempRL = getRedLineFromTable(tbl)
            If Not tempRL Is Nothing Then
                RED_LINES.addRedLine tempRL
            End If
        Next tbl
    Next WS
End Sub

Private Function isTableValid(tbl As ListObject) As Boolean
    
    isTableValid = False
    
    'check if Table name is valid
    If Left(tbl.name, 5) <> TBL_NAME_PREFIX Then
        isTableValid = False
        Exit Function
    End If
    
    'check if Table Header is valid
    Dim i As Integer
    Dim c As Range
    Dim foundColName As Boolean
    Dim inpValidArray() As String
    inpValidArray = getElemPartStringArray
    
    ReDim Preserve inpValidArray(UBound(inpValidArray) + 2)
    inpValidArray(UBound(inpValidArray) - 1) = TBL_ELEM_TYPE_COL
    inpValidArray(UBound(inpValidArray)) = TBL_ELEM_INPUT_TYPE_COL
    
    For i = 0 To UBound(inpValidArray)
        foundColName = False
        For Each c In tbl.HeaderRowRange
            If c.value = inpValidArray(i) Then
                foundColName = True
                Exit For
            End If
        Next c
        If foundColName = False Then
            isTableValid = False
            Exit Function
        End If
    Next i
    
    'if code executes till this line then table format is valid
    isTableValid = True
End Function

Private Function getRedLineFromTable(tbl As ListObject) As RedLine
    Dim outRL As RedLine
    Dim tempRLelem As RLelement
    Dim lr As ListRow
    Dim inputColl As Collection
    
    If isTableValid(tbl) Then
        Set outRL = New RedLine
        For Each lr In tbl.ListRows
            Set inputColl = tableRowToCollection(lr.index, tbl)
            Set tempRLelem = getRLelemFromCollection(inputColl)
            If Not tempRLelem Is Nothing Then
                outRL.addElem tempRLelem
             Else
                GoTo FailGetRedLine
            End If
        Next lr
    Else
        GoTo FailGetRedLine
    End If
    
    outRL.name = tbl.name
    Set getRedLineFromTable = outRL
    
    Exit Function
FailGetRedLine:
    Set outRL = Nothing
    Set getRedLineFromTable = outRL
    Exit Function
End Function

Private Function tableRowToCollection(rowIndex As Integer, tbl As ListObject) As Collection
    Dim result As New Collection
    Dim c As Range
    Dim tempVal As Variant
    
    For Each c In tbl.HeaderRowRange
        tempVal = tbl.ListRows(rowIndex).Range.Cells(1, c.Column).value
        If tempVal <> vbNullString Then
            result.Add tempVal, c.value
        End If
    Next c
    Set tableRowToCollection = result
End Function


Public Function getRLelemFromCollection(inputColl As Collection) As RLelement
    Dim outRLelem As RLelement
    Dim elemType As RL_ELEM_TYPES
    elemType = RLenums.rlElemTypeFromString(inputColl.item(TBL_ELEM_TYPE_COL))
    
    Select Case elemType
        Case RL_ELEM_TYPES.ELEM_LINE
            Dim tempRLline As New RLline
            If tempRLline.initFromCollection(inputColl) Then
                Set getRLelemFromCollection = tempRLline
             Else
                GoTo FailGetRLelem
            End If
        Case RL_ELEM_TYPES.ELEM_ARC_CIRCLE
            Dim tempRLarcCircle As New RLarcCircle
            If tempRLarcCircle.initFromCollection(inputColl) Then
                Set getRLelemFromCollection = tempRLarcCircle
             Else
                GoTo FailGetRLelem
            End If
        Case RL_ELEM_TYPES.ELEM_ARC_CLOTHOID
            Dim tempRLarcClothoid As New RLarcClothoid
            If tempRLarcClothoid.initFromCollection(inputColl) Then
                Set getRLelemFromCollection = tempRLarcClothoid
             Else
                GoTo FailGetRLelem
            End If
        Case Else
        
    End Select
    Exit Function
FailGetRLelem:
    Set outRLelem = Nothing
    Set getRLelemFromCollection = outRLelem
End Function





