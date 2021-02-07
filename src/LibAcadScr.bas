Attribute VB_Name = "LibAcadScr"
''' TopoXL: Excel UDF library for land surveyors
''' Copyright (C) 2021 Bogdan Morosanu and Cristian Buse
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
'' Store Acad script generation functions
''========================================================================

'@Folder("TopoXL.libs")

Option Explicit
Option Private Module

Private Const MODULE_NAME As String = "LibAcad"

' Autocad commands constants
Private Const ACAD_CMD_LAYER As String = "-layer"
Private Const ACAD_CMD_POINT As String = "point"
Private Const ACAD_CMD_PLINE As String = "pline"
Private Const ACAD_CMD_3DPOLY As String = "3dpoly"
Private Const ACAD_CMD_TEXT As String = "-text"
Private Const ACAD_CMD_INSERT As String = "-insert"

' Returns a string representing an AutoCad script which creates
' a point entity.
' Example: point 1,2,3
'
' Parameters:
'   - cooArr: coordinate array (dimension = 2)
' Raises error if:
'   - the second dimension of cooArr is not of size 2 or 3
'   - cooArr values are not numbers
Public Function pnt(ByVal cooArr As Variant) As String
    Dim i As Long
    Dim cooDim As Long
    Dim jLBound As Long
    jLBound = LBound(cooArr, 2)
    cooDim = UBound(cooArr, 2) - jLBound + 1
    pnt = vbNullString

    For i = LBound(cooArr, 1) To UBound(cooArr, 1)
        pnt = pnt & ACAD_CMD_POINT & " "
        Select Case cooDim
            Case 2
                pnt = pnt & numbersToCSV(cooArr(i, jLBound), cooArr(i, jLBound + 1)) & vbNewLine
            Case 3
                pnt = pnt & numbersToCSV(cooArr(i, jLBound), cooArr(i, jLBound + 1), cooArr(i, jLBound + 2)) & vbNewLine
            Case Else
                Err.Raise 5, MODULE_NAME, "Second dimension of cooArr must be of size 2 or 3!"
        End Select
    Next i
End Function

' Returns a string representing an AutoCad script which creates
' a pline/3dpoly entity
' Example:
' pline 1,2
' 3,4
'
' or
' 3dpoly 1,2,3
' 4,5,6
'
' Parameters:
'   - cooArr: coordinate array (dimension = 2)
' Raises error if:
'   - the first dimension of the array is of a size less than 2
'   - the second dimension of cooArr is not of size 2 or 3
'   - cooArr values are not numbers
Public Function pline(ByVal cooArr As Variant) As String
    ' Raise error if cooArr dimension 1 is not at least of size 2 (2 sets of coo at least)
    If UBound(cooArr, 1) - LBound(cooArr, 1) + 1 < 2 Then
        Err.Raise 5, MODULE_NAME, "Must have at least two sets of coo!"
    End If
    
    Dim i As Long
    Dim cooDim As Long
    Dim jLBound As Long
    jLBound = LBound(cooArr, 2)
    cooDim = UBound(cooArr, 2) - jLBound + 1
    pline = vbNullString

    Select Case cooDim
        Case 2
            pline = ACAD_CMD_PLINE & " "
        Case 3
            pline = ACAD_CMD_3DPOLY & " "
        Case Else
            Err.Raise 5, MODULE_NAME, "Second dimension of cooArr must be of size 2 or 3!"
    End Select

    For i = LBound(cooArr, 1) To UBound(cooArr, 1)
        Select Case cooDim
            Case 2
                pline = pline & numbersToCSV(cooArr(i, jLBound), cooArr(i, jLBound + 1)) & vbNewLine
            Case 3
                pline = pline & numbersToCSV(cooArr(i, jLBound), cooArr(i, jLBound + 1), cooArr(i, jLBound + 2)) & vbNewLine
            Case Else
                ' pass - never reaches this branch as cooDim is checked to be 2 or 3 only
        End Select
    Next i
End Function

' Returns a string representing an AutoCad script which inserts
' a block entity
' Example:
' -insert "blockName"
' 10,20,30
' 1
' 1
' 0
' Parameters:
'   - values: values array (two dimension array) (name, coordinates, x scale, y scale, rotation)
' Raises error if:
'   - the second dimension of cooArr is not of size 6 or 7
'   - coordinate, scale and rotation values are not numbers
Public Function blkInsert(ByVal values As Variant) As String
    Dim i As Long
    Dim valuesDim As Long
    Dim jLBound As Long
    jLBound = LBound(values, 2)
    valuesDim = UBound(values, 2) - jLBound + 1
    Dim cooStr As String ' insertion coordinates
    Dim lastCooIdx As Long  ' index of the first value after coordinate values (3 or 4)
    blkInsert = vbNullString
    For i = LBound(values, 1) To UBound(values, 1)
    blkInsert = blkInsert & ACAD_CMD_INSERT & " "
        Select Case valuesDim
            Case 6
                cooStr = numbersToCSV(values(i, jLBound + 1), values(i, jLBound + 2))
                lastCooIdx = 3
            Case 7
                cooStr = numbersToCSV(values(i, jLBound + 1), values(i, jLBound + 2), values(i, jLBound + 3))
                lastCooIdx = 4
            Case Else
                Err.Raise 5, MODULE_NAME, "Second dimension of values must be of size 6 or 7!"
        End Select
        
        blkInsert = blkInsert & getStringValue(values(i, jLBound + 0)) & vbNewLine  ' block name
        blkInsert = blkInsert & cooStr & vbNewLine  ' coordinates
        blkInsert = blkInsert & getNumericValue(values(i, jLBound + lastCooIdx)) & vbNewLine     ' x scale
        blkInsert = blkInsert & getNumericValue(values(i, jLBound + lastCooIdx + 1)) & vbNewLine ' y scale
        blkInsert = blkInsert & getNumericValue(values(i, jLBound + lastCooIdx + 2))              ' rotation
        If i < UBound(values, 1) Then blkInsert = blkInsert & vbNewLine                     ' add new line if multiple inserts
    Next i
End Function

' Returns a string representing an AutoCad script which creates
' a Single Line Text entity
' Example:
' -text 1,2,3 1 0 "Sample"
' Parameters:
'   - values: values array (2D array) (coordinates, height, rotation, text)
' Raises error if:
'   - the second dimension of cooArr is not of size 5 or 6
'   - cooArr values are not numbers
Public Function sText(ByVal values As Variant) As String
    Dim i As Long
    Dim valuesDim As Long
    Dim jLBound As Long
    jLBound = LBound(values, 2)
    valuesDim = UBound(values, 2) - jLBound + 1
    Dim cooStr As String ' insertion coordinates
    Dim lastCooIdx As Long  ' index of the first value after coordinate values (2 or 3)
    sText = vbNullString
    For i = LBound(values, 1) To UBound(values, 1)
    sText = sText & ACAD_CMD_TEXT & " "
        Select Case valuesDim
            Case 5
                cooStr = numbersToCSV(values(i, jLBound), values(i, jLBound + 1))
                lastCooIdx = 2
            Case 6
                cooStr = numbersToCSV(values(i, jLBound), values(i, jLBound + 1), values(i, jLBound + 2))
                lastCooIdx = 3
            Case Else
                Err.Raise 5, MODULE_NAME, "Second dimension of values must be of size 5 or 6!"
        End Select
        
        sText = sText & cooStr
        sText = sText & " " & getNumericValue(values(i, jLBound + lastCooIdx))      ' height
        sText = sText & " " & getNumericValue(values(i, jLBound + lastCooIdx + 1))  ' rotation
        sText = sText & " " & getStringValue(values(i, jLBound + lastCooIdx + 2))                 ' text
        If i < UBound(values, 1) Then sText = sText & vbNewLine                     ' add new line if multiple texts created
        
    Next i
End Function
' Returns a string representing an AutoCad script which changes
' a layer
' Example:
' -text 1,2,3 1 0 "Sample"
' Parameters:
'   - lyrName: layer name to be made active
' Raises error if:
'   - lyrName can't be casted to String
'   - lyrName is 0 length string
Public Function chngLayer(ByVal lyrName As Variant) As String
    chngLayer = ACAD_CMD_LAYER & " m " & getStringValue(lyrName) & vbNewLine
End Function


' Returns a comma delimited string of the input numbers
' Parameters:
'   - numbers: array of numbers to be used in the generation of the CSV string
' Raises error if:
'   - any of the input value is not a number
'   - array size is less than 2
Private Function numbersToCSV(ParamArray numbers() As Variant) As String
    Dim i As Integer
    Dim s As Integer    ' array size
    
    ' Raise error if array size is less than 2
    s = UBound(numbers) - LBound(numbers) + 1
    If s < 2 Then
        Err.Raise 5, MODULE_NAME, "At least two numbers must be provided!"
    End If
    
    ' Concatenate result string
    numbersToCSV = vbNullString
    For i = LBound(numbers) To UBound(numbers)
        If i = LBound(numbers) Then
            numbersToCSV = getNumericValue(numbers(i))
        Else
            numbersToCSV = numbersToCSV & "," & getNumericValue(numbers(i))
        End If
    Next i
End Function

' Returns the input value if IsNumeric evaluates to true.
' To be used with script generation functions where a numeric
' value is required (coordinates, scales, rotation)
' Parameters:
'   - v: value to be checked
' Raises error if:
'   - input value is not numeric
Private Function getNumericValue(ByVal v As Variant) As Variant
    If VBA.IsNumeric(v) Then
        getNumericValue = v
    Else
        Err.Raise 5, MODULE_NAME, "Value " & v & " must be a numeric value!"
    End If
End Function

' Returns the input value as String.
' To be used with script generation functions where a text
' value is required (block names, text content)
' Parameters:
'   - v: value to be checked
' Raises error if:
'   - input value is 0 length string or can't be read as string
Private Function getStringValue(ByVal v As Variant) As Variant
    On Error GoTo failString
    If VBA.Len(CStr(v)) > 0 Then
        getStringValue = v
    Else
        GoTo failString
    End If
Exit Function
failString:
    Err.Raise 5, MODULE_NAME, "String value or non zero length string required!"
End Function
