Attribute VB_Name = "UDF_DataManipulation"
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
''    Modules: rangeFunctions, ArrayFunctions
''    Classes: None
''=======================================================
Option Explicit
'Public Function dmFilterColumsFromRange(rng As Range, ParamArray colIndex() As Variant) As Variant
'    Dim tempIndexColl As Collection
'    Dim tempResult As New Collection
'    Dim t As Variant
'    Dim i As Long
'    Dim r As Long
'
'    t = colIndex
'    Set tempIndexColl = rangeFunctions.valuesToCollection(t)
'    'return error if no matching column is found for column index
'    For i = 1 To tempIndexColl.count
'        If tempIndexColl.item(i) > rng.columns.count Then GoTo failRef
'    Next i
'
'    For r = 1 To rng.rows.count
'        For i = 1 To tempIndexColl.count
'            tempResult.Add (rng.Cells(r, tempIndexColl.item(i)))
'        Next i
'    Next r
'
'    dmFilterColumsFromRange = rangeFunctions.collectionTo2Darray(tempIndexColl.count, tempResult)
'Exit Function
'failRef:
'    dmFilterColumsFromRange = CVErr(xlErrRef)
'End Function

'Public Function dmFilterRowsFromRange(rng As Range, ParamArray rowIndex() As Variant) As Variant
'    Dim tempIndexRow As Collection
'    Dim tempResult As New Collection
'    Dim t As Variant
'    Dim i As Long
'    Dim r As Long
'
'    t = rowIndex
'    Set tempIndexRow = rangeFunctions.valuesToCollection(t)
'    'return error if no matching row is found for row index
'    For i = 1 To tempIndexRow.count
'        If tempIndexRow.item(i) > rng.rows.count Then GoTo failRef
'    Next i
'
'    For r = 1 To tempIndexRow.count
'        For i = 1 To rng.columns.count
'            tempResult.Add (rng.Cells(tempIndexRow.item(r), i))
'        Next i
'    Next r
'
'    dmFilterRowsFromRange = rangeFunctions.collectionTo2Darray(rng.columns.count, tempResult)
'Exit Function
'failRef:
'    dmFilterRowsFromRange = CVErr(xlErrRef)
'End Function

Public Function dmFilterRowsAndColumns(v As Variant, rows As Variant, columns As Variant) As Variant
    Dim c As Integer
    Dim r As Integer
    Dim tempArr As Variant
    Dim temprows As New Collection
    Dim tempResult As New Collection
    Dim tempcolumns As New Collection
    Dim i As Integer, j As Integer
    
    If TypeName(v) = "Range" Then
        c = v.columns.count
        r = v.rows.count
    ElseIf IsArray(v) Then
        If arrayNumberOfDimensions(v) <> 2 Then
            GoTo failInput
        Else
            c = UBound(v, 2) - LBound(v, 2) + 1
            r = UBound(v, 1) - LBound(v, 1) + 1
        End If
    ElseIf IsNumeric(v) Or TypeName(v) = "String" Then
        c = 1
        r = 1
    Else
        GoTo failRef
    End If
        
    'validate the rows index "rows"
    If IsNumeric(rows) Then
        If rows = 0 Then
            For i = 1 To r
                temprows.Add (i)
            Next i
        Else
            If rows > r Or rows < 1 Then
                GoTo failRef
            Else
                temprows.Add (rows)
            End If
        End If
    Else
        Set temprows = rangeFunctions.valuesToCollection(rows)
        If Not temprows.count > 0 Then GoTo failInput
        For i = 1 To temprows.count
            If temprows.item(i) > r Or temprows.item(i) < 1 Then GoTo failRef
        Next i
    End If
    'validate the columns index "columns"
    If IsNumeric(columns) Then
        If columns = 0 Then
            For i = 1 To c
                tempcolumns.Add (i)
            Next i
        Else
            If columns > c Or columns < 1 Then
                GoTo failRef
            Else
                tempcolumns.Add (columns)
            End If
        End If
    Else
        Set tempcolumns = rangeFunctions.valuesToCollection(columns)
        If Not tempcolumns.count > 0 Then GoTo failInput
        For i = 1 To tempcolumns.count
            If tempcolumns.item(i) > c Or tempcolumns.item(i) < 1 Then GoTo failRef
        Next i
    End If
    
    tempArr = rangeFunctions.valuesTo2Darray(c, v)
    For i = 1 To temprows.count
        For j = 1 To tempcolumns.count
            tempResult.Add (tempArr(temprows.item(i) - 1, tempcolumns.item(j) - 1))
        Next j
    Next i
    
    dmFilterRowsAndColumns = rangeFunctions.collectionTo2Darray(tempcolumns.count, tempResult)
Exit Function
failInput:
    dmFilterRowsAndColumns = CVErr(xlErrNA)
    Exit Function
failRef:
    dmFilterRowsAndColumns = CVErr(xlErrRef)
End Function

Public Function dmFilterOutValuesTo2DArray(filter As Variant, colCount As Integer, ParamArray values() As Variant) As Variant
    Dim t As Variant
    Dim tempColl As Collection
    t = values
    
    Set tempColl = New Collection
    Set tempColl = rangeFunctions.valuesToCollection(t)
    Set tempColl = rangeFunctions.filterOutValueFromCollection(filter, tempColl)
    
    'return error if the number of values  is not divisible by colCount parameter
    If tempColl.count Mod colCount <> 0 Then GoTo failInput

    'return result
    dmFilterOutValuesTo2DArray = rangeFunctions.collectionTo2Darray(colCount, tempColl)
    Exit Function
failInput:
    dmFilterOutValuesTo2DArray = CVErr(xlErrNum)
    Exit Function
End Function

'Gets the first numeric value from multiple ranges.
'ranges parameter - an array of ranges that will be parsed cell by cell to search first numeric value
Public Function dmGetFirstNumericValue(ParamArray v() As Variant) As Variant
    Dim tempValue As Variant
    Dim values As Collection
    Dim t As Variant
        
    t = v
    Set values = rangeFunctions.valuesToCollection(t)
    
    For Each tempValue In values
        If Application.WorksheetFunction.IsNumber(tempValue) Then
            dmGetFirstNumericValue = tempValue
            Exit Function
        End If
    Next
    dmGetFirstNumericValue = CVErr(xlErrNA)
End Function

'Gets the first non null value from multiple ranges.
'ranges parameter - an array of ranges that will be parsed cell by cell to search first non null value
Public Function dmGetFirstNonNullValue(ParamArray v() As Variant) As Variant
    Dim tempValue As Variant
    Dim values As Collection
    Dim t As Variant
    
    t = v
    Set values = valuesToCollection(t)
    
    For Each tempValue In values
        If tempValue <> vbNullString Then
            dmGetFirstNonNullValue = tempValue
            Exit Function
        End If
    Next
    dmGetFirstNonNullValue = vbNullString
End Function

'Returns specified values in reversed order
'Parameters:
'   -values: a variant that contains the value(s) that will be returned in reversed order
'Returns #Null! error if:
'   - values parameter is empty (has no values).
Public Function dmReverseGroupedValuesTo1DArray(groupSize As Integer, ParamArray values() As Variant) As Variant
    Dim tempColl As Collection
    Dim tempValue As Variant
    Dim t As Variant
    Dim result() As Variant
    
    If groupSize < 1 Then GoTo failInput
    
    t = values
    Set tempColl = reverseCollection(groupSize, valuesToCollection(t))
    ReDim result(tempColl.count - 1)
    Dim i As Long

    For i = 0 To tempColl.count - 1
        result(i) = tempColl.item(i + 1)
    Next i
    dmReverseGroupedValuesTo1DArray = result
Exit Function
failInput:
    dmReverseGroupedValuesTo1DArray = CVErr(xlErrRef)
End Function

'Returns specified values in a 2D array format
'Parameters:
'   -colCount: number od columns that the returned array will have
'   -values: a variant that contains the value(s) that will be returned in a 2D array
'Returns #Null! error if:
'   - there are no values in parameter "values"
'   - number of values is not divisible by "colCount" parameter.
Public Function dmValuesTo2DArray(colCount As Integer, ParamArray values() As Variant) As Variant
    Dim t As Variant
    Dim tempColl As Collection
    
    t = values
    Set tempColl = valuesToCollection(t)
    
    'return error if there are no values in values parameter
    If tempColl.count = 0 Then
        dmValuesTo2DArray = CVErr(xlErrRef)
        Exit Function
    End If
    
    'return error if the number of values is not divisible by colCount parameter
    If tempColl.count Mod colCount <> 0 Then
        dmValuesTo2DArray = CVErr(xlErrNum)
        Exit Function
    End If
     
    'transform ranges to array
    dmValuesTo2DArray = valuesTo2Darray(colCount, t)
End Function
'TODO - de modificat comentariile
'Returns values from multiple ranges in a separated value format
'separator parameter - specifies the separator used between values
'ranges parameter - ranges that will be parsed cell by cell and returned by this function as a string in separated values format
Public Function dmValuesToSeparatedString(separator As String, ParamArray values() As Variant) As String
    Dim t As Variant
    Dim val As Variant
    Dim tempColl As New Collection
    Dim result As String
    
    t = values
    Set tempColl = rangeFunctions.valuesToCollection(t)
    
    'concatenate values to result string
    result = vbNullString
    For Each val In tempColl
        result = result & separator & val
    Next val
    
    'IF values has been added to result string THEN remove first separator
    If result <> vbNullString Then
        result = Right(result, Len(result) - Len(separator))
    End If
    dmValuesToSeparatedString = result
End Function


Function dmUniqueValues(ParamArray values() As Variant) As Variant

Dim Dict As Object
Dim i As Long, j As Long, NumRows As Long, NumCols As Long
Dim tempColl As Collection
Dim t As Variant

t = values
Set tempColl = rangeFunctions.valuesToCollection(t)

    'put unique data elements in a dictionay
    Set Dict = CreateObject("Scripting.Dictionary")
    
    For i = 1 To tempColl.count
        Dict(tempColl(i)) = 1
    Next i
    dmUniqueValues = WorksheetFunction.Transpose(Dict.keys)

End Function

Public Function dmValuesToArrayString(itemLen As Integer, ParamArray values() As Variant) As Variant
    Dim outArray() As String
    Dim tempColl As Collection
    Dim tempString As String
    Dim t As Variant
    
    tempString = vbNullString
    t = values
    Set tempColl = rangeFunctions.valuesToCollection(t)
    Set tempColl = rangeFunctions.filterOutValueFromCollection("", tempColl)
    
    Dim i As Long
    For i = 1 To tempColl.count
        tempString = tempString & tempColl.item(i) & Chr(10)
    Next i
        
    ReDim outArray(1 To 1, 1 To -Int(-Len(tempString) / itemLen))
    
    For i = 1 To UBound(outArray, 2)
        outArray(1, i) = Mid(tempString, (i - 1) * itemLen + 1, itemLen)
    Next i
    dmValuesToArrayString = outArray
End Function

