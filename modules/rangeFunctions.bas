Attribute VB_Name = "rangeFunctions"
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
''    Modules: UDF_DataManipulation, UDF_COGO
''    Classes: None
''Calls:
''    Modules: ArrayFunctions
''    Classes: None
''=======================================================
Option Private Module
Option Explicit

'Returns a 2D array of values
'Parameters:
'   -v: a variant that contain the value(s) that will be used to populate the returned collection.
'       This can be an Excel Range Object, simple values (numbers or strings) or an (multidimensional) array of these types.
'       Other types than those specified will be ignored.
'   -colCount: number of columns that the result array will have.
'Result array will contain all values of input parameter. Arrays of values will be parsed in their storage order.
'Excel Ranges will be parsed in wise order (lines then column).
'The size of the array will always fit all values. If the total number of values is not divisible by the colCount parameter then the last values of the array will be Empty.
Public Function valuesTo2Darray(ByVal colCount As Integer, ByVal v As Variant) As Variant
    valuesTo2Darray = collectionTo2Darray(colCount, valuesToCollection(v))
End Function

'Returns a 2D array with based on values contained in the specified collection
'Parameters:
'   -colCount: number of columns that the result array will have.
'   -c: collection that contains the values that will be added to the result
'The size of the array will always fit all values. If the total number of values is not divisible by the colCount parameter then the last values of the array will be Empty.
Public Function collectionTo2Darray(ByVal colCount As Integer, c As Collection) As Variant()
    Dim i As Long
    Dim result() As Variant: ReDim result(-Int(-(c.count - colCount) / colCount), colCount - 1)
     
    Dim t As Variant
    For Each t In c
        result(Int(i / colCount), i Mod colCount) = t
        i = i + 1
    Next t
    collectionTo2Darray = result
End Function

'Returns a collection from specified values
'Parameters:
'   -v: a variant that contain the value(s) that will be used to populate the returned collection.
'       This can be an Excel Range Object, simple values (numbers or strings) or an (multidimensional) array of these types.
'       Other types than those specified will be ignored.
'Result collection will contain all values of the input parameter. Arrays of values will be parsed in their storage order.
'Excel Ranges will be parsed in wise order (lines then column)
Public Function valuesToCollection(v As Variant) As Collection
    Set valuesToCollection = New Collection
    
    'IF input is an Excel Range THEN add all cells values to result collection
    If TypeName(v) = "Range" Then
        Dim r  As Long
        Dim c As Long
        Dim tempRange As Range
        Set tempRange = v
        For r = 1 To tempRange.rows.count
            For c = 1 To tempRange.columns.count
                valuesToCollection.Add tempRange.Cells(r, c).value
            Next c
        Next r
    'IF input is an array THEN call again this function for each item of the array
    ElseIf IsArray(v) Then
        Dim i As Long
        Dim item As Variant
        For Each item In arrayTo1DrowMajorOrder(v)
            Dim tempColl As New Collection
            Dim j As Long
            Set tempColl = valuesToCollection(item)
            For j = 1 To tempColl.count
                valuesToCollection.Add tempColl.item(j)
            Next j
        Next item
    'IF input is value (number or string) THEN add values to result collection
    ElseIf IsNumeric(v) Or TypeName(v) = "String" Then
        valuesToCollection.Add v
    'IF none of situation above is meet THEN do nothing
    Else
    
    End If
End Function

Public Function filterOutValueFromCollection(filter As Variant, ByVal c As Collection) As Collection
    Dim itm As Variant
    Dim filterItem As Variant
    Dim result As New Collection
    Dim filterColl As New Collection
    Dim found As Boolean
    found = False
    Set filterColl = valuesToCollection(filter)
    For Each itm In c
        For Each filterItem In filterColl
            If itm = filterItem Then
                found = True
                Exit For
            End If
        Next filterItem
        If Not found Then result.Add itm
        found = False
    Next itm
    Set filterOutValueFromCollection = result
End Function


'TODO - de adaugat comentarii pentru functie
'TODO - functia nu mai este valabila atunci cand este apelata si cu valori sau vectori in loc de range-uri.
'Peste tot pe unde este apelata trebuie folosita  functia valuesToCollection. Nr de elemente este dat de proprietatea count a clasei Collection
Public Function rangesCellCount(ranges As Variant) As Long
    Dim cellCount As Long
    Dim tempRange As Range
    cellCount = 0
    Dim i As Integer
    For i = 0 To UBound(ranges)
        Set tempRange = ranges(i)
        cellCount = cellCount + tempRange.Cells.count
    Next i
    rangesCellCount = cellCount
End Function

'Reverses a collection but allows the user to group the collection's values in groups of same size and reverse the order of groups from end to start
'Parameters:
'   -groupSize: a integer which indicates the number of values in each of the group that formes the collection.
'               Must be greater or equal to 1.
'   -c: a collection with values which will be reversed
'       The number of values in collection must be a multiple of groupSize.
'Result collection will contain all values of the input collection.
Public Function reverseCollection(groupSize As Integer, c As Collection) As Collection
    Dim result As New Collection
    Dim tempValue As Variant
    Dim count As Long
    
    If groupSize < 1 Then GoTo failInput
    count = c.count
    If count < 1 Or count Mod groupSize <> 0 Then
        GoTo failInput
    End If
    
    Dim i As Long
    Dim j As Long
    For i = count - groupSize + 1 To 1 Step -groupSize
        For j = i To i + groupSize - 1
            result.Add c.item(j)
        Next j
    Next i
    Set reverseCollection = result
Exit Function
failInput:
    Set reverseCollection = CVErr(xlErrNum)
End Function
