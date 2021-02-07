Attribute VB_Name = "LibUDF"
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
'' Stores general functions to be used by any UDF_ module
'' These functions are are designed to be called by function exposed in UDF_
'' so they should not be referenced in other contexts
''========================================================================

'@Folder("TopoXL.libs")

Option Explicit
Option Private Module

Private Const MODULE_NAME As String = "LibUDF"

' Returns True if input parameter can be used as a 2D Array object
' This function should be called by UDF_ functions only which are
' expecting an Excel Range or Array as input.
' Parameters:
'   - param: ByRef input parameter to be checked if it can be used as
'            array. The input reference is changed as follows:
'               * Range.Value2 array if Excel Range object
'               * no changes if Array object
'               * 1D arrays are changed to 2D arrays
'               * one item array if single value
' Returns:
'   - TRUE if 'param' is of type "Range", area count = 1. Changes the reference to Range.Values2
'   - TRUE if 'param' is 1D or 2D "array". If 1D array is changed to 2D
'   - TRUE if 'param' is a single value. Reference changed to a single item 2D array
'   - FALSE if 'param is of type Range with Area count > 1
'   - FALSE if 'param is of type array with dimension different from 2 or 1
'   - FALSE in any other scenario
Public Function getInAs2DArray(ByRef param As Variant) As Boolean
    ' Case input is Excel Range
    If VBA.TypeName(param) = "Range" Then
        If param.Areas.count > 1 Then
            getInAs2DArray = False
            Exit Function
        Else
            param = param.Value2
        End If
    End If
    
    ' Convert single value to 1-element 2D array
    If Not VBA.IsArray(param) Then
        'param = LibArr.valueToArray2D(param)
        param = Array(param)
    End If
    
    ' The param shoud be of type array at this stage
    Select Case LibArr.getArrayDimsCount(param)
        Case 2
            getInAs2DArray = True
            Exit Function
        Case 1
            Dim colsCount As Long
            colsCount = UBound(param) - LBound(param) + 1
            param = LibUDF.OneDArrayTo2DArray(param, colsCount)
            getInAs2DArray = True
            Exit Function
        Case Else
            getInAs2DArray = False
            Exit Function
    End Select
End Function

' Function copied from https://github.com/cristianbuse/VBA-ArrayTools/blob/master/Code%20Modules/LibArrayTools.bas
' Returns a 2D array (1 based index) based on values contained in the specified 1D array
' Parameters:
'   - arr: the 1D array that contains the values to be used
'   - columnsCount: the number of columns that the result 2D array will have
' Raises error:
'   -  5 if:
'       * input array is not 1D
'       * input array has no elements. Zero-length array (e.g.
'       * the number of columns is less than 1
' Notes:
'   - if the total Number of values is not divisible by columnsCount then the
'     extra values (last row) of the array are by default the value Empty
Public Function OneDArrayTo2DArray(ByRef arr As Variant, ByVal columnsCount As Long) As Variant()
    'Check Input
    If getArrayDimsCount(arr) <> 1 Then
        Err.Raise 5, MODULE_NAME, "Expected 1D Array"
    ElseIf LBound(arr) > UBound(arr) Then
        Err.Raise 5, MODULE_NAME, "Zero-length array. No elements"
    ElseIf columnsCount < 1 Then
        Err.Raise 5, MODULE_NAME, "Invalid Columns Count"
    End If
    
    Dim elemCount As Long: elemCount = UBound(arr) - LBound(arr) + 1
    Dim rowsCount As Long: rowsCount = -VBA.Int(-elemCount / columnsCount)
    Dim res() As Variant
    Dim i As Long: i = 0
    Dim r As Long
    Dim c As Long
    Dim v As Variant
    '
    'Populate result array
    ReDim res(1 To rowsCount, 1 To columnsCount)
    For Each v In arr
        r = i \ columnsCount
        c = i Mod columnsCount
        res(r + 1, c + 1) = v
        i = i + 1
    Next v
    '
    OneDArrayTo2DArray = res
End Function


