Attribute VB_Name = "IC_ENUMS"
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
''    Classes: Interval
''Calls:
''    Modules: None
''    Classes: None
''=======================================================
Option Explicit
Option Private Module

'Interval Collection input type strings
Private Const STR_INPUT_NONE As String = "None"
Private Const STR_INPUT_STV_ENV_VAL As String = "StvEnvVal"

'Interval Collection input types enum
Public Enum IC_INPUT_TYPE
    INPUT_NONE = 0
    INPUT_STV_ENV_VAL = 1
    INPUT_COUNT = 2
End Enum

'Interval Collection Interval part strings
Private Const STR_PART_NONE As String = "None"
Private Const STR_PART_STV As String = "Start Value"
Private Const STR_PART_ENV As String = "End Value"
Private Const STR_PART_VAL As String = "Value"

'Interval Collection Interval parts enum
Public Enum IC_INTVL_PARTS
    PART_NONE = 0
    PART_STV = 1
    PART_ENV = 2
    PART_VAL = 3
    PART_COUNT = 4
End Enum

'get IC_INPUT_TYPE string
Public Function icIntvlInputTypeToString(intvlPart As IC_INTVL_PARTS) As String
    Select Case intvlPart
        Case INPUT_STV_ENV_VAL
            icIntvlInputTypeToString = STR_INPUT_STV_ENV_VAL
        Case Else
            icIntvlInputTypeToString = STR_INPUT_NONE
    End Select
End Function

'get IC_INPUT_TYPE from string
Public Function icIntvlInputTypeFromString(s As String) As IC_INPUT_TYPE
    Dim arr As Variant
    Dim i As Integer
    
    arr = getInputStringArray
    For i = 0 To UBound(arr)
        If arr(i) = s Then
            icIntvlInputTypeFromString = i + 1
            Exit Function
        End If
    Next i
    icIntvlInputTypeFromString = INPUT_NONE
End Function

'get IC_INTVL_PARTS string
Public Function icIntvlPartToString(inputType As IC_INPUT_TYPE) As String
    Select Case inputType
        Case PART_STV
            icIntvlPartToString = STR_PART_STV
        Case PART_ENV
            icIntvlPartToString = STR_PART_ENV
        Case PART_VAL
            icIntvlPartToString = STR_PART_VAL
        Case Else
            icIntvlPartToString = STR_PART_NONE
    End Select
End Function
'returns an array of Interval Collection Interval Parts strings
Public Function getIntvlPartStringArray() As Variant
    Dim i As IC_INTVL_PARTS
    Dim outArray(0 To IC_INTVL_PARTS.PART_COUNT - 2) As String
    For i = 1 To IC_INTVL_PARTS.PART_COUNT - 1
        outArray(i - 1) = icIntvlPartToString(i)
    Next i
    getIntvlPartStringArray = outArray
End Function

'returns an array of Interval Collection Interval Input Type strings
Public Function getInputStringArray() As Variant
    Dim i As IC_INPUT_TYPE
    Dim outArray(0 To IC_INPUT_TYPE.INPUT_COUNT - 2) As String
    For i = 1 To IC_INPUT_TYPE.INPUT_COUNT - 1
        outArray(i - 1) = icIntvlInputTypeToString(i)
    Next i
    getInputStringArray = outArray
End Function



