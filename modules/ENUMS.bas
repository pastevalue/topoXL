Attribute VB_Name = "ENUMS"
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
''    Modules: UDF_ToAcad
''    Classes: None
''Calls:
''    Modules: None
''    Classes: None
''=======================================================
Option Explicit
Option Private Module
Public Enum AXIS_ORDER
    AXIS_NONE = 0
    AXIS_XY = 1
    AXIS_YX = 2
    AXIS_XYZ = 3
    AXIS_XZY = 4
    AXIS_YXZ = 5
    AXIS_YZX = 6
    AXIS_ZXY = 7
    AXIS_ZYX = 8
    AXIS_COUNT = 9
End Enum

Public Function axisOrderToString(axisOrder As AXIS_ORDER)
    Select Case axisOrder
    Case AXIS_XY
        axisOrderToString = "XY"
    Case AXIS_YX
        axisOrderToString = "YX"
    Case AXIS_XYZ
        axisOrderToString = "XYZ"
    Case AXIS_XZY
        axisOrderToString = "XZY"
    Case AXIS_YXZ
        axisOrderToString = "YXZ"
    Case AXIS_YZX
        axisOrderToString = "YZX"
    Case AXIS_ZXY
        axisOrderToString = "ZXY"
    Case AXIS_ZYX
        axisOrderToString = "ZYX"
    Case Else
        axisOrderToString = "NONE"
    End Select
End Function

'returns an array of AXIS_ORDER strings
Public Function getAxisOrderStringArray() As Variant
    Dim i As AXIS_ORDER
    Dim outArray(0 To AXIS_ORDER.AXIS_COUNT - 2) As String
    For i = 1 To AXIS_ORDER.AXIS_COUNT - 1
        outArray(i - 1) = axisOrderToString(i)
    Next i
    getAxisOrderStringArray = outArray
End Function

'get AXIS_ORDER from string
Public Function getAxisOrderFromString(s As String) As RL_INPUT_TYPE
    Dim arr As Variant
    Dim i As Integer
    
    arr = getAxisOrderStringArray
    For i = 0 To UBound(arr)
        If arr(i) = s Then
            getAxisOrderFromString = i + 1
            Exit Function
        End If
    Next i
    getAxisOrderFromString = AXIS_NONE
End Function

Public Function getAxisPosition(a As String, axisOrder As AXIS_ORDER)
    getAxisPosition = InStr(1, axisOrderToString(axisOrder), a)
End Function

