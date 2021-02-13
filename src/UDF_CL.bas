Attribute VB_Name = "UDF_CL"
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
'' Description
'' UDF module exposing Centerline functions
''========================================================================

'@Folder("TopoXL.UDF")

Option Explicit

Public Function clPntByMeasOffset(ByVal clName As String, _
                                  ByVal measure As Double, _
                                  ByVal offset As Double) As Variant
    Application.Volatile False
    
    Dim tmpCL As CL
    Set tmpCL = XL.getUserCLs.getCL(clName)
    If tmpCL Is Nothing Then
        clPntByMeasOffset = CVErr(xlErrNA) ' CL name not found
    Else
        Dim p As Point
        Set p = tmpCL.calcPointByMeasOffset(measure, offset)
        If p Is Nothing Then
            clPntByMeasOffset = CVErr(xlErrNum) ' Measure out of range
        Else
            clPntByMeasOffset = p.toArrayXY() ' Return calculated coordinates
        End If
    End If
End Function

Public Function clMeasOffsetOfPnt(ByVal clName As String, _
                                  ByVal x As Double, ByVal y As Double) As Variant
    Application.Volatile False
    
    Dim tmpCL As CL
    Set tmpCL = XL.getUserCLs.getCL(clName)
    If tmpCL Is Nothing Then
        clMeasOffsetOfPnt = CVErr(xlErrNA) ' CL name not found
    Else
        Dim mo As MeasOffset
        Set mo = tmpCL.calcMeasOffsetOfPoint(x, y)
        If mo Is Nothing Then
            clMeasOffsetOfPnt = CVErr(xlErrNum) ' No measure and offset for given coordinates
        Else
            clMeasOffsetOfPnt = mo.toArray() ' Return caluclated measure and offset
        End If
    End If
End Function

Public Function clXatY(ByVal clName As String, ByVal y As Double) As Variant
    Application.Volatile False
    
    Dim tmpCL As CL
    Set tmpCL = XL.getUserCLs.getCL(clName)
    If tmpCL Is Nothing Then
        clXatY = CVErr(xlErrNA)         ' CL name not found
    Else
        Dim result As Variant
        result = tmpCL.calcXatY(y)
        If IsNull(result) Then
            clXatY = CVErr(xlErrNA)     ' y value not valid
        Else
            clXatY = result             ' return calculated value
        End If
    End If
End Function

Public Function clYatX(ByVal clName As String, ByVal x As Double) As Variant
    Application.Volatile False
    
    Dim tmpCL As CL
    Set tmpCL = XL.getUserCLs.getCL(clName)
    If tmpCL Is Nothing Then
        clYatX = CVErr(xlErrNA)         ' CL name not found
    Else
        Dim result As Variant
        result = tmpCL.calcYatX(x)
        If IsNull(result) Then
            clYatX = CVErr(xlErrNA)     ' x value not valid
        Else
            clYatX = result             ' return calculated value
        End If
    End If
End Function
