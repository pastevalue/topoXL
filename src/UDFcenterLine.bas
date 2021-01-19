Attribute VB_Name = "UDFcenterLine"
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
'' UDF module used used to store CL (Centerline) related functions
'' "userCLs" variable declared in this module stores all the initialized
'' CLs in a CLs object
''========================================================================

'@Folder("TopoXL.UDF")

Option Explicit
Public userCLs As CLs

Public Sub initCL()
    Dim WS As Worksheet
    Dim tbl As ListObject
    Dim tmpCL As CL
    
    Set UDFcenterLine.userCLs = New CLs
    'loop all worksheets
    For Each WS In ThisWorkbook.Worksheets
        'loop all tables
        For Each tbl In WS.ListObjects
            Set tmpCL = FactoryCL.newCLtbl(tbl)
            If Not tmpCL Is Nothing Then
                UDFcenterLine.userCLs.addCL tmpCL
            End If
        Next tbl
    Next WS
End Sub

Public Function clPntByMeasOffset(ByVal clName As String, _
                                  ByVal measure As Double, _
                                  ByVal offset As Double) As Variant
    Application.Volatile False
    
    Dim tmpCL As CL
    Set tmpCL = userCLs.getCL(clName)
    If tmpCL Is Nothing Then
        clPntByMeasOffset = CVErr(xlErrNA) ' CL name not found
    Else
        Dim tmpP As Point
        Set tmpP = tmpCL.calcPointByMeasOffset(measure, offset)
        If tmpP Is Nothing Then
            clPntByMeasOffset = CVErr(xlErrNum) ' Measure out of range
        Else
            clPntByMeasOffset = UDFcenterLine.PointToXLarray(tmpP) ' Return calculated coordinates
        End If
    End If
End Function

Public Function clMeasOffsetOfPnt(ByVal clName As String, _
                                  ByVal x As Double, ByVal y As Double) As Variant
    Application.Volatile False
    
    Dim mo As MeasOffset
    Set mo = userCLs.getCL(clName).calcMeasOffsetOfPoint(x, y)
    If mo Is Nothing Then
        clMeasOffsetOfPnt = CVErr(xlErrNA)
    Else
        clMeasOffsetOfPnt = MeasOffsetToXLarray(mo)
    End If
End Function

Private Function PointToXLarray(p As Point) As Variant
    Dim result(1 To 1, 1) As Double
    result(1, 0) = p.x
    result(1, 1) = p.y
    PointToXLarray = result
End Function

Private Function MeasOffsetToXLarray(mo As MeasOffset) As Variant
    Dim result(1 To 1, 1) As Double
    result(1, 0) = mo.m
    result(1, 1) = mo.o
    MeasOffsetToXLarray = result
End Function

' TODO: create a function which returns values to XLarray. To replace PointToXLarray and MeasOffsetToXLarray






