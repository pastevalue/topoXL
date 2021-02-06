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
'' UDF module used used to store CL (Centerline) related functions
'' "userCLs" variable declared in this module stores all the initialized
'' CLs in a CLs object
''========================================================================

'@Folder("TopoXL.UDF")

Option Explicit
Public userCLs As CLs

' Searches all worksheets for tables which can be used to initialize a Center Line
' If errors are found in the input table (missing columns, missing values,...),
' center line is not initialized
' Initilized center lines are stored in the userCLs variable of this module
Public Sub initCLs()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim tmpCL As CL
    
    Set UDF_CL.userCLs = New CLs
    'loop all worksheets
    For Each ws In ThisWorkbook.Worksheets
        'loop all tables
        For Each tbl In ws.ListObjects
            Set tmpCL = FactoryCL.newCLtbl(tbl)
            If Not tmpCL Is Nothing Then
                UDF_CL.userCLs.addCL tmpCL
            End If
        Next tbl
    Next ws
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
    Set tmpCL = userCLs.getCL(clName)
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
    Set tmpCL = userCLs.getCL(clName)
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
    Set tmpCL = userCLs.getCL(clName)
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







