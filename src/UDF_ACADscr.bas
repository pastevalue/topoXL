Attribute VB_Name = "UDF_ACADscr"
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
'' UDF module used used to store Acad script generation functions
'' The output of these functions should be used as input for Autocad scripts
''========================================================================

'@Folder("TopoXL.UDF")
Option Explicit

Public Function acadScrPnt(ByRef arr As Variant) As Variant
    Application.Volatile False
        
    ' Case input is Excel Range
    If VBA.TypeName(arr) = "Range" Then
        If arr.Areas.count > 1 Then
            GoTo failInput
        Else
            arr = arr.Value2
        End If
    End If
    
    ' Case input is array (array may be result of the Range.Value2 content from above)
    If VBA.IsArray(arr) Then
        On Error GoTo failInput
        acadScrPnt = LibAcadScr.pnt(arr)
    Else
        GoTo failInput
    End If
    Exit Function
failInput:
    acadScrPnt = CVErr(xlErrNum)
    Exit Function
End Function


Public Function acadScrPLine(ByRef arr As Variant) As Variant
    Application.Volatile False
        
    ' Case input is Excel Range
    If VBA.TypeName(arr) = "Range" Then
        If arr.Areas.count > 1 Then
            GoTo failInput
        Else
            arr = arr.Value2
        End If
    End If
    
    ' Case input is array (array may be result of the Range.Value2 content from above)
    If VBA.IsArray(arr) Then
        On Error GoTo failInput
        acadScrPLine = LibAcadScr.pline(arr)
    Else
        GoTo failInput
    End If
    Exit Function
failInput:
    acadScrPLine = CVErr(xlErrNum)
    Exit Function
End Function

Public Function acadScrInsBl(ByRef arr As Variant) As Variant
    Application.Volatile False
        
    ' Case input is Excel Range
    If VBA.TypeName(arr) = "Range" Then
        If arr.Areas.count > 1 Then
            GoTo failInput
        Else
            arr = arr.Value2
        End If
    End If
    
    ' Case input is array (array may be result of the Range.Value2 content from above)
    If VBA.IsArray(arr) Then
        On Error GoTo failInput
        acadScrInsBl = LibAcadScr.blkInsert(arr)
    Else
        GoTo failInput
    End If
    Exit Function
failInput:
    acadScrInsBl = CVErr(xlErrNum)
    Exit Function
End Function

Public Function acadScrSText(ByRef arr As Variant) As Variant
    Application.Volatile False
        
    ' Case input is Excel Range
    If VBA.TypeName(arr) = "Range" Then
        If arr.Areas.count > 1 Then
            GoTo failInput
        Else
            arr = arr.Value2
        End If
    End If
    
    ' Case input is array (array may be result of the Range.Value2 content from above)
    If VBA.IsArray(arr) Then
        On Error GoTo failInput
        acadScrSText = LibAcadScr.sText(arr)
    Else
        GoTo failInput
    End If
    Exit Function
failInput:
    acadScrSText = CVErr(xlErrNum)
    Exit Function
End Function

Public Function acadScrChngLyr(ByVal lyr As Variant) As Variant
    On Error GoTo failInput
    acadScrChngLyr = LibAcadScr.chngLayer(lyr)
    Exit Function
failInput:
    acadScrChngLyr = CVErr(xlErrNum)
    Exit Function
End Function
