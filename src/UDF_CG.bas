Attribute VB_Name = "UDF_CG"
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
'' UDF module used used to store COGO functions
''========================================================================

'@Folder("TopoXL.UDF")

Option Explicit

Public Function cgDist2D(ByVal x1 As Double, ByVal y1 As Double, _
                         ByVal x2 As Double, ByVal y2 As Double) As Variant
    Application.Volatile False
        
    cgDist2D = LibGeom.dist2D(x1, y1, x2, y2)
End Function

Public Function cgDist3D(ByVal x1 As Double, ByVal y1 As Double, ByVal z1 As Double, _
                         ByVal x2 As Double, ByVal y2 As Double, ByVal z2 As Double) As Variant
    Application.Volatile False
        
    cgDist3D = LibGeom.dist3D(x1, y1, z1, x2, y2, z2)
End Function

Public Function cgTheta(ByVal x1 As Double, ByVal y1 As Double, _
                        ByVal x2 As Double, ByVal y2 As Double) As Variant
    Application.Volatile False
    On Error GoTo failInput
    cgTheta = LibGeom.Atn2(x2 - x1, y2 - y1)
    
    Exit Function
failInput:
    cgTheta = CVErr(xlErrNum)
End Function

Public Function cgSide(ByVal x1 As Double, ByVal y1 As Double, _
                       ByVal x2 As Double, ByVal y2 As Double, _
                       ByVal x As Double, ByVal y As Double) As Variant
    Application.Volatile False
    On Error GoTo failInput
    cgSide = LibGeom.orientationIndex(x1, y1, x2, y2, x, y)
    Exit Function
failInput:
    cgSide = CVErr(xlErrNum)
End Function

Public Function cgCooInBB(ByVal x As Double, ByVal y As Double, _
                          ByVal x1 As Double, ByVal y1 As Double, _
                          ByVal x2 As Double, ByVal y2 As Double) As Variant
    Application.Volatile False
        
    cgCooInBB = LibGeom.cooInBB(x, y, x1, y1, x2, y2)
End Function

Public Function cgIntLbyCooAndTh(ByVal x1 As Double, ByVal y1 As Double, ByVal theta1 As Double, _
                                 ByVal x2 As Double, ByVal y2 As Double, ByVal theta2 As Double) As Variant
    Application.Volatile False
    
    Dim p As Point
    Set p = LibGeom.intLbyThAndCoo(x1, y1, theta1, x2, y2, theta2)
    If p Is Nothing Then
        cgIntLbyCooAndTh = CVErr(xlErrNA)
    Else
        cgIntLbyCooAndTh = p.toArrayXY()
    End If
End Function

Public Function cgIntLSbyCoo(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double, _
                             ByVal x3 As Double, ByVal y3 As Double, ByVal x4 As Double, ByVal y4 As Double) As Variant
    Application.Volatile False
    
    Dim p As Point
    Set p = LibGeom.intLSbyCoo(x1, y1, x2, y2, x3, y3, x4, y4)
    If p Is Nothing Then
        cgIntLSbyCoo = CVErr(xlErrNA)
    Else
        cgIntLSbyCoo = p.toArrayXY()
    End If
End Function

Public Function cgIntLbyCooAndDs(ByVal x1 As Double, ByVal y1 As Double, ByVal dx1 As Double, ByVal dy1 As Double, _
                                 ByVal x2 As Double, ByVal y2 As Double, ByVal dx2 As Double, ByVal dy2 As Double) As Variant
    Application.Volatile False
    
    Dim p As Point
    Set p = LibGeom.intLbyCooAndDs(x1, y1, dx1, dy1, x2, y2, dx2, dy2)
    If p Is Nothing Then
        cgIntLbyCooAndDs = CVErr(xlErrNA)
    Else
        cgIntLbyCooAndDs = p.toArrayXY()
    End If
End Function

Public Function cgIntPLbyCoo(ByVal coos1 As Variant, ByVal coos2 As Variant) As Variant
    Application.Volatile False
    
    If Not LibUDF.getInAs2DArray(coos1) Or Not LibUDF.getInAs2DArray(coos2) Then
        GoTo failInput
    End If
    
    On Error GoTo failInput
    Dim pc As PointColl
    Set pc = LibGeom.intPLbyCoo(coos1, coos2)
    If pc.count > 0 Then
        cgIntPLbyCoo = pc.toArrayXY(True)
    Else
        cgIntPLbyCoo = CVErr(xlErrNA)   ' no intersection found
    End If
    Exit Function
failInput:
    cgIntPLbyCoo = CVErr(xlErrNum)
End Function

Public Function cgExtTrimLS(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double, _
                            ByVal length As Double, ByVal part As Integer) As Variant
                          
    Application.Volatile False
    On Error GoTo failInput
    Dim pc As PointColl
    Set pc = LibGeom.extTrimLS(x1, y1, x2, y2, length, part)
    cgExtTrimLS = pc.toArrayXY()
    Exit Function
failInput:
    cgExtTrimLS = CVErr(xlErrNum)
End Function

Public Function cgOffsetLS(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double, _
                           ByVal offset As Double) As Variant
                          
    Application.Volatile False
    On Error GoTo failInput
    Dim pc As PointColl
    Set pc = LibGeom.offsetLS(x1, y1, x2, y2, offset)
    cgOffsetLS = pc.toArrayXY()
    Exit Function
failInput:
    cgOffsetLS = CVErr(xlErrNum)
End Function

Public Function cgAreaByCoo(ByVal coos As Variant) As Variant
    Application.Volatile False
    
    If Not LibUDF.getInAs2DArray(coos) Then
        GoTo failInput
    End If
    
    On Error GoTo failInput
    cgAreaByCoo = LibGeom.areaByCoo(coos)
    Exit Function
failInput:
    cgAreaByCoo = CVErr(xlErrNum)
End Function







