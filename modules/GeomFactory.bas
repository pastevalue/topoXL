Attribute VB_Name = "GeomFactory"
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
'' Description:
'' Factory module used to create geometry related
'' classes (Point, LineSegment etc.)
''=======================================================

'@Folder("TopoXL.geom")

' Creates a Point from a pair of grid coordinates
Public Function NewPoint(ByVal X As Double, ByVal Y As Double) As Point
  Set NewPoint = New Point
  NewPoint.Init X, Y
End Function

' Creates a Point from Variant values
'
' Returns:
'  A new Point object
'  Nothing if conversion of Variant to Double fails
Public Function NewPointFromVariant(ByVal X As Variant, ByVal Y As Variant) As Point
    On Error GoTo ErrHandler
    Set NewPointFromVariant = NewPoint(CDbl(X), CDbl(Y))
    Exit Function
ErrHandler:
    Set NewPointFromVariant = Nothing
End Function

' Creates a MeasOffset from a measure distance and an offset
Public Function NewMeasOffset(ByVal m As Double, ByVal o As Double) As MeasOffset
  Set NewMeasOffset = New MeasOffset
  NewMeasOffset.Init m, o
End Function

' Creates a MeasOffset from Variant values
'
' Returns:
'  A new MeasOffset object
'  Nothing if conversion of Variant to Double fails
Public Function NewMeasOffsetFromVariant(ByVal m As Variant, ByVal o As Variant) As MeasOffset
    On Error GoTo ErrHandler
    Set NewMeasOffsetFromVariant = NewMeasOffset(CDbl(m), CDbl(o))
    Exit Function
ErrHandler:
    Set NewMeasOffsetFromVariant = Nothing
End Function


