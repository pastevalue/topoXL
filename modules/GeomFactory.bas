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

''========================================================================
'' Description
'' Factory module used to create geometry related
'' classes instances (Point, LineSegment etc.)
''========================================================================

'@Folder("TopoXL.geom")

Option Explicit
Option Private Module

' Creates a Point from a pair of grid coordinates
Public Function NewPoint(ByVal x As Double, ByVal y As Double) As Point
    Set NewPoint = New Point
    NewPoint.Init x, y
End Function

' Creates a Point from Variant values
' Returns:
'   - a new Point object
'   - Nothing if conversion of Variant to Double fails
Public Function NewPointFromVariant(ByVal x As Variant, ByVal y As Variant) As Point
    On Error GoTo ErrHandler
    Set NewPointFromVariant = NewPoint(CDbl(x), CDbl(y))
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
' Returns:
'  - a new MeasOffset object
'  - Nothing if conversion of Variant to Double fails
Public Function NewMeasOffsetFromVariant(ByVal m As Variant, ByVal o As Variant) As MeasOffset
    On Error GoTo ErrHandler
    Set NewMeasOffsetFromVariant = NewMeasOffset(CDbl(m), CDbl(o))
    Exit Function
ErrHandler:
    Set NewMeasOffsetFromVariant = Nothing
End Function

' Creates a LineSegment from two sets of grid coordinates
Public Function NewLineSegment(ByVal x1 As Double, ByVal y1 As Double, _
                               ByVal x2 As Double, ByVal y2 As Double) As LineSegment
    Set NewLineSegment = New LineSegment
    NewLineSegment.Init x1, y1, x2, y2
End Function

' Creates a LineSegment from two sets of grid coordinates defined
' as Variant values
' Returns:
'   - a new LineSegment object
'   - Nothing if conversion of Variant to Double fails
Public Function NewLineSegmentFromVariant(ByVal x1 As Variant, ByVal y1 As Variant, _
                                          ByVal x2 As Variant, ByVal y2 As Variant) As LineSegment
    On Error GoTo ErrHandler
    Set NewLineSegmentFromVariant = NewLineSegment(CDbl(x1), CDbl(y1), CDbl(x2), CDbl(y2))
    Exit Function
ErrHandler:
    Set NewLineSegmentFromVariant = Nothing
End Function


