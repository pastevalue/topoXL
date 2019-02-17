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

'@Folder("TopoXL.geom")

' Creates a Point from a pair of grid coordinates
Public Function NewPoint(ByVal X As Double, ByVal Y As Double) As Point
  Set NewPoint = New Point
  NewPoint.Init X, Y
End Function

'Creates a Point from Variant values
Public Function NewPointFromVariant(ByVal X As Variant, ByVal Y As Variant) As Point
    On Error GoTo ErrHandler
    Set NewPointFromVariant = NewPoint(CDbl(X), CDbl(Y))
    Exit Function
ErrHandler:
    Set NewPointFromVariant = Nothing
End Function
