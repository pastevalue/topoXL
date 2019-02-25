Attribute VB_Name = "GeomLib"
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
'' Description:
'' Stores basic geometry functions
''========================================================================

'@Folder("TopoXL.geom")

Option Explicit
Option Private Module

Public Const PI As Double = 3.14159265358979

' Returns the distance between two sets of 2D grid coordinates
Public Function Dist2D(ByVal x1 As Double, ByVal y1 As Double, _
                       ByVal x2 As Double, ByVal y2 As Double) As Double
    Dist2D = Sqr((x1 - x2) ^ 2 + (y1 - y2) ^ 2)
End Function

' Returns the angle in radians between the positive x-axis
' and the ray to the point (X,Y). The returned value
' is within range (-pi, pi]
'
' Raises error for (0,0)
Public Function Atn2(ByVal x As Double, ByVal y As Double) As Double
    Select Case x
    Case Is > 0
        Atn2 = Atn(y / x)
    Case Is < 0
        Dim tmpSign As Integer
        If y = 0 Then
            tmpSign = 1
        Else
            tmpSign = Sgn(y)
        End If
        Atn2 = Atn(y / x) + PI * tmpSign
    Case Is = 0
        If y = 0 Then
            Err.Raise 5, "Atan2 function", "Cant compute Atan2 on (0,0)"
        Else
            Atn2 = PI / 2 * Sgn(y)
        End If
    End Select
End Function

' Gets the orientation index (side) of a set of coordinates relative to a line
' Returns -1 if coordinates are on the left, 0 if the coordinates are on the line
' and +1 if coordinates are on the right
' Left and right are considered relative to the start and end coordinates of the line
Public Function GetOrientationIndex(ByVal x1 As Double, ByVal y1 As Double, _
                                    ByVal x2 As Double, ByVal y2 As Double, _
                                    ByVal x As Double, ByVal y As Double) As Integer
    GetOrientationIndex = Sgn((y2 - y1) * (x - x1) - (x2 - x1) * (y - y1))
End Function


