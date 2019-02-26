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
Public Const TWO_PI As Double = 6.28318530717959
Private Const MODULE_NAME As String = "GeomLib"

' Returns the distance between two sets of 2D grid coordinates
Public Function Dist2D(ByVal x1 As Double, ByVal y1 As Double, _
                       ByVal x2 As Double, ByVal y2 As Double) As Double
    Dist2D = Math.Sqr((x1 - x2) * (x1 - x2) + (y1 - y2) * (y1 - y2))
End Function

' Returns the arcsine of a number.
' Raises error if x is not within [-1,1] interval
Public Function ASin(ByVal x As Double) As Double
    Select Case x
    Case -1
        ASin = 6 * Math.Atn(1)
    Case 0:
        ASin = 0
    Case 1:
        ASin = 2 * Math.Atn(1)
    Case Else:
        ASin = Math.Atn(x / Math.Sqr(-x * x + 1))
    End Select
End Function

' Returns the arccosine of a number.
' Raises error if x is not within [-1,1] interval
Public Function ACos(ByVal x As Double) As Double
    Select Case x
    Case -1
        ACos = 4 * Math.Atn(1)
    Case 0:
        ACos = 2 * Math.Atn(1)
    Case 1:
        ACos = 0
    Case Else:
        ACos = Math.Atn(-x / Math.Sqr(-x * x + 1)) + 2 * Math.Atn(1)
    End Select
End Function

' Returns the angle in radians between the positive x-axis
' and the ray to the point (X,Y). The returned value
' is within range (-pi, pi]
' Raises error for (0,0)
Public Function Atn2(ByVal x As Double, ByVal y As Double) As Double
    Select Case x
    Case Is > 0
        Atn2 = Math.Atn(y / x)
    Case Is < 0
        If y < 0 Then
            Atn2 = Math.Atn(y / x) - PI
        Else
            Atn2 = Math.Atn(y / x) + PI
        End If
    Case Is = 0
        If y = 0 Then
            Err.Raise 5, MODULE_NAME & ".Atn2", "Can't compute Atn2 on (0,0)!"
        Else
            Atn2 = PI / 2 * Sgn(y)
        End If
    End Select
End Function

' Returns the orientation index (side) of a set of coordinates relative to a line
' Returns -1 if coordinates are on the left, 0 if the coordinates are on the line
' and +1 if coordinates are on the right
' Left and right are considered relative to the start and end coordinates of the line
Public Function GetOrientationIndex(ByVal x1 As Double, ByVal y1 As Double, _
                                    ByVal x2 As Double, ByVal y2 As Double, _
                                    ByVal x As Double, ByVal y As Double) As Integer
    GetOrientationIndex = Sgn((y2 - y1) * (x - x1) - (x2 - x1) * (y - y1))
End Function

' Returns the circumference of a circle defined by its radius
' Parameters:
'   - r: the radius of the circle
' Raises error for negative r (radius)
Public Function GetCircleCircumference(ByVal r As Double) As Double
    If r >= 0 Then
        GetCircleCircumference = TWO_PI * r
    Else
        Err.Raise 5, MODULE_NAME & ".GetCircleCircumference", _
                  "Can't compute circumference of a circle with negative radius!"
    End If
End Function

' Normalize an angle in a 2*PI wide interval around a center value.
' Parameters:
'   - a: angle to  be normalized
'   - c: center of the desired 2*PI interval for the result
Public Function NormalizeAngle(ByVal a As Double, ByVal c As Double) As Double
    NormalizeAngle = a - TWO_PI * MathLib.Floor((a + PI - c) / TWO_PI)
End Function



