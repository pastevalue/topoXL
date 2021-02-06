Attribute VB_Name = "LibGeom"
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

'@Folder("TopoXL.libs")

Option Explicit
Option Private Module

Public Const PI As Double = 3.14159265358979
' Rounded (more precise value) of 6.28318530717959 not used so it fits with math in code where 2*PI equals 6.28318530717958
Public Const TWO_PI As Double = 6.28318530717958
Private Const MODULE_NAME As String = "LibGeom"

' Returns the 2D distance between two sets of 2D grid coordinates
' Parameters:
'   - x1, y1: 2D grid coordinates of first point
'   - x2, y2: 2D grid coordinates of first point
Public Function dist2D(ByVal x1 As Double, ByVal y1 As Double, _
                       ByVal x2 As Double, ByVal y2 As Double) As Double
    dist2D = Math.Sqr((x1 - x2) * (x1 - x2) + (y1 - y2) * (y1 - y2))
End Function

' Returns the 3D distance between two sets of 3D grid coordinates
' Parameters:
'   - x1, y1, z1: 3D grid coordinates of first point
'   - x2, y2, z2: 3D grid coordinates of first point
Public Function dist3D(ByVal x1 As Double, ByVal y1 As Double, ByVal z1 As Double, _
                       ByVal x2 As Double, ByVal y2 As Double, ByVal z2 As Double) As Double
    dist3D = Math.Sqr((x1 - x2) * (x1 - x2) + (y1 - y2) * (y1 - y2) + (z1 - z2) * (z1 - z2))
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
Public Function orientationIndex(ByVal x1 As Double, ByVal y1 As Double, _
                                 ByVal x2 As Double, ByVal y2 As Double, _
                                 ByVal x As Double, ByVal y As Double, _
                                 Optional epsilon As Double = 0.000000000000001) As Integer
    Dim d As Double: d = (y2 - y1) * (x - x1) - (x2 - x1) * (y - y1)
    If Math.Abs(d) < epsilon Then d = 0
    orientationIndex = Math.Sgn(d)
End Function

' Returns the circumference of a circle defined by its radius
' Parameters:
'   - r: the radius of the circle
' Raises error for negative r (radius)
Public Function circleCircumference(ByVal r As Double) As Double
    If r >= 0 Then
        circleCircumference = TWO_PI * r
    Else
        Err.Raise 5, MODULE_NAME & ".GetCircleCircumference", _
                  "Can't compute circumference of a circle with negative radius!"
    End If
End Function

' Normalize an angle in a 2*PI wide interval around a center value.
' Parameters:
'   - a: angle to  be normalized
'   - c: center of the desired 2*PI interval for the result
Public Function normalizeAngle(ByVal a As Double, ByVal c As Double) As Double
    normalizeAngle = a - TWO_PI * Int((a + PI - c) / TWO_PI)
End Function

' Returns a Point representing the foot of the perpendicular from a point defined
' by x,y to a line segment defined by 2 points x1,y1 and x2,y2
' Returns Nothing if the 2 points defining the line segment are identical
Public Function footOfPerpendicular(ByVal x1 As Double, ByVal y1 As Double, _
                                    ByVal x2 As Double, ByVal y2 As Double, _
                                    ByVal x As Double, ByVal y As Double, _
                                    Optional epsilon As Double = 0.000000000000001) As Point
    Dim isLineVertical As Boolean: isLineVertical = LibMath.areDoublesEqual(x1, x2, epsilon)
    Dim isLineHorizontal As Boolean: isLineHorizontal = LibMath.areDoublesEqual(y1, y2, epsilon)
    
    If isLineVertical And isLineHorizontal Then Exit Function ' Points 1 and 2 are identical (not a line)
    
    Set footOfPerpendicular = New Point
    If isLineVertical Then
        footOfPerpendicular.x = x1
        footOfPerpendicular.y = y
    ElseIf isLineHorizontal Then
        footOfPerpendicular.x = x
        footOfPerpendicular.y = y1
    Else
        Dim lineSlope As Double: lineSlope = (y2 - y1) / (x2 - x1)
        footOfPerpendicular.x = x2 + (lineSlope * (y - y2) + x - x2) / (lineSlope * lineSlope + 1)
        footOfPerpendicular.y = lineSlope * (footOfPerpendicular.x - x2) + y2
    End If
End Function

' Returns a Point representing the intersection of 2 lines each defined by
' a P(x,y) point and a theta (angle in radians between the positive x-axis
' and a ray that passes through that P point - see Atn2)
' Returns Nothing if the lines are parallel
Public Function intOf2LinesByTheta(ByVal x1 As Double, ByVal y1 As Double, _
                                   ByVal theta1 As Double, _
                                   ByVal x2 As Double, ByVal y2 As Double, _
                                   ByVal theta2 As Double, _
                                   Optional epsilon As Double = 0.000000000000001) As Point
    ' Check if lines are parallel
    If LibMath.areDoublesEqual(theta1, theta2, epsilon) Then Exit Function
    Set intOf2LinesByTheta = New Point
    
    If LibMath.areDoublesEqual(theta1, 0, epsilon) _
    Or LibMath.areDoublesEqual(theta1, LibGeom.PI, epsilon) Then
        intOf2LinesByTheta.y = y1
        intOf2LinesByTheta.x = (intOf2LinesByTheta.y - y2) / Math.Tan(theta2) + x2
    ElseIf LibMath.areDoublesEqual(theta2, 0, epsilon) _
    Or LibMath.areDoublesEqual(theta2, LibGeom.PI, epsilon) Then
        intOf2LinesByTheta.y = y2
        intOf2LinesByTheta.x = (intOf2LinesByTheta.y - y1) / Math.Tan(theta1) + x1
    Else
        ' Compute Cotangents
        Dim ctg1 As Variant: ctg1 = 1 / Math.Tan(theta1)
        Dim ctg2 As Variant: ctg2 = 1 / Math.Tan(theta2)

        intOf2LinesByTheta.x = (x1 * ctg2 - x2 * ctg1 - ctg1 * ctg2 * (y1 - y2)) / (ctg2 - ctg1)
        intOf2LinesByTheta.y = (intOf2LinesByTheta.x - x1) / ctg1 + y1
    End If
End Function

' Returns TRUE if input coordinates are inside or on the edges
' of the input bounding box, FALSE otherwise
' If the input coordinates are o
' Parameters:
'   - x, y: coordinates to be checked if they are inside a bounding box
'   - x1, y1, x2, y2: coordinates to define the bounding box (oposite corners)
Public Function cooInBB(ByVal x As Double, ByVal y As Double, _
                        ByVal x1 As Double, ByVal y1 As Double, _
                        ByVal x2 As Double, ByVal y2 As Double) As Variant
    Dim xInside As Boolean
    Dim yInside As Boolean
    
    xInside = (x >= x1 And x <= x2) Or (x >= x2 And x <= x1)
    yInside = (y >= y1 And y <= y2) Or (y <= y1 And y >= y2)
    
    cooInBB = xInside And yInside
End Function
