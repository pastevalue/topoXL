VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LineSegment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
'' Represents a line segment defined by two Point(s).
'' P1 and P2 define the start and the end coordinates, respectively.
''------------------------------------------------------------------------
'' Called by:
''    Modules: None
''    Classes: None
'' Calls:
''    Modules: GeomLib
''    Classes: Point
''========================================================================

'@Folder("TopoXL.geom")

Option Explicit

Private m_P1 As Point
Private m_P2 As Point

' Class constructor
Private Sub Class_Initialize()
    Set m_P1 = New Point                         ' First Point defining the line segment
    Set m_P2 = New Point                         ' Second Point defining the line segment
End Sub

' Class initialization
Public Sub Init(ByVal x1 As Double, ByVal y1 As Double, _
                ByVal x2 As Double, ByVal y2 As Double)
    m_P1.Init x1, y1
    m_P2.Init x2, y2

End Sub

' Returns dX of this LineSegment
Public Function GetDX() As Double
    GetDX = Me.P2.x - Me.P1.x
End Function

' Returns dY of this LineSegment
Public Function GetDY() As Double
    GetDY = Me.P2.y - Me.P1.y
End Function

' Returns the length of this LineSegment
Public Function GetLength() As Double
    GetLength = GeomLib.Dist2D(m_P1.x, m_P1.y, m_P2.x, m_P2.y)
End Function

' Return the angle that the vector defined by this segment
' makes with the X-axis
Public Function GetTheta() As Double
    GetTheta = GeomLib.Atn2(Me.P2.x - Me.P1.x, Me.P2.y - Me.P1.y)
End Function

' Tests whether this LineSegment is vertical
Public Function IsVertical() As Boolean
    IsVertical = (m_P1.x = m_P2.x)
End Function

' Tests whether this LineSegment is horizontal
Public Function IsHorizontal() As Boolean
    IsHorizontal = (m_P1.y = m_P2.y)
End Function

' Returns the Projection Factor for the projection of the point p(x,y)
' onto this LineSegment.  The Projection Factor is the constant r
' by which the vector for this segment must be multiplied to
' equal the vector for the projection of p on the line
' defined by this segment. r has the following meaning:
'   r = 0 p = P1
'   r = 1 P = P2
'   r < 0 P is on the backward extension of this LineSegment
'   r > 1 P is on the forward extension of this LineSegment
'   0 < r < 1 P is interior to this LineSegment
Public Function GetProjectionFactor(ByVal x As Double, ByVal y As Double) As Double
    Dim p As New Point
    
    p.Init x, y
    If Me.P1.IsEqual(p) Then
        GetProjectionFactor = 0#
        Exit Function
    End If
    
    If Me.P2.IsEqual(p) Then
        GetProjectionFactor = 1#
        Exit Function
    End If
    
    Dim dX As Double
    Dim dY As Double
        
    dX = Me.GetDX
    dY = Me.GetDY
    GetProjectionFactor = ((p.x - Me.P1.x) * dX + (p.y - Me.P1.y) * dY) _
        / (dX * dX + dY * dY)
End Function

' Returns a Point which is at a specified offset (o) from this LineSegment
' and has its projection on this LineSegment at a specified distance (m) from
' one of its ends
' Parameters:
'   - m (measure): the distance measured along this LineSegment: positive
'                  is measured from P1 and negative from P2
'   - o (offset): the perpendicular distance measured against this LineSegment
' Result:
'   - A Point object with the computed coordinates
'   - Nothing if the distance in the m parameter is greater than this LineSegment
'   - length
Public Function GetPointByMeasOffset(ByVal m As Double, ByVal o As Double) As Point
    Dim length As Double
    length = Me.GetLength
    
    If Abs(m) > length Then
        Set GetPointByMeasOffset = Nothing
        Exit Function
    End If
    
    Dim fraction As Double
    Dim mX As Double
    Dim mY As Double
    Dim uX As Double
    Dim uY As Double
    
    If m < 0 Then
        fraction = (length + m) / length
    Else
        fraction = m / length
    End If
    
    ' the measure point
    mX = Me.P1.x + fraction * (Me.P2.x - Me.P1.x)
    mY = Me.P1.y + fraction * (Me.P2.y - Me.P1.y)
    
    ' u is the vector that is the length of the offset, in the direction of the segment
    uX = o * Me.GetDX / length
    uY = o * Me.GetDY / length
    
    Set GetPointByMeasOffset = New Point
    
    ' the result is the measure point plus "u" vector rotated 90 degrees CW
    GetPointByMeasOffset.Init mX + uY, mY - uX
    
End Function

' Returns a MeasOffset which is linear referencing a set of coordinates
' against this LineSegment
' Result:
'   - MeasOffset object with the computed measure and offset distances
'   - Nothing if the given coordinate doesn't lie within this LineSegment on this LineSegment
Public Function GetMeasOffsetOfPoint(ByVal x As Double, ByVal y As Double) As MeasOffset
    Dim r As Double
    
    r = Me.GetProjectionFactor(x, y)
    Select Case r
        'Projection point not on this LineSegment
    Case Is < 0#, Is > 1#
        Set GetMeasOffsetOfPoint = Nothing
    Case Else
        ' Projected point on this LineSegment
        Dim pX As Double
        Dim pY As Double
        pX = Me.P1.x + r * Me.GetDX
        pY = Me.P1.y + r * Me.GetDY
        
        'offset sign
        Dim os As Integer
        os = GeomLib.GetOrientationIndex(Me.P1.x, Me.P1.y, Me.P2.x, Me.P2.y, x, y)
        
        Set GetMeasOffsetOfPoint = New MeasOffset
        GetMeasOffsetOfPoint.Init r * Me.GetLength, os * GeomLib.Dist2D(pX, pY, x, y)
    End Select
End Function

' Test if this LineSegment is equivalent to another
' Returns TRUE if all members are equal
Public Function IsEqual(ls As LineSegment) As Boolean
    IsEqual = Me.P1.x = ls.P1.x And Me.P1.y = ls.P1.y And _
              Me.P2.x = ls.P2.x And Me.P2.y = ls.P2.y
End Function

' Returns a string representation of the LineSegment class
Public Function ToString() As String
    ToString = "(" & P1.ToString & "," & P2.ToString & ")"
End Function

'P1 Point property
Public Property Get P1() As Point
    Set P1 = m_P1
End Property

Public Property Let P1(P1 As Point)
    Set m_P1 = P1
End Property

'P2 Point property
Public Property Get P2() As Point
    Set P2 = m_P2
End Property

Public Property Let P2(P2 As Point)
    Set m_P2 = P2
End Property


