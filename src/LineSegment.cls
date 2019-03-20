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
'' Represents a line segment defined by coordinates of two points.
'' Full list of members
'' m_sX ,m_sY, m_eX, m_eY, m_dX, m_dY, m_length, m_theta.
'' m_dX, m_dY, m_length, m_theta members are reduntant but they are
'' all computed at the initialization stage for the purpose
'' of increased speed in the calculation of "calcPointByMeasOffset" and
'' "calcMeasOffsetOfPoint".
'' Initilization methods
'' "init". Any new initialize method must compute all members of
'' the class. If wrong parameters are used with the initialize methods,
'' error number 5 is raised.
''========================================================================

'@Folder("TopoXL.CL.geom")

Option Explicit
Implements IGeom

Private m_sX As Double                           ' start Point X
Private m_sY As Double                           ' start Point Y
Private m_eX As Double                           ' end Point X
Private m_eY As Double                           ' end Point Y

' Secondary members
Private m_dX As Double                           ' delta X
Private m_dY As Double                           ' delta Y
Private m_length As Double                       ' length
Private m_theta As Double                        ' theta

' Class constructor
Private Sub Class_Initialize()
    m_sX = 0
    m_sY = 0
    m_eX = 0
    m_eY = 0
    m_dX = 0
    m_dY = 0
    m_length = 0
    m_theta = 0
End Sub

' Initialize by: start Point and end Point
' Raises error if:
'   - start and end coordinates are equal
Public Sub init(ByVal sX As Double, ByVal sY As Double, _
                ByVal eX As Double, ByVal eY As Double)
    m_sX = sX
    m_sY = sY
    m_eX = eX
    m_eY = eY
    
    ' Raise error if start and end coordinates are equal
    If sX = eX And sY = eY Then
        Err.Raise 5, TypeName(Me), "Can't initiate line segment. Start and end point must be different!"
    End If

    ' Compute other members (m_dX, m_dY, m_length, m_theta)
    m_dX = eX - sX
    m_dY = eY - sY
    m_length = LibGeom.dist2D(sX, sY, eX, eY)
    m_theta = LibGeom.Atn2(m_dX, m_dY)
End Sub

' Tests whether this LineSegment is vertical
Public Function isVertical() As Boolean
    isVertical = (sX = eX)
End Function

' Tests whether this LineSegment is horizontal
Public Function isHorizontal() As Boolean
    isHorizontal = (sY = eY)
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
Public Function calcProjectionFactor(ByVal x As Double, ByVal y As Double) As Double
    ' Identical to start point
    If m_sX = x And m_sY = y Then
        calcProjectionFactor = 0#
        Exit Function
    End If
    ' Identical to end point
    If m_eX = x And m_eY = y Then
        calcProjectionFactor = 1#
        Exit Function
    End If
    
    calcProjectionFactor = ((x - sX) * m_dX + (y - sY) * m_dY) / (m_dX * m_dX + m_dY * m_dY)
End Function

' Returns a Point which is at a specified offset (o) from this LineSegment
' and has its projection on this LineSegment at a specified distance (m) from
' one of its ends
' Parameters:
'   - m (measure): the distance from start point measured along this LineSegment
'   - o (offset): the perpendicular distance measured against this LineSegment
' Result:
'   - A Point object with the computed coordinates
'   - Nothing if the distance in the m parameter is greater than this LineSegment
'     length or less than 0
Public Function calcPointByMeasOffset(ByVal m As Double, ByVal o As Double) As Point
    
    If m < 0 Or m > m_length Then
        Set calcPointByMeasOffset = Nothing
        Exit Function
    End If
    
    Dim fraction As Double
    Dim mX As Double
    Dim mY As Double
    Dim uX As Double
    Dim uY As Double
    
    fraction = m / m_length
    
    ' the measure point
    mX = m_sX + fraction * m_dX
    mY = m_sY + fraction * m_dY
    
    ' u is the vector that is the length of the offset, in the direction of the segment
    uX = o * m_dX / m_length
    uY = o * m_dY / m_length
    
    Set calcPointByMeasOffset = New Point
    
    ' the result is the measure point plus "u" vector rotated 90 degrees CW
    calcPointByMeasOffset.init mX + uY, mY - uX
    
End Function

' Returns a MeasOffset which is linear referencing a set of coordinates
' against this LineSegment
' Result:
'   - MeasOffset object with the computed measure and offset distances
'   - Nothing if the projection of the given coordinates doesn't lie
'     within this LineSegment
Public Function calcMeasOffsetOfPoint(ByVal x As Double, ByVal y As Double) As MeasOffset
    Dim r As Double
    
    r = Me.calcProjectionFactor(x, y)
    Select Case r
        'Projection point not on this LineSegment
    Case Is < 0#, Is > 1#
        Set calcMeasOffsetOfPoint = Nothing
    Case Else
        ' Projected point on this LineSegment
        Dim px As Double
        Dim py As Double
        px = m_sX + r * m_dX
        py = m_sY + r * m_dY
        
        'offset sign
        Dim os As Integer
        os = LibGeom.orientationIndex(m_sX, m_sY, m_eX, m_eY, x, y)
        
        Set calcMeasOffsetOfPoint = New MeasOffset
        calcMeasOffsetOfPoint.init r * m_length, os * LibGeom.dist2D(px, py, x, y)
    End Select
End Function

' Test if this LineSegment is equal to another
' Returns TRUE if all members are equal
Public Function isEqual(ls As LineSegment) As Boolean
    isEqual = m_sX = ls.sX And m_sY = ls.sY And _
              m_eX = ls.eX And m_eY = ls.eY
End Function

' Returns a string representation of the LineSegment class
Public Function toString() As String
    toString = "((" & m_sX & "," & m_sY & "),(" & m_eX & "," & m_eY & "))"
End Function

' Start point property
Public Property Get sX() As Double
    sX = m_sX
End Property

Public Property Get sY() As Double
    sY = m_sY
End Property

' End point property
Public Property Get eX() As Double
    eX = m_eX
End Property

Public Property Get eY() As Double
    eY = m_eY
End Property

' Delta X property
Public Property Get dX() As Double
    dX = m_dX
End Property

' Delta Y property
Public Property Get dY() As Double
    dY = m_dY
End Property

' Length property
Public Property Get length() As Double
    length = m_length
End Property

' Theta property
Public Property Get theta() As Double
    theta = m_theta
End Property

' IGeom interface
Private Function IGeom_calcPointByMeasOffset(ByVal m As Double, ByVal o As Double) As Point
    Set IGeom_calcPointByMeasOffset = Me.calcPointByMeasOffset(m, o)
End Function

Private Function IGeom_calcMeasOffsetOfPoint(ByVal x As Double, ByVal y As Double) As MeasOffset
    Set IGeom_calcMeasOffsetOfPoint = Me.calcMeasOffsetOfPoint(x, y)
End Function

Private Property Get IGeom_length() As Double
    IGeom_length = Me.length
End Property

Public Property Get IGeom_geomTypeName() As String
    IGeom_geomTypeName = ConstCL.LS_NAME
End Property











