VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Line2D"
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

''=======================================================
''Called by:
''    Modules: None
''    Classes: RLline
''Calls:
''    Modules: UDF_COGO
''    Classes: Coo2D
''=======================================================
Option Explicit

Private m_startCoo As Coo2D
Private m_endCoo As Coo2D
Private m_azimuth As Double
Private m_length As Double

'Class constructor
Private Sub Class_Initialize()
    Set m_startCoo = New Coo2D
    Set m_endCoo = New Coo2D
    m_azimuth = 0
    m_length = 0
End Sub

'Class initialization
Public Sub init(coo1 As Coo2D, coo2 As Coo2D)
    Set m_startCoo = coo1
    Set m_endCoo = coo2
    Call setAzimuth
    Call setlength
End Sub

Private Sub setAzimuth()
    m_azimuth = cogoAzimuth(startCoo.x, startCoo.y, endCoo.x, endCoo.y)
End Sub

Private Sub setlength()
    m_length = cogoDistance2D(m_startCoo.x, m_startCoo.y, m_endCoo.x, m_endCoo.y)
End Sub

'Length method
Public Function getLength() As Double
    getLength = m_length
End Function

'Gets coordinates at specified distance and offset
Public Function getCoo(dist As Double, offset As Double) As Coo2D
    
    If (dist > m_length) Or (dist < 0) Then
        getCoo = Null
        Debug.Print "Total length of the line is " & m_length & ". Can't calculate coordinates at a distance of " & dist & "!"
        Exit Function
    End If
    
    Dim perpendicularAzimuth As Double
    perpendicularAzimuth = m_azimuth + PI / 2 * (2 - Sgn(offset))
    
    Dim result As New Coo2D
    result.x = m_startCoo.x + Sin(m_azimuth) * dist + Sin(perpendicularAzimuth) * Abs(offset)
    result.y = m_startCoo.y + Cos(m_azimuth) * dist + Cos(perpendicularAzimuth) * Abs(offset)
    Set getCoo = result
End Function

'Gets abscissa at specified distance and offset
Public Function getCooX(dist As Double, offset As Double) As Double
    
    If (dist > m_length) Or (dist < 0) Then
        getCooX = Null
        Debug.Print "Total length of the line is " & m_length & ". Can't calculate coordinates at a distance of " & dist & "!"
    End If
    getCooX = m_startCoo.x + Sin(m_azimuth) * dist + Sin(m_azimuth + PI / 2 * (2 - Sgn(offset))) * Abs(offset)
End Function

'Gets ordinate at specified distance and offset
Public Function getCooY(dist As Double, offset As Double) As Double
    
    If (dist > m_length) Or (dist < 0) Then
        getCooY = Null
        Debug.Print "Total length of the line is " & m_length & ". Can't calculate coordinates at a distance of " & dist & "!"
    End If
    getCooY = m_startCoo.y + Cos(m_azimuth) * dist + Cos(m_azimuth + PI / 2 * (2 - Sgn(offset))) * Abs(offset)
End Function

'Gets abscissa at specified ordinate
Public Function getX(y As Double) As Double
    If y = m_startCoo.y Then
        getX = m_startCoo.x
        Exit Function
    ElseIf y = m_endCoo.y Then
        getX = m_endCoo.x
        Exit Function
    End If
    If ((m_startCoo.y < y) And (m_endCoo.y > y)) Or ((m_startCoo.y > y) And (m_endCoo.y < y)) Then
        getX = m_startCoo.x + (m_endCoo.x - m_startCoo.x) * (y - m_startCoo.y) / (m_endCoo.y - m_startCoo.y)
    End If
End Function

'Gets ordinate at specified abscissa
Public Function getY(x As Double) As Double
    If x = m_startCoo.x Then
        getY = m_startCoo.y
        Exit Function
    ElseIf x = m_endCoo.x Then
        getY = m_endCoo.y
        Exit Function
    End If
    If ((m_startCoo.x < x) And (m_endCoo.x > x)) Or ((m_startCoo.x > x) And (m_endCoo.x < x)) Then
        getY = m_startCoo.y + (m_endCoo.y - m_startCoo.y) * (x - m_startCoo.x) / (m_endCoo.x - m_startCoo.x)
    End If
End Function

'Gets abscissa at specified ordinate
Public Function getCooXbasedOnCooY(CooY As Double) As Double
    
    If (m_startCoo.y = m_endCoo.y) And (m_startCoo.y = CooY) Then
        getCooXbasedOnCooY = m_startCoo.x
    ElseIf ((m_startCoo.y <= CooY) And (m_endCoo.y >= CooY)) Or ((m_startCoo.y >= CooY) And (m_endCoo.y <= CooY)) Then
        getCooXbasedOnCooY = m_startCoo.x + (m_endCoo.x - m_startCoo.x) * (CooY - m_startCoo.y) / (m_endCoo.y - m_startCoo.y)
    End If
End Function

'Gets ordinate at specified abscissa
Public Function getCooYbasedOnCooX(CooX As Double) As Double
    
    If (m_startCoo.x = m_endCoo.x) And (m_startCoo.x = CooX) Then
        getCooYbasedOnCooX = m_startCoo.y
    ElseIf ((m_startCoo.x <= CooX) And (m_endCoo.x >= CooX)) Or ((m_startCoo.x >= CooX) And (m_endCoo.x <= CooX)) Then
        getCooYbasedOnCooX = m_startCoo.y + (m_endCoo.y - m_startCoo.y) * (CooX - m_startCoo.x) / (m_endCoo.x - m_startCoo.x)
    End If
End Function

'For a set of coordinates returns the distance to the line and the distance from perpendicular coordinates to the start of the line
Public Function getPerpDistanceFromPoint(ByVal px As Double, ByVal py As Double, ByRef distFromStart As Double, ByRef distFromLine As Double) As Boolean
    Dim lineSlope As Double
    Dim perpPoint As New Coo2D
    
    If m_startCoo.x <> m_endCoo.x Then
        lineSlope = (m_endCoo.y - m_startCoo.y) / (m_endCoo.x - m_startCoo.x)
        perpPoint.x = m_endCoo.x + (py - m_endCoo.y + 1 / lineSlope * (px - m_endCoo.x)) / (lineSlope + 1 / lineSlope)
        perpPoint.y = lineSlope * (perpPoint.x - m_endCoo.x) + m_endCoo.y
    Else
        perpPoint.x = m_startCoo.x
        perpPoint.y = py
    End If
    
    If isInsideBoundBox(perpPoint) Then
        distFromLine = cogoDistance2D(px, py, perpPoint.x, perpPoint.y) * cogoGetSide(m_startCoo.x, m_startCoo.y, m_endCoo.x, m_endCoo.y, px, py)
        distFromStart = cogoDistance2D(m_startCoo.x, m_startCoo.y, perpPoint.x, perpPoint.y)
        getPerpDistanceFromPoint = True
    Else
        distFromStart = 0
        distFromLine = 0
        getPerpDistanceFromPoint = False
    End If
End Function

'Return TRUE if a point is inside the bounding box of the line
Private Function isInsideBoundBox(ByVal p As Coo2D) As Boolean
    Dim validX As Boolean
    Dim validY As Boolean
    
    validX = (p.x >= m_startCoo.x And p.x <= m_endCoo.x) Or (p.x >= m_endCoo.x And p.x <= m_startCoo.x)
    validY = (p.y >= m_startCoo.y And p.y <= m_endCoo.y) Or (p.y <= m_startCoo.y And p.y >= m_endCoo.y)
    
    isInsideBoundBox = validX And validY
End Function

'Azimuth property
Public Property Get getAzimuth() As Double
    getAzimuth = m_azimuth
End Property

'Start coordinate property
Public Property Get startCoo() As Coo2D
    Set startCoo = m_startCoo
End Property
Public Property Let startCoo(startCoo As Coo2D)
    Set m_startCoo = startCoo
    Call setAzimuth
    Call setlength
End Property

'End coordinate property
Public Property Get endCoo() As Coo2D
    Set endCoo = m_endCoo
End Property
Public Property Let endCoo(endCoo As Coo2D)
    Set m_endCoo = endCoo
    Call setAzimuth
    Call setlength
End Property



