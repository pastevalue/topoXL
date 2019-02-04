VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ArcCircle"
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
''    Classes: RLarcCircle
''Calls:
''    Modules: RLenums, UDF_COGO, UDF_Math
''    Classes: Coo2D
''=======================================================
Option Explicit

'primary members
Private m_startCoo As Coo2D 'start coordinates of the arc circle
Private m_centerCoo As Coo2D 'center coordinates of the arc circle
Private m_length As Double 'length of the arc circle
Private m_curveDir As CURVE_DIRECTION 'arc circle curve direction

'secondary members
Private m_radius As Double ' radius of the arc circle
Private m_cenToStartAz As Double 'azimuth from center to start
Private m_cenToEndAz As Double 'azimuth from center to end
Private m_endCoo As Coo2D 'end coordinates of the arc circle

'Class constructor with default values
Private Sub Class_Initialize()
    Set m_startCoo = New Coo2D
    Set m_centerCoo = New Coo2D
    m_cenToStartAz = 0
    m_length = 0
    m_curveDir = CURVE_DIRECTION.NONE
    
    Set m_endCoo = New Coo2D
End Sub

'Initialize class from CooIn, CooCenter, Length and CurveDirection
Public Sub initFromInCenLenDir(inCoo As Coo2D, cenCoo As Coo2D, length As Double, curveDirection As CURVE_DIRECTION)
    Set m_startCoo = inCoo
    Set m_centerCoo = cenCoo
    m_length = length
    m_curveDir = curveDirection
    Call setRadius
    Call setCenterAzimuths
    Call setEndCoo
End Sub

'Initialize class from CooIn, CooOut, Radius and CurveDirection
Public Sub initFromInOutRadDir(inCoo As Coo2D, outCoo As Coo2D, radius As Double, curveDirection As CURVE_DIRECTION)
    Dim tempInOutAz As Double 'CooIn-CooOut azimuth
    Dim tempInOutDist As Double 'CooIn-CooOut chord distance
    Dim tempChordMiddleToCenDist 'chord middle to center distance
    Dim tempCen As Coo2D 'arc circle center coo
    Dim tempLength As Double 'arc circle length
        
    tempInOutAz = cogoAzimuth(inCoo.x, inCoo.y, outCoo.x, outCoo.y)
    tempInOutDist = cogoDistance2D(inCoo.x, inCoo.y, outCoo.x, outCoo.y)
    tempChordMiddleToCenDist = Sqr(radius ^ 2 - tempInOutDist ^ 2 / 4)
    Set tempCen = New Coo2D
    tempCen.x = inCoo.x + Sin(tempInOutAz) * tempInOutDist / 2 + Sin(tempInOutAz + curveDirection * PI / 2) * tempChordMiddleToCenDist
    tempCen.y = inCoo.y + Cos(tempInOutAz) * tempInOutDist / 2 + Cos(tempInOutAz + curveDirection * PI / 2) * tempChordMiddleToCenDist
    tempLength = 2 * radius * mathASin(tempInOutDist / (2 * radius))
    Call initFromInCenLenDir(inCoo, tempCen, tempLength, curveDirection)
End Sub

'Gets the length of the arc circle
Public Function getLength() As Double
    getLength = m_length
End Function

'Sets the radius of the circle
Private Sub setRadius()
    m_radius = cogoDistance2D(m_centerCoo.x, m_centerCoo.y, m_startCoo.x, m_startCoo.y)
End Sub

'Sets the azimuths for the center to start line and center to end line
Private Sub setCenterAzimuths()
    m_cenToStartAz = cogoAzimuth(m_centerCoo.x, m_centerCoo.y, m_startCoo.x, m_startCoo.y)
    m_cenToEndAz = m_cenToStartAz + m_length / m_radius * m_curveDir
End Sub
Private Sub setEndCoo()
    Set m_endCoo = getCoo(m_length, 0)
    m_endCoo.x = m_centerCoo.x + m_radius * Sin(m_cenToEndAz)
    m_endCoo.y = m_centerCoo.y + m_radius * Cos(m_cenToEndAz)
End Sub

'Gets coordinates at specified distance and offset
Public Function getCoo(dist As Double, offset As Double) As Coo2D
    
    If (dist > m_length) Or (dist < 0) Then
        GoTo ErrHandler
    End If
    
    Dim centerToPointAzimuth As Double
    Dim centerToPointDist As Double
    
    centerToPointAzimuth = m_cenToStartAz + dist / m_radius * m_curveDir
    centerToPointDist = m_radius - offset * m_curveDir
    
    Dim result As New Coo2D
    result.x = m_centerCoo.x + centerToPointDist * Sin(centerToPointAzimuth)
    result.y = m_centerCoo.y + centerToPointDist * Cos(centerToPointAzimuth)
    Set getCoo = result
Exit Function
ErrHandler:
  Debug.Print "Total length of the arc circle is " & m_length & ". Can't calculate coordinates at a distance of " & dist & "!"
End Function

'Gets abscissa at specified distance and offset
Public Function getCooX(dist As Double, offset As Double) As Double
    
    If (dist > m_length) Or (dist < 0) Then
        GoTo ErrHandler
    End If
        
    getCooX = m_centerCoo.x + (m_radius - offset * m_curveDir) * Sin(m_cenToStartAz + dist / m_radius * m_curveDir)

Exit Function
ErrHandler:
  Debug.Print "Total length of the arc circle is " & m_length & ". Can't calculate coordinates at a distance of " & dist & "!"
End Function

'Gets ordinate at specified distance and offset
Public Function getCooY(dist As Double, offset As Double) As Double
    
    If (dist > m_length) Or (dist < 0) Then
        GoTo ErrHandler
    End If
        
    getCooY = m_centerCoo.y + (m_radius - offset * m_curveDir) * Cos(m_cenToStartAz + dist / m_radius * m_curveDir)

Exit Function
ErrHandler:
  Debug.Print "Total length of the arc circle is " & m_length & ". Can't calculate coordinates at a distance of " & dist & "!"
End Function

'Gets abscissa at specified ordinate
Public Function getX(y As Double) As Double
    Dim delta As Double
    Dim tempX As Double
    
    If y = m_startCoo.y Then
        getX = m_startCoo.x
        Exit Function
    ElseIf y = m_endCoo.y Then
        getX = m_endCoo.x
        Exit Function
    End If
    
    delta = m_radius ^ 2 - (y - m_centerCoo.y) ^ 2
    Select Case delta
        Case Is > 0
            tempX = m_centerCoo.x - Sqr(delta)
            If cogoGetSide(m_startCoo.x, m_startCoo.y, m_endCoo.x, m_endCoo.y, tempX, y) = -m_curveDir Then
                getX = tempX
            Else
                tempX = m_centerCoo.x + Sqr(delta)
                If cogoGetSide(m_startCoo.x, m_startCoo.y, m_endCoo.x, m_endCoo.y, tempX, y) = -m_curveDir Then
                    getX = tempX
                Else
                    getX = 0
                End If
            End If
        Case 0
            getX = m_centerCoo.x
        Case Is < 0
            getX = 0
    End Select
End Function

'Gets ordinate at specified abscissa
Public Function getY(x As Double) As Double
    Dim delta As Double
    Dim tempY As Double
    
    If x = m_startCoo.x Then
        getY = m_startCoo.y
        Exit Function
    ElseIf x = m_endCoo.x Then
        getY = m_endCoo.y
        Exit Function
    End If
    
    delta = m_radius ^ 2 - (x - m_centerCoo.x) ^ 2
    Select Case delta
        Case Is > 0
            tempY = m_centerCoo.y - Sqr(delta)
            If cogoGetSide(m_startCoo.x, m_startCoo.y, m_endCoo.x, m_endCoo.y, x, tempY) = -m_curveDir Then
                getY = tempY
            Else
                tempY = m_centerCoo.y + Sqr(delta)
                If cogoGetSide(m_startCoo.x, m_startCoo.y, m_endCoo.x, m_endCoo.y, x, tempY) = -m_curveDir Then
                    getY = tempY
                Else
                    getY = 0
                End If
            End If
        Case 0
            getY = m_centerCoo.y
        Case Is < 0
            getY = 0
    End Select
End Function

'For a set of coordinates returns the radial distance to the arc circle and the arc length from radial point to the start of the ar circle
Public Function getRadialDistFromPoint(ByVal px As Double, ByVal py As Double, ByRef distFromStart As Double, ByRef distFromArcCircle As Double) As Boolean
    
    Dim sideOfPntCenterToStart As Integer
    Dim sideOfPntCenterToEnd As Integer
    Dim pointBetweenCSandCE As Integer
    Dim intOnArc As Double 'used to establish if the intersection point is on the arc circle
    Dim int1 As New Coo2D 'first point of intersection between circle(extension of arc circle) and center to input point line
    Dim int2 As New Coo2D 'second point of intersection between circle(extension of arc circle) and center to input point line
    Dim tempDistCenterToPoint As Double 'distance from the center of the arc circle to the input point
    sideOfPntCenterToStart = cogoGetSide(m_centerCoo.x, m_centerCoo.y, m_startCoo.x, m_startCoo.y, px, py)
    sideOfPntCenterToEnd = cogoGetSide(m_centerCoo.x, m_centerCoo.y, m_endCoo.x, m_endCoo.y, px, py)
    
    pointBetweenCSandCE = sideOfPntCenterToStart * sideOfPntCenterToEnd
    If (m_length < PI * m_radius) Then 'arc circle length is smaller than half a circle length
        'Select case for radial point from input point
        Select Case pointBetweenCSandCE
        
        'point is not on the arc circle
        Case 1
            distFromStart = 0
            distFromArcCircle = 0
            getRadialDistFromPoint = False
    
        'point is on the arc circle
        Case -1
            tempDistCenterToPoint = cogoDistance2D(m_centerCoo.x, m_centerCoo.y, px, py)
    
            int1.x = m_centerCoo.x + (m_centerCoo.x - px) * m_radius / tempDistCenterToPoint
            int1.y = m_centerCoo.y + (m_centerCoo.y - py) * m_radius / tempDistCenterToPoint
            int2.x = 2 * m_centerCoo.x - int1.x
            int2.y = 2 * m_centerCoo.y - int1.y
           
            'IF first intersection point is on arc circle ELSE go to second point
            If cogoGetSide(m_startCoo.x, m_startCoo.y, m_endCoo.x, m_endCoo.y, int1.x, int1.y) = -m_curveDir Then
                distFromArcCircle = cogoDistance2D(int1.x, int1.y, px, py) _
                                    * cogoGetSide(m_startCoo.x, m_startCoo.y, int1.x, int1.y, px, py)
                distFromStart = getArcLengthOfChord(cogoDistance2D(m_startCoo.x, m_startCoo.y, int1.x, int1.y))
                getRadialDistFromPoint = True
            Else
                distFromArcCircle = cogoDistance2D(int2.x, int2.y, px, py) _
                                    * cogoGetSide(m_startCoo.x, m_startCoo.y, int2.x, int2.y, px, py)
                distFromStart = getArcLengthOfChord(cogoDistance2D(m_startCoo.x, m_startCoo.y, int2.x, int2.y))
                getRadialDistFromPoint = True
            End If
    
        'point is on the center to start line or center to end line
        Case 0
            'point is on the center to start line
            If sideOfPntCenterToStart = 0 Then
                distFromStart = 0
                distFromArcCircle = cogoDistance2D(m_startCoo.x, m_startCoo.y, px, py) _
                                    * cogoGetSide(m_startCoo.x, m_startCoo.y, m_endCoo.x, m_endCoo.y, px, py)
                getRadialDistFromPoint = True
            'point is on the center to end line
            Else
                distFromStart = m_length
                distFromArcCircle = cogoDistance2D(m_endCoo.x, m_endCoo.y, px, py) _
                                    * cogoGetSide(m_startCoo.x, m_startCoo.y, m_endCoo.x, m_endCoo.y, px, py)
                getRadialDistFromPoint = True
            End If
        End Select
    Else 'arc circle length is larger (or equal) than half a circle length
        If pointBetweenCSandCE = 0 Then 'point is on the center to start line
            Dim dist1, dist2 As Double
            Dim oppositePoint As Coo2D
            Set oppositePoint = New Coo2D
            
            If sideOfPntCenterToStart = 0 Then
                dist1 = cogoDistance2D(m_startCoo.x, m_startCoo.y, px, py)
                oppositePoint.x = 2 * m_centerCoo.x - m_startCoo.x
                oppositePoint.y = 2 * m_centerCoo.y - m_startCoo.y
                dist2 = cogoDistance2D(oppositePoint.x, oppositePoint.y, px, py)
                If dist1 <= dist2 Then
                    distFromStart = 0
                    distFromArcCircle = -dist1 * cogoGetSide(m_startCoo.x, m_startCoo.y, m_endCoo.x, m_endCoo.y, px, py)
                Else
                    distFromStart = PI * m_radius
                    distFromArcCircle = -dist2 * cogoGetSide(m_startCoo.x, m_startCoo.y, m_endCoo.x, m_endCoo.y, px, py)
                End If
                getRadialDistFromPoint = True
            'point is on the center to end line
            Else
                dist1 = cogoDistance2D(m_endCoo.x, m_endCoo.y, px, py)
                oppositePoint.x = 2 * m_centerCoo.x - m_endCoo.x
                oppositePoint.y = 2 * m_centerCoo.y - m_endCoo.y
                dist2 = cogoDistance2D(oppositePoint.x, oppositePoint.y, px, py)
                If dist1 <= dist2 Then
                    distFromStart = m_length
                    distFromArcCircle = -dist1 * cogoGetSide(m_startCoo.x, m_startCoo.y, m_endCoo.x, m_endCoo.y, px, py)
                Else
                    distFromStart = m_length - PI * m_radius
                    distFromArcCircle = -dist2 * cogoGetSide(m_startCoo.x, m_startCoo.y, m_endCoo.x, m_endCoo.y, px, py)
                End If
                getRadialDistFromPoint = True
            End If
        Else 'point is on the arc circle but not on the center-start line or on center-end line
            tempDistCenterToPoint = cogoDistance2D(m_centerCoo.x, m_centerCoo.y, px, py)
            int1.x = m_centerCoo.x + (m_centerCoo.x - px) * m_radius / tempDistCenterToPoint
            int1.y = m_centerCoo.y + (m_centerCoo.y - py) * m_radius / tempDistCenterToPoint
            int2.x = 2 * m_centerCoo.x - int1.x
            int2.y = 2 * m_centerCoo.y - int1.y
            
            'IF first intersection point is on arc circle ELSE go to second point
            intOnArc = cogoGetSide(m_startCoo.x, m_startCoo.y, m_endCoo.x, m_endCoo.y, int1.x, int1.y)
            If intOnArc = -m_curveDir Or intOnArc = 0 Then
                distFromStart = getArcLengthOfChord(cogoDistance2D(m_startCoo.x, m_startCoo.y, int1.x, int1.y))
                distFromArcCircle = cogoDistance2D(int1.x, int1.y, px, py)
                If intOnArc <> 0 Then
                    distFromArcCircle = distFromArcCircle * cogoGetSide(m_startCoo.x, m_startCoo.y, int1.x, int1.y, px, py)
                Else
                    distFromArcCircle = -distFromArcCircle * cogoGetSide(m_startCoo.x, m_startCoo.y, m_endCoo.x, m_endCoo.y, px, py)
                End If
                'if the length from start to intersection point is larger than PI*R
                If cogoGetSide(m_startCoo.x, m_startCoo.y, m_centerCoo.x, m_centerCoo.y, int1.x, int1.y) = m_curveDir Then
                    distFromStart = 2 * PI * m_radius - distFromStart
                    distFromArcCircle = -distFromArcCircle
                End If
                getRadialDistFromPoint = True
                
                Dim tempDistFromArcCIrcle As Double
                'If second intersection point is also on arc circle
                intOnArc = cogoGetSide(m_startCoo.x, m_startCoo.y, m_endCoo.x, m_endCoo.y, int2.x, int2.y)
                If intOnArc = -m_curveDir Or intOnArc = 0 Then
                        tempDistFromArcCIrcle = cogoDistance2D(int2.x, int2.y, px, py)
                    
                        'If input point is closer to the second intersection point than to the first one
                        If tempDistFromArcCIrcle < Abs(distFromArcCircle) Then
                            If intOnArc <> 0 Then
                                distFromArcCircle = tempDistFromArcCIrcle * cogoGetSide(m_startCoo.x, m_startCoo.y, int2.x, int2.y, px, py)
                            Else
                                distFromArcCircle = -tempDistFromArcCIrcle * cogoGetSide(m_startCoo.x, m_startCoo.y, m_endCoo.x, m_endCoo.y, px, py)
                            End If
                            distFromStart = getArcLengthOfChord(cogoDistance2D(m_startCoo.x, m_startCoo.y, int2.x, int2.y))
                            'if the length from start to intersection point is larger than PI*R
                            If cogoGetSide(m_startCoo.x, m_startCoo.y, m_centerCoo.x, m_centerCoo.y, int2.x, int2.y) = m_curveDir Then
                                distFromStart = 2 * PI * m_radius - distFromStart
                                distFromArcCircle = -distFromArcCircle
                            End If
                        End If
                End If
            Else
                intOnArc = cogoGetSide(m_startCoo.x, m_startCoo.y, m_endCoo.x, m_endCoo.y, int2.x, int2.y)
                distFromArcCircle = cogoDistance2D(int2.x, int2.y, px, py)
                If intOnArc <> 0 Then
                    distFromArcCircle = distFromArcCircle * cogoGetSide(m_startCoo.x, m_startCoo.y, int2.x, int2.y, px, py)
                Else
                    distFromArcCircle = -distFromArcCircle * cogoGetSide(m_startCoo.x, m_startCoo.y, m_endCoo.x, m_endCoo.y, px, py)
                End If
                distFromStart = getArcLengthOfChord(cogoDistance2D(m_startCoo.x, m_startCoo.y, int2.x, int2.y))
                'if the length from start to intersection point is larger than PI*R
                If cogoGetSide(m_startCoo.x, m_startCoo.y, m_centerCoo.x, m_centerCoo.y, int2.x, int2.y) = m_curveDir Then
                    distFromStart = 2 * PI * m_radius - distFromStart
                    distFromArcCircle = -distFromArcCircle
                End If
                getRadialDistFromPoint = True
            End If
        End If
    End If
End Function

Private Function getArcLengthOfChord(chord As Double) As Double
    getArcLengthOfChord = CDbl(mathGetArcCircleLengthFromChord(chord, m_radius))
End Function

' startCoo property
Public Property Get startCoo() As Coo2D
    Set startCoo = m_startCoo
End Property

' endCoo property
Public Property Get endCoo() As Coo2D
    Set endCoo = m_endCoo
End Property





