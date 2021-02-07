VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CircularArc"
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
'' Represents a circular arc.
'' Full list of members
'' m_sX, m_sY, m_cX, m_cY, m_eX, m_eY, m_length, m_curveDir, m_rad,
'' m_sTheta, m_eTheta, m_isCircle. Some of these members are redundant
'' but they are all computed at the initialization stage for the purpose
'' of increased speed in the calculation of "calcPointByMeasOffset" and
'' "calcMeasOffsetOfPoint".
'' Initilization methods
'' "initFromSCLD" and "initFromSERD". Any new initialize
'' method must compute all members of the class. If wrong parameters
'' are used with the initialize methods, error number 5 is raised.
''========================================================================

'@Folder("TopoXL.CL.geom")

Option Explicit
Implements IGeom

Private m_sX As Double                           ' start Point X
Private m_sY As Double                           ' start Point Y
Private m_cX As Double                           ' center Point X
Private m_cY As Double                           ' center Point Y
Private m_eX As Double                           ' end Point X
Private m_eY As Double                           ' end Point Y
Private m_length As Double                       ' length of the circular arc
Private m_curveDir As CURVE_DIR                  ' circular arc curve direction (CW or CCW)
Private m_rad As Double                          ' radius of the circular arc
Private m_sTheta As Double                       ' theta of the center to start line
Private m_eTheta As Double                       ' theta of the center to end line
Private m_isCircle As Boolean                    ' the circular arc is a full circle

' Class constructor with default values
Private Sub Class_Initialize()
    m_sX = 0
    m_sY = 0
    m_cX = 0
    m_cY = 0
    m_eX = 0
    m_eY = 0
    m_length = 0
    m_curveDir = CURVE_DIR.CD_NONE
    m_rad = 0
    m_sTheta = 0
    m_eTheta = 0
    m_isCircle = False
End Sub

' Initialize by: start Point, center Point, length and curve direction
' If length is greater than 2 * PI * radius then circular arc is
' considered to be a full circle and length is reduced to the
' 2 * PI * radius value
' Raises error if:
'   - start and center coordinates are equal
'   - length is less than or equal to 0
'   - curveDir is not CW (clocwise) or CCW (counter-clockwise)
Public Sub initFromSCLD(ByVal sX As Double, ByVal sY As Double, _
                        ByVal cX As Double, ByVal cY As Double, _
                        ByVal length As Double, ByVal curveDir As CURVE_DIR)
                        
    ' Raise error if start and center coordinates are equal
    If sX = cX And sY = cY Then
        Err.Raise 5, TypeName(Me), "Can't initiate circular arc. Start and center point must be different!"
    End If
    
    ' Raise error if length is less than or equal to 0
    If length <= 0 Then
        Err.Raise 5, TypeName(Me), "Can't initiate circular arc. Length is less than or equal to 0!"
    End If
    
    ' Raise error if curveDir is not defined
    If curveDir <> CD_CW And curveDir <> CD_CCW Then
        Err.Raise 5, TypeName(Me), "Can't initiate circular arc. Curve direction not defined!"
    End If
                        
    m_sX = sX
    m_sY = sY
    m_cX = cX
    m_cY = cY
    m_curveDir = curveDir
    
    ' Compute other members (m_e, m_rad, m_cToStheta, m_cToEtheta)
    m_rad = LibGeom.dist2D(sX, sY, cX, cY)
    m_sTheta = LibGeom.Atn2(sX - cX, sY - cY)
    
    ' If is circle
    If length >= LibGeom.TWO_PI * m_rad Then
        m_isCircle = True
        m_length = LibGeom.TWO_PI * m_rad
        m_eTheta = m_sTheta
        m_eX = m_sX
        m_eY = m_sY
    Else
        m_isCircle = False
        m_length = length
        m_eTheta = LibGeom.normalizeAngle(m_sTheta, LibGeom.PI) - length / m_rad * curveDir
        m_eTheta = LibGeom.normalizeAngle(m_eTheta, 0#) ' normalize theta in 0.0
        m_eX = m_cX + m_rad * Math.Cos(m_eTheta)
        m_eY = m_cY + m_rad * Math.Sin(m_eTheta)
    End If
End Sub

' Initialize by: start Point, end Point, radius and curve direction
' Raises error if:
'   - start and end coordinates are equal
'   - radius is less than or equal to 0
'   - curveDir is not CW (clocwise) or CCW (counter-clockwise)
Public Sub initFromSERD(ByVal sX As Double, ByVal sY As Double, _
                        ByVal eX As Double, ByVal eY As Double, _
                        ByVal rad As Double, ByVal curveDir As CURVE_DIR)
                        
    ' Raise error if start and end coordinates are equal
    If sX = eX And sY = eY Then
        Err.Raise 5, TypeName(Me), "Can't initiate circular arc. Start and end point must be different!"
    End If
    
    ' Raise error if length is less than or equal to 0
    If rad <= 0 Then
        Err.Raise 5, TypeName(Me), "Can't initiate circular arc. Radius is less than or equal to 0!"
    End If
    
    ' Raise error if curveDir is not defined
    If curveDir <> CD_CW And curveDir <> CD_CCW Then
        Err.Raise 5, TypeName(Me), "Can't initiate circular arc. Curve direction not defined!"
    End If
    
    m_sX = sX
    m_sY = sY
    m_eX = eX
    m_eY = eY
    m_rad = rad
    m_isCircle = False
    m_curveDir = curveDir
    
    ' Compute other members (m_c, m_length, m_cToStheta, m_cToEtheta)
    Dim q As Double                              ' Distance between s (start) point and e (end) point
    Dim x As Double                              ' X of q mid point
    Dim y As Double                              ' Y of q mid point
    
    q = LibGeom.dist2D(sX, sY, eX, eY)
    x = (sX + eX) / 2
    y = (sY + eY) / 2
    
    m_cX = x - curveDir * Math.Sqr(rad * rad - (q / 2) * (q / 2)) * ((sY - eY) / q)
    m_cY = y - curveDir * Math.Sqr(rad * rad - (q / 2) * (q / 2)) * ((eX - sX) / q)
    
    m_length = rad * LibGeom.ACos(1 - q * q / (2 * rad * rad))
    
    m_sTheta = LibGeom.Atn2(sX - m_cX, sY - m_cY)
    m_eTheta = LibGeom.Atn2(eX - m_cX, eY - m_cY)
End Sub

' Returns True if the theta of the ray with the start point in the
' center of this CircularArc intersects this CircularArc,
' False otherwise
Public Function isThetaOnArc(ByVal t As Double) As Boolean
    If m_isCircle Then
        isThetaOnArc = True
        Exit Function
    End If
    
    Dim st As Double                             ' start theta
    Dim et As Double                             ' end theta
    
    ' If curve direction is CW then swap start and end theta
    Select Case m_curveDir
    Case CD_CW
        st = m_eTheta
        et = m_sTheta
    Case CD_CCW
        st = m_sTheta
        et = m_eTheta
    Case Else
        isThetaOnArc = False
        Exit Function
    End Select
    
    If st < et Then
        isThetaOnArc = st <= t And t <= et
    Else                                         ' wrap angle
        isThetaOnArc = (st <= t And t <= LibGeom.PI) _
        Or (-LibGeom.PI <= t And t <= et)
    End If
End Function

' Returns a Point which is at a specified offset (o) from this CircularArc
' and has its projection on this CircularArc at a specified length (m) from
' one of its ends
' Parameters:
'   - m (measure): the distance measured along this CircularArc: positive
'                  is measured from Start and negative from End
'   - o (offset): the perpendicular distance measured against this CircularArc
' Result:
'   - A Point object with the computed coordinates
'   - Nothing if the distance in the m parameter is greater than this CircularArc
'     length or less than 0
Public Function calcPointByMeasOffset(ByVal m As Double, ByVal o As Double) As Point

    If m < 0 Or m > m_length Then
        Set calcPointByMeasOffset = Nothing
        Exit Function
    End If
    
    Dim pTheta As Double                         ' Theta from the arc center to the result point
    Dim pDist As Double                          ' Distance from the arc center to the result point
        
    pTheta = LibGeom.normalizeAngle(m_sTheta, LibGeom.PI) - m / m_rad * curveDirection
    pTheta = LibGeom.normalizeAngle(pTheta, 0#)  ' normalize theta back to 0.0
    pDist = m_rad - o * m_curveDir
    
    Set calcPointByMeasOffset = New Point
    calcPointByMeasOffset.x = m_cX + pDist * Math.Cos(pTheta)
    calcPointByMeasOffset.y = m_cY + pDist * Math.Sin(pTheta)
End Function

' Returns a MeasOffset which is linear referencing a set of coordinates
' against this CircularArc
' Result:
'   - MeasOffset object with the computed measure and offset distances
'   - Nothing if the projection of the given coordinates doesn't lie
'     within this CircularArc
Public Function calcMeasOffsetOfPoint(ByVal x As Double, ByVal y As Double) As MeasOffset
    Dim pTheta1 As Double
    Dim pTheta2 As Double
    Dim refTheta As Double
    
    ' Return Nothing if input point is the center of the circular arc
    If x = m_cX And y = m_cY Then
        Set calcMeasOffsetOfPoint = Nothing
        Exit Function
    End If
    
    ' Set reference theta
    If m_curveDir = CD_CW Then
        refTheta = m_eTheta
    Else
        refTheta = m_sTheta
    End If
    
    
    pTheta1 = LibGeom.Atn2(x - m_cX, y - m_cY)   ' Theta of arc center to input point line
    pTheta2 = LibGeom.normalizeAngle(pTheta1 + LibGeom.PI, 0) ' Theta of input point to arc center line
                                      
    Dim p1OnArc As Boolean                       ' pTheta1 intersects circular arc
    Dim p2OnArc As Boolean                       ' pTheta2 intersects circular arc
    
    If m_isCircle Then
        p1OnArc = True
        p2OnArc = True
    Else
        p1OnArc = Me.isThetaOnArc(pTheta1)
        p2OnArc = Me.isThetaOnArc(pTheta2)
    End If
       
    ' NO intersection
    If Not p1OnArc And Not p2OnArc Then
        Set calcMeasOffsetOfPoint = Nothing
        Exit Function
    End If
    
    Set calcMeasOffsetOfPoint = New MeasOffset
    Dim mAngle As Double                         ' measure angle
    Dim p1NotOnArcAndp2OnArcFlag As Integer      ' indicator for p1 not on arc and p2 on arc
    p1NotOnArcAndp2OnArcFlag = 1
    
    If p1OnArc And Not p2OnArc Then              ' ONE intersection, p1
        mAngle = LibGeom.normalizeAngle(pTheta1, refTheta) - refTheta
    ElseIf Not p1OnArc And p2OnArc Then          ' ONE intersection, p2
        mAngle = LibGeom.normalizeAngle(pTheta2, refTheta) - refTheta
        p1NotOnArcAndp2OnArcFlag = -1
    Else                                         ' TWO intersections
        mAngle = LibGeom.normalizeAngle(pTheta1, refTheta) - refTheta
    End If
            
    mAngle = LibGeom.normalizeAngle(mAngle, LibGeom.PI)
    
    ' compute mTheta as difference of total angle and actual angle if start and end were swapped
    If m_curveDir = CD_CW Then mAngle = m_length / m_rad - mAngle
    
    ' if measure equals to full circle back to start point
    If Abs(mAngle) >= LibGeom.TWO_PI Then mAngle = 0
    
    calcMeasOffsetOfPoint.m = Abs(m_rad * mAngle) ' set measure

    Dim d As Double                              ' distance from center to input point
    d = LibGeom.dist2D(m_cX, m_cY, x, y)
    calcMeasOffsetOfPoint.o = (m_rad - d * p1NotOnArcAndp2OnArcFlag) * m_curveDir ' set offset
End Function

' Returns the X value at a given Y value
' Computation based on the circle equation (X-Xc)^2 + (Y-Yc)^2 = r^2
' Result:
'   - X value computed at the given Y value
'   - Null if Y value is out of range
'   - Null if multiple Xs can be returned (Y line intersects circular arc in two points)
Public Function calcXatY(ByVal y As Double) As Variant
    If m_isCircle Then
        calcXatY = Null
        Exit Function
    End If
        
    ' IF Y is between Ys of the circle defining the arc THEN compute X ELSE return Null
    If m_cY - m_rad <= y And y <= m_cY + m_rad Then
        Dim dXabs As Double ' delta X modulus (absolute value)
        Dim x1 As Double ' first solution of circle equation
        Dim t1 As Double ' theta to x1
        Dim x1onArc As Boolean
        Dim x2 As Double ' second solution of circle equation
        Dim t2 As Double ' theta to x2
        Dim x2onArc As Boolean
        
        dXabs = Math.Sqr(m_rad * m_rad - (y - m_cY) * (y - m_cY))
        x1 = m_cX - dXabs
        t1 = LibGeom.Atn2(x1 - m_cX, y - m_cY)
        x1onArc = Me.isThetaOnArc(t1)
        x2 = m_cX + dXabs
        t2 = LibGeom.Atn2(x2 - m_cX, y - m_cY)
        x2onArc = Me.isThetaOnArc(t2)
        
        If x1onArc And Not x2onArc Then         ' one intersection: x1
            calcXatY = x1
        ElseIf Not x1onArc And x2onArc Then     ' one intersection: x2
            calcXatY = x2
        ElseIf x1onArc And x2onArc Then
            If Abs(x1 - x2) < 0.000000000000001 Then
                calcXatY = x1                   ' two intersection almost coincident (1E-15): return one of them
            Else
                calcXatY = Null                 ' two distinct intersections: return Null
            End If
        Else                                    ' no intersection: return Null
            calcXatY = Null
        End If
    Else
        calcXatY = Null                         ' out of range: return Null
    End If
End Function

' Returns the Y value at a given X value
' Computation based on the circle equation (X-Xc)^2 + (Y-Yc)^2 = r^2
' Result:
'   - Y value computed at the given X value
'   - Null if X value is out of range
'   - Null if multiple Ys can be returned (Y line intersects circular arc in two points)
Public Function calcYatX(ByVal x As Double) As Variant
    If m_isCircle Then
        calcYatX = Null
        Exit Function
    End If
        
    ' IF X is between Xs of the circle defining the arc THEN compute Y ELSE return Null
    If m_cX - m_rad <= x And x <= m_cX + m_rad Then
        Dim dYabs As Double ' delta X modulus (absolute value)
        Dim y1 As Double ' first solution of circle equation
        Dim t1 As Double ' theta to x1
        Dim y1onArc As Boolean
        Dim y2 As Double ' second solution of circle equation
        Dim t2 As Double ' theta to x2
        Dim y2onArc As Boolean
        
        dYabs = Math.Sqr(m_rad * m_rad - (x - m_cX) * (x - m_cX))
        y1 = m_cY - dYabs
        t1 = LibGeom.Atn2(x - m_cX, y1 - m_cY)
        y1onArc = Me.isThetaOnArc(t1)
        y2 = m_cY + dYabs
        t2 = LibGeom.Atn2(x - m_cX, y2 - m_cY)
        y2onArc = Me.isThetaOnArc(t2)
        
        If y1onArc And Not y2onArc Then         ' one intersection: y1
            calcYatX = y1
        ElseIf Not y1onArc And y2onArc Then     ' one intersection: y2
            calcYatX = y2
        ElseIf y1onArc And y2onArc Then
            If Abs(y1 - y2) < 0.000000000000001 Then
                calcYatX = y1                   ' two intersection almost coincident (1E-15): return one of them
            Else
                calcYatX = Null                 ' two distinct intersections: return Null
            End If
        Else                                    ' no intersection: return Null
            calcYatX = Null
        End If
    Else
        calcYatX = Null                         ' out of range: return Null
    End If
End Function

' Test if this CircularArc is equal to another CircularArc object
' Returns TRUE if all members are equal
Public Function isEqual(other As CircularArc) As Boolean
    isEqual = m_sX = other.sX And m_sY = other.sY And _
              m_cX = other.cX And m_cY = other.cY And _
              m_eX = other.eX And m_eY = other.eY And _
              m_length = other.length And _
              m_curveDir = other.curveDirection And _
              m_rad = other.radius
    m_sTheta = other.sTheta And m_eTheta = other.eTheta And _
               m_isCircle = other.isCircle
End Function

' Returns a string representation of the this CircularArc
Public Function toString() As String
    Dim s As New Point
    Dim e As New Point
    Dim c As New Point
    
    s.init m_sX, m_sY
    e.init m_eX, m_eY
    c.init m_cX, m_cY
    
    dir = ConstCL.curveDirToString(m_curveDir)
    toString = "Start: " & s.toStringXY & ";" & _
               "End: " & e.toStringXY & ";" & _
               "Center: " & c.toStringXY & ";" & _
               "Radius: " & m_rad & ";" & _
               "Start Theta: " & m_sTheta & ";" & _
               "End Theta: " & m_eTheta & ";" & _
               "Length: " & m_length & ";" & _
               "Curve Direction: " & ConstCL.curveDirToString(m_curveDir) & ";" & _
               "Is Circle: " & m_isCircle
End Function

' Start Point X of circular arc property
Public Property Get sX() As Double
    sX = m_sX
End Property

' Start Point Y of circular arc property
Public Property Get sY() As Double
    sY = m_sY
End Property

' Center Point X of circular arc property
Public Property Get cX() As Double
    cX = m_cX
End Property

' Center Point Y of circular arc property
Public Property Get cY() As Double
    cY = m_cY
End Property

' End Point X of circular arc property
Public Property Get eX() As Double
    eX = m_eX
End Property

' End Point Y of circular arc property
Public Property Get eY() As Double
    eY = m_eY
End Property

' Length of circular arc property
Public Property Get length() As Double
    length = m_length
End Property

' Length of circular arc property
Public Property Get curveDirection() As CURVE_DIR
    curveDirection = m_curveDir
End Property

' Radius of circular arc property
Public Property Get radius() As Double
    radius = m_rad
End Property

' Theta of center to start point line
Public Property Get sTheta() As Double
    sTheta = m_sTheta
End Property

' Theta of center to start point line
Public Property Get eTheta() As Double
    eTheta = m_eTheta
End Property

' Is circle
Public Property Get isCircle() As Double
    isCircle = m_isCircle
End Property

' IGeom interface
Private Function IGeom_calcPointByMeasOffset(ByVal m As Double, ByVal o As Double) As Point
    Set IGeom_calcPointByMeasOffset = Me.calcPointByMeasOffset(m, o)
End Function

Private Function IGeom_calcMeasOffsetOfPoint(ByVal x As Double, ByVal y As Double) As MeasOffset
    Set IGeom_calcMeasOffsetOfPoint = Me.calcMeasOffsetOfPoint(x, y)
End Function

Private Function IGeom_calcXatY(ByVal y As Double) As Variant
    IGeom_calcXatY = Me.calcXatY(y)
End Function

Private Function IGeom_calcYatX(ByVal x As Double) As Variant
    IGeom_calcYatX = Me.calcYatX(x)
End Function

Private Property Get IGeom_length() As Double
    IGeom_length = Me.length
End Property

Public Property Get IGeom_geomTypeName() As String
    IGeom_geomTypeName = ConstCL.CA_NAME
End Property





