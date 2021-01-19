VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClothoidArc"
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
'' Theory:
''   A clothoid arc (also known as Euler Spiral or Cornu Spiral) is a type
''     of curve that has its curvature proportional to it's length. In
''     other words the product of length and curvature radius is constant
''     along the clothoid.
''   The clothoid has a util (useful) arc comprised between the start
''     point (origin) and the point where the tangent to the clothoid is
''     normal (perpendicular) on the abscissa axis.
''   Wiki: https://en.wikipedia.org/wiki/Fresnel_integral#Euler_spiral
''
'' Class Description:
''   Represents a clothoid arc that cannot be larger than the clothoid's
''     util arc - the angle formed by the abscissa axis and the tangent
''     on the endpoint is less than or equal to PI/2. For the util arc a
''     Taylor Series Expansion provides sufficient precision (2E-7) while
''     for larger values a set of efficient approximations based on
''     rational functions would be more suitable
''========================================================================

'@Folder("TopoXL.CL.geom")

Option Explicit
Implements IGeom

' Internal Private Types
Private Type ANGLE_STRUCT
    val_ As Double
    sin_ As Double
    cos_ As Double
End Type

Private Type POINT_STRUCT
    x As Double
    y As Double
End Type

' Class Members
'*******************************************************************************
Private m_startPt As POINT_STRUCT                ' origin
Private m_endPt As POINT_STRUCT                  ' clothoid point at m_length
Private m_intPt As POINT_STRUCT                  ' intersection of perpendiculars on start and end point
Private m_peakPt As POINT_STRUCT                 ' any point on the internal X (abscissa) axis that forms m_startTheta with m_startPt
Private m_startTheta As ANGLE_STRUCT             ' internal X (abscissa) axis theta
Private m_length As Double                       ' length of the clothoid arc
Private m_endRadius As Double                    ' radius value at the end point
Private m_curveDir As CURVE_DIR                  ' clothoid arc curve direction (CW or CCW)
Private m_const As Double                        ' clothoid constant = 2 * rad * len
'*******************************************************************************

'*******************************************************************************
' Initialize by: start Point, start theta, length, end radius and curve direction
' Raises error 5 if:
'   - length/endRadius is less than or equal to 0
'   - curveDir is not CW (clocwise) or CCW (counter-clockwise)
'   - the arc is not a util arc (tangent at end Point creates an angle larger
'     than PI/2 with then internal X axis)
'*******************************************************************************
Public Sub init(ByVal startX As Double, ByVal startY As Double, _
                ByVal startTheta As Double, _
                ByVal length As Double, ByVal endRadius As Double, _
                ByVal curveDirection As CURVE_DIR)
    Const methodName As String = "init"
                
    If length <= 0 Then
        Err.Raise 5, TypeName(Me) & "." & methodName, "Length is zero or a negative number!"
    ElseIf endRadius <= 0 Then
        Err.Raise 5, TypeName(Me) & "." & methodName, "EndRadius is zero or a negative number!"
    ElseIf curveDirection <> CD_CW And curveDirection <> CD_CCW Then
        Err.Raise 5, TypeName(Me) & "." & methodName, "Invalid curve direction!"
    ElseIf length / endRadius > LibGeom.PI Then  'Equivalent of: alphaAt(length) > PI / 2
        Err.Raise 5, TypeName(Me) & "." & methodName, "The arc exceeds the util arc of the clothoid"
    End If

    m_startPt.x = startX
    m_startPt.y = startY

    m_startTheta.val_ = startTheta
    m_startTheta.sin_ = Math.Sin(startTheta)
    m_startTheta.cos_ = Math.Cos(startTheta)

    m_length = length
    m_endRadius = endRadius
    m_curveDir = curveDirection
    m_const = 2 * m_endRadius * m_length
    
    Dim endPt As Point: Set endPt = Me.calcPointByMeasOffset(m_length, 0#)
    m_endPt.x = endPt.x
    m_endPt.y = endPt.y
    
    Dim intPt As Point
    Set intPt = LibGeom.intOf2LinesByTheta( _
                m_startPt.x, m_startPt.y, _
                m_startTheta.val_ - m_curveDir * LibGeom.PI / 2, _
                m_endPt.x, m_endPt.y, _
                m_startTheta.val_ + m_curveDir * (LibGeom.PI / 2 - alphaAt(m_length)))
    m_intPt.x = intPt.x
    m_intPt.y = intPt.y

    m_peakPt.x = m_startPt.x + m_startTheta.cos_ * m_length
    m_peakPt.y = m_startPt.y + m_startTheta.sin_ * m_length
End Sub

'*******************************************************************************
' Returns a Point which is at a specified length along the clothoid arc (m) and
'   at a specified offset (o)
' Parameters:
'   - m (measure): the distance measured along the ClothoidArc -always positive
'   - o (offset): the perpendicular distance from the clothoid's tangent line in
'                 the point defined my measure
' Notes:
'   - Returns Nothing if the measure is negative or a number larger than the
'     total length of the clothoid arc
'*******************************************************************************
Public Function calcPointByMeasOffset(ByVal m As Double, ByVal o As Double) As Point
    If m < 0 Or m > m_length Then Exit Function
    
    Set calcPointByMeasOffset = New Point
    
    If m = 0 Then
        calcPointByMeasOffset.x = m_startPt.x + o * m_startTheta.sin_
        calcPointByMeasOffset.y = m_startPt.y - o * m_startTheta.cos_
    Else
        Dim alpha As Double: alpha = alphaAt(m) * -m_curveDir
        Dim dX As Double: dX = m * fresnelCosine(alpha) ' abscissa (local)
        Dim dY As Double: dY = m * fresnelSine(alpha) ' ordinate (local)
    
        calcPointByMeasOffset.x = m_startPt.x - dX * m_startTheta.sin_ + dY * m_startTheta.cos_ + o * Math.Sin(m_startTheta.val_ + alpha)
        calcPointByMeasOffset.y = m_startPt.y + dX * m_startTheta.cos_ + dY * m_startTheta.sin_ - o * Math.Cos(m_startTheta.val_ + alpha)
    End If
End Function

'*******************************************************************************
' Returns a MeasOffset which is linearly referencing a set of coordinates
'   against this ClothoidArc
' An initial guess of the measure is computed and then a number (capped) of
'   iterations improve the measure until a geometric condition is fulfilled.
' Geometric condition (used the same variable names as in the method below):
'   - startToFootDist must equal the sum of 3 distances:
'     1) dy
'     2) dx * Math.Tan(tempAlpha)
'     3) perpDist * Math.Tan(tempAlpha)
'   - Depending on the location of the input point, the 3 distances above
'       have + or - sign in calculus
'   - If the input point is located on the start to peak line that the
'       geometric condition above uses only the first 2 distances with a + sign
' Notes:
'   - Returns Nothing if the projection of the given coordinates doesn't lie
'       within this ClothoidArc or the required precision was not fulfilled
'*******************************************************************************
Public Function calcMeasOffsetOfPoint(ByVal x As Double, ByVal y As Double) As MeasOffset
    Const epsilon As Double = 0.00000001
    Const maxIterations As Integer = 100
    
    Dim sideIntStart As Integer                  ' side of the input point related to intToStart line
    Dim sideIntEnd As Integer                    ' side of the input point related to intToEnd line
    Dim sideStartEnd As Integer                  ' side of the input point related to startToEnd line
    
    sideIntStart = LibGeom.orientationIndex(m_intPt.x, m_intPt.y, m_startPt.x, m_startPt.y, x, y, epsilon)
    sideIntEnd = LibGeom.orientationIndex(m_intPt.x, m_intPt.y, m_endPt.x, m_endPt.y, x, y, epsilon)
    
    If sideIntStart * sideIntEnd = 1 Then
        If sideIntStart = -m_curveDir Then Exit Function ' not on the clothoid. Return Nothing
    ElseIf sideIntStart * sideIntEnd = 0 Then
        sideStartEnd = LibGeom.orientationIndex(m_startPt.x, m_startPt.y, m_endPt.x, m_endPt.y, x, y, epsilon)
        Set calcMeasOffsetOfPoint = New MeasOffset
        If sideIntStart = 0 Then
            calcMeasOffsetOfPoint.m = 0
            calcMeasOffsetOfPoint.o = LibGeom.dist2D(m_startPt.x, m_startPt.y, x, y) * sideStartEnd
        ElseIf sideIntEnd = 0 Then
            calcMeasOffsetOfPoint.m = m_length
            calcMeasOffsetOfPoint.o = LibGeom.dist2D(m_endPt.x, m_endPt.y, x, y) * sideStartEnd
        End If
        Exit Function
    End If

    Dim sideStartPeek As Integer                 ' side of the input point related to startToPeak line
    Dim startToFootDist As Double                ' the distance from the start point to the projection of the input point on the start to peak line
    Dim measure As Double                        ' the length of the clothoid arc (improved by iterations)
    Dim tempAlpha As Double                      ' the angle formed by clothoid's tangent and startToPeak line
    Dim dX As Double                             ' abscissa (local)
    Dim dY As Double                             ' ordinate (local)
    Dim sumDist As Double                        ' explained in the function header above
    Dim i As Integer                             ' iterator
    Dim isMeasFound As Boolean                   ' indicator that measure has been found in the desired precision
    
    sideStartPeek = -m_curveDir * LibGeom.orientationIndex(m_startPt.x, m_startPt.y, m_peakPt.x, m_peakPt.y, x, y, epsilon)
    
    If sideStartPeek = 0 Then
        ' The input point is on the startToPeak line
        startToFootDist = LibGeom.dist2D(m_startPt.x, m_startPt.y, x, y)
        measure = LibMath.min(startToFootDist, m_length)
        For i = 1 To maxIterations
            tempAlpha = alphaAt(measure)
            dX = measure * fresnelCosine(tempAlpha)
            dY = measure * fresnelSine(tempAlpha)
            sumDist = dX * Math.Tan(tempAlpha) + dY
            measure = measure + startToFootDist - sumDist
            isMeasFound = LibMath.areDoublesEqual(sumDist, startToFootDist, epsilon)
            If isMeasFound Then Exit For
        Next i
        If Not isMeasFound Then Exit Function    ' return Nothing
        '
        Set calcMeasOffsetOfPoint = New MeasOffset
        calcMeasOffsetOfPoint.m = measure
        calcMeasOffsetOfPoint.o = -m_curveDir * dX / Math.Cos(tempAlpha)
        Exit Function
    End If

    Dim footPt As Point                          ' the foot of the perpendicular from the input point to startToPeak line
    Dim perpDist As Double                       ' the distance from the input point to footPt point
    Dim tempMeas As Double                       ' temporary measure
    Dim diff As Double                           ' the difference between the previous and the current calculated measure
    Dim j As Integer                             ' sign for perpDist
    Dim k As Integer                             ' sign for dX and dY
    
    Set footPt = LibGeom.footOfPerpendicular(m_startPt.x, m_startPt.y, m_peakPt.x, m_peakPt.y, x, y, epsilon)
    startToFootDist = LibGeom.dist2D(m_startPt.x, m_startPt.y, footPt.x, footPt.y)
    perpDist = LibGeom.dist2D(x, y, footPt.x, footPt.y)
    measure = LibMath.min(startToFootDist, m_length)
    If perpDist > measure * 5 And m_length / 2 > measure Then measure = m_length / 2 ' adjust for a sharp angle
    k = sideIntStart * m_curveDir
    j = k * sideStartPeek
    
    For i = 1 To maxIterations
        tempAlpha = alphaAt(measure)
        dX = measure * fresnelCosine(tempAlpha)
        dY = measure * fresnelSine(tempAlpha)
        sumDist = Math.Tan(tempAlpha) * (j * perpDist + k * dX) + k * dY
        tempMeas = Math.Sqr(tempAlpha * startToFootDist / Math.Abs(sumDist) * m_const)
        diff = Math.Abs(measure - tempMeas)
        measure = LibMath.min(tempMeas, m_length)
        isMeasFound = LibMath.areDoublesEqual(sumDist, startToFootDist, epsilon)
        If isMeasFound Then Exit For
        If diff = 0 Then Exit For                ' stuck. Result will not improve anymore
    Next i
    If Not isMeasFound Then Exit Function        ' return nothing
    
    Set calcMeasOffsetOfPoint = New MeasOffset
    calcMeasOffsetOfPoint.m = measure
    calcMeasOffsetOfPoint.o = -m_curveDir * sideStartPeek / Cos(tempAlpha) * (perpDist + sideStartPeek * dX)
End Function

'*******************************************************************************
' The angle formed by the clothoid's tangent (at a certain length) and the start
'   to peek theta (the local clothoid abscissa axis)
' ---
' (1) The product of length and curvature radius is constant along the clothoid,
'   so for any 2 points 'p1' and 'p2' the following is true:
'   length_p1 * radius_p1 = length_p2 * radius_p2 = m_len * m_endRad
' (2) For any point 'pt' along the clothoid, the alpha angle is:
'   alpha_pt = length_pt / (2 * radius_pt)
' Combining (1) + (2):
'   alpha_pt = length_pt / (2 * m_endRad * m_rad / length_pt)
'   alpha_pt = length_pt / (m_const / length_pt)
'   alpha_pt = length_pt * length_pt / m_const
'*******************************************************************************
Private Function alphaAt(ByVal length As Double) As Double
    alphaAt = length * length / m_const
End Function

'*******************************************************************************
' Fresnel integral sine S(x)=integral[0,x]sin(y^2)dy aproximated by
'   Taylor series expansion with first 7 elements
' Represents the ordinate of the clothoid in local system
'*******************************************************************************
Private Function fresnelSine(ByVal alpha As Double) As Double
    fresnelSine = 1 - alpha ^ 2 / 10# + _
                  alpha ^ 4 / 216# - _
                  alpha ^ 6 / 9360# + _
                  alpha ^ 8 / 685440# - _
                  alpha ^ 10 / 76204800# + _
                  alpha ^ 12 / 11975040000#
End Function

'*******************************************************************************
' Fresnel integral cosine C(x)=integral[0,x]cos(y^2)dy aproximated by
'   Taylor series expansion with first 6 elements
' Represents the abscissa of the clothoid in local system
'*******************************************************************************
Private Function fresnelCosine(ByVal alpha As Double) As Double
    fresnelCosine = alpha / 3# - _
                    alpha ^ 3 / 42# + _
                    alpha ^ 5 / 1320# - _
                    alpha ^ 7 / 75600# + _
                    alpha ^ 9 / 6894720# - _
                    alpha ^ 11 / 918086400#
End Function

' Start Point X of clothoid arc
Public Property Get sX() As Double
    sX = m_startPt.x
End Property

' Start Point Y of clothoid arc
Public Property Get sY() As Double
    sY = m_startPt.y
End Property

' Length of clothoid arc
Public Property Get length() As Double
    length = m_length
End Property

' End Radius of clothoid arc
Public Property Get endRadius() As Double
    endRadius = m_endRadius
End Property

' Length of clothoid arc
Public Property Get curveDirection() As CURVE_DIR
    curveDirection = m_curveDir
End Property

' Theta of internal X (abscissa) axis
Public Property Get startTheta() As Double
    startTheta = m_startTheta.val_
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
    IGeom_geomTypeName = ConstCL.CLA_NAME
End Property
