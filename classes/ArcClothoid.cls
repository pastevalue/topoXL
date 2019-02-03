VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ArcClothoid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''=======================================================
''Called by:
''    Modules: None
''    Classes: RLarcClothoid
''Calls:
''    Modules: RLenums, UDF_COGO
''    Classes: Coo2D
''=======================================================
Option Explicit

'primary members
Private m_startCoo As Coo2D 'arc clothoid start point (tangent in)
Private m_peak As Coo2D 'arc clothoid peak
Private m_rad As Double 'arc clothoid radius
Private m_len As Double 'arc clothoid length
Private m_curveDir As CURVE_DIRECTION

'secondary members
Private m_clothConst As Double 'arc clothoid constant
Private m_startToPeakAz As Double 'azimuth from start point to peak point
Private m_endCoo As Coo2D 'arc clothoid end point (tangent out)
Private m_center As Coo2D 'arc clothoid center

'Class constructor with default values
Private Sub Class_Initialize()
    Set m_peak = New Coo2D
    Set m_startCoo = New Coo2D
    m_len = 0
    m_rad = 0
    m_clothConst = 0
    m_curveDir = CURVE_DIRECTION.NONE
    Set m_endCoo = New Coo2D
    Set m_center = New Coo2D
    m_startToPeakAz = 0
End Sub

'Initialize class
Public Sub init(clothoidIn As Coo2D, peak As Coo2D, length As Double, radius As Double, curveDirection As CURVE_DIRECTION)
    Set m_startCoo = clothoidIn
    Set m_peak = peak
    m_rad = radius
    m_len = length
    m_clothConst = 2 * m_rad * m_len
    m_curveDir = curveDirection
    Call setClothoidStartToPeakAzimuth
    Call setClothoidEndCoo
    Call setArcClothoidCenter
End Sub

'Gets arc clothoid length
Public Function getLength() As Double
    getLength = m_len
End Function

'Sets azimuth of start point to peak line
Private Sub setClothoidStartToPeakAzimuth()
    m_startToPeakAz = cogoAzimuth(m_startCoo.x, m_startCoo.y, m_peak.x, m_peak.y)
End Sub

'Sets arc clothoid end coordinates
Private Sub setClothoidEndCoo()
    Set m_endCoo = getCoo(m_len, 0)
End Sub

'Sets arc clothoid center coordinates
Private Sub setArcClothoidCenter()
    Dim startToCenAz As Double 'start to center azimuth
    Dim endToCenAz As Double 'end to center azimuth
    Dim alpha As Double

    startToCenAz = m_startToPeakAz + PI / 2 * m_curveDir
    alpha = m_len ^ 2 / m_clothConst
    endToCenAz = m_startToPeakAz + m_curveDir * (alpha - PI / 2)
    
    Dim tmpCenterX As Double
    Dim tmpCenterY As Double
    
    cogoIntersectionOf2LinesByCooAndAzimuth m_startCoo.x, m_startCoo.y, startToCenAz, m_endCoo.x, m_endCoo.y, endToCenAz, tmpCenterX, tmpCenterY
     m_center.x = tmpCenterX
     m_center.y = tmpCenterY
End Sub

'Gets coordinates at specified distance and offset
Public Function getCoo(dist As Double, offset As Double) As Coo2D
    
    If (dist > m_len) Or (dist < 0) Then
        getCoo = Null
        Debug.Print "Total length of the clothoid is " & m_len & ". Can't calculate coordinates at a distance of " & dist & "!"
        Exit Function
    End If
    
    Dim alpha As Double
    Dim dX As Double 'abscissa of clothoid in local system
    Dim dY As Double 'ordinate of clothoid in local system
    Dim temp_azimuth As Double
    Dim result As New Coo2D
    
    alpha = dist ^ 2 / m_clothConst
    dX = clothoidFresnelCosineDx(dist, alpha)
    dY = clothoidFresnelSineDy(dist, alpha)
    
    temp_azimuth = m_startToPeakAz + PI / 2 * m_curveDir

    result.x = m_startCoo.x + dY * Sin(m_startToPeakAz) + dX * Sin(temp_azimuth)
    result.y = m_startCoo.y + dY * Cos(m_startToPeakAz) + dX * Cos(temp_azimuth)
    
    If offset <> 0 Then
        temp_azimuth = m_startToPeakAz - PI / 2 * m_curveDir + alpha * m_curveDir
        result.x = result.x - offset * Sin(temp_azimuth) * m_curveDir
        result.y = result.y - offset * Cos(temp_azimuth) * m_curveDir
    End If
    Set getCoo = result
End Function

'Gets abscissa at specified distance and offset
Public Function getCooX(dist As Double, offset As Double) As Double
    
    If (dist > m_len) Or (dist < 0) Then
        GoTo ErrHandler
    End If
    
    Dim alpha As Double
    Dim dX As Double 'abscissa of clothoid in local system
    Dim dY As Double 'ordinate of clothoid in local system
    Dim temp_azimuth As Double
    
    alpha = dist ^ 2 / m_clothConst
    dX = clothoidFresnelCosineDx(dist, alpha)
    dY = clothoidFresnelSineDy(dist, alpha)
    
    temp_azimuth = m_startToPeakAz + PI / 2 * m_curveDir

    getCooX = m_startCoo.x + dY * Sin(m_startToPeakAz) + dX * Sin(temp_azimuth)
    
    If offset <> 0 Then
        temp_azimuth = m_startToPeakAz - PI / 2 * m_curveDir + alpha * m_curveDir
        getCooX = getCooX - offset * Sin(temp_azimuth) * m_curveDir
    End If
    
Exit Function
ErrHandler:
  Debug.Print "Total length of the clothoid is " & m_len & ". Can't calculate coordinates at a distance of " & dist & "!"
End Function


'Gets ordinate at specified distance and offset
Public Function getCooY(dist As Double, offset As Double) As Double
    
    If (dist > m_len) Or (dist < 0) Then
        GoTo ErrHandler
    End If
    
    Dim alpha As Double
    Dim dX As Double 'abscissa of clothoid in local system
    Dim dY As Double 'ordinate of clothoid in local system
    Dim temp_azimuth As Double
    
    alpha = dist ^ 2 / m_clothConst
    dX = clothoidFresnelCosineDx(dist, alpha)
    dY = clothoidFresnelSineDy(dist, alpha)
    
    temp_azimuth = m_startToPeakAz + PI / 2 * m_curveDir

    getCooY = m_startCoo.y + dY * Cos(m_startToPeakAz) + dX * Cos(temp_azimuth)
    
    If offset <> 0 Then
        temp_azimuth = m_startToPeakAz - PI / 2 * m_curveDir + alpha * m_curveDir
        getCooY = getCooY - offset * Cos(temp_azimuth) * m_curveDir
    End If
    
Exit Function
ErrHandler:
  Debug.Print "Total length of the clothoid is " & m_len & ". Can't calculate coordinates at a distance of " & dist & "!"
End Function

'Gets abscissa at specified ordinate
'TODO - de implementat functia
Public Function getX(y As Double) As Double
    getX = Null
End Function

'Gets ordinate at specified abscissa
'TODO - de implementat functia
Public Function getY(x As Double) As Double
    getY = Null
End Function

'For a set of coordinates returns the radial distance to the arc clothoid and the arc length from radial point to the start of the arc clothoid
Public Function getRadialDistFromPoint(ByVal px As Double, ByVal py As Double, ByRef distFromStart As Double, ByRef distFromArcClothoid As Double) As Boolean
    Const doublePrecision As Double = 0.000000001
    
    Dim sideOfPntCenterToStart As Integer 'specifies on which side of the CenterToStart line the input point (px, py) is (-1 to the left, 1 to the right, 0 on the line)
    Dim sideOfPntCenterToEnd As Integer   'specifies on which side of the CenterToEnd line the input point is (-1 to the left, 1 to the right, 0 on the line)
    
    sideOfPntCenterToStart = cogoGetSide(m_center.x, m_center.y, m_startCoo.x, m_startCoo.y, px, py)
    sideOfPntCenterToEnd = cogoGetSide(m_center.x, m_center.y, m_endCoo.x, m_endCoo.y, px, py)
    
    'Select case for radial point from input point
    Select Case sideOfPntCenterToStart * sideOfPntCenterToEnd
        
    'point is not on the arc clothoid
    Case 1
        distFromStart = 0
        distFromArcClothoid = 0
        getRadialDistFromPoint = False
    
    'point is on the arc clothoid
    Case -1
        Dim tempStartToPeakSlope As Double         'the slope of startToPeak line (for finding the foot of the perpendicular from input point to this line)
        Dim tempPerpOnStartToPeakLine As New Coo2D 'the foot of the perpendicular from the input point to startToPeak line
        Dim tempArc As Double                      'the distance from the start point to the tempPerpOnStartToPeakLine point (used in calculation and _
                                                        as first aproximation for arcClothoid arc length from start to radial point)
        Dim tempOffset As Double                   'the distance from the input point to tempPerpOnStartToPeakLine point
        Dim tempAlpha As Double                    'the angle formed by inputPointToRadialPoint and perpendicularOnStartToPeakLine directions _
                                                        or the angle of curve from beginning of spiral (infinite Radius) to a particular point on the spiral
        Dim tempSide As Integer                    'side of the input point related to startToPeak line
        Dim dX As Double                           'the distance from the radial point to the startToPeakLine also known as the Fresnel cosine integral _
                                                        C(x)=integral[0,x]cos(y^2)dy
        Dim dY As Double                           'the distance from the start point to the foot of the perpendicular from radial point to startToPeakLine _
                                                        also known as the Fresnel sine integral S(x)=integral[0,x]sin(y^2)dy
        Dim sumDist As Double                      'the sum of 3 distances which must equal the tempArc. These 3 distances are: _
                                                        dy, dx*tan(tempAlpha), tempOffset*tan(tempAlpha). Depending on the location of the input point related _
                                                        to the clothoid elements, these 3 distances have + or - sign in calculus
        Dim varLength As Double                    'the length of the clothoid arc from start to radial point which is improved by iterations
        Dim inputPointIsBeyondCenter As Boolean         'the input point has corresponding radial point on the clothoid, but if the input point is _
                                                        beyond the intersection of the centerToStart and centerToEnd lines related to the clothoid, _
                                                        then some signs will change in calculus
        
        If sideOfPntCenterToStart <> cogoGetSide(m_center.x, m_center.y, m_startCoo.x, m_startCoo.y, m_endCoo.x, m_endCoo.y) Then _
            inputPointIsBeyondCenter = True Else inputPointIsBeyondCenter = False
        tempSide = -cogoGetSide(m_startCoo.x, m_startCoo.y, m_peak.x, m_peak.y, px, py) * m_curveDir
        If tempSide <> 0 Then 'the input point is not on the startToPeak line
            tempStartToPeakSlope = (m_peak.y - m_startCoo.y) / (m_peak.x - m_startCoo.x)
            tempPerpOnStartToPeakLine.x = m_peak.x + (py - m_peak.y + 1 / tempStartToPeakSlope * (px - m_peak.x)) / (tempStartToPeakSlope + 1 / tempStartToPeakSlope)
            tempPerpOnStartToPeakLine.y = tempStartToPeakSlope * (tempPerpOnStartToPeakLine.x - m_peak.x) + m_peak.y
            tempArc = cogoDistance2D(m_startCoo.x, m_startCoo.y, tempPerpOnStartToPeakLine.x, tempPerpOnStartToPeakLine.y)
            
            'the first assumed length of the clothoid arc from start to radial point (varLength)
            If tempArc > m_len Then
                varLength = m_len
             Else
                varLength = tempArc
            End If
            
            tempAlpha = varLength ^ 2 / m_clothConst
            tempOffset = cogoDistance2D(px, py, tempPerpOnStartToPeakLine.x, tempPerpOnStartToPeakLine.y)
            
            'the Fresnel sine and cosine integrals are aproximated by Taylor series expansion
            dX = clothoidFresnelCosineDx(varLength, tempAlpha)
            dY = clothoidFresnelSineDy(varLength, tempAlpha)
            
            'the difference between the previous and the current calculated arc length
            Dim diff As Double
            'temporary variables to keep track if the iterations are improving the result and if not then make the necessary adjustements
            Dim tempVarLength As Double
            Dim tempDiff As Double
            
            diff = 0
            tempVarLength = 0
            tempDiff = 0
            
            'the varible used to count the steps of the iterations
            Dim i As Integer
            'temp variables for signs
            Dim j, k As Integer
            j = tempSide
            k = 1
            If inputPointIsBeyondCenter Then
                j = 1
                k = -1
            End If
                        
            sumDist = Tan(tempAlpha) * (j * tempOffset + k * dX) + k * dY
            varLength = Sqr(tempAlpha * tempArc / Abs(sumDist) * m_clothConst)
            For i = 0 To 100
                If Abs(sumDist - tempArc) < doublePrecision Then Exit For 'exit for if precision is achieved
                tempAlpha = varLength ^ 2 / m_clothConst
                dX = clothoidFresnelCosineDx(varLength, tempAlpha)
                dY = clothoidFresnelSineDy(varLength, tempAlpha)
                sumDist = Tan(tempAlpha) * (j * tempOffset + k * dX) + k * dY
                tempVarLength = Sqr(tempAlpha * tempArc / Abs(sumDist) * m_clothConst)
                If i = 0 Then
                    diff = Abs(varLength - tempVarLength)
                Else
                    tempDiff = Abs(varLength - tempVarLength)
                    If tempDiff > diff Then
                        If varLength > tempVarLength Then
                            varLength = varLength - diff / 2
                        Else
                            varLength = varLength + diff / 2
                        End If
                    Else
                        varLength = tempVarLength
                    End If
                End If
            Next i
            distFromArcClothoid = -m_curveDir * tempSide / Cos(tempAlpha) * (tempOffset + tempSide * dX)
         Else 'the input point is on the startToPeak line
            tempArc = cogoDistance2D(m_startCoo.x, m_startCoo.y, px, py)
            If tempArc > m_len Then
                varLength = m_len
             Else
                varLength = tempArc
            End If
            tempAlpha = varLength ^ 2 / m_clothConst
            dX = clothoidFresnelCosineDx(varLength, tempAlpha)
            dY = clothoidFresnelSineDy(varLength, tempAlpha)
            sumDist = Tan(tempAlpha) * dX + dY
            varLength = varLength + tempArc - sumDist
            For i = 0 To 100
                If Abs(sumDist - tempArc) < doublePrecision Then Exit For ' exit for if precision is achieved
                tempAlpha = varLength ^ 2 / m_clothConst
                dX = clothoidFresnelCosineDx(varLength, tempAlpha)
                dY = clothoidFresnelSineDy(varLength, tempAlpha)
                sumDist = Tan(tempAlpha) * dX + dY
                varLength = varLength + tempArc - sumDist
            Next i
            distFromArcClothoid = -m_curveDir * dX / Cos(tempAlpha)
        End If
        distFromStart = varLength
        getRadialDistFromPoint = True
        
    'point is on the center to start line or center to end line
    Case 0
        'point is on the center to start line
        If sideOfPntCenterToStart = 0 Then
            distFromStart = 0
            distFromArcClothoid = -m_curveDir * cogoDistance2D(m_startCoo.x, m_startCoo.y, px, py) * _
                                  cogoGetSide(m_startCoo.x, m_startCoo.y, m_endCoo.x, m_endCoo.y, px, py)
            getRadialDistFromPoint = True
         'point is on the center to end line
         Else
            distFromStart = m_len
            distFromArcClothoid = -m_curveDir * cogoDistance2D(m_endCoo.x, m_endCoo.y, px, py) * _
                                  cogoGetSide(m_startCoo.x, m_startCoo.y, m_endCoo.x, m_endCoo.y, px, py)
            getRadialDistFromPoint = True
        End If
    End Select
End Function

' startCoo property
Public Property Get startCoo() As Coo2D
    Set startCoo = m_startCoo
End Property


' endCoo property
Public Property Get endCoo() As Coo2D
    Set endCoo = m_endCoo
End Property

' Fresnel integral sine S(x)=integral[0,x]sin(y^2)dy aproximated by Taylor series expansion with first 7 elements
Private Function clothoidFresnelSineDy(arcLength As Double, alpha As Double) As Double
    clothoidFresnelSineDy = arcLength * (1 - alpha ^ 2 / 10 + alpha ^ 4 / 216 - alpha ^ 6 / 9360 + alpha ^ 8 / 685440 - alpha ^ 10 / 76204800 + alpha ^ 12 / 1197504000 / 10)
End Function

' Fresnel integral cosine C(x)=integral[0,x]cos(y^2)dy aproximated by Taylor series expansion with first 6 elements
Private Function clothoidFresnelCosineDx(arcLength As Double, alpha As Double) As Double
    clothoidFresnelCosineDx = arcLength * (alpha / 3 - alpha ^ 3 / 42 + alpha ^ 5 / 1320 - alpha ^ 7 / 75600 + alpha ^ 9 / 6894720 - alpha ^ 11 / 918086400)
End Function


