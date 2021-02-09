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
                                 Optional ByVal epsilon As Double = 0.000000000000001) As Integer
    If x1 = x2 And y1 = y2 Then
        Err.Raise 5, MODULE_NAME & ".orientationIndex", "Can't compute orientation index using a reference line defined by the same coordinates!"
    End If
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
                                    Optional ByVal epsilon As Double = 0.000000000000001) As Point
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
Public Function intLbyThAndCoo(ByVal x1 As Double, ByVal y1 As Double, _
                               ByVal theta1 As Double, _
                               ByVal x2 As Double, ByVal y2 As Double, _
                               ByVal theta2 As Double, _
                               Optional ByVal epsilon As Double = 0.000000000000001) As Point
    ' Check if lines are parallel
    If LibMath.areDoublesEqual(theta1, theta2, epsilon) Then Exit Function
    Set intLbyThAndCoo = New Point
    
    If LibMath.areDoublesEqual(theta1, 0, epsilon) _
        Or LibMath.areDoublesEqual(theta1, LibGeom.PI, epsilon) Then
        intLbyThAndCoo.y = y1
        intLbyThAndCoo.x = (intLbyThAndCoo.y - y2) / Math.Tan(theta2) + x2
    ElseIf LibMath.areDoublesEqual(theta2, 0, epsilon) _
        Or LibMath.areDoublesEqual(theta2, LibGeom.PI, epsilon) Then
        intLbyThAndCoo.y = y2
        intLbyThAndCoo.x = (intLbyThAndCoo.y - y1) / Math.Tan(theta1) + x1
    Else
        ' Compute Cotangents
        Dim ctg1 As Variant: ctg1 = 1 / Math.Tan(theta1)
        Dim ctg2 As Variant: ctg2 = 1 / Math.Tan(theta2)

        intLbyThAndCoo.x = (x1 * ctg2 - x2 * ctg1 - ctg1 * ctg2 * (y1 - y2)) / (ctg2 - ctg1)
        intLbyThAndCoo.y = (intLbyThAndCoo.x - x1) / ctg1 + y1
    End If
End Function

' Rreturns the intersection of two line segments defined by start and end coordinates
' Parameters:
'   - x1, y1, x2, y2: coordinates of the first line segment
'   - x3, y3, x4, y5: coordinates of the second line segment
' Returns:
'   - The intersection point if exists
'   - Nothing if line segments don't intersect
'   - Nothing if line segments are identical
Public Function intLSbyCoo(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double, _
                           ByVal x3 As Double, ByVal y3 As Double, ByVal x4 As Double, ByVal y4 As Double) As Point
    Dim dx1 As Double
    Dim dy1 As Double
    Dim dx2 As Double
    Dim dy2 As Double
    Dim det As Double   ' determinant
    
    dx1 = x2 - x1
    dy1 = y2 - y1
    dx2 = x4 - x3
    dy2 = y4 - y3
    det = dx1 * dy2 - dy1 * dx2
        
    If det = 0 Then Exit Function   ' Return Nothing - line segments are parallel
        
    Dim dX As Double
    Dim dY As Double
    Dim t As Double
    
    dX = x3 - x1
    dY = y3 - y1
    t = (dX * dy2 - dY * dx2) / det
    If t < 0 Or t > 1 Then Exit Function    ' Return Nothing - no intersection

    Dim u As Double
    u = (dX * dy1 - dY * dx1) / det
    If u < 0 Or u > 1 Then Exit Function    ' Return Nothing - no intersection
    
    Set intLSbyCoo = FactoryGeom.newPnt(x1 + t * dx1, y1 + t * dy1)
End Function

' Rreturns the intersection of two lines defined by a set of coordinates and offsets (coordinate deltas)
' Parameters:
'   - x1, y1, dx1, dy1: coordinates and offsets/deltas of the first line
'   - x2, y2, dx2, dy2: coordinates and offsets/deltas of the second line
' Returns:
'   - The intersection point if exists
'   - Nothing if lines don't intersect
'   - Nothing if lines are identical
Public Function intLbyCooAndDs(ByVal x1 As Double, ByVal y1 As Double, ByVal dx1 As Double, ByVal dy1 As Double, _
                               ByVal x2 As Double, ByVal y2 As Double, ByVal dx2 As Double, ByVal dy2 As Double) As Point
                                
    If dx1 = 0 And dx2 = 0 Then Exit Function    ' Return Nothing - parallel vertical lines
    
    If dx1 = 0 Then
        Set intLbyCooAndDs = FactoryGeom.newPnt(x1, y2 + dy2 / dx2 * (x1 - x2))
    ElseIf dx2 = 0 Then
        Set intLbyCooAndDs = FactoryGeom.newPnt(x2, y1 + dy1 / dx1 * (x2 - x1))
    Else
        Dim p1 As Double
        Dim p2 As Double
            
        p1 = dy1 / dx1
        p2 = dy2 / dx2
        If p1 = p2 Then
            Exit Function                        ' Return Nothing - no intersection
        Else
            Dim resX As Double
            Dim resY As Double
            resY = ((x1 - x2) * p1 * p2 - y1 * p2 + y2 * p1) / (p1 - p2)
            If p1 = 0 Then
                resX = (resY - y2) / p2 + x2
            Else
                resX = (resY - y1) / p1 + x1
            End If
            Set intLbyCooAndDs = FactoryGeom.newPnt(resX, resY)
        End If
    End If
End Function

' Returns coordinates of intersections between the 2D polylines (PLs) defined by the two
' input coordinate sets
Public Function intPLbyCoo(ByVal coo2DArr1 As Variant, ByVal coo2DArr2 As Variant) As PointColl



    Dim cooDim1 As Long
    Dim cooDim2 As Long
    Dim j1LBound As Long                         '2nd dim lower bound of coosArr1
    Dim j2LBound As Long                         '2nd dim lower bound of coosArr1
    j1LBound = LBound(coo2DArr1, 2)
    j2LBound = LBound(coo2DArr2, 2)
    cooDim1 = UBound(coo2DArr1, 2) - j1LBound + 1
    cooDim2 = UBound(coo2DArr2, 2) - j2LBound + 1
    If cooDim1 <> 2 Or cooDim2 <> 2 Then
        Err.Raise 5, MODULE_NAME, "Second dimension of input arrays must be of size 2. XY coordinates expected!"
    Else
        Dim i As Long
        Dim j As Long
        Dim p As Point
        
        If UBound(coo2DArr1, 1) - LBound(coo2DArr1, 1) + 1 < 2 Or UBound(coo2DArr2, 1) - LBound(coo2DArr2, 1) + 1 < 2 Then
            Err.Raise 5, MODULE_NAME, "1st dimension of input arrays must be of size greater than 2. At least two sets of coordinates expected!"
        End If
        Set intPLbyCoo = New PointColl
        For i = LBound(coo2DArr1, 1) To UBound(coo2DArr1, 1) - 1
            For j = LBound(coo2DArr2, 1) To UBound(coo2DArr2, 1) - 1
                Set p = LibGeom.intLSbyCoo(coo2DArr1(i, j1LBound), coo2DArr1(i, j1LBound + 1), coo2DArr1(i + 1, j1LBound), coo2DArr1(i + 1, j1LBound + 1), _
                                           coo2DArr2(j, j2LBound), coo2DArr2(j, j2LBound + 1), coo2DArr2(j + 1, j2LBound), coo2DArr2(j + 1, j2LBound + 1))
                If Not p Is Nothing Then
                    intPLbyCoo.add p
                End If
            Next j
        Next i
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

' Returns coordinates of extended/trimed line segment
' Parameters:
'   - x1, y1, x2, y2: coordinates of the line segment that will be extended/trimed (x1,y2 - start; x2,y2 - end)
'   - length: length of extend/trim:
'       * > 0 = Extend
'       * < 0 = Trim
'   - part: specifies the part (start or end) on which to extend/trim
'       * -1 start of line segment
'       * 0 start and end of line segment
'       * 1 end of line segment
' Raises error if:
'   - part parameter is not -1, 0 or 1
'   - input line segment is zero length (start = end)
Public Function extTrimLS(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double, _
                          ByVal length As Double, ByVal part As Integer) As PointColl

    If part < -1 Or part > 1 Then
        Err.Raise 5, MODULE_NAME, "Part value must be -1, 0 or 1!"
    End If
    
    Dim dX As Double
    Dim dY As Double
    Dim d As Double
    Dim sp As Point ' Result start point
    Dim ep As Point ' Result end point
    
    dX = x2 - x1
    dY = y2 - y1
    d = Math.Sqr(dX * dX + dY * dY)
    If d = 0 Then
        Err.Raise 5, MODULE_NAME, "Line segment must be different from zero length line segment!"
    End If
        
    Select Case part
        Case -1
            Set sp = FactoryGeom.newPnt(x1 - length * dX / d, y1 - length * dY / d)
            Set ep = FactoryGeom.newPnt(x2, y2)
        Case 0
            Set sp = FactoryGeom.newPnt(x1 - length * dX / d, y1 - length * dY / d)
            Set ep = FactoryGeom.newPnt(x2 + length * dX / d, y2 + length * dY / d)
        Case 1
            Set sp = FactoryGeom.newPnt(x1, y1)
            Set ep = FactoryGeom.newPnt(x2 + length * dX / d, y2 + length * dY / d)
    End Select
    Set extTrimLS = FactoryGeom.newPntColl(sp, ep)
End Function

' Returns coordinates of offseted line segment
' Parameters:
'   - x1, y1, x2, y2: coordinates of the line segment that will be offseted (x1,y2 - start; x2,y2 - end)
'   - offset: offset distance.
'       * > 0 = offset on the left
'       * < 0 = offset on the right
' Raises error if:
'   - input line segment is zero length (start = end)
Public Function offsetLS(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double, _
                         ByVal offset As Double) As PointColl
    Dim dX As Double
    Dim dY As Double
    Dim d As Double
    Dim offX As Double
    Dim offY As Double
    Dim sp As Point                              ' Result start point
    Dim ep As Point                              ' Result end point
    
    dX = x2 - x1
    dY = y2 - y1
    d = Math.Sqr(dX * dX + dY * dY)
    If d = 0 Then
        Err.Raise 5, MODULE_NAME, "Line segment must be different from zero length line segment!"
    End If
    offX = offset * dX / d
    offY = offset * dY / d
    Set sp = FactoryGeom.newPnt(x1 + offY, y1 - offX)
    Set ep = FactoryGeom.newPnt(x2 + offY, y2 - offX)
    Set offsetLS = FactoryGeom.newPntColl(sp, ep)
End Function

' Returns the area defined by the input coordinates
' Analytical fomrula used: SUM(Xi(Yi+1 - Yi-1)) / 2, for i = 0 -> n, n = pairs of coordinates
' Formula fails if coordinates generate a self intersecting polygon
' Parameters:
'   - coosArr: 2D array of coordinates
' Returns error if:
'   - coosArr 2nd dimension size is different from 2 (XY coos are expected)
'   - coosArr 1st dimension size is less than 3 (minimum 3 points required)
Public Function areaByCoo(ByVal coo2Darr As Variant) As Double
    Dim cooDim As Long
    Dim jLBound As Long '2nd dim lower bound
    jLBound = LBound(coo2Darr, 2)
    cooDim = UBound(coo2Darr, 2) - jLBound + 1
    If cooDim <> 2 Then
        Err.Raise 5, MODULE_NAME, "Second dimension of input array must be of size 2. XY coordinates expected!"
    Else
        Dim pSum As Double ' partial/cumulative sum
        Dim Lidx As Long
        Dim Uidx As Long
        Dim i As Long
        
        Lidx = LBound(coo2Darr, 1)
        Uidx = UBound(coo2Darr, 1)
        If Uidx - Lidx + 1 < 3 Then
            Err.Raise 5, MODULE_NAME, "1st dimension of input array must be of size greater than 3. At least 3 sets of coordinates expected!"
        End If
        pSum = 0
        For i = Lidx To Uidx
            If i = Lidx Then
                pSum = pSum + coo2Darr(i, jLBound) * (coo2Darr(i + 1, jLBound + 1) - coo2Darr(Uidx, jLBound + 1))
            ElseIf i = Uidx Then
                pSum = pSum + coo2Darr(i, jLBound) * (coo2Darr(Lidx, jLBound + 1) - coo2Darr(i - 1, jLBound + 1))
            Else
                pSum = pSum + coo2Darr(i, jLBound) * (coo2Darr(i + 1, jLBound + 1) - coo2Darr(i - 1, jLBound + 1))
            End If
        Next i
    End If
    areaByCoo = Math.Abs(pSum / 2)
End Function

