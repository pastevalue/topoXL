Attribute VB_Name = "UDF_COGO"
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
''    Modules: UDF_Math
''    Classes: ArcCircle, ArcClothoid, Line2D, RLarcCircle
''Calls:
''    Modules: COGOfunctions, MathFunctions, rangeFunctions
''    Classes: None
''=======================================================
Option Explicit

'get distance between two sets of 2D coordinates
Public Function cogoDistance2D(x1 As Double, y1 As Double, x2 As Double, y2 As Double) As Double
    cogoDistance2D = Sqr((x1 - x2) ^ 2 + (y1 - y2) ^ 2)
End Function

'get distance between two sets of 3D coordinates
Public Function cogoDistance3D(x1 As Double, y1 As Double, z1 As Double, x2 As Double, y2 As Double, z2 As Double) As Double
    cogoDistance3D = Sqr((x1 - x2) ^ 2 + (y1 - y2) ^ 2 + (z1 - z2) ^ 2)
End Function

'Gets the side of a set of coordinates relative to a line
'Returns -1 if coordinates are on the left, 0 if the coordinates are on the line and +1 if coordinates are on the right
'Left and right are considered relative to the start and end coordinates of the line
Public Function cogoGetSide(startX As Double, startY As Double, endX As Double, endY As Double, px As Double, py As Double) As Integer
    Application.Volatile False
    startX = Round(startX, 8)
    startY = Round(startY, 8)
    endX = Round(endX, 8)
    endY = Round(endY, 8)
    px = Round(px, 8)
    py = Round(py, 8)
    cogoGetSide = Sgn((endY - startY) * (px - startX) - (endX - startX) * (py - startY))
End Function

'get azimuth between two sets of 2D coordinates
Public Function cogoAzimuth(x1 As Double, y1 As Double, x2 As Double, y2 As Double) As Variant
    Application.Volatile False
    Dim dX As Double
    Dim dY As Double
    
    If x1 = x2 And y1 = y2 Then GoTo failInput
    dX = x2 - x1
    dY = y2 - y1
    Select Case dY
        Case Is > 0
            cogoAzimuth = Atn(dX / dY)
        Case Is < 0
            cogoAzimuth = Atn(dX / dY) + PI * Sgn(dX)
            If dX = 0 Then cogoAzimuth = cogoAzimuth + PI
        Case Is = 0
            cogoAzimuth = PI / 2 * Sgn(dX)
    End Select
Exit Function
failInput:
    cogoAzimuth = CVErr(xlErrNum)
End Function

'get the intersection of 2 lines which are defined with start point and azimuth
'the function returns TRUE if an intersection is found and FALSE if not
'if TRUE then the function also returns by reference both X and Y coordinates of intersection
Public Function cogoIntersectionOf2LinesByCooAndAzimuth(ByVal x1 As Double, ByVal y1 As Double, ByVal azimuth1 As Double, ByVal x2 As Double, ByVal y2 As Double, ByVal azimuth2 As Double, _
                                        ByRef xInt As Double, ByRef yInt As Double) As Boolean
    Application.Volatile False
    
    Dim tg1 As Double
    Dim tg2 As Double
    
    tg1 = Tan(azimuth1)
    tg2 = Tan(azimuth2)
    If tg1 <> tg2 Then
        xInt = (x1 * tg2 - x2 * tg1 - tg1 * tg2 * (y1 - y2)) / (tg2 - tg1)
        yInt = (xInt - x1) / tg1 + y1
        cogoIntersectionOf2LinesByCooAndAzimuth = True
     Else
        xInt = 0
        yInt = 0
        cogoIntersectionOf2LinesByCooAndAzimuth = False
    End If
End Function

Public Function cogoIntersectionOf2LinesByCooAndAzimuthArr(ByVal x1 As Double, ByVal y1 As Double, ByVal azimuth1 As Double, ByVal x2 As Double, ByVal y2 As Double, ByVal azimuth2 As Double) As Variant()
    Application.Volatile False
    
    Dim xInt As Double, yInt As Double
    Dim arr(0 To 1) As Variant
    
    Dim tg1 As Double
    Dim tg2 As Double
    
    tg1 = Tan(azimuth1)
    tg2 = Tan(azimuth2)
    If tg1 <> tg2 Then
        xInt = (x1 * tg2 - x2 * tg1 - tg1 * tg2 * (y1 - y2)) / (tg2 - tg1)
        yInt = (xInt - x1) / tg1 + y1
     Else
        xInt = CVErr(xlErrNA)
        yInt = CVErr(xlErrNA)
    End If
    arr(0) = xInt
    arr(1) = yInt
    cogoIntersectionOf2LinesByCooAndAzimuthArr = arr
End Function



'get the intersection of 2 lines which are defined with start point and offsets
'offsets are considered relative to the start point of the line
'if intersection is found the function returns an array with both X and Y coordinates of intersection
'if intersection not found the function returns the appropriate error
Public Function cogoIntersectionOf2LinesByCooAndOffset(ByVal x1 As Double, ByVal y1 As Double, ByVal offsetX1 As Double, ByVal offsetY1 As Double, _
                                                            ByVal x2 As Double, ByVal y2 As Double, ByVal offsetX2 As Double, ByVal offsetY2 As Double) As Variant
    Application.Volatile False
    Dim result(1 To 1, 1) As Double
    
    If offsetX1 = 0 And offsetX2 = 0 Then
        cogoIntersectionOf2LinesByCooAndOffset = CVErr(xlErrNum)
    ElseIf offsetX1 = 0 Then
        result(1, 0) = x1
        result(1, 1) = y2 + offsetY2 / offsetX2 * (x1 - x2)
        cogoIntersectionOf2LinesByCooAndOffset = result
    ElseIf offsetX2 = 0 Then
        result(1, 0) = x2
        result(1, 1) = y1 + offsetY1 / offsetX1 * (x2 - x1)
        cogoIntersectionOf2LinesByCooAndOffset = result
    Else
        Dim p1 As Double
        Dim p2 As Double
            
        p1 = offsetY1 / offsetX1
        p2 = offsetY2 / offsetX2
        If p1 = p2 Then
            cogoIntersectionOf2LinesByCooAndOffset = CVErr(xlErrValue)
        Else
            result(1, 1) = ((x1 - x2) * p1 * p2 - y1 * p2 + y2 * p1) / (p1 - p2)
            If p1 = 0 Then
                result(1, 0) = (result(1, 1) - y2) / p2 + x2
            Else
                result(1, 0) = (result(1, 1) - y1) / p1 + x1
            End If
            cogoIntersectionOf2LinesByCooAndOffset = result
        End If
    End If
End Function

'calculates the area between 3 or more points defined by abscissa and ordinate (cartesian coordinates)
'area formula is module divided by 2 of: sum(Xi(Yi+1 - Yi-1)) for i=0->n where n is the number of points
'be sure the points provided are in the right order to avoid overlapping lines in the poligon
Public Function cogoCalculateEnclosedArea(ParamArray values() As Variant)
    Application.Volatile False
    Dim t As Variant
    Dim i As Integer
    Dim tempColl As Collection
    
    t = values
    Set tempColl = rangeFunctions.valuesToCollection(t)
    
    'IF the number of cells in ranges is not divisible by 2 or the number of points is less then 3 THEN return error
    If tempColl.count Mod 2 <> 0 Or tempColl.count < 6 Then GoTo failInput
    
    cogoCalculateEnclosedArea = Abs(calculateEnclosedAreaFromCollection(tempColl))
    
Exit Function
failInput:
    cogoCalculateEnclosedArea = CVErr(xlErrNum)
End Function

Public Function cogoSegmentIntersection(x1 As Double, y1 As Double, x2 As Double, y2 As Double, _
                                        x3 As Double, y3 As Double, x4 As Double, y4 As Double) As Variant
    Application.Volatile False
    Dim result(1 To 1, 1) As Double
    If segmentIntersection(x1, y1, x2, y2, x3, y3, x4, y4, result(1, 0), result(1, 1)) Then
        cogoSegmentIntersection = result
     Else
        cogoSegmentIntersection = CVErr(xlErrNA)
    End If
End Function

Private Function segmentIntersection(x1 As Double, y1 As Double, x2 As Double, y2 As Double, _
                                     x3 As Double, y3 As Double, x4 As Double, y4 As Double, _
                                     ByRef outX As Double, ByRef outY As Double) As Boolean
    Application.Volatile False
    Dim dx1 As Double, dy1 As Double
    Dim dx2 As Double, dy2 As Double
    segmentIntersection = False
                                                                              
    dx1 = x2 - x1
    dy1 = y2 - y1
    dx2 = x4 - x3
    dy2 = y4 - y3
    Dim temp As Double
    temp = dx1 * dy2 - dy1 * dx2
    If temp = 0 Then GoTo NoIntersection
    
    Dim dX As Double, dY As Double
    Dim t As Double
    
    dX = x3 - x1
    dY = y3 - y1
    t = (dX * dy2 - dY * dx2) / temp
    If t < 0 Or t > 1 Then GoTo NoIntersection

    Dim u As Double
    u = (dX * dy1 - dY * dx1) / temp
    If u < 0 Or u > 1 Then GoTo NoIntersection
    
    outX = x1 + t * dx1
    outY = y1 + t * dy1
    segmentIntersection = True

    Exit Function
NoIntersection:
    segmentIntersection = False
    Exit Function
End Function
'Returns coordinates of extended/trimed segment line at specified length.
'Parameters:
'   -values: values that represent the coordinates of the segment that will be extended (start X, start Y, end X, end Y).
'   -side: specifies on which side of segment line to make the extension
'       * -1 before segment line;
'       * 0 before and after segment line;
'       * 1 after segment line;
'   -length: lenght of the extension. Extends or trims of value is positive or negative
'Use as array formula to get all coordinates. Result contains the values for coordinates of the new line segment (start X, start Y, end X, end Y).
Public Function cogoExtendTrimSegment(side As Integer, length As Double, ParamArray values() As Variant) As Variant
    Application.Volatile False
    Dim tempColl As Collection
    Dim t As Variant
    
    t = values
    Set tempColl = rangeFunctions.valuesToCollection(t)
    'return error if there aren't 4 values
    If tempColl.count <> 4 Then GoTo failInput
    
    Const i_startX As Integer = 1 'start X index in values parameter
    Const i_startY As Integer = 2 'start Y index in values parameter
    Const i_endX As Integer = 3 'end X index in values parameter
    Const i_endY As Integer = 4 'end Y index in values parameter
    
    'select extend side
    Dim result(1 To 1, 1 To 4) As Double
    Dim dX As Double
    Dim dY As Double
    Dim D As Double
    
    dX = tempColl(i_endX) - tempColl(i_startX)
    dY = tempColl(i_endY) - tempColl(i_startY)
    D = Sqr(dX ^ 2 + dY ^ 2)
    
    Select Case side
        'before start of line segment
        Case -1
            result(1, 1) = tempColl(i_startX) - length * dX / D
            result(1, 2) = tempColl(i_startY) - length * dY / D
            result(1, 3) = tempColl(i_endX)
            result(1, 4) = tempColl(i_endY)
        ' before and after line segment
        Case 0
            result(1, 1) = tempColl(i_startX) - length * dX / D
            result(1, 2) = tempColl(i_startY) - length * dY / D
            result(1, 3) = tempColl(i_endX) + length * dX / D
            result(1, 4) = tempColl(i_endY) + length * dY / D
        'after end of line segment
        Case 1
            result(1, 1) = tempColl(i_startX)
            result(1, 2) = tempColl(i_startY)
            result(1, 3) = tempColl(i_endX) + length * dX / D
            result(1, 4) = tempColl(i_endY) + length * dY / D
        Case Else
            GoTo failInput
    End Select
    cogoExtendTrimSegment = result
    Exit Function
failInput:
    cogoExtendTrimSegment = CVErr(xlErrNum)
    Exit Function
End Function


'Returns coordinates of a offseted segment line with specified value.
'Parameters:
'   -values: values that represent the coordinates of the segment that will be extended (start X, start Y, end X, end Y).
'   -offset: offset distance. Negative value will offset on the left and positive value will offset on the right
'Use as array formula to get all coordinates
'Returns #DIV/0! error if:
'   - start and end point of segment are identical
Public Function cogoOffsetSegment(offset As Double, ParamArray values() As Variant) As Variant
    Application.Volatile False
    Dim tempColl As Collection
    Dim t As Variant
    
    t = values
    Set tempColl = rangeFunctions.valuesToCollection(t)
    'return error if there aren't 4 values
    If tempColl.count <> 4 Then GoTo failInput
    
    Const i_startX As Integer = 1 'start X index in values parameter
    Const i_startY As Integer = 2 'start Y index in values parameter
    Const i_endX As Integer = 3 'end X index in values parameter
    Const i_endY As Integer = 4 'end Y index in values parameter
    
    Dim result(1 To 1, 1 To 4) As Double
    Dim dX As Double
    Dim dY As Double
    Dim D As Double 'segment lenght

    
    
    dX = tempColl(i_endX) - tempColl(i_startX)
    dY = tempColl(i_endY) - tempColl(i_startY)
    D = Sqr(dX ^ 2 + dY ^ 2)
    
    'return error if start and end point of segment are identical
    If D = 0 Then
        cogoOffsetSegment = CVErr(xlErrDiv0)
        Exit Function
    End If
    
    result(1, 1) = tempColl(i_startX) + offset * dY / D
    result(1, 2) = tempColl(i_startY) - offset * dX / D
    result(1, 3) = tempColl(i_endX) + offset * dY / D
    result(1, 4) = tempColl(i_endY) - offset * dX / D
    cogoOffsetSegment = result
    Exit Function
failInput:
    cogoOffsetSegment = CVErr(xlErrNum)
    Exit Function
End Function

Public Function cogoPLineIntersection(line1 As Variant, line2 As Variant) As Variant
    Application.Volatile False
    Dim line1Arr() As Variant
    Dim line2Arr() As Variant
    Dim tempX As Double
    Dim tempY As Double
    Dim outCooColl As New Collection
       
    line1Arr = rangeFunctions.valuesTo2Darray(2, line1)
    line2Arr = rangeFunctions.valuesTo2Darray(2, line2)

    Dim i As Integer, j As Integer
    Dim tempX1 As Double, tempY1 As Double, tempX2 As Double, tempY2 As Double
    Dim tempX3 As Double, tempY3 As Double, tempX4 As Double, tempY4 As Double
    
    
    For i = 0 To UBound(line1Arr) - 1
        On Error GoTo failInput
        tempX1 = line1Arr(i, 0)
        tempY1 = line1Arr(i, 1)
        tempX2 = line1Arr(i + 1, 0)
        tempY2 = line1Arr(i + 1, 1)
        For j = 0 To UBound(line2Arr) - 1
            tempX3 = line2Arr(j, 0)
            tempY3 = line2Arr(j, 1)
            tempX4 = line2Arr(j + 1, 0)
            tempY4 = line2Arr(j + 1, 1)
            On Error GoTo 0
            If segmentIntersection(tempX1, tempY1, tempX2, tempY2, tempX3, tempY3, tempX4, tempY4, tempX, tempY) Then
                outCooColl.Add tempX
                outCooColl.Add tempY
            End If
        Next j
    Next i
    
    If outCooColl.count >= 2 Then
        cogoPLineIntersection = rangeFunctions.collectionTo2Darray(2, outCooColl)
     Else
        cogoPLineIntersection = CVErr(xlErrNA)
    End If
    Exit Function
failInput:
    cogoPLineIntersection = CVErr(xlErrNum)
    Exit Function
End Function

'returns the cut and fill surfaces between two coplanar (base and comparision) polylines
Public Function cogoCutAndFill2D(basePline As Variant, comparisionPline As Variant) As Variant
    Dim tempBaseColl As New Collection
    Dim tempCompColl As New Collection
    Dim tempIntColl As New Collection
    Dim tempBaseArr() As Variant
    Dim tempCompArr() As Variant
    Dim tempCutFill(0 To 0, 0 To 1) As Variant
    Dim partialSum As Double
    Dim i As Long, j As Long
    Const precision = 0.000001
    
    'check if the plines have at least 2 points and the coordinates are in pairs of 2
    Set tempBaseColl = rangeFunctions.valuesToCollection(basePline)
    If tempBaseColl.count < 4 Or tempBaseColl.count Mod 2 <> 0 Then GoTo failInput
    
    Set tempCompColl = rangeFunctions.valuesToCollection(comparisionPline)
    If tempCompColl.count < 4 Or tempCompColl.count Mod 2 <> 0 Then GoTo failInput
    
    'check if the vertexes of plines are ordered on X axis and if so then if the order is right to left reverse the pline into left to right
    If Not areGroupedValuesSorted(2, 2, 1, tempBaseColl) Then
        If areGroupedValuesSorted(-2, 2, 1, tempBaseColl) Then
            Set tempBaseColl = reverseCollection(2, tempBaseColl)
        Else
            'values are not sorted
            GoTo failInput
        End If
    End If
    If Not areGroupedValuesSorted(2, 2, 1, tempCompColl) Then
        If areGroupedValuesSorted(-2, 2, 1, tempCompColl) Then
            Set tempCompColl = reverseCollection(2, tempCompColl)
        Else
            'values are not sorted
            GoTo failInput
        End If
    End If
    
    'establish the boundaries on X axis for the cutFill calculation
    Dim leftX As Double, rightX As Double
    If tempBaseColl.item(1) > tempCompColl.item(1) Then
        leftX = tempBaseColl.item(1)
    Else
        leftX = tempCompColl.item(1)
    End If
    If tempBaseColl.item(tempBaseColl.count - 1) < tempCompColl.item(tempCompColl.count - 1) Then
        rightX = tempBaseColl.item(tempBaseColl.count - 1)
    Else
        rightX = tempCompColl.item(tempCompColl.count - 1)
    End If
    
    'check if the plines do not overlapp
    If tempBaseColl.item(tempBaseColl.count - 1) <= tempCompColl.item(1) Or tempCompColl.item(tempCompColl.count - 1) <= tempBaseColl.item(1) Then GoTo failOverlapp
    
    'remove the unnedeed points from plines
    Set tempBaseColl = trim2DPlineToCollection(1, leftX, rightX, tempBaseColl)
    Set tempCompColl = trim2DPlineToCollection(1, leftX, rightX, tempCompColl)
    
    tempBaseArr = rangeFunctions.collectionTo2Darray(2, tempBaseColl)
    tempCompArr = rangeFunctions.collectionTo2Darray(2, tempCompColl)
    
    'find all the intersection points between the 2 plines
    Dim tempBaseX1 As Double, tempBaseY1 As Double, tempBaseX2 As Double, tempBaseY2 As Double
    Dim tempCompX1 As Double, tempCompY1 As Double, tempCompX2 As Double, tempCompY2 As Double
    Dim tempIntX As Double, tempIntY As Double
    
    For i = 0 To UBound(tempBaseArr) - 1
        On Error GoTo failInput
        tempBaseX1 = CDbl(tempBaseArr(i, 0))
        tempBaseY1 = CDbl(tempBaseArr(i, 1))
        tempBaseX2 = CDbl(tempBaseArr(i + 1, 0))
        tempBaseY2 = CDbl(tempBaseArr(i + 1, 1))
        For j = 0 To UBound(tempCompArr) - 1
            tempCompX1 = CDbl(tempCompArr(j, 0))
            tempCompY1 = CDbl(tempCompArr(j, 1))
            tempCompX2 = CDbl(tempCompArr(j + 1, 0))
            tempCompY2 = CDbl(tempCompArr(j + 1, 1))
            On Error GoTo 0
            If segmentIntersection(tempBaseX1, tempBaseY1, tempBaseX2, tempBaseY2, tempCompX1, tempCompY1, tempCompX2, tempCompY2, tempIntX, tempIntY) Then
                If tempIntX - precision > leftX And tempIntX + precision < rightX Then
                    tempIntColl.Add tempIntX
                    tempIntColl.Add tempIntY
                    tempIntColl.Add i
                    tempIntColl.Add j
                End If
            End If
        Next j
    Next i
    
    tempCutFill(0, 0) = 0
    tempCutFill(0, 1) = 0
    Dim tempColl As New Collection
    'if points of intersection between base and comparision plines are found
    If tempIntColl.count >= 4 Then
        Dim tempIntArr() As Variant
        tempIntArr = rangeFunctions.collectionTo2Darray(4, tempIntColl)
        
        'area until first intersection point
        For i = 0 To tempIntArr(0, 2)
            tempColl.Add tempBaseArr(i, 0)
            tempColl.Add tempBaseArr(i, 1)
        Next i
        tempColl.Add tempIntArr(0, 0)
        tempColl.Add tempIntArr(0, 1)
        For i = tempIntArr(0, 3) To 0 Step -1
            tempColl.Add tempCompArr(i, 0)
            tempColl.Add tempCompArr(i, 1)
        Next i
        partialSum = calculateEnclosedAreaFromCollection(tempColl)
        If partialSum < 0 Then
            tempCutFill(0, 0) = tempCutFill(0, 0) + partialSum
        Else
            tempCutFill(0, 1) = tempCutFill(0, 1) + partialSum
        End If
        'area between 2 consecutive intersection points
        If UBound(tempIntArr) >= 1 Then
            For i = 0 To UBound(tempIntArr) - 1
                Set tempColl = New Collection
                tempColl.Add tempIntArr(i, 0)
                tempColl.Add tempIntArr(i, 1)
                If tempIntArr(i + 1, 2) <> tempIntArr(i, 2) Then
                    For j = tempIntArr(i, 2) + 1 To tempIntArr(i + 1, 2)
                        tempColl.Add tempBaseArr(j, 0)
                        tempColl.Add tempBaseArr(j, 1)
                    Next j
                End If
                tempColl.Add tempIntArr(i + 1, 0)
                tempColl.Add tempIntArr(i + 1, 1)
                If tempIntArr(i + 1, 3) <> tempIntArr(i, 3) Then
                    For j = tempIntArr(i + 1, 3) To tempIntArr(i, 3) + 1 Step -1
                        tempColl.Add tempCompArr(j, 0)
                        tempColl.Add tempCompArr(j, 1)
                    Next j
                End If
                partialSum = calculateEnclosedAreaFromCollection(tempColl)
                If partialSum < 0 Then
                    tempCutFill(0, 0) = tempCutFill(0, 0) + partialSum
                Else
                    tempCutFill(0, 1) = tempCutFill(0, 1) + partialSum
                End If
            Next i
        End If
        'area after last intersection point
        Set tempColl = New Collection
        tempColl.Add tempIntArr(UBound(tempIntArr), 0)
        tempColl.Add tempIntArr(UBound(tempIntArr), 1)
        For i = tempIntArr(UBound(tempIntArr), 2) + 1 To UBound(tempBaseArr)
            tempColl.Add tempBaseArr(i, 0)
            tempColl.Add tempBaseArr(i, 1)
        Next i
        For i = UBound(tempCompArr) To tempIntArr(UBound(tempIntArr), 3) + 1 Step -1
            tempColl.Add tempCompArr(i, 0)
            tempColl.Add tempCompArr(i, 1)
        Next i
        partialSum = calculateEnclosedAreaFromCollection(tempColl)
        If partialSum < 0 Then
            tempCutFill(0, 0) = tempCutFill(0, 0) + partialSum
        Else
            tempCutFill(0, 1) = tempCutFill(0, 1) + partialSum
        End If
    Else 'no intersection points are found
        For i = 0 To UBound(tempBaseArr)
            tempColl.Add tempBaseArr(i, 0)
            tempColl.Add tempBaseArr(i, 1)
        Next i
        For i = UBound(tempCompArr) To 0 Step -1
            tempColl.Add tempCompArr(i, 0)
            tempColl.Add tempCompArr(i, 1)
        Next i
        partialSum = calculateEnclosedAreaFromCollection(tempColl)
        If partialSum < 0 Then
            tempCutFill(0, 0) = partialSum
        Else
            tempCutFill(0, 1) = partialSum
        End If
    End If
           
    cogoCutAndFill2D = tempCutFill
    
Exit Function
failInput:
    cogoCutAndFill2D = CVErr(xlErrNum)
    Exit Function
failOverlapp:
    cogoCutAndFill2D = CVErr(xlErrNA)
End Function

'Returns an array which represents a polyline on 2 axis
'The returned polyline is trimmed from another polyline between a minimum and a maximum value on a chosen axis
'The initial polyline is assumed to be ordered on the axis chosen for trimming from smallest to largest
'Parameters:
'   -index: the index of the axis chosen for trimming. Must be 1 or 2 (X or Y)
'   -min: the minimum value for the trimming on the chosen axis
'   -max: the maximum value for the trimming on the chosen axis
'   -c: the collection containing the initial sorted polyline
Public Function cogoTrimPline(index As Integer, min As Double, max As Double, ParamArray values() As Variant) As Variant
    Dim tempColl As Collection
    Dim tempArray() As Variant
    Dim i As Long
    Dim t As Variant
    
    'IF the index of axis is not 1 or 2 (X or Y) then return error
    If index < 1 Or index > 2 Then GoTo failInput
    'IF min is greater or equal to max then return error
    If min >= max Then GoTo failInput
    t = values
    Set tempColl = rangeFunctions.valuesToCollection(t)
    
    'IF the number of values is not divisible by number of coordinates THEN return error
    If tempColl.count Mod 2 <> 0 Then GoTo failInput
    
    If Not areGroupedValuesSorted(2, 2, index, tempColl) Then
        If areGroupedValuesSorted(-2, 2, index, tempColl) Then
            Set tempColl = reverseCollection(2, tempColl)
        Else
            'values are not sorted
            GoTo failInput
        End If
    End If
    
    Set tempColl = trim2DPlineToCollection(index, min, max, tempColl)
    cogoTrimPline = rangeFunctions.collectionTo2Darray(1, tempColl)
    
Exit Function
failInput:
    cogoTrimPline = CVErr(xlErrNum)
End Function

'returns a single polyline from two coplanar polylines ("base" and "paste")
'to the "paste" polyline are added the segments of the "base" polyline which do not overlapp with the "paste" polyline
'the 2 polylines are assumed to have ordered vertexes on the X axis
Public Function cogoPastePline(basePline As Variant, pastePline As Variant) As Variant
    Dim tempBaseColl As New Collection
    Dim tempPasteColl As New Collection
    Dim tempResultColl As New Collection
    Dim i As Long, j As Long
    
    'check if the plines have at least 2 points and the coordinates are in pairs of 2
    Set tempBaseColl = rangeFunctions.valuesToCollection(basePline)
    If tempBaseColl.count < 4 Or tempBaseColl.count Mod 2 <> 0 Then GoTo failInput
    
    Set tempPasteColl = rangeFunctions.valuesToCollection(pastePline)
    If tempPasteColl.count < 4 Or tempPasteColl.count Mod 2 <> 0 Then GoTo failInput
    
    'check if the vertexes of plines are ordered on X axis and if so then if the order is right to left reverse the pline into left to right
    If Not areGroupedValuesSorted(2, 2, 1, tempBaseColl) Then
        If areGroupedValuesSorted(-2, 2, 1, tempBaseColl) Then
            Set tempBaseColl = reverseCollection(2, tempBaseColl)
        Else
            'values are not sorted
            GoTo failInput
        End If
    End If
    If Not areGroupedValuesSorted(2, 2, 1, tempPasteColl) Then
        If areGroupedValuesSorted(-2, 2, 1, tempPasteColl) Then
            Set tempPasteColl = reverseCollection(2, tempPasteColl)
        Else
            'values are not sorted
            GoTo failInput
        End If
    End If
    
    'create the plines (from "basePline") situated on left and right of "pastePline" and add all items to result
    Dim leftOfPasteColl As New Collection
    Dim rightOfPasteColl As New Collection
    
    If tempBaseColl.item(1) < tempPasteColl.item(1) Then
        Set leftOfPasteColl = trim2DPlineToCollection(1, tempBaseColl.item(1), tempPasteColl.item(1), tempBaseColl)
        For i = 1 To leftOfPasteColl.count - 2
            tempResultColl.Add leftOfPasteColl.item(i)
        Next i
        If leftOfPasteColl.item(leftOfPasteColl.count) <> tempPasteColl.item(2) Then
            tempResultColl.Add leftOfPasteColl.item(leftOfPasteColl.count - 1)
            tempResultColl.Add leftOfPasteColl.item(leftOfPasteColl.count)
        End If
    End If
    For i = 1 To tempPasteColl.count
        tempResultColl.Add tempPasteColl.item(i)
    Next i
    If tempPasteColl.item(tempPasteColl.count - 1) < tempBaseColl.item(tempBaseColl.count - 1) Then
        Set rightOfPasteColl = trim2DPlineToCollection(1, tempPasteColl.item(tempPasteColl.count - 1), tempBaseColl.item(tempBaseColl.count - 1), tempBaseColl)
        If rightOfPasteColl.item(2) <> tempPasteColl.item(tempPasteColl.count) Then
            tempResultColl.Add rightOfPasteColl.item(1)
            tempResultColl.Add rightOfPasteColl.item(2)
        End If
        For i = 3 To rightOfPasteColl.count
            tempResultColl.Add rightOfPasteColl.item(i)
        Next i
    End If
    
    cogoPastePline = rangeFunctions.collectionTo2Darray(1, tempResultColl)
    
Exit Function
failInput:
    cogoPastePline = CVErr(xlErrNum)
 End Function
 
'For a set of coordinates returns the distance to the line and the distance from perpendicular coordinates to the start of the line
Public Function cogoGetPerpDistanceFromPoint(side As Integer, startCooX As Double, startCooY As Double, endCooX As Double, endCooY As Double, pCooX As Double, pCooY As Double) As Variant
    Dim lineSlope As Double
    Dim perpPointX As Double
    Dim perpPointY As Double
    Dim result(1) As Variant
    Dim distFromLine As Double
    Dim distFromPerpToLine As Double
    
    If side <> -1 And side <> 1 Then GoTo failInput
    If startCooX <> endCooX Then
        If startCooY <> endCooY Then
            lineSlope = (endCooY - startCooY) / (endCooX - startCooX)
            perpPointX = endCooX + (pCooY - endCooY + 1 / lineSlope * (pCooX - endCooX)) / (lineSlope + 1 / lineSlope)
            perpPointY = lineSlope * (perpPointX - endCooX) + endCooY
            distFromLine = cogoDistance2D(pCooX, pCooY, perpPointX, perpPointY)
            Select Case side
                Case -1
                    distFromPerpToLine = cogoDistance2D(startCooX, startCooY, perpPointX, perpPointY)
                Case 1
                    distFromPerpToLine = cogoDistance2D(endCooX, endCooY, perpPointX, perpPointY)
            End Select
            If cogoPointIsInsideBoundBox(startCooX, startCooY, endCooX, endCooY, perpPointX, perpPointY) Then distFromPerpToLine = -distFromPerpToLine
        Else
            distFromLine = Abs(pCooY - startCooY)
            Select Case side
                Case -1
                    distFromPerpToLine = Abs(pCooX - startCooX)
                Case 1
                    distFromPerpToLine = Abs(pCooX - endCooX)
            End Select
            If cogoPointIsInsideBoundBox(startCooX, startCooY, endCooX, endCooY, pCooX, startCooY) Then distFromPerpToLine = -distFromPerpToLine
        End If
    Else
        If startCooY = endCooY Then
            GoTo failInput
        Else
            distFromLine = Abs(pCooX - startCooX)
            Select Case side
                Case -1
                    distFromPerpToLine = Abs(pCooY - startCooY)
                Case 1
                    distFromPerpToLine = Abs(pCooY - endCooY)
            End Select
            If cogoPointIsInsideBoundBox(startCooX, startCooY, endCooX, endCooY, startCooX, pCooY) Then distFromPerpToLine = -distFromPerpToLine
        End If
    End If
    
    distFromLine = distFromLine * cogoGetSide(startCooX, startCooY, endCooX, endCooY, pCooX, pCooY)
    result(0) = distFromLine
    result(1) = distFromPerpToLine
    cogoGetPerpDistanceFromPoint = result
Exit Function
failInput:
    Debug.Print "Start & end points must be 2 distinct points"
    cogoGetPerpDistanceFromPoint = CVErr(xlErrNum)
End Function

'Return TRUE if a point is inside the bounding box of a line
Public Function cogoPointIsInsideBoundBox(startCooX As Double, startCooY As Double, endCooX As Double, endCooY As Double, pCooX As Double, pCooY As Double) As Variant
    Dim validX As Boolean
    Dim validY As Boolean
    
    If startCooX = endCooX And startCooY = endCooY Then GoTo failInput
    
    validX = (pCooX >= startCooX And pCooX <= endCooX) Or (pCooX >= endCooX And pCooX <= startCooX)
    validY = (pCooY >= startCooY And pCooY <= endCooY) Or (pCooY <= startCooY And pCooY >= endCooY)
    
    cogoPointIsInsideBoundBox = validX And validY
Exit Function
failInput:
    Debug.Print "Start & end points must be 2 distinct points"
    cogoPointIsInsideBoundBox = CVErr(xlErrNum)
End Function

