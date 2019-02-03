Attribute VB_Name = "UDF_Math"
''=======================================================
''Called by:
''    Modules: None
''    Classes: ArcCircle
''Calls:
''    Modules: UDF_COGO, MathFunctions
''    Classes: None
''=======================================================
Option Explicit
Public Const PI As Double = 3.14159265358979
'Returns the arc sine of a number.
'Value parameter must be in [-1,1] interval.
Public Function mathASin(value As Double) As Double
    value = Round(value, 10)
    If Abs(value) = 1 Then
        mathASin = value * PI / 2
     Else
        mathASin = Atn(value / Sqr(1 - value * value))
    End If
End Function

'Returns the arc cosine of a number.
'Value parameter must be in [-1,1] interval.
Public Function mathACos(value As Double) As Double
    Application.Volatile False
    mathACos = PI / 2 - mathASin(value)
End Function

'Returns the length of an arc circle if provided with chord length and radius
'Chord length and radius must be positive values and chord length must be less or equal to radius
Public Function mathGetArcCircleLengthFromChord(chordDist As Double, radius As Double) As Variant
    If chordDist > 2 * radius Or chordDist <= 0 Then GoTo failInput
    mathGetArcCircleLengthFromChord = 2 * radius * mathASin(chordDist / (2 * radius))
    Exit Function
failInput:
    mathGetArcCircleLengthFromChord = CVErr(xlErrNum)
    Debug.Print "Chord length and radius must be positive values and chord length must be less or equal to radius"
End Function

'Returns TRUE or FALSE if specified values are sorted or not
'Parameters:
'   -order: a number that indicates how to make the comparasion between values:
'       - -2 for descending order with  equal values accepted;
'       - -1 for descending order with equal values not accepted;
'       - 1 for ascending order with  equal values accepted;
'       - 2 for ascending order with equal values not accepted.
'   -values: a variant that contains the value(s) that will be used to be checked if they are sorted.
'Returns #N/A error if:
'   - order parameter is not -2,-1, 1 or 2;
'   - values parameter has less than 2 elements.
Public Function mathAreValuesSorted(order As Integer, ParamArray values() As Variant) As Variant
    Application.Volatile False
    Dim t As Variant
    Dim tempColl As Collection
    Dim i As Long
    
    t = values
    Set tempColl = rangeFunctions.valuesToCollection(t)
    
    'fail if there are less than two values to check
    If tempColl.count < 2 Then GoTo failInput
    
     mathAreValuesSorted = areGroupedValuesSorted(order, 1, 1, tempColl)

    Exit Function
failInput:
    mathAreValuesSorted = CVErr(xlErrNum)
    Exit Function
End Function


'TODO - de coafat functia impreuna cu Cristi.

'Returns linear interpolated value of Y for supplied X, based on supplied ranges of known X and Y values.
'Parameters:
'   -axisOrder: specifies the order in wich X and Y values will be parsed. 1 is for XY and -1 is for YX;
'   -interpolationType: controls how function responds to X arguments outside the known X values
'       - 0: returns error;
'       - 1: extrapolates based on last two X-Y pairs (either two highest or two lowest);
'       - 2: extrapolates based on first and last X-Y pair (full range of supplied values).
'   -xValue: X values for wich the interpolated Y is calculated;
'   -xyKnownValues: values of known X's and Y's.
'X and Y values can be arrays of values that will be parsed in their storage order, Excel Ranges that will be parsed in wise order (lines then column) or values.
'If values total number is not divisible by 2 or is smaller than 4 then an error will occur.
Public Function mathInterpLinear(axisOrder As Integer, interpolationType As Integer, xValue As Double, ParamArray xyKnownValues() As Variant) As Variant
    Dim tempColl As Collection
    Dim valArray() As Variant
    Dim t As Variant
    Dim cellCount As Long
    Dim xIndex As Integer
    Dim yIndex As Integer
    Dim n As Integer
         
    t = xyKnownValues
    Set tempColl = rangeFunctions.valuesToCollection(t)
    cellCount = tempColl.count
    If cellCount Mod 2 <> 0 Or cellCount < 4 Then GoTo failInput
    
    'get index for X and Y
    Select Case axisOrder
        Case 1
            xIndex = 0
            yIndex = 1
        Case -1
            xIndex = 1
            yIndex = 0
        Case Else
            GoTo failInput
    End Select
    valArray = rangeFunctions.collectionTo2Darray(2, tempColl)
  
    Dim beforeX As Double, afterX As Double
    Dim beforeXfound As Boolean, afterXfound As Boolean
    Dim beforeIndex As Integer, afterIndex As Integer
    Dim i As Integer
    Dim Dict As Object
    Dim val As Double
    Dim sameXfound As Boolean
    
    'create the scripting dictionary to count the X values provided that are unique
    Set Dict = CreateObject("Scripting.Dictionary")
    beforeX = xValue
    afterX = xValue
    beforeXfound = False
    afterXfound = False
    sameXfound = False
    beforeIndex = -1
    afterIndex = -1
    
    'go through all the elements and check for the closest before and after values of X
    n = cellCount / 2
    For i = 0 To n - 1
        val = valArray(i, xIndex)
        Dict(val) = 1
        If val < xValue Then
            If beforeXfound Then
                If val > beforeX Then
                    beforeX = val
                    beforeIndex = i
                End If
            Else
                beforeXfound = True
                beforeX = val
                beforeIndex = i
            End If
        ElseIf val > xValue Then
            If afterXfound Then
                If val < afterX Then
                    afterX = val
                    afterIndex = i
                End If
            Else
                afterXfound = True
                afterX = val
                afterIndex = i
            End If
        Else
            'if the wanted X already exists then just return the corresponding Y
            mathInterpLinear = valArray(i, yIndex)
            sameXfound = True
        End If
    Next i
    
    'check if there are duplicate values of X and return error if true
    'if needed the values of the scripting dictionary are returned with .keys
    'if all the X values are unique check if the wanted Y was already found because the X already existed
    If Dict.count < n Then
        GoTo failInput
    ElseIf sameXfound Then
        Exit Function
    End If
    
    Dim beforeY As Double
    Dim afterY As Double
    
    
    If beforeXfound And afterXfound Then
        'if for the wanted X there are both before and after values then just interpolate between those
        beforeY = valArray(beforeIndex, yIndex)
        afterY = valArray(afterIndex, yIndex)
        mathInterpLinear = interpolate2D(xValue, beforeX, beforeY, afterX, afterY)
    Else
        'if either before or after X was not found then treat the interpolation based on the interpolationType argument
        Dim maxBeforeOfBefore As Double, minAfterOfAfter As Double
        Dim maxBeforeOfBeforeFound As Boolean, minAfterOfAfterFound As Boolean
        Dim maxBeforeOfBeforeIndex As Integer, minAfterOfAfterIndex As Integer
        Select Case interpolationType
            Case 0
                mathInterpLinear = CVErr(xlErrNA)
                Exit Function
            Case 1
                If beforeXfound Then
                    maxBeforeOfBefore = beforeX
                    maxBeforeOfBeforeFound = False
                    maxBeforeOfBeforeIndex = -1
                    For i = 0 To n - 1
                        val = valArray(i, xIndex)
                        If val < beforeX Then
                            If maxBeforeOfBeforeFound Then
                                If val > maxBeforeOfBefore Then
                                    maxBeforeOfBefore = val
                                    maxBeforeOfBeforeIndex = i
                                End If
                            Else
                                maxBeforeOfBeforeFound = True
                                maxBeforeOfBefore = val
                                maxBeforeOfBeforeIndex = i
                            End If
                        End If
                    Next i
                    beforeY = valArray(maxBeforeOfBeforeIndex, yIndex)
                    afterY = valArray(beforeIndex, yIndex)
                    mathInterpLinear = interpolate2D(xValue, maxBeforeOfBefore, beforeY, beforeX, afterY)
                Else
                    minAfterOfAfter = afterX
                    minAfterOfAfterFound = False
                    minAfterOfAfterIndex = -1
                    For i = 0 To n - 1
                        val = valArray(i, xIndex)
                        If val > afterX Then
                            If minAfterOfAfterFound Then
                                If val < minAfterOfAfter Then
                                    minAfterOfAfter = val
                                    minAfterOfAfterIndex = i
                                End If
                            Else
                                minAfterOfAfterFound = True
                                minAfterOfAfter = val
                                minAfterOfAfterIndex = i
                            End If
                        End If
                    Next i
                    beforeY = valArray(afterIndex, yIndex)
                    afterY = valArray(minAfterOfAfterIndex, yIndex)
                    mathInterpLinear = interpolate2D(xValue, afterX, beforeY, minAfterOfAfter, afterY)
                End If
            Case 2
                Dim min As Double, max As Double
                Dim minIndex As Integer, maxIndex As Integer
                min = valArray(0, xIndex)
                max = valArray(0, xIndex)
                minIndex = 0
                maxIndex = 0
                For i = 1 To n - 1
                        val = valArray(i, xIndex)
                        If val < min Then
                            min = val
                            minIndex = i
                        ElseIf val > max Then
                            max = val
                            maxIndex = i
                        End If
                Next i
                beforeY = valArray(minIndex, yIndex)
                afterY = valArray(maxIndex, yIndex)
                mathInterpLinear = interpolate2D(xValue, min, beforeY, max, afterY)
            Case Else
                GoTo failInput
        End Select
    End If
    
    Exit Function

failInput:
    mathInterpLinear = CVErr(xlErrNum)
    Exit Function
End Function

'Returns linear interpolated value of Y for supplied X, based on supplied known multiple intervals of X and Y values.
'Intervals are parsed in specified order. Interpolated value is given for the first interval where X value fits.
'Parameters:
'   -axisOrder: specifies the order in wich X and Y values will be parsed. 1 is for XY and -1 is for YX;
'   -xValue: X values for wich the interpolated Y is calculated;
'   -intervals: intervals with known X's and Y's.
'X and Y values can be arrays of values that will be parsed in their storage order, Excel Ranges that will be parsed in wise order (lines then column) or values.
'Returns #N/A! error if:
'   - X value is not included in any of the specified intervals
'Returns #Num! error if:
'   - invalid parameters are suplied
Public Function mathInterpLinearMultiIntervals(axisOrder As Integer, xValue As Double, ParamArray intervals() As Variant) As Variant
    Dim tempColl As Collection
    Dim t As Variant
    Dim tempResult As Variant
    Dim Interval As Variant
    
    For Each Interval In intervals
        t = Interval
        Set tempColl = rangeFunctions.valuesToCollection(t)
        tempResult = mathInterpLinear(axisOrder, 0, xValue, Interval)
        If IsError(tempResult) Then
            If tempResult = CVErr(xlErrNum) Then GoTo failInput
        End If
        If IsNumeric(tempResult) Then
            mathInterpLinearMultiIntervals = tempResult
            Exit Function
        End If
    Next Interval
    'return #N/A! If xValue is not in any interval
    mathInterpLinearMultiIntervals = CVErr(xlErrNA)
    Exit Function
failInput:
    mathInterpLinearMultiIntervals = CVErr(xlErrNum)
End Function

'returns a 2 element array which represents the coordinates of the center of circumscribed circle of a triangle
'the parameters are the 2D coordinates of the triangle vertexes
Public Function mathCenterOfCircumscribedCircle(x1 As Double, y1 As Double, x2 As Double, y2 As Double, x3 As Double, y3 As Double) As Variant
    Dim result(0 To 1) As Double
    Dim coefficient As Double
    
    coefficient = 2 * (x1 * (y2 - y3) + x2 * (y3 - y1) + x3 * (y1 - y2))
    If coefficient = 0 Then GoTo failInput
    result(0) = ((x1 ^ 2 + y1 ^ 2) * (y2 - y3) + (x2 ^ 2 + y2 ^ 2) * (y3 - y1) + (x3 ^ 2 + y3 ^ 2) * (y1 - y2)) / coefficient
    result(1) = ((x1 ^ 2 + y1 ^ 2) * (x3 - x2) + (x2 ^ 2 + y2 ^ 2) * (x1 - x3) + (x3 ^ 2 + y3 ^ 2) * (x2 - x1)) / coefficient
    mathCenterOfCircumscribedCircle = result
    Exit Function
failInput:
    mathCenterOfCircumscribedCircle = CVErr(xlErrNum)
End Function

'returns a 2 element array which represents the values in radians of the 2 angles between 2 coplanar segments
'the parameters are the 2D coordinates of the vertexes of the 2 lines
Public Function mathAngleBetweenLines(xL1P1 As Double, yL1P1 As Double, xL1P2 As Double, yL1P2 As Double, xL2P1 As Double, yL2P1 As Double, xL2P2 As Double, yL2P2 As Double) As Variant
    Dim result(0 To 1) As Double
    Dim orientation1 As Double
    Dim orientation2 As Double
    
    If xL1P1 = xL1P2 And yL1P1 = yL1P2 Then GoTo failInput
    If xL2P1 = xL2P2 And yL2P1 = yL2P2 Then GoTo failInput
    
    orientation1 = cogoAzimuth(xL1P1, yL1P1, xL1P2, yL1P2)
    orientation2 = cogoAzimuth(xL2P1, yL2P1, xL2P2, yL2P2)
    
    result(0) = Abs(orientation1 - orientation2)
    If result(0) > PI Then result(0) = result(0) - PI
    If result(0) <> 0 Then
        result(1) = PI - result(0)
    Else
        result(1) = 0
    End If
    mathAngleBetweenLines = result
Exit Function
failInput:
    mathAngleBetweenLines = CVErr(xlErrNum)
End Function

