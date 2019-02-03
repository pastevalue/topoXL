Attribute VB_Name = "MathFunctions"
''=======================================================
''Called by:
''    Modules: COGOfunctions, UDF_Math, UDF_COGO
''    Classes: None
''Calls:
''    Modules: None
''    Classes: None
''=======================================================
Option Private Module
Option Explicit

Public Function interpolate2D(x As Double, x1 As Double, y1 As Double, x2 As Double, y2 As Double) As Double
    interpolate2D = (x - x1) / (x2 - x1) * (y2 - y1) + y1
End Function

'Returns TRUE or FALSE if specified values are sorted or not
'Parameters:
'   -order: a number that indicates how to make the comparasion between values:
'       - -2 for descending order with  equal values accepted;
'       - -1 for descending order with equal values not accepted;
'       - 1 for ascending order with equal values not accepted;
'       - 2 for ascending order with equal values accepted.
'   -groupSize: a integer which indicates the step that the values are compared. For example if groupSize=3 then only values with indexes 1, 4, 7, ... will be compared
'   -groupIndex: a integer which indicates the starting position for comparision. Must be smaller or equal to groupSize.
'               For example if groupSize is 2 and groupIndex is 2 then only values 2, 4, 6, ... will be compared.
'   -c: a collection that contains the value(s) that will be used to be checked if they are sorted.
Public Function areGroupedValuesSorted(order As Integer, groupSize As Integer, groupIndex As Integer, c As Collection) As Boolean
    Dim i As Long
    'in case needed the precison of comparision can be set
    
    'select by order parameter
    Select Case order
        'check for descending order with equal values accepted
        Case -2
            For i = groupIndex To c.count - 2 * groupSize + groupIndex Step groupSize
                If c(i) < c(i + groupSize) Then
                    areGroupedValuesSorted = False
                    Exit Function
                End If
            Next i
        'check for descending order
        Case -1
            For i = groupIndex To c.count - 2 * groupSize + groupIndex Step groupSize
                If c(i) <= c(i + groupSize) Then
                    areGroupedValuesSorted = False
                    Exit Function
                End If
            Next i
        'check for ascending order
        Case 1
            For i = groupIndex To c.count - 2 * groupSize + groupIndex Step groupSize
                If c(i) >= c(i + groupSize) Then
                    areGroupedValuesSorted = False
                    Exit Function
                End If
            Next i
        'check for ascending order with equal values accepted
        Case 2
            For i = groupIndex To c.count - 2 * groupSize + groupIndex Step groupSize
                If c(i) > c(i + groupSize) Then
                    areGroupedValuesSorted = False
                    Exit Function
                End If
            Next i
    End Select
    areGroupedValuesSorted = True
End Function
