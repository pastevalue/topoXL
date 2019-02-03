VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IntvlColl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''=======================================================
''Called by:
''    Modules: IntervalCollectionsInit
''    Classes: IntvlColls
''Calls:
''    Modules: None
''    Classes: IntvlColl
''=======================================================
Option Explicit
Private m_name As String
Private m_intervals As Collection

'constructor
Private Sub Class_Initialize()
    m_name = vbNullString
    Set m_intervals = New Collection
End Sub

'add new interval to Interval Collection
Public Function addIntvl(intvl As Interval) As Boolean
    Dim i As Long
    Dim tempInterval As Interval
   
    'check if the new interval is connected to previous intervals and does not overlapp any of them
    If m_intervals.count > 0 Then
        For i = 1 To m_intervals.count
            Set tempInterval = m_intervals.item(i)
            If intvl.startVal = tempInterval.endVal Then
                If intvl.endVal > tempInterval.startVal Then
                    intvl.startConnectedIndex = i
                    tempInterval.endConnectedIndex = m_intervals.count + 1
                Else
                    GoTo failAdd
                End If
            ElseIf intvl.startVal < Me.getLastIntvl.endVal Then
                If intvl.endVal = tempInterval.startVal Then
                    intvl.endConnectedIndex = i
                    tempInterval.startConnectedIndex = m_intervals.count + 1
                ElseIf intvl.endVal > tempInterval.startVal Then
                    GoTo failAdd
                End If
            End If
        Next i
    End If
    'add interval
    m_intervals.Add intvl
    addIntvl = True
    Exit Function
failAdd:
    addIntvl = False
End Function

'returns a value corresponding the position specified from a collection of intervals
'Parameters:
'   -searchType: controls how the function behaves if the position is located at the exact start or end of an interval;
'       - -1: if the position is located exactly between two intervals then return the value from the first one (considering stored order);
'       - 0: ignores if the position is located at the start or end of an interval;
'            The function returns the value from the first interval found;
'       - 1: if the position is located exactly between two intervals then return the value from the second one (considering stored order).
'   -position: a value that the function uses to find the desired interval.
Public Function getValue(searchType As Integer, position As Double) As Variant
    Dim i As Long
    Dim j As Long
    Dim tempValue As String
    Dim tempInterval As Interval
    
    For i = 1 To m_intervals.count
        Set tempInterval = m_intervals.item(i)
        If position >= tempInterval.startVal And position <= tempInterval.endVal Then
            Select Case searchType
                Case -1
                    If position = tempInterval.startVal Then
                        If tempInterval.startConnectedIndex <> -1 Then
                            tempValue = m_intervals.item(tempInterval.startConnectedIndex).value
                        Else
                            tempValue = tempInterval.value
                        End If
                    Else
                        tempValue = tempInterval.value
                    End If
                Case 0
                    If position <> tempInterval.startVal And position <> tempInterval.endVal Then
                        tempValue = tempInterval.value
                    ElseIf position = tempInterval.startVal And tempInterval.startConnectedIndex = -1 Then
                        tempValue = tempInterval.value
                    ElseIf position = tempInterval.endVal And tempInterval.endConnectedIndex = -1 Then
                        tempValue = tempInterval.value
                    Else
                        tempValue = vbNullString
                    End If
                Case 1
                    If position = tempInterval.endVal Then
                        If tempInterval.endConnectedIndex <> -1 Then
                            tempValue = m_intervals.item(tempInterval.endConnectedIndex).value
                        Else
                            tempValue = tempInterval.value
                        End If
                    Else
                        tempValue = tempInterval.value
                    End If
                Case Else
                    GoTo failInput
            End Select
            Exit For
        End If
    Next i
    If Not Format(tempValue) = vbNullString Then
        getValue = tempValue
    Else
        getValue = CVErr(xlErrNA)
    End If
    Exit Function
failInput:
    getValue = CVErr(xlErrNum)
End Function

Public Function getIntvl(index As Integer) As Interval
    Set getIntvl = m_intervals.item(index)
End Function

Public Function getLastIntvl() As Interval
    Set getLastIntvl = m_intervals.item(m_intervals.count)
End Function

Public Function getIntervals() As Collection
    Set getIntervals = m_intervals
End Function

Public Function getIntvlCount() As Integer
    getIntvlCount = m_intervals.count
End Function

Public Sub clear()
    Dim i As Integer
    For i = 1 To m_intervals.count
        m_intervals.remove i
    Next i
End Sub

Public Property Get name() As String
    name = m_name
End Property

Public Property Let name(n As String)
    m_name = n
End Property







