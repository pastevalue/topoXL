VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Interval"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''=======================================================
''Called by:
''    Modules: IntervalCollectionsInit
''    Classes: IntvlColl
''Calls:
''    Modules: IC_ENUMS
''    Classes: None
''=======================================================
Option Explicit

Private m_startVal As Double
Private m_endVal As Double
Private m_value As String
Private m_startConnectedIndex As Integer
Private m_endConnectedIndex As Integer

'Class constructor
Private Sub Class_Initialize()
    m_startVal = 0
    m_endVal = 0
    m_value = vbNullString
    m_startConnectedIndex = -1 'index of interval to wich start of current is connected (-1 if is not connected)
    m_endConnectedIndex = -1 'index of interval to wich end of current is connected (-1 if is not connected)
End Sub

'Class initialization
Public Sub init(startVal As Double, endVal As Double, value As String)
    m_startVal = startVal
    m_endVal = endVal
    m_value = value
End Sub

'Class initialization from string
Public Function initFromStrings(startVal As String, endVal As String, value As String)
    On Error GoTo ErrHandler
    Me.init CDbl(startVal), CDbl(endVal), value
    initFromStrings = True
    Exit Function
ErrHandler:
    Debug.Print "The ends " & startVal & " and " & endVal & " couldn't be converted to Double type!"
End Function

Public Function initFromCollection(coll As Collection) As Boolean
    Dim tempVal As String
    Dim tempInputType As IC_INPUT_TYPE
    Dim startVal As Double
    Dim endVal As Double
    Dim value As String
    
    On Error GoTo FailInit

    tempInputType = IC_ENUMS.icIntvlInputTypeFromString(coll.item(IntervalCollectionsInit.TBL_INTVL_INPUT_TYPE_COL))
    
    Select Case tempInputType
        Case IC_INPUT_TYPE.INPUT_STV_ENV_VAL
            startVal = coll.item(IC_ENUMS.icIntvlPartToString(IC_INTVL_PARTS.PART_STV))
            endVal = coll.item(IC_ENUMS.icIntvlPartToString(IC_INTVL_PARTS.PART_ENV))
            value = coll.item(IC_ENUMS.icIntvlPartToString(IC_INTVL_PARTS.PART_VAL))
            
            'initialize class
            If startVal <= endVal Then
                Call init(startVal, endVal, value)
            Else
                Call init(endVal, startVal, value)
            End If
            initFromCollection = True
        Case Else
              GoTo FailInit
    End Select

Exit Function
FailInit:
    initFromCollection = False
End Function

'startVal property
Public Property Get startVal() As Double
    startVal = m_startVal
End Property
Public Property Let startVal(startVal As Double)
    m_startVal = startVal
End Property

'endVal property
Public Property Get endVal() As Double
    endVal = m_endVal
End Property
Public Property Let endVal(endVal As Double)
    m_endVal = endVal
End Property

'value property
Public Property Get value() As String
    value = m_value
End Property
Public Property Let value(value As String)
    m_value = value
End Property

'startIsConnected property
Public Property Get startConnectedIndex() As Integer
    startConnectedIndex = m_startConnectedIndex
End Property
Public Property Let startConnectedIndex(startConnectedIndex As Integer)
    m_startConnectedIndex = startConnectedIndex
End Property

'endIsConnected property
Public Property Get endConnectedIndex() As Integer
    endConnectedIndex = m_endConnectedIndex
End Property
Public Property Let endConnectedIndex(endConnectedIndex As Integer)
    m_endConnectedIndex = endConnectedIndex
End Property

