VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Coo2D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''=======================================================
''Called by:
''    Modules: UDF_RedLine
''    Classes: Line2D, ArcCircle, ArcClothoid, RLelement,
''             RLarcCircle, RLarcClothoid,RedLine, RLline
''Calls:
''    Modules: None
''    Classes: None
''=======================================================
Option Explicit

Private m_X As Double 'Abscissa
Private m_Y As Double 'Ordinate

'Class constructor
Private Sub Class_Initialize()
    m_X = 0
    m_Y = 0
End Sub

'Class initialization
Public Sub init(x As Double, y As Double)
    m_X = x
    m_Y = y
End Sub

'Class initialization from string
Public Function initFromStrings(x As String, y As String)
    On Error GoTo ErrHandler
    Me.init CDbl(x), CDbl(y)
    initFromStrings = True
    Exit Function
ErrHandler:
    Debug.Print "Coordinates " & x & " and " & y & " couldn't be converted to Double type!"
End Function

Public Function toArray() As Variant
    Dim result(1 To 1, 1) As Double
    result(1, 0) = m_X
    result(1, 1) = m_Y
    toArray = result
End Function

Public Function toText() As String
    If Me Is Nothing Then
        toText = ""
    Else
        toText = "X=" & m_X & " " & "Y=" & m_Y
    End If
End Function

'X property
Public Property Get x() As Double
    x = m_X
End Property
Public Property Let x(x As Double)
    m_X = x
End Property

'Y property
Public Property Get y() As Double
    y = m_Y
End Property
Public Property Let y(y As Double)
    m_Y = y
End Property


