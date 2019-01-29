VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RedLines"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_redLines As Collection

'Constructor
Private Sub Class_Initialize()
    Set m_redLines = New Collection
End Sub

'add new Red Line
Public Function addRedLine(rl As RedLine) As Boolean
    On Error GoTo ErrHandler
    m_redLines.Add rl, rl.name
    addRedLine = True
    Exit Function
    
ErrHandler:
    addRedLine = False
End Function

'gets a Rl object by it's name
Public Function getRedLine(name As String) As RedLine
    On Error GoTo ErrHandler
    Set getRedLine = m_redLines.item(name)
    Exit Function
ErrHandler:
    Set getRedLine = Nothing
End Function

