VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IntvlColls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_intvlColls As Collection

'Constructor
Private Sub Class_Initialize()
    Set m_intvlColls = New Collection
End Sub

'add new Red Line
Public Function addIntvlColl(ic As IntvlColl) As Boolean
    On Error GoTo ErrHandler
    m_intvlColls.Add ic, ic.name
    addIntvlColl = True
    Exit Function
    
ErrHandler:
    addIntvlColl = False
End Function

'gets a IC object by it's name
Public Function getIntvlColl(name As String) As IntvlColl
    On Error GoTo ErrHandler
    Set getIntvlColl = m_intvlColls.item(name)
    Exit Function
ErrHandler:
    Set getIntvlColl = Nothing
End Function


