Attribute VB_Name = "UDF_RedLine"
''=======================================================
''Called by:
''    Modules: None
''    Classes: None
''Calls:
''    Modules: None
''    Classes: Coo2D
''=======================================================
Option Explicit

Public Function rlGetCooArray(rlName As String, station As Double, offset As Double) As Variant
    Application.Volatile False
    Dim tempCoo As Coo2D
    Set tempCoo = RED_LINES.getRedLine(rlName).getCoo(station, offset)
    If tempCoo Is Nothing Then
        rlGetCooArray = CVErr(xlErrNum)
     Else
        rlGetCooArray = tempCoo.toArray
    End If
End Function
    
Public Function rlGetStAndOffsetArray(rlName As String, px As Double, py As Double) As Variant
    Application.Volatile False
    Dim result(1 To 1, 1) As Double
    If RED_LINES.getRedLine(rlName).getPerpOrRadFromPoint(px, py, result(1, 0), result(1, 1)) Then
        rlGetStAndOffsetArray = result
     Else
        rlGetStAndOffsetArray = CVErr(xlErrNum)
    End If
End Function

Public Function rlGetX(rlName As String, y As Double) As Double
    Application.Volatile False
    rlGetX = RED_LINES.getRedLine(rlName).getX(y)
    If IsNull(rlGetX) Then rlGetX = CVErr(xlErrNum)
End Function

Public Function rlGetY(rlName As String, x As Double) As Double
    Application.Volatile False
    rlGetY = RED_LINES.getRedLine(rlName).getY(x)
    If IsNull(rlGetY) Then rlGetY = CVErr(xlErrNum)
End Function

Public Function rlGetLength(rlName As String) As Variant
    Application.Volatile False
    rlGetLength = RED_LINES.getRedLine(rlName).getLength
End Function

Public Function rlGetFirstElemStartX(rlName As String) As Variant
    Application.Volatile False
    rlGetFirstElemStartX = RED_LINES.getRedLine(rlName).getFirstElement.startX
End Function

Public Function rlGetFirstElemStartY(rlName As String) As Variant
    Application.Volatile False
    rlGetFirstElemStartY = RED_LINES.getRedLine(rlName).getFirstElement.startY
End Function

Public Function rlGetLastElemEndX(rlName As String) As Variant
    Application.Volatile False
    rlGetLastElemEndX = RED_LINES.getRedLine(rlName).getLastElement.endX
End Function

Public Function rlGetLastElemEndY(rlName As String) As Variant
    Application.Volatile False
    rlGetLastElemEndY = RED_LINES.getRedLine(rlName).getLastElement.endY
End Function




