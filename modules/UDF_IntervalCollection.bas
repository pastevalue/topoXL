Attribute VB_Name = "UDF_IntervalCollection"
Option Explicit

Public Function icGetValue(icName As String, searchType As Integer, position As Double) As Variant
    Application.Volatile False
    Dim tempValue As Variant
    tempValue = INTVL_COLLS.getIntvlColl(icName).getValue(searchType, position)
    icGetValue = tempValue
End Function
