Attribute VB_Name = "UDF_Analyze"
''=======================================================
''Called by:
''    Modules: None
''    Classes: None
''Calls:
''    Modules: None
''    Classes: None
''=======================================================
Option Explicit

Public Function anlHasFormula(ParamArray ranges() As Variant) As Variant
    Dim r As Variant
    Dim c As Variant
    Dim temp As Range
    anlHasFormula = False
    For Each r In ranges
        For Each c In r
            If Not (c.HasFormula) Then
                anlHasFormula = False
                Exit Function
            End If
        Next c
    Next r
    anlHasFormula = True
End Function

