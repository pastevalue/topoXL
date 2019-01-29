Attribute VB_Name = "registerFunctions"
Option Private Module
Sub registerUDF(UDFname As String, description As String, category As Integer)
    'Application.MacroOptions Macro:=UDFname, description:=description, category:=category
End Sub

Sub unregisterUDF(UDFname As String)
    'Application.MacroOptions Macro:=UDFname, description:=Empty, category:=Empty
End Sub

Sub registerUDFs()
   Call registerUDF("toAcadPoint", "Raporteaza punctele in autocad!", 9)
End Sub


Sub unRegisterUDFs()
    Call unregisterUDF("toAcadPoint")
End Sub









