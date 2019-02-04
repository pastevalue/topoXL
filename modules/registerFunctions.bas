Attribute VB_Name = "registerFunctions"
''' TopoXL: Excel UDF library for land surveyors
''' Copyright (C) 2019 Bogdan Morosanu and Cristian Buse
''' This program is free software: you can redistribute it and/or modify
''' it under the terms of the GNU General Public License as published by
''' the Free Software Foundation, either version 3 of the License, or
''' (at your option) any later version.
'''
''' This program is distributed in the hope that it will be useful,
''' but WITHOUT ANY WARRANTY; without even the implied warranty of
''' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
''' GNU General Public License for more details.
'''
''' You should have received a copy of the GNU General Public License
''' along with this program.  If not, see <https://www.gnu.org/licenses/>.

''=======================================================
''Called by:
''    Modules: None
''    Classes: None
''Calls:
''    Modules: None
''    Classes: None
''=======================================================
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









