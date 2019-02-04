Attribute VB_Name = "UDF_Analyze"
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

