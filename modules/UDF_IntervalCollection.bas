Attribute VB_Name = "UDF_IntervalCollection"
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
''    Modules: IntervalCollectionsInit
''    Classes: None
''=======================================================
Option Explicit

Public Function icGetValue(icName As String, searchType As Integer, position As Double) As Variant
    Application.Volatile False
    Dim tempValue As Variant
    tempValue = INTVL_COLLS.getIntvlColl(icName).getValue(searchType, position)
    icGetValue = tempValue
End Function
