Attribute VB_Name = "LibArr"
''' TopoXL: Excel UDF library for land surveyors
''' Copyright (C) 2021 Bogdan Morosanu and Cristian Buse
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

''========================================================================
'' Description:
'' Stores array functions
''========================================================================

'@Folder("TopoXL.libs")
Option Explicit

' Function copied from https://github.com/cristianbuse/VBA-ArrayTools/blob/master/Code%20Modules/LibArrayTools.bas
Public Function GetArrayDimsCount(ByRef arr As Variant) As Long
    'In Visual Basic, you can declare arrays with up to 60 dimensions
    Const MAX_DIMENSION As Long = 60
    Dim dimension As Long
    Dim tempBound As Long
    '
    'A zero-length array has 1 dimension! Ex. Array() returns (0 to -1)
    '
    'Check the lower (or the upper) bound while increasing the dimension in a
    '   loop until an error occurs (when the dimension checked is invalid)
    On Error GoTo FinalDimension
    For dimension = 1 To MAX_DIMENSION
        tempBound = LBound(arr, dimension)
    Next dimension
Exit Function 'Good practice but not needed. Code will never reach this line
FinalDimension:
    GetArrayDimsCount = dimension - 1
End Function
