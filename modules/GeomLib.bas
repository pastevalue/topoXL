Attribute VB_Name = "GeomLib"
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
''Description
''Stores basic geometry functions
''=======================================================

'@Folder("TopoXL.geom")
Option Explicit

'Returns the distance between two sets of 2D grid coordinates
Public Function Dist2D(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double) As Double
    Dist2D = Sqr((X1 - X2) ^ 2 + (Y1 - Y2) ^ 2)
End Function


