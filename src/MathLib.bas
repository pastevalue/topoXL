Attribute VB_Name = "MathLib"
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
'' Description:
'' Stores basic math functions
''=======================================================

'@Folder("TopoXL")
Option Explicit
Option Private Module

' Compares to double values for equality.
' Returns TRUE if the variance is less than tollerance (epsilon)
Function AreDoublesEqual(ByVal d1 As Double, ByVal d2 As Double, _
                         Optional epsilon As Double = 0.000000000000001) As Boolean
    Dim absDiff As Double
    absDiff = Math.Abs(d1 - d2)
    AreDoublesEqual = absDiff < epsilon
    
End Function

' Returns a rounded a number down to the nearest integer or to the nearest
' multiple of significance
' Parameters:
'   - n: the number to be rounded down
'   - f: fhe multiple to which you want to round
Public Function Floor(ByVal n As Double, Optional ByVal f As Double = 1) As Double
    Floor = Int(n / f) * f
End Function
