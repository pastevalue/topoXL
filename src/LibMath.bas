Attribute VB_Name = "LibMath"
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

''========================================================================
'' Description:
'' Stores basic math functions
''========================================================================

'@Folder("TopoXL.libs")

Option Explicit
Option Private Module

' Compares to double values for equality
' Parameters:
'   - d1, d2: doubles to be compared
'   - epsilon: tollerance with default value of 1E-15
' Returns TRUE if the variance is less than tollerance (epsilon)
Function areDoublesEqual(ByVal d1 As Double, ByVal d2 As Double, _
                         Optional ByVal epsilon As Double = 0.000000000000001) As Boolean
    Dim absDiff As Double
    absDiff = Math.Abs(d1 - d2)
    areDoublesEqual = absDiff < epsilon
End Function

' Returns a rounded a number down to the nearest integer or to the nearest
' multiple of significance
' Parameters:
'   - n: the number to be rounded down
'   - f: fhe multiple to which you want to round
Public Function floor(ByVal n As Double, Optional ByVal f As Double = 1) As Double
    floor = Int(n / f) * f
End Function

' Returns a rounded a number up to the nearest integer or to the nearest
' multiple of significance
' Parameters:
'   - n: the number to be rounded up
'   - f: fhe multiple to which you want to round
Public Function ceiling(ByVal n As Double, Optional ByVal f As Double = 1) As Double
    ceiling = -Int(-n / f) * f
End Function

' Returns the maximum of two doubles
Public Function max(ByVal d1 As Double, ByVal d2 As Double) As Double
    max = IIf(d1 > d2, d1, d2)
End Function

' Returns the minimum of two values
Public Function min(ByVal d1 As Double, ByVal d2 As Double) As Double
    min = IIf(d1 < d2, d1, d2)
End Function

' Returns a random number between lower bound and upper bound
Public Function rndBetween(ByVal lowerbound As Double, ByVal upperbound As Double) As Double
    rndBetween = (upperbound - lowerbound) * Rnd + lowerbound
End Function


