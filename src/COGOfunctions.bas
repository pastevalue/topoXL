Attribute VB_Name = "COGOfunctions"
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
''    Modules: UDF_COGO
''    Classes: None
''Calls:
''    Modules: MathFunctions
''    Classes: None
''=======================================================
Option Private Module
Option Explicit

'Returns a collection which represents a polyline on 2 axis
'The returned polyline is trimmed from another polyline between a minimum and a maximum value on a chosen axis
'The initial polyline is assumed to be ordered on the axis chosen for trimming from smallest to largest
'Parameters:
'   -index: the index of the axis chosen for trimming. Must be 1 or 2 (X or Y)
'   -min: the minimum value for the trimming on the chosen axis
'   -max: the maximum value for the trimming on the chosen axis
'   -c: the collection containing the initial sorted polyline
Public Function trim2DPlineToCollection(index As Integer, min As Double, max As Double, c As Collection) As Collection
    Dim result As New Collection
    Dim i As Long
    Dim tempInterpolate As Double
          
    If c.item(index) < min Or c.item(c.count - 2 + index) > max Then
        For i = index To c.count - 4 + index Step 2
            If c.item(i) < min And c.item(i + 2) > min Then
                If index = 1 Then
                    tempInterpolate = interpolate2D(min, c.item(i), c.item(i + 1), c.item(i + 2), c.item(i + 3))
                    result.Add min
                    result.Add tempInterpolate
                Else
                    tempInterpolate = interpolate2D(min, c.item(i), c.item(i - 1), c.item(i + 2), c.item(i + 1))
                    result.Add tempInterpolate
                    result.Add min
                End If
                If c.item(i) < max And c.item(i + 2) > max Then
                    If index = 1 Then
                        result.Add c.item(i)
                        result.Add c.item(i + 1)
                        tempInterpolate = interpolate2D(max, c.item(i), c.item(i + 1), c.item(i + 2), c.item(i + 3))
                        result.Add max
                        result.Add tempInterpolate
                    Else
                        result.Add c.item(i - 1)
                        result.Add c.item(i)
                        tempInterpolate = interpolate2D(max, c.item(i), c.item(i - 1), c.item(i + 2), c.item(i + 1))
                        result.Add tempInterpolate
                        result.Add max
                    End If
                End If
            ElseIf c.item(i) < max And c.item(i + 2) > max Then
                If index = 1 Then
                    result.Add c.item(i)
                    result.Add c.item(i + 1)
                    tempInterpolate = interpolate2D(max, c.item(i), c.item(i + 1), c.item(i + 2), c.item(i + 3))
                    result.Add max
                    result.Add tempInterpolate
                Else
                    result.Add c.item(i - 1)
                    result.Add c.item(i)
                    tempInterpolate = interpolate2D(max, c.item(i), c.item(i - 1), c.item(i + 2), c.item(i + 1))
                    result.Add tempInterpolate
                    result.Add max
                End If
            ElseIf c.item(i) >= min And c.item(i) <= max Then
                If index = 1 Then
                    result.Add c.item(i)
                    result.Add c.item(i + 1)
                Else
                    result.Add c.item(i - 1)
                    result.Add c.item(i)
                End If
            End If
        Next i
        i = c.count - 2 + index
        If c.item(i) >= min And c.item(i) <= max Then
            If index = 1 Then
                result.Add c.item(i)
                result.Add c.item(i + 1)
            Else
                result.Add c.item(i - 1)
                result.Add c.item(i)
            End If
        End If
        Set trim2DPlineToCollection = result
        Exit Function
    End If
    Set trim2DPlineToCollection = c
End Function

'calculates the area between 3 or more points defined by abscissa and ordinate (cartesian coordinates)
'area formula is sum(Xi(Yi+1 - Yi-1)) divided by 2 for i=0->n where n is the number of points
'be sure the points provided are in the right order to avoid overlapping lines in the poligon
Public Function calculateEnclosedAreaFromCollection(c As Collection)
    Dim partialSum As Double
    Dim i As Integer
    Dim tempColl As Collection

    Set tempColl = c
    partialSum = 0
    For i = 3 To tempColl.count - 2 Step 2
        partialSum = partialSum + tempColl.item(i) * (tempColl.item(i + 3) - tempColl.item(i - 1))
    Next i
    
    'complete with first and last loop
    partialSum = partialSum + tempColl.item(1) * (tempColl.item(4) - tempColl.item(tempColl.count)) + tempColl.item(tempColl.count - 1) * (tempColl.item(2) - tempColl.item(tempColl.count - 2))
    
    'finalize the area formula by dividing by 2
    calculateEnclosedAreaFromCollection = partialSum / 2
End Function

