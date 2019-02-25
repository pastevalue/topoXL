Attribute VB_Name = "ArrayFunctions"
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
''    Modules: rangeFunctions, UDF_DataManipulation
''    Classes: None
''Calls:
''    Modules: None
''    Classes: None
''=======================================================
Option Private Module
Option Explicit

'Returns the passed multidimensional array as a 1 dimension array with all the elements in row-major order
Public Function arrayTo1DrowMajorOrder(arr As Variant) As Variant
    Dim arrDim As Integer 'number of dimensions
    Dim arrElem As Long 'number of elements
    Dim counter As Integer
    Dim item
    Dim i As Long
    
    arrDim = arrayNumberOfDimensions(arr)
    If arrDim = 1 Then
        arrayTo1DrowMajorOrder = arr
        Exit Function
    End If
    If arrDim < 1 Then GoTo failInput
    
    'boundArr: a two dimensional array with 2 colums and arrDim rows where the first column represents the number
    '          of elements for each of the arrDim dimensions and the second column represents the dimension multipicators.
    '          The multiplicators are calculated like following: for the first dimension, the total number of elements
    '          divided by the number of elements of the first dimension, and the other multiplicators are calculated the
    '          same but instead of the total number of elements, the multiplicator from the previous dimension is divided
    '          to the number of elements of current dimension
    'the number of elements of each dimension is equal to the difference between UBound and LBound of that dimension
    ReDim boundArr(1 To arrDim, 1 To 2)
        For i = 1 To arrDim
        boundArr(i, 1) = UBound(arr, i) - LBound(arr, i) + 1
    Next i
    
    'the total number of elements is the product of the number of elements for each dimension
    arrElem = 1
    For i = 1 To arrDim
        arrElem = arrElem * boundArr(i, 1)
    Next i
    
    'the calculation of multiplicators
    boundArr(1, 2) = arrElem / boundArr(1, 1)
    If arrDim > 1 Then
        For i = 2 To arrDim
            boundArr(i, 2) = boundArr(i - 1, 2) / boundArr(i, 1)
        Next i
    End If
    
    ReDim arrIndex(1 To arrElem)
    counter = 0
    Call arrayOrder(arrDim, boundArr, arrDim, 0, arrIndex, counter)
    
    ReDim arrElements(1 To arrElem)
    i = 0
    For Each item In arr
        i = i + 1
        arrElements(arrIndex(i)) = item
    Next item
    
    arrayTo1DrowMajorOrder = arrElements
Exit Function
failInput:
    Debug.Print "Nr. dimensiuni array invalid"
End Function

'Returns a 1D array of indexes which specify the order in which the elements of a multidimensional array is read without passing the
'               multidimensional array
'Parameters:
'   -arrDim: the number of dimensions for the multidimensional array
'   -boundArr: a two dimensional array with 2 colums and arrDim rows where the first column represents the number
'               of elements for each of the arrDim dimensions and the second column represents the dimension multipicators.
'               The multiplicators are calculated like following: for the first dimension, the total number of elements
'               divided by the number of elements of the first dimension, and the other multiplicators are calculated the
'               same but instead of the total number of elements, the multiplicator from the previous dimension is divided
'               to the number of elements of current dimension
'               The multiplicators are previously calculated so the total number of elements is not passed to this function
'   -currentDim: variable to keep track on which dimensions the function currently is, during iteration. In the initial call
'               this must be equal to arrDim.
'   -index: variable to calculate the index for elements
'   -arrIndex: array to store all the indexes. The array is previously declared with the number of elements equal to the
'               total number of elements in the multidimensional array
'   -counter: variable to keep track to which element the index should be assigned
Public Function arrayOrder(arrDim, boundArr As Variant, currentDim As Integer, index As Integer, ByRef arrIndex, ByRef counter As Integer)
    Dim i As Integer
    If arrDim = 1 Then
        arrIndex(1) = 1
        Exit Function
    End If
    If currentDim <> 1 Then
        If currentDim = arrDim Then
            For i = 1 To boundArr(currentDim, 1)
                Call arrayOrder(arrDim, boundArr, currentDim - 1, index + i, arrIndex, counter)
            Next i
        Else
            For i = 1 To boundArr(currentDim, 1)
                Call arrayOrder(arrDim, boundArr, currentDim - 1, index + (i - 1) * boundArr(currentDim, 2), arrIndex, counter)
            Next i
        End If
    Else
        For i = 1 To boundArr(currentDim, 1)
            counter = counter + 1
            arrIndex(counter) = index + (i - 1) * boundArr(currentDim, 2)
        Next i
    End If
End Function

'Returns the number of dimensions for a multidimensional array
Public Function arrayNumberOfDimensions(arr As Variant)
        Dim dimnum As Long
        Dim errorCheck As Long
      On Error GoTo FinalDimension
      For dimnum = 1 To 60000
         errorCheck = LBound(arr, dimnum)
      Next dimnum
      Exit Function
FinalDimension:
      arrayNumberOfDimensions = dimnum - 1
End Function





