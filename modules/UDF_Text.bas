Attribute VB_Name = "UDF_Text"
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

'returns the sum of integers from the received ranges
'each cell in particular represents an individual string
'the integer is in fact the concatenation of all found figures in the string(each cell) ignoring all other characters
Public Function SumIntegersFromTexts(ParamArray ranges() As Variant)
    Application.Volatile False
    Dim tempArray() As Variant
    Dim t As Variant
    Dim partialSum As Integer
    Dim i As Integer
    Dim Length_of_String As Integer
    Dim Current_Pos As Integer
    Dim currentChar As String
    Dim temp As String
    Dim ExtractIntegerfromText As Integer
    Dim Phrase As String

    t = ranges
    partialSum = 0
    tempArray = rangeFunctions.valuesTo2Darray(1, t)
    For i = 0 To UBound(tempArray)
        Phrase = CStr(tempArray(i, 0))
        Length_of_String = Len(Phrase)
        temp = ""
        For Current_Pos = 1 To Length_of_String
            currentChar = Mid(Phrase, Current_Pos, 1)
            If (IsNumeric(currentChar)) = True Then temp = temp & currentChar
        Next Current_Pos
        If Len(temp) = 0 Then
            ExtractIntegerfromText = 0
        Else
            ExtractIntegerfromText = CInt(temp)
        End If
        partialSum = partialSum + ExtractIntegerfromText
    Next i
    SumIntegersFromTexts = partialSum
End Function

Public Function textShowSexagesimalDegreesMinutesAndSecondsFromDec(dec As Double) As String
    Application.Volatile False
    Dim remaining As Double
    Dim sexagesimal As String
    Dim sign As Integer
    
    sign = Sgn(dec)
    dec = Abs(dec)
    remaining = dec - WorksheetFunction.Quotient(dec, 1)
    dec = dec - remaining
    sexagesimal = WorksheetFunction.Text(dec, "00") & "° "
    remaining = remaining * 60
    dec = remaining
    remaining = dec - WorksheetFunction.Quotient(dec, 1)
    dec = dec - remaining
    sexagesimal = sexagesimal & WorksheetFunction.Text(dec, "00") & "' "
    remaining = remaining * 60
    sexagesimal = sexagesimal & WorksheetFunction.Text(remaining, "00.00") & "''"
    If sign = -1 Then sexagesimal = "-" & sexagesimal
    textShowSexagesimalDegreesMinutesAndSecondsFromDec = sexagesimal
End Function
