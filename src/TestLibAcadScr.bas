Attribute VB_Name = "TestLibAcadScr"
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

'@TestModule
'@Folder("Tests")

Option Explicit
Option Private Module

Private Assert As Object
Private Fakes As Object

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'This method runs before every test in the module..
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

'@TestMethod("PntScr")
Private Sub TestPntScr1pntValid()
    On Error GoTo TestFail
    
    'Arrange:
    Dim coos2d(0, 1) As Variant
    Dim coos3D(0, 2) As Variant
    Dim coosNon0Idx(2 To 2, 10 To 12) As Variant
    
    Dim expected2D As String
    Dim expected3D As String

    'Act:
    coos2d(0, 0) = 1
    coos2d(0, 1) = 2
    
    coos3D(0, 0) = 1
    coos3D(0, 1) = 2
    coos3D(0, 2) = 3
    
    coosNon0Idx(2, 10) = 1
    coosNon0Idx(2, 11) = 2
    coosNon0Idx(2, 12) = 3
    
    expected2D = "point 1,2" & vbNewLine
    expected3D = "point 1,2,3" & vbNewLine
    
    'Assert:
    Assert.AreEqual expected2D, LibAcadScr.pnt(coos2d)
    Assert.AreEqual expected3D, LibAcadScr.pnt(coos3D)
    Assert.AreEqual expected3D, LibAcadScr.pnt(coosNon0Idx)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("PntScr")
Private Sub TestPntScr2pntsValid()
    On Error GoTo TestFail
    
    'Arrange:
    Dim coos(1, 1) As Variant
    Dim expected As String

    'Act:
    coos(0, 0) = 1
    coos(0, 1) = 2
    coos(1, 0) = 3
    coos(1, 1) = 4
    
    expected = "point 1,2" & vbNewLine & "point 3,4" & vbNewLine
    
    'Assert:
    Assert.AreEqual expected, LibAcadScr.pnt(coos)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("PntScr")
Private Sub TestPntScrInvalidDimOfArray()
    Const ExpectedError As Long = 5
    On Error GoTo TestFail
    
    'Arrange:
    Dim coos(1) As Variant
    Dim result As String

    'Act:
    coos(0) = 1
    coos(1) = 2
    
    result = LibAcadScr.pnt(coos)

Assert:
    Assert.Fail "Expected error was not raised."

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@TestMethod("PntScr")
Private Sub TestPntScrInvalidSize1OfArray()
    Const ExpectedError As Long = 5
    On Error GoTo TestFail
    
    'Arrange:
    Dim coos(0, 0) As Variant
    Dim result As String

    'Act:
    coos(0, 0) = 1
    
    result = LibAcadScr.pnt(coos)

Assert:
    Assert.Fail "Expected error was not raised."

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@TestMethod("PntScr")
Private Sub TestPntScrInvalidSize4OfArray()
    Const ExpectedError As Long = 5
    On Error GoTo TestFail
    
    'Arrange:
    Dim coos(0, 3) As Variant
    Dim result As String

    'Act:
    coos(0, 0) = 1
    coos(0, 1) = 2
    coos(0, 2) = 3
    coos(0, 3) = 4
    
    result = LibAcadScr.pnt(coos)

Assert:
    Assert.Fail "Expected error was not raised."

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@TestMethod("PntScr")
Private Sub TestPntScrInvalidNumber()
    Const ExpectedError As Long = 5
    On Error GoTo TestFail
    
    'Arrange:
    Dim coos(0, 1) As Variant
    Dim result As String

    'Act:
    coos(0, 0) = 1
    coos(0, 1) = "x"

    result = LibAcadScr.pnt(coos)

Assert:
    Assert.Fail "Expected error was not raised."

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@TestMethod("PntScr")
Private Sub TestPntScrInputNotArray()
    Const ExpectedError As Long = 5
    On Error GoTo TestFail
    
    'Arrange:
    Dim coos As Variant
    Dim result As String

    'Act:
    coos = 1

    result = LibAcadScr.pnt(coos)

Assert:
    Assert.Fail "Expected error was not raised."

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@TestMethod("PlineScr")
Private Sub TestPlineScrValid()
    On Error GoTo TestFail
    
    'Arrange:
    Dim coos2d(1, 1) As Variant
    Dim coos3D(1, 2) As Variant
    Dim coosNon0Idx(2 To 3, 10 To 12) As Variant
    
    Dim expected2D As String
    Dim expected3D As String

    'Act:
    coos2d(0, 0) = 1
    coos2d(0, 1) = 2
    coos2d(1, 0) = 3
    coos2d(1, 1) = 4
    
    coos3D(0, 0) = 1
    coos3D(0, 1) = 2
    coos3D(0, 2) = 3
    coos3D(1, 0) = 4
    coos3D(1, 1) = 5
    coos3D(1, 2) = 6
    
    coosNon0Idx(2, 10) = 1
    coosNon0Idx(2, 11) = 2
    coosNon0Idx(2, 12) = 3
    coosNon0Idx(3, 10) = 4
    coosNon0Idx(3, 11) = 5
    coosNon0Idx(3, 12) = 6
    
    expected2D = "pline 1,2" & vbNewLine & "3,4" & vbNewLine
    expected3D = "3dpoly 1,2,3" & vbNewLine & "4,5,6" & vbNewLine
    
    'Assert:
    Assert.AreEqual expected2D, LibAcadScr.pline(coos2d)
    Assert.AreEqual expected3D, LibAcadScr.pline(coos3D)
    Assert.AreEqual expected3D, LibAcadScr.pline(coosNon0Idx)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("PlineScr")
Private Sub TestPlineScrInvalidDimOfArray()
    Const ExpectedError As Long = 5
    On Error GoTo TestFail
    
    'Arrange:
    Dim coos(1) As Variant
    Dim result As String

    'Act:
    coos(0) = 1
    coos(1) = 2
    
    result = LibAcadScr.pline(coos)

Assert:
    Assert.Fail "Expected error was not raised."

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@TestMethod("PlineScr")
Private Sub TestPlineScrInvalidCooCountOfArray()
    Const ExpectedError As Long = 5
    On Error GoTo TestFail
    
    'Arrange:
    Dim coos(0, 1) As Variant
    Dim result As String

    'Act:
    coos(0, 0) = 1
    coos(0, 1) = 2
    
    result = LibAcadScr.pline(coos)

Assert:
    Assert.Fail "Expected error was not raised."

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@TestMethod("PlineScr")
Private Sub TestPlineScrInvalidSize4OfArray()
    Const ExpectedError As Long = 5
    On Error GoTo TestFail
    
    'Arrange:
    Dim coos(1, 3) As Variant
    Dim result As String

    'Act:
    coos(0, 0) = 1
    coos(0, 1) = 2
    coos(0, 2) = 3
    coos(0, 3) = 4
    coos(1, 0) = 5
    coos(1, 1) = 6
    coos(1, 2) = 7
    coos(1, 3) = 8
    
    result = LibAcadScr.pline(coos)

Assert:
    Assert.Fail "Expected error was not raised."

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@TestMethod("PlineScr")
Private Sub TestPlineScrInvalidNumber()
    Const ExpectedError As Long = 5
    On Error GoTo TestFail
    
    'Arrange:
    Dim coos(1, 1) As Variant
    Dim result As String

    'Act:
    coos(0, 0) = 1
    coos(0, 1) = 2
    coos(1, 0) = 3
    coos(1, 1) = "x"

    result = LibAcadScr.pline(coos)

Assert:
    Assert.Fail "Expected error was not raised."

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@TestMethod("PlineScr")
Private Sub TestPlineScrInputNotArray()
    Const ExpectedError As Long = 5
    On Error GoTo TestFail
    
    'Arrange:
    Dim coos As Variant
    Dim result As String

    'Act:
    coos = 1

    result = LibAcadScr.pline(coos)

Assert:
    Assert.Fail "Expected error was not raised."

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

