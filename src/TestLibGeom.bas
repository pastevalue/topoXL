Attribute VB_Name = "TestLibGeom"
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

'@TestModule
'@Folder("Tests")

Option Explicit
Option Private Module

Private Assert As Object
Private Fakes As Object

'@ModuleInitialize
Public Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
End Sub

'@ModuleCleanup
Public Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestMethod
Public Sub TestDist2D()
    On Error GoTo TestFail

    'Assert:
    Assert.AreEqual 5#, dist2D(0, 0, 3, 4)
    Assert.AreEqual 0#, dist2D(0, 0, 0, 0)
    Assert.AreEqual 0#, dist2D(1#, 1#, 1#, 1#)
    Assert.AreEqual 0#, dist2D(-1#, -1#, -1#, -1#)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub TestAtan2()
    On Error GoTo TestFail
    'Arrange:
    Dim eps As Double
    
    'Act
    eps = 0.000000000001                         '1E-12
    
    'Assert:
    
    Assert.IsTrue LibMath.areDoublesEqual(LibGeom.PI / 2, LibGeom.Atn2(0, 1), eps), "Expected PI/2!"
    Assert.IsTrue LibMath.areDoublesEqual(LibGeom.PI / 4, LibGeom.Atn2(1, 1), eps), "Expected PI/4!"
    Assert.IsTrue LibMath.areDoublesEqual(0#, LibGeom.Atn2(1, 0), eps), "Expected 0!"
    Assert.IsTrue LibMath.areDoublesEqual(-LibGeom.PI / 4, LibGeom.Atn2(1, -1), eps), "Expected -PI/4"
    Assert.IsTrue LibMath.areDoublesEqual(-LibGeom.PI / 2, LibGeom.Atn2(0, -1), eps), "Expected -PI/2"
    Assert.IsTrue LibMath.areDoublesEqual(-3 * LibGeom.PI / 4, LibGeom.Atn2(-1, -1), eps), "Expected -3/4*PI"
    Assert.IsTrue LibMath.areDoublesEqual(LibGeom.PI, LibGeom.Atn2(-1, 0), eps), "Expected PI"
    Assert.IsTrue LibMath.areDoublesEqual(3 * LibGeom.PI / 4, LibGeom.Atn2(-1, 1), eps), "Expected 3/4*PI"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub TestAtan2on0deltas()
    Const ExpectedError As Long = 5
    On Error GoTo TestFail
    
    'Arrange:
    Dim sut As Double
    'Act:
    sut = LibGeom.Atn2(0, 0)

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

'@TestMethod
Public Sub TestOrientationIndex()
    On Error GoTo TestFail
    'Arrange:
    Dim eps As Double
    
    'Act
    eps = 0.000000000001 '1E-12
    
    'Assert:
    ' Horizontal line
    Assert.AreEqual -1, LibGeom.orientationIndex(0, 0, 1, 0, 0, 1), "Point is on the left of the horizontal line!"
    Assert.AreEqual -1, LibGeom.orientationIndex(1, 0, 0, 0, 0, -1), "Point is on the left of the horizontal line!"
    Assert.AreEqual 0, LibGeom.orientationIndex(0, 0, 1, 0, 0.5, 0), "Point is on the horizontal line!"
    Assert.AreEqual 1, LibGeom.orientationIndex(0, 0, 1, 0, 0, -1), "Point is on the right of the horizontal line!"
    Assert.AreEqual 1, LibGeom.orientationIndex(1, 0, 0, 0, 0, 1), "Point is on the right of the horizontal line!"
    ' Vertical line
    Assert.AreEqual -1, LibGeom.orientationIndex(0, 0, 0, 1, -1, 0), "Point is on the left of the vertical line!"
    Assert.AreEqual -1, LibGeom.orientationIndex(0, 1, 0, 0, 1, 0), "Point is on the left of the vertical line!"
    Assert.AreEqual 0, LibGeom.orientationIndex(0, 0, 0, 1, 0, 0.5), "Point is on the vertical line!"
    Assert.AreEqual 1, LibGeom.orientationIndex(0, 0, 0, 1, 1, 0), "Point is on the right of the vertical line!"
    Assert.AreEqual 1, LibGeom.orientationIndex(0, 1, 0, 0, -1, 0), "Point is on the right of the vertical line!"
    ' Sloped line
    Assert.AreEqual -1, LibGeom.orientationIndex(0, 0, 1, 1, 0, 1), "Point is on the left of the sloped line!"
    Assert.AreEqual -1, LibGeom.orientationIndex(1, 1, 0, 0, 1, 0), "Point is on the left of the sloped line!"
    Assert.AreEqual 0, LibGeom.orientationIndex(0, 0, 1, 1, 0.5, 0.5), "Point is on the sloped line!"
    Assert.AreEqual 1, LibGeom.orientationIndex(0, 0, 1, 1, 1, 0), "Point is on the right of the sloped line!"
    Assert.AreEqual 1, LibGeom.orientationIndex(1, 1, 0, 0, 0, 1), "Point is on the right of the sloped line!"
    
    ' Point close to the input line
    Assert.AreEqual -1, LibGeom.orientationIndex(0, 0, 0, 1, 0 - eps * 10, 0.5, eps), "Point is very close on the left of the vertical line!"
    Assert.AreEqual 0, LibGeom.orientationIndex(0, 0, 0, 1, 0 + eps / 10, 0.5, eps), "Point is very close on the vertical line!"
    Assert.AreEqual 1, LibGeom.orientationIndex(0, 0, 0, 1, 0 + eps * 10, 0.5, eps), "Point is very close on the right of the vertical line!"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

