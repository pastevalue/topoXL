Attribute VB_Name = "TestLineSegment"
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
Public Sub TestInitInvalidSE()
    Const ExpectedError As Long = 5
    On Error GoTo TestFail
    
    'Arrange:
    Dim sut As New LineSegment
    
    'Act:
    sut.init 0, 0, 0, 0

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
Public Sub TestInit()
    On Error GoTo TestFail
    
    'Arrange:
    Dim sut As New LineSegment
    Dim epsilon As Double
    
    'Act:
    epsilon = 0.000000000000001                  '1E-15
    

    ' First quadrant, CW, PI/6 length
    sut.init 0, 0, 3, 4
    
    'Assert:
    Assert.AreEqual 3#, sut.dX, "dX must be 3!"
    Assert.AreEqual 4#, sut.dY, "dy must be 4!"
    Assert.AreEqual 5#, sut.length, "Length must be 4!"
    Assert.IsTrue LibMath.areDoublesEqual(LibGeom.Atn2(3, 4), sut.theta, epsilon), "Theta of line segment must be" & LibGeom.Atn2(3, 4) & "!"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub TestLength()
    On Error GoTo TestFail
    
    'Arrange:
    Dim sut As New LineSegment
    
    'Act:
    sut.init 0, 0, 3, 4
    
    'Assert:
    Assert.AreEqual 5#, sut.length, "Line segment length is wrong!"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub TestTheta()
    On Error GoTo TestFail
    
    'Arrange:
    Dim eps As Double
    Dim sut As New LineSegment
    
    'Act
    eps = 0.000000000001                         '1E-12
    
    'Assert:
    sut.init 0, 0, 1, 0
    Assert.IsTrue LibMath.areDoublesEqual(0#, sut.theta, eps), "Expected 0!"
    
    sut.init 0, 0, 1, 1
    Assert.IsTrue LibMath.areDoublesEqual(LibGeom.PI / 4, sut.theta, eps), "Expected PI/4!"
    
    sut.init 0, 0, 0, 1
    Assert.IsTrue LibMath.areDoublesEqual(LibGeom.PI / 2, sut.theta, eps), "Expected PI/2!"
    
    sut.init 0, 0, -1, 1
    Assert.IsTrue LibMath.areDoublesEqual(3 * LibGeom.PI / 4, sut.theta, eps), "Expected 3/4*PI"
    
    sut.init 0, 0, -1, 0
    Assert.IsTrue LibMath.areDoublesEqual(LibGeom.PI, sut.theta, eps), "Expected PI"
    
    sut.init 0, 0, 1, -1
    Assert.IsTrue LibMath.areDoublesEqual(-LibGeom.PI / 4, sut.theta, eps), "Expected -PI/4"
    
    sut.init 0, 0, 0, -1
    Assert.IsTrue LibMath.areDoublesEqual(-LibGeom.PI / 2, sut.theta, eps), "Expected -PI/2"
    
    sut.init 0, 0, -1, -1
    Assert.IsTrue LibMath.areDoublesEqual(-3 * LibGeom.PI / 4, sut.theta, eps), "Expected -3/4*PI"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub TestIsHorizontal()
    On Error GoTo TestFail
    
    'Arrange:
    Dim sut As New LineSegment
    sut.init 0, 1.33, 10, 1.33
    
    'Assert:
    Assert.IsTrue sut.isHorizontal, "Line segment is not horizontal!"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub TestIsVertical()
    On Error GoTo TestFail
    
    'Arrange:
    Dim sut As New LineSegment
    sut.init 1.33, 0, 1.33, 10

    'Assert:
    Assert.IsTrue sut.isVertical, "Line segment is not vertical!"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub TestCalcProjectionFactor()
    On Error GoTo TestFail
    
    'Arrange:
    Dim sut As New LineSegment
    Dim x As Double
    Dim y As Double
    Dim e_min As Double
    Dim e_max As Double
    
    'Act:
    e_min = 0.000000000000001                    '1E-15
    e_max = 100000000000000#
    x = 1 / 3
    y = 0
    sut.init 0, 0, x, y

    'Assert:
    Assert.AreEqual 0#, sut.calcProjectionFactor(0, y + e_max), "Projection point must be start point of Line Segment!"
    Assert.AreEqual 1#, sut.calcProjectionFactor(x, y + e_max), "Projection point must be end point of Line Segment!"
    Assert.IsTrue 1 < sut.calcProjectionFactor(x + e_min, y + e_max)
    Assert.IsTrue 0 > sut.calcProjectionFactor(0 - e_min, y + e_max)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub TestIsEqual()
    On Error GoTo TestFail
    
    'Arrange:
    Dim sut As New LineSegment
    Dim els As New LineSegment                   'equivalent LineSegment
    Dim x1 As Double
    Dim y1 As Double
    Dim x2 As Double
    Dim y2 As Double
    
    'Act:
    x1 = 1 / 3
    y1 = Math.Sqr(2)
    
    x2 = 100000000000000#
    y2 = 0.000000000000001
    
    sut.init x1, y1, x2, y2
    els.init x1, y1, x2, y2
    
    'Assert:
    Assert.IsTrue sut.isEqual(els), "X and Y of equivalent LineSegments are different!"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub TestCalcPointByMeasOffset()
    On Error GoTo TestFail
    
    'Arrange:
    Dim ls As New LineSegment
    Dim sut As New Point
    Dim expected As New Point
    
    ' Horizontal line segment
    ls.init 0, 0, 1, 0
    
    ' left offset in start point
    Set sut = ls.calcPointByMeasOffset(0, -1)
    expected.init 0, 1
    Assert.IsTrue sut.isEqual(expected), "Expected: " & expected.toStringXY & ", sut: " & sut.toStringXY
    
    ' no offset in start point
    Set sut = ls.calcPointByMeasOffset(0, 0)
    expected.init 0, 0
    Assert.IsTrue sut.isEqual(expected), "Expected: " & expected.toStringXY & ", sut: " & sut.toStringXY
    
    ' right offset in start point
    Set sut = ls.calcPointByMeasOffset(0, 1)
    expected.init 0, -1
    Assert.IsTrue sut.isEqual(expected), "Expected: " & expected.toStringXY & ", sut: " & sut.toStringXY
    
    ' left offset in mid point
    Set sut = ls.calcPointByMeasOffset(0.5, -1)
    expected.init 0.5, 1
    Assert.IsTrue sut.isEqual(expected), "Expected: " & expected.toStringXY & ", sut: " & sut.toStringXY
    
    ' no offset in mid point
    Set sut = ls.calcPointByMeasOffset(0.5, 0)
    expected.init 0.5, 0
    Assert.IsTrue sut.isEqual(expected), "Expected: " & expected.toStringXY & ", sut: " & sut.toStringXY
    
    ' right offset in mid point
    Set sut = ls.calcPointByMeasOffset(0.5, 1)
    expected.init 0.5, -1
    Assert.IsTrue sut.isEqual(expected), "Expected: " & expected.toStringXY & ", sut: " & sut.toStringXY
    
    ' left offset in end point
    Set sut = ls.calcPointByMeasOffset(1, -1)
    expected.init 1, 1
    Assert.IsTrue sut.isEqual(expected), "Expected: " & expected.toStringXY & ", sut: " & sut.toStringXY
    
    ' no offset in end point
    Set sut = ls.calcPointByMeasOffset(1, 0)
    expected.init 1, 0
    Assert.IsTrue sut.isEqual(expected), "Expected: " & expected.toStringXY & ", sut: " & sut.toStringXY
    
    ' right offset in end point
    Set sut = ls.calcPointByMeasOffset(1, 1)
    expected.init 1, -1
    Assert.IsTrue sut.isEqual(expected), "Expected: " & expected.toStringXY & ", sut: " & sut.toStringXY
    
    
    ' Vertical line segment
    ls.init 0, 1, 0, 0
    
    ' left offset in start point
    Set sut = ls.calcPointByMeasOffset(0, -1)
    expected.init 1, 1
    Assert.IsTrue sut.isEqual(expected), "Expected: " & expected.toStringXY & ", sut: " & sut.toStringXY
    
    ' no offset in start point
    Set sut = ls.calcPointByMeasOffset(0, 0)
    expected.init 0, 1
    Assert.IsTrue sut.isEqual(expected), "Expected: " & expected.toStringXY & ", sut: " & sut.toStringXY
    
    ' right offset in start point
    Set sut = ls.calcPointByMeasOffset(0, 1)
    expected.init -1, 1
    Assert.IsTrue sut.isEqual(expected), "Expected: " & expected.toStringXY & ", sut: " & sut.toStringXY
    
    ' left offset in mid point
    Set sut = ls.calcPointByMeasOffset(0.5, -1)
    expected.init 1, 0.5
    Assert.IsTrue sut.isEqual(expected), "Expected: " & expected.toStringXY & ", sut: " & sut.toStringXY
    
    ' no offset in mid point
    Set sut = ls.calcPointByMeasOffset(0.5, 0)
    expected.init 0, 0.5
    Assert.IsTrue sut.isEqual(expected), "Expected: " & expected.toStringXY & ", sut: " & sut.toStringXY
    
    ' right offset in mid point
    Set sut = ls.calcPointByMeasOffset(0.5, 1)
    expected.init -1, 0.5
    Assert.IsTrue sut.isEqual(expected), "Expected: " & expected.toStringXY & ", sut: " & sut.toStringXY
    
    ' left offset in end point
    Set sut = ls.calcPointByMeasOffset(1, -1)
    expected.init 1, 0
    Assert.IsTrue sut.isEqual(expected), "Expected: " & expected.toStringXY & ", sut: " & sut.toStringXY
    
    ' no offset in end point
    Set sut = ls.calcPointByMeasOffset(1, 0)
    expected.init 0, 0
    Assert.IsTrue sut.isEqual(expected), "Expected: " & expected.toStringXY & ", sut: " & sut.toStringXY
    
    ' right offset in end point
    Set sut = ls.calcPointByMeasOffset(1, 1)
    expected.init -1, 0
    Assert.IsTrue sut.isEqual(expected), "Expected: " & expected.toStringXY & ", sut: " & sut.toStringXY
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub TestCalcPointByMeasOffsetOutside()
    On Error GoTo TestFail
    
    'Arrange:
    Dim ls As New LineSegment
    Dim e As Double
    
    'Act:
    e = 0.000000000000001                        '1E-15
    ls.init 1, 1, 0, 0
    
    'Assert:
    Assert.IsNothing ls.calcPointByMeasOffset(ls.length + e, 0) ' positive measure
    Assert.IsNothing ls.calcPointByMeasOffset(-(ls.length + e), 0) ' negative measure

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub TestCalcMeasOffsetOfPoint()
    On Error GoTo TestFail
    
    'Arrange:
    Dim ls As New LineSegment
    Dim sut As New MeasOffset
    Dim expected As New MeasOffset
    
    ' Horizontal line segment
    ls.init 0, 0, 1, 0
    
    ' left offset in start point
    Set sut = ls.calcMeasOffsetOfPoint(0, 1)
    expected.init 0, -1
    Assert.IsTrue sut.isEqual(expected), "Expected: " & expected.toString & ", sut: " & sut.toString
    
    ' no offset in start point
    Set sut = ls.calcMeasOffsetOfPoint(0, 0)
    expected.init 0, 0
    Assert.IsTrue sut.isEqual(expected), "Expected: " & expected.toString & ", sut: " & sut.toString
    
    ' ritght offset in start point
    Set sut = ls.calcMeasOffsetOfPoint(0, -1)
    expected.init 0, 1
    Assert.IsTrue sut.isEqual(expected), "Expected: " & expected.toString & ", sut: " & sut.toString
    
    ' left offset in mid point
    Set sut = ls.calcMeasOffsetOfPoint(0.5, 1)
    expected.init 0.5, -1
    Assert.IsTrue sut.isEqual(expected), "Expected: " & expected.toString & ", sut: " & sut.toString
    
    ' no offset in mid point
    Set sut = ls.calcMeasOffsetOfPoint(0.5, 0)
    expected.init 0.5, 0
    Assert.IsTrue sut.isEqual(expected), "Expected: " & expected.toString & ", sut: " & sut.toString
    
    ' right offset in mid point
    Set sut = ls.calcMeasOffsetOfPoint(0.5, -1)
    expected.init 0.5, 1
    Assert.IsTrue sut.isEqual(expected), "Expected: " & expected.toString & ", sut: " & sut.toString
    
    ' left offset in end point
    Set sut = ls.calcMeasOffsetOfPoint(1, 1)
    expected.init 1, -1
    Assert.IsTrue sut.isEqual(expected), "Expected: " & expected.toString & ", sut: " & sut.toString
    
    ' no offset in end point
    Set sut = ls.calcMeasOffsetOfPoint(1, 0)
    expected.init 1, 0
    Assert.IsTrue sut.isEqual(expected), "Expected: " & expected.toString & ", sut: " & sut.toString
    
    ' ritght offset in end point
    Set sut = ls.calcMeasOffsetOfPoint(1, -1)
    expected.init 1, 1
    Assert.IsTrue sut.isEqual(expected), "Expected: " & expected.toString & ", sut: " & sut.toString
    
    
    ' Vertical line segment
    ls.init 0, 1, 0, 0
    
    ' left offset in start point
    Set sut = ls.calcMeasOffsetOfPoint(1, 1)
    expected.init 0, -1
    Assert.IsTrue sut.isEqual(expected), "Expected: " & expected.toString & ", sut: " & sut.toString
    
    ' no offset in start point
    Set sut = ls.calcMeasOffsetOfPoint(0, 1)
    expected.init 0, 0
    Assert.IsTrue sut.isEqual(expected), "Expected: " & expected.toString & ", sut: " & sut.toString
    
    ' ritght offset in start point
    Set sut = ls.calcMeasOffsetOfPoint(-1, 1)
    expected.init 0, 1
    Assert.IsTrue sut.isEqual(expected), "Expected: " & expected.toString & ", sut: " & sut.toString
    
    ' left offset in mid point
    Set sut = ls.calcMeasOffsetOfPoint(1, 0.5)
    expected.init 0.5, -1
    Assert.IsTrue sut.isEqual(expected), "Expected: " & expected.toString & ", sut: " & sut.toString
    
    ' no offset in mid point
    Set sut = ls.calcMeasOffsetOfPoint(0, 0.5)
    expected.init 0.5, 0
    Assert.IsTrue sut.isEqual(expected), "Expected: " & expected.toString & ", sut: " & sut.toString
    
    ' right offset in mid point
    Set sut = ls.calcMeasOffsetOfPoint(-1, 0.5)
    expected.init 0.5, 1
    Assert.IsTrue sut.isEqual(expected), "Expected: " & expected.toString & ", sut: " & sut.toString
    
    ' left offset in end point
    Set sut = ls.calcMeasOffsetOfPoint(1, 0)
    expected.init 1, -1
    Assert.IsTrue sut.isEqual(expected), "Expected: " & expected.toString & ", sut: " & sut.toString
    
    ' no offset in end point
    Set sut = ls.calcMeasOffsetOfPoint(0, 0)
    expected.init 1, 0
    Assert.IsTrue sut.isEqual(expected), "Expected: " & expected.toString & ", sut: " & sut.toString
    
    ' ritght offset in end point
    Set sut = ls.calcMeasOffsetOfPoint(-1, 0)
    expected.init 1, 1
    Assert.IsTrue sut.isEqual(expected), "Expected: " & expected.toString & ", sut: " & sut.toString
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub TestCalcMeasOffsetOfPointOutside()
    On Error GoTo TestFail
    
    'Arrange:
    Dim ls As New LineSegment
    Dim e As Double
    
    'Act:
    e = 0.000000000000001                        '1E-15
    ls.init 1, 1, 0, 0
    
    'Assert:
    Assert.IsNothing ls.calcMeasOffsetOfPoint(1 + e, 1)
    Assert.IsNothing ls.calcMeasOffsetOfPoint(1, 1 + e)
    Assert.IsNothing ls.calcMeasOffsetOfPoint(0 - e, 0)
    Assert.IsNothing ls.calcMeasOffsetOfPoint(0, 0 - e)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub TestCalcXatYonHorizontal()
    On Error GoTo TestFail
    
    'Arrange:
    Dim lsH As New LineSegment
    Dim lsHe As New LineSegment
    Dim e As Double
    
    'Act:
    e = 0.000000000000001            '1E-15
    lsH.init 0, 1, 1, 1
    lsHe.init 0, 1, 1, 1 + e
    
    'Assert:
    Assert.IsTrue IsNull(lsH.calcXatY(1)), "Null expected for horizontal lines"
    Assert.IsFalse IsNull(lsHe.calcXatY(1)), "Not Null expected for non-horizontal line"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub TestCalcXatYonVertical()
    On Error GoTo TestFail
    
    'Arrange:
    Dim lsV As New LineSegment
    Dim lsVe As New LineSegment
    Dim e As Double
    
    'Act:
    e = 0.000000000000001 '1E-15
    lsV.init 1, 0, 1, 1
    lsVe.init 1, 0, 1 + e, 1
    
    'Assert:
    Assert.AreEqual 1#, lsV.calcXatY(0), "Same value expected for all Ys on vertical lines"
    Assert.AreEqual 1#, lsV.calcXatY(0.5), "Same value expected for all Ys on vertical lines"
    Assert.AreEqual 1#, lsV.calcXatY(1), "Same value expected for all Ys on vertical lines"
    Assert.AreNotEqual 1#, lsVe.calcXatY(0.5), "Different value expected for if line is not perfectly vertical"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub TestCalcXatYoutOfRange()
    On Error GoTo TestFail
    
    'Arrange:
    Dim ls As New LineSegment
    Dim e As Double
    
    'Act:
    e = 0.000000000000001            '1E-15
    ls.init 0, 0, 1, 1
    
    'Assert:
    Assert.IsTrue IsNull(ls.calcXatY(0 - e)), "Null expected - outside of Y range on left"
    Assert.IsTrue IsNull(ls.calcXatY(1 + e)), "Null expected - outside of Y range on right"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub TestCalcXatYvalid()
    On Error GoTo TestFail
    
    'Arrange:
    Dim ls As New LineSegment
    Dim e As Double
    
    'Act:
    ls.init 0, 0, 2, 1
    
    'Assert:
    Assert.AreEqual 1#, ls.calcXatY(0.5)
    Assert.AreEqual 0#, ls.calcXatY(0)
    Assert.AreEqual 2#, ls.calcXatY(1)
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
'''''
'@TestMethod
Public Sub TestCalcYatXonVertical()
    On Error GoTo TestFail
    
    'Arrange:
    Dim lsV As New LineSegment
    Dim lsVe As New LineSegment
    Dim e As Double
    
    'Act:
    e = 0.000000000000001            '1E-15
    lsV.init 1, 0, 1, 1
    lsVe.init 1, 0, 1 + e, 1
    
    'Assert:
    Assert.IsTrue IsNull(lsV.calcYatX(1)), "Null expected for horizontal lines"
    Assert.IsFalse IsNull(lsVe.calcYatX(1)), "Not Null expected for non-horizontal line"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub TestCalcYatXonHorizontal()
    On Error GoTo TestFail
    
    'Arrange:
    Dim lsH As New LineSegment
    Dim lsHe As New LineSegment
    Dim e As Double
    
    'Act:
    e = 0.000000000000001 '1E-15
    lsH.init 0, 1, 1, 1
    lsHe.init 0, 1, 1, 1 + e
    
    'Assert:
    Assert.AreEqual 1#, lsH.calcYatX(0), "Same value expected for all Ys on vertical lines"
    Assert.AreEqual 1#, lsH.calcYatX(0.5), "Same value expected for all Ys on vertical lines"
    Assert.AreEqual 1#, lsH.calcYatX(1), "Same value expected for all Ys on vertical lines"
    Assert.AreNotEqual 1#, lsHe.calcYatX(0.5), "Different value expected for if line is not perfectly vertical"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub TestCalcYatXoutOfRange()
    On Error GoTo TestFail
    
    'Arrange:
    Dim ls As New LineSegment
    Dim e As Double
    
    'Act:
    e = 0.000000000000001            '1E-15
    ls.init 0, 0, 1, 1
    
    'Assert:
    Assert.IsTrue IsNull(ls.calcYatX(0 - e)), "Null expected - outside of X range on left"
    Assert.IsTrue IsNull(ls.calcYatX(1 + e)), "Null expected - outside of X range on right"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub TestCalcYatXvalid()
    On Error GoTo TestFail
    
    'Arrange:
    Dim ls As New LineSegment
    Dim e As Double
    
    'Act:
    ls.init 0, 0, 1, 2
    
    'Assert:
    Assert.AreEqual 1#, ls.calcYatX(0.5)
    Assert.AreEqual 0#, ls.calcYatX(0)
    Assert.AreEqual 2#, ls.calcYatX(1)
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

