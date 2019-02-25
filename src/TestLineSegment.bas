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
Public Sub TestGetLength()
    On Error GoTo TestFail
    
    'Arrange:
    Dim sut As New LineSegment
    
    'Act:
    sut.Init 0, 0, 3, 4
    
    'Assert:
    Assert.AreEqual 5#, sut.GetLength, "Line segment length is wrong!"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub TestGetTheta()
    On Error GoTo TestFail
    
    'Arrange:
    Dim eps As Double
    Dim sut As New LineSegment
    
    'Act
    eps = 0.000000000001                         '1E-12
    
    'Assert:
    sut.Init 0, 0, 1, 0
    Assert.IsTrue MathLib.AreDoublesEqual(0#, sut.GetTheta(), eps), "Expected 0!"
    
    sut.Init 0, 0, 1, 1
    Assert.IsTrue MathLib.AreDoublesEqual(GeomLib.PI / 4, sut.GetTheta(), eps), "Expected PI/4!"
    
    sut.Init 0, 0, 0, 1
    Assert.IsTrue MathLib.AreDoublesEqual(GeomLib.PI / 2, sut.GetTheta(), eps), "Expected PI/2!"
    
    sut.Init 0, 0, -1, 1
    Assert.IsTrue MathLib.AreDoublesEqual(3 * GeomLib.PI / 4, sut.GetTheta(), eps), "Expected 3/4*PI"
    
    sut.Init 0, 0, -1, 0
    Assert.IsTrue MathLib.AreDoublesEqual(GeomLib.PI, sut.GetTheta(), eps), "Expected PI"
    
    sut.Init 0, 0, 1, -1
    Assert.IsTrue MathLib.AreDoublesEqual(-GeomLib.PI / 4, sut.GetTheta(), eps), "Expected -PI/4"
    
    sut.Init 0, 0, 0, -1
    Assert.IsTrue MathLib.AreDoublesEqual(-GeomLib.PI / 2, sut.GetTheta(), eps), "Expected -PI/2"
    
    sut.Init 0, 0, -1, -1
    Assert.IsTrue MathLib.AreDoublesEqual(-3 * GeomLib.PI / 4, sut.GetTheta(), eps), "Expected -3/4*PI"

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
    sut.Init 0, 1.33, 10, 1.33
    
    'Assert:
    Assert.IsTrue sut.IsHorizontal, "Line segment is not horizontal!"

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
    sut.Init 1.33, 0, 1.33, 10

    'Assert:
    Assert.IsTrue sut.IsVertical, "Line segment is not vertical!"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub TestGetProjectionFactor()
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
    sut.Init 0, 0, x, y

    'Assert:
    Assert.AreEqual 0#, sut.GetProjectionFactor(0, y + e_max), "Projection point must be start point of Line Segment!"
    Assert.AreEqual 1#, sut.GetProjectionFactor(x, y + e_max), "Projection point must be end point of Line Segment!"
    Assert.IsTrue 1 < sut.GetProjectionFactor(x + e_min, y + e_max)
    Assert.IsTrue 0 > sut.GetProjectionFactor(0 - e_min, y + e_max)

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
    
    sut.Init x1, y1, x2, y2
    els.Init x1, y1, x2, y2
    
    'Assert:
    Assert.IsTrue sut.IsEqual(els), "X and Y of equivalent LineSegments are different!"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub TestGetPointByMeasOffset()
    On Error GoTo TestFail
    
    'Arrange:
    Dim ls As New LineSegment
    Dim sut As New Point
    Dim expected As New Point
    
    ' Horizontal line segment
    ls.Init 0, 0, 1, 0
    
    ' left offset in start point
    Set sut = ls.GetPointByMeasOffset(0, -1)
    expected.Init 0, 1
    Assert.IsTrue sut.IsEqual(expected), "Expected: " & expected.ToString & ", sut: " & sut.ToString
    
    ' no offset in start point
    Set sut = ls.GetPointByMeasOffset(0, 0)
    expected.Init 0, 0
    Assert.IsTrue sut.IsEqual(expected), "Expected: " & expected.ToString & ", sut: " & sut.ToString
    
    ' right offset in start point
    Set sut = ls.GetPointByMeasOffset(0, 1)
    expected.Init 0, -1
    Assert.IsTrue sut.IsEqual(expected), "Expected: " & expected.ToString & ", sut: " & sut.ToString
    
    ' left offset in mid point
    Set sut = ls.GetPointByMeasOffset(0.5, -1)
    expected.Init 0.5, 1
    Assert.IsTrue sut.IsEqual(expected), "Expected: " & expected.ToString & ", sut: " & sut.ToString
    
    ' no offset in mid point
    Set sut = ls.GetPointByMeasOffset(0.5, 0)
    expected.Init 0.5, 0
    Assert.IsTrue sut.IsEqual(expected), "Expected: " & expected.ToString & ", sut: " & sut.ToString
    
    ' right offset in mid point
    Set sut = ls.GetPointByMeasOffset(0.5, 1)
    expected.Init 0.5, -1
    Assert.IsTrue sut.IsEqual(expected), "Expected: " & expected.ToString & ", sut: " & sut.ToString
    
    ' left offset in end point
    Set sut = ls.GetPointByMeasOffset(1, -1)
    expected.Init 1, 1
    Assert.IsTrue sut.IsEqual(expected), "Expected: " & expected.ToString & ", sut: " & sut.ToString
    
    ' no offset in end point
    Set sut = ls.GetPointByMeasOffset(1, 0)
    expected.Init 1, 0
    Assert.IsTrue sut.IsEqual(expected), "Expected: " & expected.ToString & ", sut: " & sut.ToString
    
    ' right offset in end point
    Set sut = ls.GetPointByMeasOffset(1, 1)
    expected.Init 1, -1
    Assert.IsTrue sut.IsEqual(expected), "Expected: " & expected.ToString & ", sut: " & sut.ToString
    
    
    ' Vertical line segment
    ls.Init 0, 1, 0, 0
    
    ' left offset in start point
    Set sut = ls.GetPointByMeasOffset(0, -1)
    expected.Init 1, 1
    Assert.IsTrue sut.IsEqual(expected), "Expected: " & expected.ToString & ", sut: " & sut.ToString
    
    ' no offset in start point
    Set sut = ls.GetPointByMeasOffset(0, 0)
    expected.Init 0, 1
    Assert.IsTrue sut.IsEqual(expected), "Expected: " & expected.ToString & ", sut: " & sut.ToString
    
    ' right offset in start point
    Set sut = ls.GetPointByMeasOffset(0, 1)
    expected.Init -1, 1
    Assert.IsTrue sut.IsEqual(expected), "Expected: " & expected.ToString & ", sut: " & sut.ToString
    
    ' left offset in mid point
    Set sut = ls.GetPointByMeasOffset(0.5, -1)
    expected.Init 1, 0.5
    Assert.IsTrue sut.IsEqual(expected), "Expected: " & expected.ToString & ", sut: " & sut.ToString
    
    ' no offset in mid point
    Set sut = ls.GetPointByMeasOffset(0.5, 0)
    expected.Init 0, 0.5
    Assert.IsTrue sut.IsEqual(expected), "Expected: " & expected.ToString & ", sut: " & sut.ToString
    
    ' right offset in mid point
    Set sut = ls.GetPointByMeasOffset(0.5, 1)
    expected.Init -1, 0.5
    Assert.IsTrue sut.IsEqual(expected), "Expected: " & expected.ToString & ", sut: " & sut.ToString
    
    ' left offset in end point
    Set sut = ls.GetPointByMeasOffset(1, -1)
    expected.Init 1, 0
    Assert.IsTrue sut.IsEqual(expected), "Expected: " & expected.ToString & ", sut: " & sut.ToString
    
    ' no offset in end point
    Set sut = ls.GetPointByMeasOffset(1, 0)
    expected.Init 0, 0
    Assert.IsTrue sut.IsEqual(expected), "Expected: " & expected.ToString & ", sut: " & sut.ToString
    
    ' right offset in end point
    Set sut = ls.GetPointByMeasOffset(1, 1)
    expected.Init -1, 0
    Assert.IsTrue sut.IsEqual(expected), "Expected: " & expected.ToString & ", sut: " & sut.ToString
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub TestGetPointByNegMeasOffset()
    On Error GoTo TestFail
    
    'Arrange:
    Dim ls As New LineSegment
    Dim sut As New Point
    Dim expected As New Point
    
    'Act:
    ls.Init 3, 0, 0, 0
    
    'Assert:
    Set sut = ls.GetPointByMeasOffset(-2, 0)
    expected.Init 2, 0
    Assert.IsTrue sut.IsEqual(expected), "Expected: " & expected.ToString & ", sut: " & sut.ToString
        
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub TestGetPointByMeasOffsetOutside()
    On Error GoTo TestFail
    
    'Arrange:
    Dim ls As New LineSegment
    Dim e As Double
    
    'Act:
    e = 0.000000000000001                        '1E-15
    ls.Init 1, 1, 0, 0
    
    'Assert:
    Assert.IsNothing ls.GetPointByMeasOffset(ls.GetLength + e, 0) ' positive measure
    Assert.IsNothing ls.GetPointByMeasOffset(-(ls.GetLength + e), 0) ' negative measure

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub TestGetMeasOffsetOfPoint()
    On Error GoTo TestFail
    
    'Arrange:
    Dim ls As New LineSegment
    Dim sut As New MeasOffset
    Dim expected As New MeasOffset
    
    ' Horizontal line segment
    ls.Init 0, 0, 1, 0
    
    ' left offset in start point
    Set sut = ls.GetMeasOffsetOfPoint(0, 1)
    expected.Init 0, -1
    Assert.IsTrue sut.IsEqual(expected), "Expected: " & expected.ToString & ", sut: " & sut.ToString
    
    ' no offset in start point
    Set sut = ls.GetMeasOffsetOfPoint(0, 0)
    expected.Init 0, 0
    Assert.IsTrue sut.IsEqual(expected), "Expected: " & expected.ToString & ", sut: " & sut.ToString
    
    ' ritght offset in start point
    Set sut = ls.GetMeasOffsetOfPoint(0, -1)
    expected.Init 0, 1
    Assert.IsTrue sut.IsEqual(expected), "Expected: " & expected.ToString & ", sut: " & sut.ToString
    
    ' left offset in mid point
    Set sut = ls.GetMeasOffsetOfPoint(0.5, 1)
    expected.Init 0.5, -1
    Assert.IsTrue sut.IsEqual(expected), "Expected: " & expected.ToString & ", sut: " & sut.ToString
    
    ' no offset in mid point
    Set sut = ls.GetMeasOffsetOfPoint(0.5, 0)
    expected.Init 0.5, 0
    Assert.IsTrue sut.IsEqual(expected), "Expected: " & expected.ToString & ", sut: " & sut.ToString
    
    ' right offset in mid point
    Set sut = ls.GetMeasOffsetOfPoint(0.5, -1)
    expected.Init 0.5, 1
    Assert.IsTrue sut.IsEqual(expected), "Expected: " & expected.ToString & ", sut: " & sut.ToString
    
    ' left offset in end point
    Set sut = ls.GetMeasOffsetOfPoint(1, 1)
    expected.Init 1, -1
    Assert.IsTrue sut.IsEqual(expected), "Expected: " & expected.ToString & ", sut: " & sut.ToString
    
    ' no offset in end point
    Set sut = ls.GetMeasOffsetOfPoint(1, 0)
    expected.Init 1, 0
    Assert.IsTrue sut.IsEqual(expected), "Expected: " & expected.ToString & ", sut: " & sut.ToString
    
    ' ritght offset in end point
    Set sut = ls.GetMeasOffsetOfPoint(1, -1)
    expected.Init 1, 1
    Assert.IsTrue sut.IsEqual(expected), "Expected: " & expected.ToString & ", sut: " & sut.ToString
    
    
    ' Vertical line segment
    ls.Init 0, 1, 0, 0
    
    ' left offset in start point
    Set sut = ls.GetMeasOffsetOfPoint(1, 1)
    expected.Init 0, -1
    Assert.IsTrue sut.IsEqual(expected), "Expected: " & expected.ToString & ", sut: " & sut.ToString
    
    ' no offset in start point
    Set sut = ls.GetMeasOffsetOfPoint(0, 1)
    expected.Init 0, 0
    Assert.IsTrue sut.IsEqual(expected), "Expected: " & expected.ToString & ", sut: " & sut.ToString
    
    ' ritght offset in start point
    Set sut = ls.GetMeasOffsetOfPoint(-1, 1)
    expected.Init 0, 1
    Assert.IsTrue sut.IsEqual(expected), "Expected: " & expected.ToString & ", sut: " & sut.ToString
    
    ' left offset in mid point
    Set sut = ls.GetMeasOffsetOfPoint(1, 0.5)
    expected.Init 0.5, -1
    Assert.IsTrue sut.IsEqual(expected), "Expected: " & expected.ToString & ", sut: " & sut.ToString
    
    ' no offset in mid point
    Set sut = ls.GetMeasOffsetOfPoint(0, 0.5)
    expected.Init 0.5, 0
    Assert.IsTrue sut.IsEqual(expected), "Expected: " & expected.ToString & ", sut: " & sut.ToString
    
    ' right offset in mid point
    Set sut = ls.GetMeasOffsetOfPoint(-1, 0.5)
    expected.Init 0.5, 1
    Assert.IsTrue sut.IsEqual(expected), "Expected: " & expected.ToString & ", sut: " & sut.ToString
    
    ' left offset in end point
    Set sut = ls.GetMeasOffsetOfPoint(1, 0)
    expected.Init 1, -1
    Assert.IsTrue sut.IsEqual(expected), "Expected: " & expected.ToString & ", sut: " & sut.ToString
    
    ' no offset in end point
    Set sut = ls.GetMeasOffsetOfPoint(0, 0)
    expected.Init 1, 0
    Assert.IsTrue sut.IsEqual(expected), "Expected: " & expected.ToString & ", sut: " & sut.ToString
    
    ' ritght offset in end point
    Set sut = ls.GetMeasOffsetOfPoint(-1, 0)
    expected.Init 1, 1
    Assert.IsTrue sut.IsEqual(expected), "Expected: " & expected.ToString & ", sut: " & sut.ToString
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub TestGetMeasOffsetOfPointOutside()
    On Error GoTo TestFail
    
    'Arrange:
    Dim ls As New LineSegment
    Dim e As Double
    
    'Act:
    e = 0.000000000000001                        '1E-15
    ls.Init 1, 1, 0, 0
    
    'Assert:
    Assert.IsNothing ls.GetMeasOffsetOfPoint(1 + e, 1)
    Assert.IsNothing ls.GetMeasOffsetOfPoint(1, 1 + e)
    Assert.IsNothing ls.GetMeasOffsetOfPoint(0 - e, 0)
    Assert.IsNothing ls.GetMeasOffsetOfPoint(0, 0 - e)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

