Attribute VB_Name = "TestGeomFactory"
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
Public Sub TestNewPointFromValidVariant()
    On Error GoTo TestFail
    
    'Arrange:
    Dim sut As Point
    Dim expected As New Point
    Dim x As Variant
    Dim y As Variant
    
    'Act:
    x = "3.33"
    y = "-6.0"
    expected.x = 3.33
    expected.y = -6#
    
    Set sut = GeomFactory.NewPointFromVariant(x, y)
    
    'Assert:
    Assert.AreEqual expected.x, sut.x, "Failed to extract X Double value from Variant!"
    Assert.AreEqual expected.y, sut.y, "Failed to extract Y Double values from Variant!"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub TestNewPointFromInvalidVariant()
    On Error GoTo TestFail
    
    'Arrange:
    Dim sut As Point
    Dim x As Variant
    Dim y As Variant
    
    'Act:
    x = "3.33"
    y = "-6.0abc"
    Set sut = GeomFactory.NewPointFromVariant(x, y)

    'Assert:
    Assert.IsTrue sut Is Nothing, "Nothing expected on Point initialize with invalid arguments!"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub TestNewMeasOffsetFromValidVariant()
    On Error GoTo TestFail
    
    'Arrange:
    Dim sut As MeasOffset
    Dim expected As New MeasOffset
    Dim m As Variant
    Dim o As Variant
    
    'Act:
    m = "3.33"
    o = "-6.0"
    expected.m = 3.33
    expected.o = -6#
    
    Set sut = GeomFactory.NewMeasOffsetFromVariant(m, o)
    
    'Assert:
    Assert.AreEqual expected.m, sut.m, "Failed to extract Measure Double value from Variant!"
    Assert.AreEqual expected.o, sut.o, "Failed to extract Offset Double values from Variant!"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub TestNewMeasOffsetFromInvalidVariant()
    On Error GoTo TestFail
    
    'Arrange:
    Dim sut As MeasOffset
    Dim m As Variant
    Dim o As Variant
    
    'Act:
    m = "3.33"
    o = "-6.0abc"
    
    Set sut = GeomFactory.NewMeasOffsetFromVariant(m, o)
    
    'Assert:
    Assert.IsTrue sut Is Nothing, "Nothing expected on MeasOffset initialize with invalid arguments!"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub TestNewLineSegmentFromValidVariant()
    On Error GoTo TestFail
    
    'Arrange:
    Dim sut As LineSegment
    Dim expected As New LineSegment
    Dim x1 As Variant
    Dim y1 As Variant
    Dim x2 As Variant
    Dim y2 As Variant
    
    'Act:
    x1 = "3.33"
    y1 = "-6.0"
    x2 = "3.0"
    y2 = "6.66"
    expected.Init 3.33, -6, 3, 6.66
    
    Set sut = GeomFactory.NewLineSegmentFromVariant(x1, y1, x2, y2)
    
    'Assert:
    Assert.AreEqual expected.P1.x, sut.P1.x
    Assert.AreEqual expected.P1.y, sut.P1.y
    Assert.AreEqual expected.P2.x, sut.P2.x
    Assert.AreEqual expected.P2.y, sut.P2.y

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub TestNewLineSegmentFromInvalidVariant()
    On Error GoTo TestFail
    
    'Arrange:
    Dim sut As LineSegment
    Dim expected As New LineSegment
    Dim x1 As Variant
    Dim y1 As Variant
    Dim x2 As Variant
    Dim y2 As Variant
    
    'Act:
    x1 = "3.33"
    y1 = "-6.0"
    x2 = "3.0"
    y2 = "6.66abc"
    expected.Init 3.33, -6, 3, 6.66
    
    Set sut = GeomFactory.NewLineSegmentFromVariant(x1, y1, x2, y2)
    
    'Assert:
    Assert.IsTrue sut Is Nothing, "Nothing expected on LineString initialize with invalid arguments!"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

