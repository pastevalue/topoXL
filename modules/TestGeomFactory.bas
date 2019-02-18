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

Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")

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
    Dim X As Variant
    Dim Y As Variant
    
    'Act:
    X = "3.33"
    Y = "-6.0"
    expected.X = 3.33
    expected.Y = -6#
    
    Set sut = GeomFactory.NewPointFromVariant(X, Y)
    
    'Assert:
    Assert.AreEqual expected.X, sut.X, "Failed to extract X Double value from Variant!"
    Assert.AreEqual expected.Y, sut.Y, "Failed to extract Y Double values from Variant!"

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
    Dim X As Variant
    Dim Y As Variant
    
    'Act:
    X = "3.33"
    Y = "-6.0abc"
    Set sut = GeomFactory.NewPointFromVariant(X, Y)

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
Public Sub TestNewMeasOffsetFromInvalidVariant()
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
