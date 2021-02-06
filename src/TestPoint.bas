Attribute VB_Name = "TestPoint"
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
Public Sub TestToString()
    On Error GoTo TestFail
    
    'Arrange:
    Dim sut2D As New Point
    Dim sut3D As New Point
    Dim expected2D As Variant
    Dim expected3D As Variant
    
    'Act:
    sut2D.init 3.33, -3#
    expected2D = "(3.33,-3)"
    
    sut3D.init 3.33, -3#, 1
    expected3D = "(3.33,-3,1)"
    
    'Assert:
    Assert.AreEqual expected2D, sut2D.toStringXY
    Assert.AreEqual expected3D, sut3D.toStringXYZ

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub TestIsEqual()
    On Error GoTo TestFail
    
    'Arrange:
    Dim sut2D As New Point
    Dim eP2D As New Point                          'equivalent Point
    Dim sut3D As New Point
    Dim eP3D As New Point                          'equivalent Point
    Dim x As Double
    Dim y As Double
    Dim z As Double
    
    'Act:
    x = 1 / 3
    y = Math.Sqr(2)
    z = 1
    
    sut2D.init x, y
    eP2D.init x, y
    
    sut3D.init x, y, z
    eP3D.init x, y, z

    'Assert:
    Assert.IsTrue sut2D.isEqual(eP2D), "X or Y of equivalent points are different!"
    Assert.IsTrue sut3D.isEqual(eP3D), "X or Y or Z of equivalent points are different!"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

