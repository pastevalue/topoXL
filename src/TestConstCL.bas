Attribute VB_Name = "TestConstCL"
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
Public Sub TestCurveDirFromVariant()
    On Error GoTo TestFail

    'Assert:
    Assert.AreEqual CURVE_DIR.CD_CCW, ConstCL.curveDirFromVariant(-1), "Failed number to curve direction!"
    Assert.AreEqual CURVE_DIR.CD_CCW, ConstCL.curveDirFromVariant(-1#), "Failed number to curve direction!"
    Assert.AreEqual CURVE_DIR.CD_CCW, ConstCL.curveDirFromVariant("-1"), "Failed number as string to curve direction!"
    Assert.AreEqual CURVE_DIR.CD_CCW, ConstCL.curveDirFromVariant("CCW"), "Failed string to curve direction!"
    Assert.AreEqual CURVE_DIR.CD_CW, ConstCL.curveDirFromVariant("CW"), "Failed string to curve direction!"
    
    Assert.AreEqual CURVE_DIR.CD_NONE, ConstCL.curveDirFromVariant(-2), "Failed CD_NONE on value out of named values!"
    Assert.AreEqual CURVE_DIR.CD_NONE, ConstCL.curveDirFromVariant(-1.1), "Failed CD_NONE on value out of named values!"
    Assert.AreEqual CURVE_DIR.CD_NONE, ConstCL.curveDirFromVariant("-2"), "Failed CD_NONE on string value out of named values!"
    Assert.AreEqual CURVE_DIR.CD_NONE, ConstCL.curveDirFromVariant("abc"), "Failed CD_NONE on string not assign for a named value!"
    Dim my_obj As New Collection
    Assert.AreEqual CURVE_DIR.CD_NONE, ConstCL.curveDirFromVariant(my_obj), "Failed CD_NONE on object!"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub TestCurveDirToString()
    On Error GoTo TestFail
    
    'Assert:
    Assert.AreEqual "none", ConstCL.curveDirToString(-2), "Failed none string on valueout of curve direction enum values!"
    Assert.AreEqual "none", ConstCL.curveDirToString(0), "Failed none string on valueout of curve direction enum values!"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
