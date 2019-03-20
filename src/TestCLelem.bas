Attribute VB_Name = "TestCLelem"
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
Public Sub TestInit()
    On Error GoTo TestFail
    ' TODO: Include ClothoidArc test
    
    'Arrange:
    Dim ls As New LineSegment
    Dim ca As New CircularArc
    Dim elem_ls As New CLelem
    Dim elem_ca As New CLelem
    Dim startM As Double
    Dim endM As Double

    'Act:
    startM = 1
    ls.init 0, 0, 3, 4
    elem_ls.init ls, startM
    
    ca.initFromSCLD 1, 0, 0, 0, LibGeom.PI / 2, CURVE_DIR.CD_CCW
    elem_ca.init ca, startM
    
    'Assert:
    Assert.AreEqual startM, elem_ls.startM
    Assert.AreEqual startM, elem_ca.startM
    Assert.AreEqual startM + ls.length, elem_ls.endM
    Assert.AreEqual startM + ca.length, elem_ca.endM

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub TestInitInvalidGeom()
    Const ExpectedError As Long = 5
    On Error GoTo TestFail
    
    'Arrange:
    Dim sut As New CLelem
    
    'Act:
    sut.init Nothing, 0

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
Public Sub TestGeomTypeName()
    On Error GoTo TestFail
    ' TODO: Include ClothoidArc test
    'Arrange:
    Dim ls As New LineSegment
    Dim ca As New CircularArc
    Dim elem_ls As New CLelem
    Dim elem_ca As New CLelem
    Dim elem_none As New CLelem
    
    elem_ls.init ls, 0
    elem_ca.init ca, 0
    
    'Assert:
    Assert.AreEqual ConstCL.LS_NAME, elem_ls.geomTypeName
    Assert.AreEqual ConstCL.CA_NAME, elem_ca.geomTypeName
    ' TODO: Add ClothoidArc case
    Assert.AreEqual "None", elem_none.geomTypeName

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub TestSetMeasures()
    On Error GoTo TestFail

    'Arrange:
    Dim ls As New LineSegment
    Dim elem As New CLelem
    Dim startM As Double
  
    'Act:
    ls.init 0, 0, 3, 4
    elem.init ls, 0
    startM = 2#
    
    'Assert:
    Assert.AreEqual 0#, elem.startM
    Assert.AreEqual ls.length, elem.endM
    
    elem.setMeasures startM
    Assert.AreEqual startM, elem.startM
    Assert.AreEqual startM + ls.length, elem.endM

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
    Dim elem As New CLelem
    Dim p As Point
    Dim startM As Double
  
    'Act:
    startM = 2#
    ls.init 0, 0, 3, 4
    elem.init ls, startM
    Set p = elem.calcPointByMeasOffset(7, 0)
    
    'Assert:
    Assert.AreEqual 3#, p.x
    Assert.AreEqual 4#, p.y

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
    Dim elem As New CLelem
    Dim mo As MeasOffset
    Dim startM As Double
  
    'Act:
    startM = 2#
    ls.init 0, 0, 3, 4
    elem.init ls, startM
    Set mo = elem.calcMeasOffsetOfPoint(3, 4)
    
    'Assert:
    Assert.AreEqual 7#, mo.m
    Assert.AreEqual 0#, mo.o

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


