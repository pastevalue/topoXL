Attribute VB_Name = "TestCL"
'@IgnoreModule LineLabelNotUsed
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

Private cl_tst As CL ' sample test ax "L" shaped matching positive coordinate axis

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

'@TestInitialize
Public Sub TestInitialize()
    Dim ls1 As New LineSegment
    Dim ls2 As New LineSegment
    
    Dim cle1 As New CLelem
    Dim cle2 As New CLelem
    
    Dim currentStartM As Double
    
    ls1.init 0, 1, 0, 0
    ls2.init 0, 0, 1, 0
    
    Set cl_tst = New CL
    cl_tst.init "TstCenterLine"
    
    currentStartM = 0
    cle1.init ls1, currentStartM
    cl_tst.addElem cle1
    currentStartM = currentStartM + cle1.length
    
    cle2.init ls2, currentStartM
    cl_tst.addElem cle2

End Sub

'@TestCleanup
Public Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

'@TestMethod
Public Sub TestCalcPointByMeasOffsetValidMeasure()
    On Error GoTo TestFail
    
    'Arrange:
    Dim sut As Point
    Dim expected As New Point
    Dim tmpM As Double

    
    'Assert:
    tmpM = 0.5
    Set sut = cl_tst.calcPointByMeasOffset(tmpM, 0)
    expected.init 0, 0.5
    Assert.IsTrue expected.isEqual(sut)
    
    tmpM = tmpM + 1
    Set sut = cl_tst.calcPointByMeasOffset(tmpM, 0)
    expected.init 0.5, 0
    Assert.IsTrue expected.isEqual(sut)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub TestCalcPointByMeasOffsetOutMeasure()
    On Error GoTo TestFail
    
    'Arrange:
    Dim sut As Point
    Dim expected As New Point
    Dim tmpM As Double
    
    'Assert:
    tmpM = -1
    Set sut = cl_tst.calcPointByMeasOffset(tmpM, 0)
    Assert.IsNothing sut
    
    tmpM = 3
    Set sut = cl_tst.calcPointByMeasOffset(tmpM, 0)
    Assert.IsNothing sut

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub TestCalcMeasOffsetOfPointValidCoo()
    On Error GoTo TestFail
    
    'Arrange:
    Dim sut As MeasOffset
    Dim expected As New MeasOffset
    
    'Assert:
    Set sut = cl_tst.calcMeasOffsetOfPoint(0, 1)
    expected.init 0, 0
    Assert.IsTrue expected.isEqual(sut)
    
    Set sut = cl_tst.calcMeasOffsetOfPoint(0, 0)
    expected.init 1, 0
    Assert.IsTrue expected.isEqual(sut)
    
    Set sut = cl_tst.calcMeasOffsetOfPoint(1, 0)
    expected.init 2, 0
    Assert.IsTrue expected.isEqual(sut)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub TestCalcMeasOffsetOfPointOutCoo()
    Dim sut As MeasOffset
    
    'Assert:
    Set sut = cl_tst.calcMeasOffsetOfPoint(-1, 2)
    Assert.IsNothing sut
    
    Set sut = cl_tst.calcMeasOffsetOfPoint(-1, -1)
    Assert.IsNothing sut
    
    Set sut = cl_tst.calcMeasOffsetOfPoint(2, -1)
    Assert.IsNothing sut

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
