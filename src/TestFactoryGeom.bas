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
Public Sub TestNewPntValidVariant()
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
    
    Set sut = GeomFactory.newPntVar(x, y)
    
    'Assert:
    Assert.AreEqual expected.x, sut.x, "Failed to extract X Double value from Variant!"
    Assert.AreEqual expected.y, sut.y, "Failed to extract Y Double values from Variant!"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub TestNewPntInvalidVariant()
    On Error GoTo TestFail
    
    'Arrange:
    Dim sut As Point
    Dim x As Variant
    Dim y As Variant
    
    'Act:
    x = "3.33"
    y = "-6.0abc"
    Set sut = GeomFactory.newPntVar(x, y)

    'Assert:
    Assert.IsTrue sut Is Nothing, "Nothing expected on Point initialized with invalid arguments!"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub TestNewMeasOffsetValidVariant()
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
    
    Set sut = GeomFactory.NewMOvar(m, o)
    
    'Assert:
    Assert.AreEqual expected.m, sut.m, "Failed to extract Measure Double value from Variant!"
    Assert.AreEqual expected.o, sut.o, "Failed to extract Offset Double values from Variant!"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub TestNewMeasOffsetInvalidVariant()
    On Error GoTo TestFail
    
    'Arrange:
    Dim sut As MeasOffset
    Dim m As Variant
    Dim o As Variant
    
    'Act:
    m = "3.33"
    o = "-6.0abc"
    
    Set sut = GeomFactory.NewMOvar(m, o)
    
    'Assert:
    Assert.IsTrue sut Is Nothing, "Nothing expected on MeasOffset initialized with invalid arguments!"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub TestNewLnSegValidVariant()
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
    expected.init 3.33, -6, 3, 6.66
    
    Set sut = GeomFactory.newLnSegVar(x1, y1, x2, y2)
    
    'Assert:
    Assert.AreEqual expected.sX, sut.sX
    Assert.AreEqual expected.sY, sut.sY
    Assert.AreEqual expected.eX, sut.eX
    Assert.AreEqual expected.eY, sut.eY

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub TestNewLnSegInvalidVariant()
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
    expected.init 3.33, -6, 3, 6.66
    
    Set sut = GeomFactory.newLnSegVar(x1, y1, x2, y2)
    
    'Assert:
    Assert.IsTrue sut Is Nothing, "Nothing expected on LineString initialized with invalid arguments!"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub TestNewCircArcSCLDvalidVariant()
    On Error GoTo TestFail
    
    'Arrange:
    Dim sut As CircularArc
    Dim expected As New CircularArc
    Dim sX As Variant
    Dim sY As Variant
    Dim cX As Variant
    Dim cY As Variant
    Dim l As Variant
    Dim dir As Variant
    
    'Act:
    sX = "0"
    sY = "1.0"
    cX = "0.0"
    cY = "0"
    l = "1"
    dir = 1
    
    expected.initFromSCLD 0, 1, 0, 0, 1, CD_CW
    
    Set sut = GeomFactory.newCircArcSCLDvar(sX, sY, cX, cY, l, dir)
    
    'Assert:
    Assert.AreEqual expected.sX, sut.sX
    Assert.AreEqual expected.sY, sut.sY
    Assert.AreEqual expected.cX, sut.cX
    Assert.AreEqual expected.cY, sut.cY
    Assert.AreEqual expected.length, sut.length
    Assert.AreEqual expected.curveDirection, sut.curveDirection

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub TestNewCircArcSCLDinvalidVariant()
    On Error GoTo TestFail
    
    Dim sut As CircularArc
    
    Set sut = GeomFactory.newCircArcSCLDvar("1abc", 1, 0, 0, 1, 1)
    Assert.IsTrue sut Is Nothing, "Nothing expected on CircularArc initialized with invalid arguments!"
    
    Set sut = GeomFactory.newCircArcSCLDvar("1", "1abc", 0, 0, 1, 1)
    Assert.IsTrue sut Is Nothing, "Nothing expected on CircularArc initialized with invalid arguments!"
    
    Set sut = GeomFactory.newCircArcSCLDvar("1", "1", "0abc", 0, 1, 1)
    Assert.IsTrue sut Is Nothing, "Nothing expected on CircularArc initialized with invalid arguments!"
    
    Set sut = GeomFactory.newCircArcSCLDvar("1", "1", "0", "0abc", 1, 1)
    Assert.IsTrue sut Is Nothing, "Nothing expected on CircularArc initialized with invalid arguments!"
    
    Set sut = GeomFactory.newCircArcSCLDvar("1", "1", "0", "0", "1abc", 1)
    Assert.IsTrue sut Is Nothing, "Nothing expected on CircularArc initialized with invalid arguments!"
    
    Set sut = GeomFactory.newCircArcSCLDvar("1", "1", "0", "0", "1", -2)
    Assert.IsTrue sut Is Nothing, "Nothing expected on CircularArc initialized with invalid arguments!"
    
    Set sut = GeomFactory.newCircArcSCLDvar("1", "1", "0", "0", "1", 0)
    Assert.IsTrue sut Is Nothing, "Nothing expected on CircularArc initialized with invalid arguments!"
    
    Set sut = GeomFactory.newCircArcSCLDvar("1", "1", "0", "0", "1", 2)
    Assert.IsTrue sut Is Nothing, "Nothing expected on CircularArc initialized with invalid arguments!"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub TestNewCircArcSCLDinvalidParams()
    On Error GoTo TestFail
    
    Dim sut As CircularArc
    
    Set sut = GeomFactory.newCircArcSCLD(0, 0, 0, 0, 1, CD_CW)
    Assert.IsTrue sut Is Nothing, "Nothing expected on CircularArc initialized with invalid arguments!"
    
    Set sut = GeomFactory.newCircArcSCLD(1, 0, 0, 0, -1, CD_CW)
    Assert.IsTrue sut Is Nothing, "Nothing expected on CircularArc initialized with invalid arguments!"
    
    Set sut = GeomFactory.newCircArcSCLD(1, 0, 0, 0, 1, CD_NONE)
    Assert.IsTrue sut Is Nothing, "Nothing expected on CircularArc initialized with invalid arguments!"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub TestNewCircArcSERDvalidVariant()
    On Error GoTo TestFail
    
    'Arrange:
    Dim sut As CircularArc
    Dim expected As New CircularArc
    Dim sX As Variant
    Dim sY As Variant
    Dim eX As Variant
    Dim eY As Variant
    Dim r As Variant
    Dim dir As Variant
    
    'Act:
    sX = "0"
    sY = "1.0"
    eX = "1.0"
    eY = "0"
    r = "1"
    dir = 1
    
    expected.initFromSERD 0, 1, 1, 0, 1, CD_CW
    
    Set sut = GeomFactory.newCircArcSERDvar(sX, sY, eX, eY, r, dir)
    
    'Assert:
    Assert.AreEqual expected.sX, sut.sX
    Assert.AreEqual expected.sY, sut.sY
    Assert.AreEqual expected.eX, sut.eX
    Assert.AreEqual expected.eY, sut.eY
    Assert.AreEqual expected.length, sut.length
    Assert.AreEqual expected.curveDirection, sut.curveDirection

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub TestNewCircArcSERDinvalidVariant()
    On Error GoTo TestFail
    
    Dim sut As CircularArc
    
    Set sut = GeomFactory.newCircArcSERDvar("1abc", 1, 0, 0, 1, 1)
    Assert.IsTrue sut Is Nothing, "Nothing expected on CircularArc initialized with invalid arguments!"
    
    Set sut = GeomFactory.newCircArcSERDvar("1", "1abc", 0, 0, 1, 1)
    Assert.IsTrue sut Is Nothing, "Nothing expected on CircularArc initialized with invalid arguments!"
    
    Set sut = GeomFactory.newCircArcSERDvar("1", "1", "0abc", 0, 1, 1)
    Assert.IsTrue sut Is Nothing, "Nothing expected on CircularArc initialized with invalid arguments!"
    
    Set sut = GeomFactory.newCircArcSERDvar("1", "1", "0", "0abc", 1, 1)
    Assert.IsTrue sut Is Nothing, "Nothing expected on CircularArc initialized with invalid arguments!"
    
    Set sut = GeomFactory.newCircArcSERDvar("1", "1", "0", "0", "1abc", 1)
    Assert.IsTrue sut Is Nothing, "Nothing expected on CircularArc initialized with invalid arguments!"
    
    Set sut = GeomFactory.newCircArcSERDvar("1", "1", "0", "0", "1", -2)
    Assert.IsTrue sut Is Nothing, "Nothing expected on CircularArc initialized with invalid arguments!"
    
    Set sut = GeomFactory.newCircArcSERDvar("1", "1", "0", "0", "1", 0)
    Assert.IsTrue sut Is Nothing, "Nothing expected on CircularArc initialized with invalid arguments!"
    
    Set sut = GeomFactory.newCircArcSCLDvar("1", "1", "0", "0", "1", 2)
    Assert.IsTrue sut Is Nothing, "Nothing expected on CircularArc initialized with invalid arguments!"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub TestNewCircArcSERDinvalidParams()
    On Error GoTo TestFail
    
    Dim sut As CircularArc
    
    Set sut = GeomFactory.newCircArcSERDvar(0, 0, 0, 0, 1, CD_CCW)
    Assert.IsTrue sut Is Nothing, "Nothing expected on CircularArc initialized with invalid arguments!"
    
    Set sut = GeomFactory.newCircArcSERDvar(0, 1, 1, 0, -1, CD_CCW)
    Assert.IsTrue sut Is Nothing, "Nothing expected on CircularArc initialized with invalid arguments!"
    
    Set sut = GeomFactory.newCircArcSERDvar(0, 1, 1, 0, 1, CD_NONE)
    Assert.IsTrue sut Is Nothing, "Nothing expected on CircularArc initialized with invalid arguments!"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
