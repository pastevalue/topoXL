Attribute VB_Name = "TestFactoryGeom"
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
    
    Set sut = FactoryGeom.newPntVar(x, y)
    
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
    Set sut = FactoryGeom.newPntVar(x, y)

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
    
    Set sut = FactoryGeom.NewMOvar(m, o)
    
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
    
    Set sut = FactoryGeom.NewMOvar(m, o)
    
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
    
    Set sut = FactoryGeom.newLnSegVar(x1, y1, x2, y2)
    
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
    
    Set sut = FactoryGeom.newLnSegVar(x1, y1, x2, y2)
    
    'Assert:
    Assert.IsTrue sut Is Nothing, "Nothing expected on LineString initialized with invalid arguments!"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub TestNewLnSegValidColl()
    On Error GoTo TestFail
    
    'Arrange:
    Dim sut As LineSegment
    Dim expected As New LineSegment
    Dim coll As New Collection
    Dim sX As Variant
    Dim sY As Variant
    Dim eX As Variant
    Dim eY As Variant
    
    'Act:
    sX = "3.33"
    sY = "-6.0"
    eX = "3.0"
    eY = "6.66"
    expected.init sX, sY, eX, eY
    
    coll.add ConstCL.LS_NAME, ConstCL.GEOM_TYPE
    coll.add ConstCL.LS_INIT_SE, ConstCL.GEOM_INIT_TYPE
    coll.add sX, ConstCL.LS_M_START_X
    coll.add sY, ConstCL.LS_M_START_Y
    coll.add eX, ConstCL.LS_M_END_X
    coll.add eY, ConstCL.LS_M_END_Y
    Set sut = FactoryGeom.newLnSegColl(coll)
    
    'Assert:
    Assert.IsTrue sut.isEqual(expected)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub TestNewLnSegInvalidColl()
    On Error GoTo TestFail
    
    'Arrange:
    Dim sut As LineSegment
    Dim coll As New Collection
    Dim sX As Variant
    Dim sY As Variant
    Dim eX As Variant
    Dim eY As Variant
    
    'Act:
    sX = "3.33"
    sY = "-6.0"
    eX = "3.0"
    eY = "6.66"
    
    coll.add sX, ConstCL.LS_M_START_X
    coll.add sY, ConstCL.LS_M_START_Y
    coll.add eX, ConstCL.LS_M_END_X
    coll.add eY, ConstCL.LS_M_END_Y
    coll.add ConstCL.LS_INIT_SE, ConstCL.GEOM_INIT_TYPE
    coll.add "Wrong Geom Type", ConstCL.GEOM_TYPE
    Set sut = FactoryGeom.newLnSegColl(coll)
    Assert.IsTrue sut Is Nothing, "LineSegment initialized with wrong GeomType!"
        
    coll.Remove ConstCL.GEOM_TYPE
    coll.Remove ConstCL.GEOM_INIT_TYPE
    coll.add ConstCL.LS_NAME, ConstCL.GEOM_TYPE
    coll.add "Wrong init type", ConstCL.GEOM_INIT_TYPE
    Set sut = FactoryGeom.newLnSegColl(coll)
    Assert.IsTrue sut Is Nothing, "LineSegment initialized with wrong Init Type GeomType!"
    
    coll.Remove ConstCL.LS_M_START_X
    coll.Remove ConstCL.GEOM_INIT_TYPE
    coll.add "1abc", ConstCL.LS_M_START_X
    coll.add ConstCL.LS_INIT_SE, ConstCL.GEOM_INIT_TYPE
    Set sut = FactoryGeom.newLnSegColl(coll)
    Assert.IsTrue sut Is Nothing, "LineSegment initialized wrong value for member GeomType!"
    
    coll.Remove ConstCL.LS_M_START_X
    Set sut = FactoryGeom.newLnSegColl(coll)
    Assert.IsTrue sut Is Nothing, "LineSegment initialized without a value for a member key GeomType!"

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
    
    Set sut = FactoryGeom.newCircArcSCLDvar(sX, sY, cX, cY, l, dir)
    
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
    
    Set sut = FactoryGeom.newCircArcSCLDvar("1abc", 1, 0, 0, 1, 1)
    Assert.IsTrue sut Is Nothing, "Nothing expected on CircularArc initialized with invalid arguments!"
    
    Set sut = FactoryGeom.newCircArcSCLDvar("1", "1abc", 0, 0, 1, 1)
    Assert.IsTrue sut Is Nothing, "Nothing expected on CircularArc initialized with invalid arguments!"
    
    Set sut = FactoryGeom.newCircArcSCLDvar("1", "1", "0abc", 0, 1, 1)
    Assert.IsTrue sut Is Nothing, "Nothing expected on CircularArc initialized with invalid arguments!"
    
    Set sut = FactoryGeom.newCircArcSCLDvar("1", "1", "0", "0abc", 1, 1)
    Assert.IsTrue sut Is Nothing, "Nothing expected on CircularArc initialized with invalid arguments!"
    
    Set sut = FactoryGeom.newCircArcSCLDvar("1", "1", "0", "0", "1abc", 1)
    Assert.IsTrue sut Is Nothing, "Nothing expected on CircularArc initialized with invalid arguments!"
    
    Set sut = FactoryGeom.newCircArcSCLDvar("1", "1", "0", "0", "1", -2)
    Assert.IsTrue sut Is Nothing, "Nothing expected on CircularArc initialized with invalid arguments!"
    
    Set sut = FactoryGeom.newCircArcSCLDvar("1", "1", "0", "0", "1", 0)
    Assert.IsTrue sut Is Nothing, "Nothing expected on CircularArc initialized with invalid arguments!"
    
    Set sut = FactoryGeom.newCircArcSCLDvar("1", "1", "0", "0", "1", 2)
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
    
    Set sut = FactoryGeom.newCircArcSCLD(0, 0, 0, 0, 1, CD_CW)
    Assert.IsTrue sut Is Nothing, "Nothing expected on CircularArc initialized with invalid arguments!"
    
    Set sut = FactoryGeom.newCircArcSCLD(1, 0, 0, 0, -1, CD_CW)
    Assert.IsTrue sut Is Nothing, "Nothing expected on CircularArc initialized with invalid arguments!"
    
    Set sut = FactoryGeom.newCircArcSCLD(1, 0, 0, 0, 1, CD_NONE)
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
    dir = "1"
    
    expected.initFromSERD 0, 1, 1, 0, 1, CD_CW
    
    Set sut = FactoryGeom.newCircArcSERDvar(sX, sY, eX, eY, r, dir)
    
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
    
    Set sut = FactoryGeom.newCircArcSERDvar("1abc", 1, 0, 0, 1, 1)
    Assert.IsTrue sut Is Nothing, "Nothing expected on CircularArc initialized with invalid arguments!"
    
    Set sut = FactoryGeom.newCircArcSERDvar("1", "1abc", 0, 0, 1, 1)
    Assert.IsTrue sut Is Nothing, "Nothing expected on CircularArc initialized with invalid arguments!"
    
    Set sut = FactoryGeom.newCircArcSERDvar("1", "1", "0abc", 0, 1, 1)
    Assert.IsTrue sut Is Nothing, "Nothing expected on CircularArc initialized with invalid arguments!"
    
    Set sut = FactoryGeom.newCircArcSERDvar("1", "1", "0", "0abc", 1, 1)
    Assert.IsTrue sut Is Nothing, "Nothing expected on CircularArc initialized with invalid arguments!"
    
    Set sut = FactoryGeom.newCircArcSERDvar("1", "1", "0", "0", "1abc", 1)
    Assert.IsTrue sut Is Nothing, "Nothing expected on CircularArc initialized with invalid arguments!"
    
    Set sut = FactoryGeom.newCircArcSERDvar("1", "1", "0", "0", "1", -2)
    Assert.IsTrue sut Is Nothing, "Nothing expected on CircularArc initialized with invalid arguments!"
    
    Set sut = FactoryGeom.newCircArcSERDvar("1", "1", "0", "0", "1", 0)
    Assert.IsTrue sut Is Nothing, "Nothing expected on CircularArc initialized with invalid arguments!"
    
    Set sut = FactoryGeom.newCircArcSCLDvar("1", "1", "0", "0", "1", 2)
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
    
    Set sut = FactoryGeom.newCircArcSERDvar(0, 0, 0, 0, 1, CD_CCW)
    Assert.IsTrue sut Is Nothing, "Nothing expected on CircularArc initialized with invalid arguments!"
    
    Set sut = FactoryGeom.newCircArcSERDvar(0, 1, 1, 0, -1, CD_CCW)
    Assert.IsTrue sut Is Nothing, "Nothing expected on CircularArc initialized with invalid arguments!"
    
    Set sut = FactoryGeom.newCircArcSERDvar(0, 1, 1, 0, 1, CD_NONE)
    Assert.IsTrue sut Is Nothing, "Nothing expected on CircularArc initialized with invalid arguments!"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub TestNewCircArcSCLDValidColl()
    On Error GoTo TestFail
    
    'Arrange:
    Dim sut As CircularArc
    Dim expected As New CircularArc
    Dim coll As New Collection
    Dim sX As Variant
    Dim sY As Variant
    Dim cX As Variant
    Dim cY As Variant
    Dim l As Variant
    Dim dir As Variant
    
    'Act:
    sX = "1"
    sY = "0"
    cX = "0.0"
    cY = "0.0"
    l = LibGeom.PI / 2
    dir = CURVE_DIR.CD_CCW
    expected.initFromSCLD sX, sY, cX, cY, l, dir
    
    coll.add ConstCL.CA_NAME, ConstCL.GEOM_TYPE
    coll.add ConstCL.CA_INIT_SCLD, ConstCL.GEOM_INIT_TYPE
    coll.add sX, ConstCL.CA_M_START_X
    coll.add sY, ConstCL.CA_M_START_Y
    coll.add cX, ConstCL.CA_M_CENTER_X
    coll.add cY, ConstCL.CA_M_CENTER_Y
    coll.add l, ConstCL.CA_M_LENGTH
    coll.add dir, ConstCL.CA_M_CURVE_DIR
    Set sut = FactoryGeom.newCircArcColl(coll)
    
    'Assert:
    Assert.IsTrue sut.isEqual(expected)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub TestNewCircArcSCLDInvalidColl()
    On Error GoTo TestFail
    
    'Arrange:
    Dim sut As CircularArc
    Dim coll As New Collection
    Dim sX As Variant
    Dim sY As Variant
    Dim cX As Variant
    Dim cY As Variant
    Dim l As Variant
    Dim dir As Variant
    
    'Act:
    sX = "1"
    sY = "0"
    cX = "0.0"
    cY = "0.0"
    l = LibGeom.PI / 2
    dir = CURVE_DIR.CD_CCW
    
    coll.add sX, ConstCL.CA_M_START_X
    coll.add sY, ConstCL.CA_M_START_Y
    coll.add cX, ConstCL.CA_M_CENTER_X
    coll.add cY, ConstCL.CA_M_CENTER_Y
    coll.add l, ConstCL.CA_M_LENGTH
    coll.add dir, ConstCL.CA_M_CURVE_DIR
    coll.add "Wrong Geometry Type", ConstCL.GEOM_TYPE
    coll.add ConstCL.CA_INIT_SCLD, ConstCL.GEOM_INIT_TYPE
    Set sut = FactoryGeom.newCircArcColl(coll)
    Assert.IsTrue sut Is Nothing, "CircularArc initialized with wrong GeomType!"
   
    coll.Remove ConstCL.GEOM_TYPE
    coll.Remove ConstCL.GEOM_INIT_TYPE
    coll.add ConstCL.CA_NAME, ConstCL.GEOM_TYPE
    coll.add "Wrong init type", ConstCL.GEOM_INIT_TYPE
    Set sut = FactoryGeom.newCircArcColl(coll)
    Assert.IsTrue sut Is Nothing, "CircularArc initialized with wrong Init Type GeomType!"
    
    coll.Remove ConstCL.CA_M_START_X
    coll.Remove ConstCL.GEOM_INIT_TYPE
    coll.add "1abc", ConstCL.CA_M_START_X
    coll.add ConstCL.CA_INIT_SCLD, ConstCL.GEOM_INIT_TYPE
    Set sut = FactoryGeom.newCircArcColl(coll)
    Assert.IsTrue sut Is Nothing, "CircularArc initialized wrong value for member GeomType!"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub TestNewCircArcSERDValidColl()
    On Error GoTo TestFail
    
    'Arrange:
    Dim sut As CircularArc
    Dim expected As New CircularArc
    Dim coll As New Collection
    Dim sX As Variant
    Dim sY As Variant
    Dim eX As Variant
    Dim eY As Variant
    Dim r As Variant
    Dim dir As Variant
    
    'Act:
    sX = "1"
    sY = "0"
    eX = "0.0"
    eY = "1"
    r = 1
    dir = CURVE_DIR.CD_CCW
    expected.initFromSERD sX, sY, eX, eY, r, dir
    
    coll.add ConstCL.CA_NAME, ConstCL.GEOM_TYPE
    coll.add ConstCL.CA_INIT_SERD, ConstCL.GEOM_INIT_TYPE
    coll.add sX, ConstCL.CA_M_START_X
    coll.add sY, ConstCL.CA_M_START_Y
    coll.add eX, ConstCL.CA_M_END_X
    coll.add eY, ConstCL.CA_M_END_Y
    coll.add r, ConstCL.CA_M_RADIUS
    coll.add dir, ConstCL.CA_M_CURVE_DIR
    Set sut = FactoryGeom.newCircArcColl(coll)
    
    'Assert:
    Assert.IsTrue sut.isEqual(expected)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub TestNewCircArcSERDInvalidColl()
    On Error GoTo TestFail
    
    'Arrange:
    Dim sut As CircularArc
    Dim coll As New Collection
    Dim sX As Variant
    Dim sY As Variant
    Dim eX As Variant
    Dim eY As Variant
    Dim r As Variant
    Dim dir As Variant
    
    'Act:
    sX = "1"
    sY = "0"
    eX = "0.0"
    eY = "1.0"
    r = 1
    dir = CURVE_DIR.CD_CCW
    
    coll.add sX, ConstCL.CA_M_START_X
    coll.add sY, ConstCL.CA_M_START_Y
    coll.add eX, ConstCL.CA_M_END_X
    coll.add eY, ConstCL.CA_M_END_Y
    coll.add r, ConstCL.CA_M_LENGTH
    coll.add dir, ConstCL.CA_M_CURVE_DIR
    coll.add "Wrong Geometry Type", ConstCL.GEOM_TYPE
    coll.add ConstCL.CA_INIT_SERD, ConstCL.GEOM_INIT_TYPE
    Set sut = FactoryGeom.newCircArcColl(coll)
    Assert.IsTrue sut Is Nothing, "CircularArc initialized with wrong GeomType!"
   
    coll.Remove ConstCL.GEOM_TYPE
    coll.Remove ConstCL.GEOM_INIT_TYPE
    coll.add ConstCL.CA_NAME, ConstCL.GEOM_TYPE
    coll.add "Wrong init type", ConstCL.GEOM_INIT_TYPE
    Set sut = FactoryGeom.newCircArcColl(coll)
    Assert.IsTrue sut Is Nothing, "CircularArc initialized with wrong Init Type GeomType!"
    
    coll.Remove ConstCL.CA_M_START_X
    coll.Remove ConstCL.GEOM_INIT_TYPE
    coll.add "1abc", ConstCL.CA_M_START_X
    coll.add ConstCL.CA_INIT_SERD, ConstCL.GEOM_INIT_TYPE
    Set sut = FactoryGeom.newCircArcColl(coll)
    Assert.IsTrue sut Is Nothing, "CircularArc initialized wrong value for member GeomType!"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

