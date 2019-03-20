Attribute VB_Name = "TestCircularArc"
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
Public Sub TestInitFromSCLDinvalidSC()
    Const ExpectedError As Long = 5
    On Error GoTo TestFail
    
    'Arrange:
    Dim sut As New CircularArc
    
    'Act:
    sut.initFromSCLD 0, 0, 0, 0, PI / 6, CURVE_DIR.CD_CW

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
Public Sub TestInitFromSCLDinvalidL()
    Const ExpectedError As Long = 5
    On Error GoTo TestFail
    
    'Arrange:
    Dim sut As New CircularArc
    
    'Act:
    sut.initFromSCLD 0, 0, 1, 0, 0, CURVE_DIR.CD_CW

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
Public Sub TestInitFromSCLDinvalidD()
    Const ExpectedError As Long = 5
    On Error GoTo TestFail
    
    'Arrange:
    Dim sut As New CircularArc
    
    'Act:
    sut.initFromSCLD 0, 0, 1, 0, 1, CURVE_DIR.CD_NONE

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
Public Sub TestInitFromSCLD()
    On Error GoTo TestFail
    
    'Arrange:
    Dim sut As New CircularArc
    Dim epsilon As Double
    
    'Act:
    epsilon = 0.000000000000001                  '1E-15
    

    ' First quadrant, CW, PI/6 length
    sut.initFromSCLD 0, 1, 0, 0, PI / 6, CURVE_DIR.CD_CW
    
    'Assert:
    Assert.AreEqual 1#, sut.radius, "Radius must be 1!"
    Assert.IsTrue LibMath.areDoublesEqual(PI / 2, sut.sTheta, epsilon), "Theta of center to start point line must be PI/2!"
    Assert.IsTrue LibMath.areDoublesEqual(PI / 3, sut.eTheta, epsilon), "Theta of center to end point line must be PI/3!"
    Assert.IsTrue LibMath.areDoublesEqual(Cos(PI / 3), sut.eX, epsilon), "X of end point must be" & Cos(PI / 3) & "!"
    Assert.IsTrue LibMath.areDoublesEqual(Sin(PI / 3), sut.eY, epsilon), "Y of end point must be" & Sin(PI / 3) & "!"
    
    ' Second quadrant, CW, PI/6 length
    sut.initFromSCLD -1, 0, 0, 0, PI / 6, CURVE_DIR.CD_CW

    'Assert:
    Assert.AreEqual 1#, sut.radius, "Radius must be 1!"
    Assert.IsTrue LibMath.areDoublesEqual(PI, sut.sTheta, epsilon), "Theta of center to start point line must be PI!"
    Assert.IsTrue LibMath.areDoublesEqual(5 * PI / 6, sut.eTheta, epsilon), "Theta of center to end point line must be 5*PI/6!"
    Assert.IsTrue LibMath.areDoublesEqual(Cos(5 * PI / 6), sut.eX, epsilon), "X of end point must be" & Cos(5 * PI / 6) & "!"
    Assert.IsTrue LibMath.areDoublesEqual(Sin(5 * PI / 6), sut.eY, epsilon), "Y of end point must be" & Sin(5 * PI / 6) & "!"
    
    ' Third quadrant, CW, PI/6 length
    sut.initFromSCLD 0, -1, 0, 0, PI / 6, CURVE_DIR.CD_CW
    
    'Assert:
    Assert.AreEqual 1#, sut.radius, "Radius must be 1!"
    Assert.IsTrue LibMath.areDoublesEqual(-PI / 2, sut.sTheta, epsilon), "Theta of center to start point line must be -PI/2!"
    Assert.IsTrue LibMath.areDoublesEqual(-2 * PI / 3, sut.eTheta, epsilon), "Theta of center to end point line must be -2*PI/3!"
    Assert.IsTrue LibMath.areDoublesEqual(Cos(-2 * PI / 3), sut.eX, epsilon), "X of end point must be" & Cos(-2 * PI / 3) & "!"
    Assert.IsTrue LibMath.areDoublesEqual(Sin(-2 * PI / 3), sut.eY, epsilon), "Y of end point must be" & Sin(-2 * PI / 3) & "!"
    
    ' Fourth quadrant, CW, PI/6 length
    sut.initFromSCLD 1, 0, 0, 0, PI / 6, CURVE_DIR.CD_CW
    
    'Assert:
    Assert.AreEqual 1#, sut.radius, "Radius must be 1!"
    Assert.IsTrue LibMath.areDoublesEqual(0, sut.sTheta, epsilon), "Theta of center to start point line must be 0.0!"
    Assert.IsTrue LibMath.areDoublesEqual(-PI / 6, sut.eTheta, epsilon), "Theta of center to end point line must be -PI/6!"
    Assert.IsTrue LibMath.areDoublesEqual(Cos(-PI / 6), sut.eX, epsilon), "X of end point must be" & Cos(-PI / 6) & "!"
    Assert.IsTrue LibMath.areDoublesEqual(Sin(-PI / 6), sut.eY, epsilon), "Y of end point must be" & Sin(-PI / 6) & "!"
    
    
    ' First quadrant, CCW, PI/6 length
    sut.initFromSCLD 1, 0, 0, 0, PI / 6, CURVE_DIR.CD_CCW
    
    'Assert:
    Assert.AreEqual 1#, sut.radius, "Radius must be 1!"
    Assert.IsTrue LibMath.areDoublesEqual(0, sut.sTheta, epsilon), "Theta of center to start point line must be 0!"
    Assert.IsTrue LibMath.areDoublesEqual(PI / 6, sut.eTheta, epsilon), "Theta of center to end point line must be PI/6!"
    Assert.IsTrue LibMath.areDoublesEqual(Cos(PI / 6), sut.eX, epsilon), "X of end point must be" & Cos(PI / 6) & "!"
    Assert.IsTrue LibMath.areDoublesEqual(Sin(PI / 6), sut.eY, epsilon), "Y of end point must be" & Sin(PI / 6) & "!"
    
    ' Second quadrant, CCW, PI/6 length
    sut.initFromSCLD 0, 1, 0, 0, PI / 6, CURVE_DIR.CD_CCW
    
    'Assert:
    Assert.AreEqual 1#, sut.radius, "Radius must be 1!"
    Assert.IsTrue LibMath.areDoublesEqual(PI / 2, sut.sTheta, epsilon), "Theta of center to start point line must be PI/2!"
    Assert.IsTrue LibMath.areDoublesEqual(2 * PI / 3, sut.eTheta, epsilon), "Theta of center to end point line must be 2*PI/3!"
    Assert.IsTrue LibMath.areDoublesEqual(Cos(2 * PI / 3), sut.eX, epsilon), "X of end point must be" & Cos(2 * PI / 3) & "!"
    Assert.IsTrue LibMath.areDoublesEqual(Sin(2 * PI / 3), sut.eY, epsilon), "Y of end point must be" & Sin(2 * PI / 3) & "!"
    
    ' Third quadrant, CCW, PI/6 length
    sut.initFromSCLD -1, 0, 0, 0, PI / 6, CURVE_DIR.CD_CCW
    
    'Assert:
    Assert.AreEqual 1#, sut.radius, "Radius must be 1!"
    Assert.IsTrue LibMath.areDoublesEqual(PI, sut.sTheta, epsilon), "Theta of center to start point line must be PI!"
    Assert.IsTrue LibMath.areDoublesEqual(-5 * PI / 6, sut.eTheta, epsilon), "Theta of center to end point line must be -5*PI/6!"
    Assert.IsTrue LibMath.areDoublesEqual(Cos(-5 * PI / 6), sut.eX, epsilon), "X of end point must be" & Cos(-5 * PI / 6) & "!"
    Assert.IsTrue LibMath.areDoublesEqual(Sin(-5 * PI / 6), sut.eY, epsilon), "Y of end point must be" & Sin(-5 * PI / 6) & "!"
    
    ' Fourth quadrant, CCW, PI/6 length
    sut.initFromSCLD 0, -1, 0, 0, PI / 6, CURVE_DIR.CD_CCW
    
    'Assert:
    Assert.AreEqual 1#, sut.radius, "Radius must be 1!"
    Assert.IsTrue LibMath.areDoublesEqual(-PI / 2, sut.sTheta, epsilon), "Theta of center to start point line must be -PI/2!"
    Assert.IsTrue LibMath.areDoublesEqual(-PI / 3, sut.eTheta, epsilon), "Theta of center to end point line must be -PI/3!"
    Assert.IsTrue LibMath.areDoublesEqual(Cos(-PI / 3), sut.eX, epsilon), "X of end point must be" & Cos(-PI / 3) & "!"
    Assert.IsTrue LibMath.areDoublesEqual(Sin(-PI / 3), sut.eY, epsilon), "Y of end point must be" & Sin(-PI / 3) & "!"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub TestInitFromSERDinvalidSE()
    Const ExpectedError As Long = 5
    On Error GoTo TestFail
    
    'Arrange:
    Dim sut As New CircularArc
    
    'Act:
    sut.initFromSERD 0, 0, 0, 0, 1, CURVE_DIR.CD_CW

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
Public Sub TestInitFromSERDinvalidR()
    Const ExpectedError As Long = 5
    On Error GoTo TestFail
    
    'Arrange:
    Dim sut As New CircularArc
    
    'Act:
    sut.initFromSERD 0, 0, 1, 0, 0, CURVE_DIR.CD_CW

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
Public Sub TestInitFromSERDinvalidD()
    Const ExpectedError As Long = 5
    On Error GoTo TestFail
    
    'Arrange:
    Dim sut As New CircularArc
    
    'Act:
    sut.initFromSERD 0, 0, 1, 0, 1, CURVE_DIR.CD_NONE

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
Public Sub TestInitFromSERD()
    On Error GoTo TestFail
    
    'Arrange:
    Dim sut As New CircularArc
    Dim epsilon As Double
    Dim sqr3 As Double
    Dim sqr2 As Double
    
    'Act:
    epsilon = 0.00000000000001                  '1E-14
    sqr3 = Sqr(3)
    sqr2 = Sqr(2)
    
    ' First quadrant, CW, PI/6 length
    sut.initFromSERD 0, 1, 1 / 2, sqr3 / 2, 1, CURVE_DIR.CD_CW
    
    'Assert:
    Assert.IsTrue LibMath.areDoublesEqual(0, sut.cX, epsilon), "X of end point must be 0!"
    Assert.IsTrue LibMath.areDoublesEqual(0, sut.cY, epsilon), "Y of end point must be 0!"
    Assert.IsTrue LibMath.areDoublesEqual(PI / 6, sut.length, epsilon), "Length of circular arc must be be PI/6!"
    Assert.IsTrue LibMath.areDoublesEqual(PI / 2, sut.sTheta, epsilon), "Theta of center to start point line must be PI/2!"
    Assert.IsTrue LibMath.areDoublesEqual(PI / 3, sut.eTheta, epsilon), "Theta of center to end point line must be PI/3!"
    
    ' Second quadrant, CW, PI/6 length
    sut.initFromSERD -1, 0, -sqr3 / 2, 1 / 2, 1, CURVE_DIR.CD_CW

    'Assert:
    Assert.IsTrue LibMath.areDoublesEqual(0, sut.cX, epsilon), "X of end point must be 0!"
    Assert.IsTrue LibMath.areDoublesEqual(0, sut.cY, epsilon), "Y of end point must be 0!"
    Assert.IsTrue LibMath.areDoublesEqual(PI / 6, sut.length, epsilon), "Length of circular arc must be be PI/6!"
    Assert.IsTrue LibMath.areDoublesEqual(PI, sut.sTheta, epsilon), "Theta of center to start point line must be PI!"
    Assert.IsTrue LibMath.areDoublesEqual(5 * PI / 6, sut.eTheta, epsilon), "Theta of center to end point line must be 5*PI/6!"

    ' Third quadrant, CW, PI/6 length
    sut.initFromSERD 0, -1, -1 / 2, -sqr3 / 2, 1, CURVE_DIR.CD_CW

    'Assert:
    Assert.IsTrue LibMath.areDoublesEqual(0, sut.cX, epsilon), "X of end point must be 0!"
    Assert.IsTrue LibMath.areDoublesEqual(0, sut.cY, epsilon), "Y of end point must be 0!"
    Assert.IsTrue LibMath.areDoublesEqual(PI / 6, sut.length, epsilon), "Length of circular arc must be be PI/6!"
    Assert.IsTrue LibMath.areDoublesEqual(-PI / 2, sut.sTheta, epsilon), "Theta of center to start point line must be -PI/2!"
    Assert.IsTrue LibMath.areDoublesEqual(-2 * PI / 3, sut.eTheta, epsilon), "Theta of center to end point line must be -2*PI/3!"

    ' Fourth quadrant, CW, PI/6 length
    sut.initFromSERD 1, 0, sqr3 / 2, -1 / 2, 1, CURVE_DIR.CD_CW

    'Assert:
    Assert.IsTrue LibMath.areDoublesEqual(0, sut.cX, epsilon), "X of end point must be 0!"
    Assert.IsTrue LibMath.areDoublesEqual(0, sut.cY, epsilon), "Y of end point must be 0!"
    Assert.IsTrue LibMath.areDoublesEqual(PI / 6, sut.length, epsilon), "Length of circular arc must be be PI/6!"
    Assert.IsTrue LibMath.areDoublesEqual(0#, sut.sTheta, epsilon), "Theta of center to start point line must be 0.0!"
    Assert.IsTrue LibMath.areDoublesEqual(-PI / 6, sut.eTheta, epsilon), "Theta of center to end point line must be -PI/6!"


    ' First quadrant, CCW, PI/6 length
    sut.initFromSERD 1, 0, sqr3 / 2, 1 / 2, 1, CURVE_DIR.CD_CCW

    'Assert:
    Assert.IsTrue LibMath.areDoublesEqual(0, sut.cX, epsilon), "X of end point must be 0!"
    Assert.IsTrue LibMath.areDoublesEqual(0, sut.cY, epsilon), "Y of end point must be 0!"
    Assert.IsTrue LibMath.areDoublesEqual(PI / 6, sut.length, epsilon), "Length of circular arc must be be PI/6!"
    Assert.IsTrue LibMath.areDoublesEqual(0#, sut.sTheta, epsilon), "Theta of center to start point line must be 0.0!"
    Assert.IsTrue LibMath.areDoublesEqual(PI / 6, sut.eTheta, epsilon), "Theta of center to end point line must be PI/6!"

    ' Second quadrant, CCW, PI/6 length
    sut.initFromSERD 0, 1, -1 / 2, sqr3 / 2, 1, CURVE_DIR.CD_CCW

    'Assert:
    Assert.IsTrue LibMath.areDoublesEqual(0, sut.cX, epsilon), "X of end point must be 0!"
    Assert.IsTrue LibMath.areDoublesEqual(0, sut.cY, epsilon), "Y of end point must be 0!"
    Assert.IsTrue LibMath.areDoublesEqual(PI / 6, sut.length, epsilon), "Length of circular arc must be be PI/6!"
    Assert.IsTrue LibMath.areDoublesEqual(PI / 2, sut.sTheta, epsilon), "Theta of center to start point line must be PI/2!"
    Assert.IsTrue LibMath.areDoublesEqual(2 * PI / 3, sut.eTheta, epsilon), "Theta of center to end point line must be 2*PI/3!"

    ' Third quadrant, CCW, PI/6 length
    sut.initFromSERD -1, 0, -sqr3 / 2, -1 / 2, 1, CURVE_DIR.CD_CCW

    'Assert:
    Assert.IsTrue LibMath.areDoublesEqual(0, sut.cX, epsilon), "X of end point must be 0!"
    Assert.IsTrue LibMath.areDoublesEqual(0, sut.cY, epsilon), "Y of end point must be 0!"
    Assert.IsTrue LibMath.areDoublesEqual(PI / 6, sut.length, epsilon), "Length of circular arc must be be PI/6!"
    Assert.IsTrue LibMath.areDoublesEqual(-PI, sut.sTheta, epsilon), "Theta of center to start point line must be -PI!"
    Assert.IsTrue LibMath.areDoublesEqual(-5 * PI / 6, sut.eTheta, epsilon), "Theta of center to end point line must be -5*PI/6!"

    ' Fourth quadrant, CCW, PI/6 length
    sut.initFromSERD 0, -1, 1 / 2, -sqr3 / 2, 1, CURVE_DIR.CD_CCW

    'Assert:
    Assert.IsTrue LibMath.areDoublesEqual(0, sut.cX, epsilon), "X of end point must be 0!"
    Assert.IsTrue LibMath.areDoublesEqual(0, sut.cY, epsilon), "Y of end point must be 0!"
    Assert.IsTrue LibMath.areDoublesEqual(PI / 6, sut.length, epsilon), "Length of circular arc must be be PI/6!"
    Assert.IsTrue LibMath.areDoublesEqual(-PI / 2, sut.sTheta, epsilon), "Theta of center to start point line must be -PI/2!"
    Assert.IsTrue LibMath.areDoublesEqual(-PI / 3, sut.eTheta, epsilon), "Theta of center to end point line must be -PI/3!"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub TestIsCircle()
    On Error GoTo TestFail
    
    'Arrange:
    Dim sut As New CircularArc  ' circular arc used for testing
    Dim e As Double
    
    'Act:
    ' very small circle
    e = 0.000000000000001
    sut.initFromSCLD e, 0, 0, 0, LibGeom.TWO_PI * e, CURVE_DIR.CD_CCW
    Assert.IsTrue sut.isCircle, "Failed to define circula arc as a very small circle"
    
    ' very large circle
    e = 1E+15
    sut.initFromSCLD e, 0, 0, 0, LibGeom.TWO_PI * e, CURVE_DIR.CD_CW
    Assert.IsTrue sut.isCircle, "Failed to define circula arc as a very small circle"

    'Assert:
    

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub TestIsThetaOnArc()
    On Error GoTo TestFail
    
    'Arrange:
    Dim sut As New CircularArc
    
    ' First quadrant, CCW
    sut.initFromSCLD 1, 0, 0, 0, LibGeom.PI / 2, CURVE_DIR.CD_CCW
    
    Assert.IsTrue sut.isThetaOnArc(0)
    Assert.IsTrue sut.isThetaOnArc(LibGeom.PI / 4)
    Assert.IsTrue sut.isThetaOnArc(LibGeom.PI / 2)
    Assert.IsFalse sut.isThetaOnArc(LibGeom.PI)
    Assert.IsFalse sut.isThetaOnArc(-LibGeom.PI * 3 / 4)
    Assert.IsFalse sut.isThetaOnArc(-LibGeom.PI / 2)
    
    ' Second quadrant, CW
    sut.initFromSCLD -1, 0, 0, 0, LibGeom.PI / 2, CURVE_DIR.CD_CW
    
    Assert.IsTrue sut.isThetaOnArc(LibGeom.PI / 2)
    Assert.IsTrue sut.isThetaOnArc(LibGeom.PI * 3 / 4)
    Assert.IsTrue sut.isThetaOnArc(LibGeom.PI)
    Assert.IsFalse sut.isThetaOnArc(0)
    Assert.IsFalse sut.isThetaOnArc(-LibGeom.PI / 4)
    Assert.IsFalse sut.isThetaOnArc(-LibGeom.PI / 2)
    
    ' Third quadrant, CCW
    sut.initFromSCLD -1, 0, 0, 0, LibGeom.PI / 2, CURVE_DIR.CD_CCW
    
    Assert.IsTrue sut.isThetaOnArc(-LibGeom.PI)
    Assert.IsTrue sut.isThetaOnArc(-LibGeom.PI * 3 / 4)
    Assert.IsTrue sut.isThetaOnArc(-LibGeom.PI / 2)
    Assert.IsFalse sut.isThetaOnArc(LibGeom.PI / 2)
    Assert.IsFalse sut.isThetaOnArc(-LibGeom.PI / 4)
    Assert.IsFalse sut.isThetaOnArc(0)
    
    ' Fourth quadrant, CW
    sut.initFromSCLD 1, 0, 0, 0, LibGeom.PI / 2, CURVE_DIR.CD_CW

    Assert.IsTrue sut.isThetaOnArc(0)
    Assert.IsTrue sut.isThetaOnArc(-LibGeom.PI / 4)
    Assert.IsTrue sut.isThetaOnArc(-LibGeom.PI / 2)
    Assert.IsFalse sut.isThetaOnArc(LibGeom.PI)
    Assert.IsFalse sut.isThetaOnArc(LibGeom.PI * 3 / 4)
    Assert.IsFalse sut.isThetaOnArc(LibGeom.PI / 2)

    ' Angle wrap CCW
    sut.initFromSCLD 0, 1, 0, 0, LibGeom.PI, CURVE_DIR.CD_CCW
    
    Assert.IsTrue sut.isThetaOnArc(LibGeom.PI * 3 / 4)
    Assert.IsTrue sut.isThetaOnArc(LibGeom.PI)
    Assert.IsTrue sut.isThetaOnArc(-LibGeom.PI)
    Assert.IsTrue sut.isThetaOnArc(-LibGeom.PI * 3 / 4)
    Assert.IsFalse sut.isThetaOnArc(LibGeom.PI / 4)
    Assert.IsFalse sut.isThetaOnArc(0)
    Assert.IsFalse sut.isThetaOnArc(-LibGeom.PI / 4)
    
    ' Angle wrap CW
    sut.initFromSCLD 0, -1, 0, 0, LibGeom.PI, CURVE_DIR.CD_CW
    
    Assert.IsTrue sut.isThetaOnArc(LibGeom.PI * 3 / 4)
    Assert.IsTrue sut.isThetaOnArc(LibGeom.PI)
    Assert.IsTrue sut.isThetaOnArc(-LibGeom.PI)
    Assert.IsTrue sut.isThetaOnArc(-LibGeom.PI * 3 / 4)
    Assert.IsFalse sut.isThetaOnArc(LibGeom.PI / 4)
    Assert.IsFalse sut.isThetaOnArc(0)
    Assert.IsFalse sut.isThetaOnArc(-LibGeom.PI / 4)
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub TestGetPointByMeasOffset()
    On Error GoTo TestFail
    
    'Arrange:
    Dim ca As New CircularArc  ' circular arc used for testing
    Dim sut As Point
    Dim epsilon As Double
    Dim r As Double
    
    'Act:
    epsilon = 0.00000000000001                  '1E-14
    r = 0.5
    
    ' Full circle, CCW, right offset
    ca.initFromSCLD 0.5, 0, 0, 0, LibGeom.TWO_PI * r, CURVE_DIR.CD_CCW
    
    Set sut = ca.calcPointByMeasOffset(0, 0.5)
    Assert.IsTrue LibMath.areDoublesEqual(1, sut.x, epsilon)
    Assert.IsTrue LibMath.areDoublesEqual(0, sut.y, epsilon)
    
    Set sut = ca.calcPointByMeasOffset(LibGeom.PI / 2 * r, 0.5)
    Assert.IsTrue LibMath.areDoublesEqual(0, sut.x, epsilon)
    Assert.IsTrue LibMath.areDoublesEqual(1, sut.y, epsilon)
    
    Set sut = ca.calcPointByMeasOffset(LibGeom.PI * r, 0.5)
    Assert.IsTrue LibMath.areDoublesEqual(-1, sut.x, epsilon)
    Assert.IsTrue LibMath.areDoublesEqual(0, sut.y, epsilon)
    
    Set sut = ca.calcPointByMeasOffset(LibGeom.PI * 3 / 2 * r, 0.5)
    Assert.IsTrue LibMath.areDoublesEqual(0, sut.x, epsilon)
    Assert.IsTrue LibMath.areDoublesEqual(-1, sut.y, epsilon)
    
    Set sut = ca.calcPointByMeasOffset(LibGeom.TWO_PI * r, 0.5)
    Assert.IsTrue LibMath.areDoublesEqual(1, sut.x, epsilon)
    Assert.IsTrue LibMath.areDoublesEqual(0, sut.y, epsilon)
    
    ' Full circle, CCW, left offset
    ca.initFromSCLD 0.5, 0, 0, 0, LibGeom.TWO_PI * r, CURVE_DIR.CD_CCW
    
    Set sut = ca.calcPointByMeasOffset(0, -0.5)
    Assert.IsTrue LibMath.areDoublesEqual(0, sut.x, epsilon)
    Assert.IsTrue LibMath.areDoublesEqual(0, sut.y, epsilon)
    
    Set sut = ca.calcPointByMeasOffset(LibGeom.PI / 2 * r, -0.5)
    Assert.IsTrue LibMath.areDoublesEqual(0, sut.x, epsilon)
    Assert.IsTrue LibMath.areDoublesEqual(0, sut.y, epsilon)
    
    Set sut = ca.calcPointByMeasOffset(LibGeom.PI * r, -0.5)
    Assert.IsTrue LibMath.areDoublesEqual(0, sut.x, epsilon)
    Assert.IsTrue LibMath.areDoublesEqual(0, sut.y, epsilon)
    
    Set sut = ca.calcPointByMeasOffset(LibGeom.PI * 3 / 2 * r, -0.5)
    Assert.IsTrue LibMath.areDoublesEqual(0, sut.x, epsilon)
    Assert.IsTrue LibMath.areDoublesEqual(0, sut.y, epsilon)
    
    Set sut = ca.calcPointByMeasOffset(LibGeom.TWO_PI * r, -0.5)
    Assert.IsTrue LibMath.areDoublesEqual(0, sut.x, epsilon)
    Assert.IsTrue LibMath.areDoublesEqual(0, sut.y, epsilon)


    ' Full circle, CW, right offset
    ca.initFromSCLD 0.5, 0, 0, 0, LibGeom.TWO_PI * r, CURVE_DIR.CD_CW
    
    Set sut = ca.calcPointByMeasOffset(0, 0.5)
    Assert.IsTrue LibMath.areDoublesEqual(0, sut.x, epsilon)
    Assert.IsTrue LibMath.areDoublesEqual(0, sut.y, epsilon)
    
    Set sut = ca.calcPointByMeasOffset(LibGeom.PI / 2 * r, 0.5)
    Assert.IsTrue LibMath.areDoublesEqual(0, sut.x, epsilon)
    Assert.IsTrue LibMath.areDoublesEqual(0, sut.y, epsilon)
    
    Set sut = ca.calcPointByMeasOffset(LibGeom.PI * r, 0.5)
    Assert.IsTrue LibMath.areDoublesEqual(0, sut.x, epsilon)
    Assert.IsTrue LibMath.areDoublesEqual(0, sut.y, epsilon)
    
    Set sut = ca.calcPointByMeasOffset(LibGeom.PI * 3 / 2 * r, 0.5)
    Assert.IsTrue LibMath.areDoublesEqual(0, sut.x, epsilon)
    Assert.IsTrue LibMath.areDoublesEqual(0, sut.y, epsilon)
    
    Set sut = ca.calcPointByMeasOffset(LibGeom.TWO_PI * r, 0.5)
    Assert.IsTrue LibMath.areDoublesEqual(0, sut.x, epsilon)
    Assert.IsTrue LibMath.areDoublesEqual(0, sut.y, epsilon)
    
    ' Full circle, CW, left offset
    ca.initFromSCLD 0.5, 0, 0, 0, LibGeom.TWO_PI * r, CURVE_DIR.CD_CW
    
    Set sut = ca.calcPointByMeasOffset(0, -0.5)
    Assert.IsTrue LibMath.areDoublesEqual(1, sut.x, epsilon)
    Assert.IsTrue LibMath.areDoublesEqual(0, sut.y, epsilon)
    
    Set sut = ca.calcPointByMeasOffset(LibGeom.PI / 2 * r, -0.5)
    Assert.IsTrue LibMath.areDoublesEqual(0, sut.x, epsilon)
    Assert.IsTrue LibMath.areDoublesEqual(-1, sut.y, epsilon)
    
    Set sut = ca.calcPointByMeasOffset(LibGeom.PI * r, -0.5)
    Assert.IsTrue LibMath.areDoublesEqual(-1, sut.x, epsilon)
    Assert.IsTrue LibMath.areDoublesEqual(0, sut.y, epsilon)
    
    Set sut = ca.calcPointByMeasOffset(LibGeom.PI * 3 / 2 * r, -0.5)
    Assert.IsTrue LibMath.areDoublesEqual(0, sut.x, epsilon)
    Assert.IsTrue LibMath.areDoublesEqual(1, sut.y, epsilon)
    
    Set sut = ca.calcPointByMeasOffset(LibGeom.TWO_PI * r, -0.5)
    Assert.IsTrue LibMath.areDoublesEqual(1, sut.x, epsilon)
    Assert.IsTrue LibMath.areDoublesEqual(0, sut.y, epsilon)
    
    ' PI/3 circular arc, CW, out of range measure
    ca.initFromSCLD 1, 0, 0, 0, LibGeom.PI / 3 * r, CURVE_DIR.CD_CCW
    Assert.IsNothing ca.calcPointByMeasOffset(LibGeom.PI / 3 + epsilon, 0)
    Assert.IsNothing ca.calcPointByMeasOffset(0 - epsilon, 0)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub TestGetMeasOffsetOfPointOnFullCircle()
    On Error GoTo TestFail
    
    'Arrange:
    Dim ca As New CircularArc  ' circular arc used for testing
    Dim sut As MeasOffset
    Dim r As Double
    Dim epsilon As Double
    Dim sqr2 As Double
    Dim sqr3 As Double

    'Act:
    epsilon = 0.00000000000001       '1E-14
    sqr3 = Sqr(3)
    sqr2 = Sqr(2)
    
    'Assert:
    ' Full circle, CCW, right offset
    r = 0.5
    ca.initFromSCLD 0.5, 0, 0, 0, LibGeom.TWO_PI * r, CURVE_DIR.CD_CCW
    
    Set sut = ca.calcMeasOffsetOfPoint(1, 0)
    Assert.IsTrue LibMath.areDoublesEqual(0, sut.m, epsilon)
    Assert.IsTrue LibMath.areDoublesEqual(0.5, sut.o, epsilon)

    Set sut = ca.calcMeasOffsetOfPoint(0, 1)
    Assert.IsTrue LibMath.areDoublesEqual(LibGeom.PI / 2 * r, sut.m, epsilon)
    Assert.IsTrue LibMath.areDoublesEqual(0.5, sut.o, epsilon)
    
    Set sut = ca.calcMeasOffsetOfPoint(-1, 0)
    Assert.IsTrue LibMath.areDoublesEqual(LibGeom.PI * r, sut.m, epsilon)
    Assert.IsTrue LibMath.areDoublesEqual(0.5, sut.o, epsilon)
    
    Set sut = ca.calcMeasOffsetOfPoint(0, -1)
    Assert.IsTrue LibMath.areDoublesEqual(LibGeom.PI * 3 / 2 * r, sut.m, epsilon)
    Assert.IsTrue LibMath.areDoublesEqual(0.5, sut.o, epsilon)
    
    ' Full circle, CW, right offset
    r = 1
    ca.initFromSCLD 1, 0, 0, 0, LibGeom.TWO_PI * r, CURVE_DIR.CD_CW
    
    Set sut = ca.calcMeasOffsetOfPoint(0.5, 0)
    Assert.IsTrue LibMath.areDoublesEqual(0, sut.m, epsilon)
    Assert.IsTrue LibMath.areDoublesEqual(0.5, sut.o, epsilon)

    Set sut = ca.calcMeasOffsetOfPoint(0, -0.5)
    Assert.IsTrue LibMath.areDoublesEqual(LibGeom.PI / 2 * r, sut.m, epsilon)
    Assert.IsTrue LibMath.areDoublesEqual(0.5, sut.o, epsilon)

    Set sut = ca.calcMeasOffsetOfPoint(-0.5, 0)
    Assert.IsTrue LibMath.areDoublesEqual(LibGeom.PI * r, sut.m, epsilon)
    Assert.IsTrue LibMath.areDoublesEqual(0.5, sut.o, epsilon)

    Set sut = ca.calcMeasOffsetOfPoint(0, 0.5)
    Assert.IsTrue LibMath.areDoublesEqual(3 * LibGeom.PI / 2 * r, sut.m, epsilon)
    Assert.IsTrue LibMath.areDoublesEqual(0.5, sut.o, epsilon)
    
    ' Full circle, CCW, left offset
    r = 1
    ca.initFromSCLD 1, 0, 0, 0, LibGeom.TWO_PI * r, CURVE_DIR.CD_CCW
    
    Set sut = ca.calcMeasOffsetOfPoint(0.5, 0)
    Assert.IsTrue LibMath.areDoublesEqual(0, sut.m, epsilon)
    Assert.IsTrue LibMath.areDoublesEqual(-0.5, sut.o, epsilon)

    Set sut = ca.calcMeasOffsetOfPoint(0, 0.5)
    Assert.IsTrue LibMath.areDoublesEqual(LibGeom.PI / 2 * r, sut.m, epsilon)
    Assert.IsTrue LibMath.areDoublesEqual(-0.5, sut.o, epsilon)

    Set sut = ca.calcMeasOffsetOfPoint(-0.5, 0)
    Assert.IsTrue LibMath.areDoublesEqual(LibGeom.PI * r, sut.m, epsilon)
    Assert.IsTrue LibMath.areDoublesEqual(-0.5, sut.o, epsilon)

    Set sut = ca.calcMeasOffsetOfPoint(0, -0.5)
    Assert.IsTrue LibMath.areDoublesEqual(3 * LibGeom.PI / 2 * r, sut.m, epsilon)
    Assert.IsTrue LibMath.areDoublesEqual(-0.5, sut.o, epsilon)

    
    ' Full circle, CW, left offset
    r = 0.5
    ca.initFromSCLD 0.5, 0, 0, 0, LibGeom.TWO_PI * r, CURVE_DIR.CD_CW
    
    Set sut = ca.calcMeasOffsetOfPoint(1, 0)
    Assert.IsTrue LibMath.areDoublesEqual(0, sut.m, epsilon)
    Assert.IsTrue LibMath.areDoublesEqual(-0.5, sut.o, epsilon)

    Set sut = ca.calcMeasOffsetOfPoint(0, -1)
    Assert.IsTrue LibMath.areDoublesEqual(LibGeom.PI / 2 * r, sut.m, epsilon)
    Assert.IsTrue LibMath.areDoublesEqual(-0.5, sut.o, epsilon)
    
    Set sut = ca.calcMeasOffsetOfPoint(-1, 0)
    Assert.IsTrue LibMath.areDoublesEqual(LibGeom.PI * r, sut.m, epsilon)
    Assert.IsTrue LibMath.areDoublesEqual(-0.5, sut.o, epsilon)
    
    Set sut = ca.calcMeasOffsetOfPoint(0, 1)
    Assert.IsTrue LibMath.areDoublesEqual(LibGeom.PI * 3 / 2 * r, sut.m, epsilon)
    Assert.IsTrue LibMath.areDoublesEqual(-0.5, sut.o, epsilon)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub TestGetMeasOffsetOfPoint()
    On Error GoTo TestFail
    
    'Arrange:
    Dim ca As New CircularArc  ' circular arc used for testing
    Dim sut As MeasOffset
    Dim r As Double
    Dim epsilon As Double
    Dim sqr2 As Double
    Dim sqr3 As Double

    'Act:
    epsilon = 0.00000000000001       '1E-14
    sqr3 = Sqr(3)
    sqr2 = Sqr(2)
    
    'Assert:
    
    ' First quadrant, CCW, PI/2 length, Pi/3 measure, right offset
    r = 0.5
    ca.initFromSCLD 0.5, 0, 0, 0, LibGeom.PI / 2 * r, CURVE_DIR.CD_CCW

    Set sut = ca.calcMeasOffsetOfPoint(1 / 2, sqr3 / 2)
    Assert.IsTrue LibMath.areDoublesEqual(LibGeom.PI / 3 * r, sut.m, epsilon), "Failed measure on first quadrant, CCW, right offset!"
    Assert.IsTrue LibMath.areDoublesEqual(0.5, sut.o, epsilon), "Failed offset on first quadrant, CCW, right offset!"

    ' Second quadrant, CCW, PI/2 length, Pi/3 measure, right offset
    ca.initFromSCLD 0, 0.5, 0, 0, LibGeom.PI / 2 * r, CURVE_DIR.CD_CCW

    Set sut = ca.calcMeasOffsetOfPoint(-sqr3 / 2, 1 / 2)
    Assert.IsTrue LibMath.areDoublesEqual(LibGeom.PI / 3 * r, sut.m, epsilon), "Failed measure on second quadrant, CCW, right offset!"
    Assert.IsTrue LibMath.areDoublesEqual(0.5, sut.o, epsilon), "Failed offset on second quadrant, CCW, right offset!"

    ' Third quadrant, CCW, PI/2 length, Pi/3 measure, right offset
    ca.initFromSCLD -0.5, 0, 0, 0, LibGeom.PI / 2 * r, CURVE_DIR.CD_CCW

    Set sut = ca.calcMeasOffsetOfPoint(-1 / 2, -sqr3 / 2)
    Assert.IsTrue LibMath.areDoublesEqual(LibGeom.PI / 3 * r, sut.m, epsilon), "Failed measure on third quadrant, CCW, right offset!"
    Assert.IsTrue LibMath.areDoublesEqual(0.5, sut.o, epsilon), "Failed offset on third quadrant, CCW, right offset!"

    ' Fourth quadrant, CCW, PI/2 length, Pi/3 measure, right offset
    ca.initFromSCLD 0, -0.5, 0, 0, LibGeom.PI / 2 * r, CURVE_DIR.CD_CCW

    Set sut = ca.calcMeasOffsetOfPoint(sqr3 / 2, -1 / 2)
    Assert.IsTrue LibMath.areDoublesEqual(LibGeom.PI / 3 * r, sut.m, epsilon), "Failed measure on fourth quadrant, CCW, right offset!"
    Assert.IsTrue LibMath.areDoublesEqual(0.5, sut.o, epsilon), "Failed offset on fourth quadrant, CCW, right offset!"

    ' First quadrant, CW, PI/2 length, Pi/3 measure, left offset
    r = 0.5
    ca.initFromSCLD 0, 0.5, 0, 0, LibGeom.PI / 2 * r, CURVE_DIR.CD_CW

    Set sut = ca.calcMeasOffsetOfPoint(sqr3 / 2, 1 / 2)
    Assert.IsTrue LibMath.areDoublesEqual(LibGeom.PI / 3 * r, sut.m, epsilon), "Failed measure on first quadrant, CW, left offset!"
    Assert.IsTrue LibMath.areDoublesEqual(-0.5, sut.o, epsilon), "Failed offset on first quadrant, CW, left offset!"

    ' Second quadrant, CW, PI/2 length, Pi/3 measure, left offset
    ca.initFromSCLD -0.5, 0, 0, 0, LibGeom.PI / 2 * r, CURVE_DIR.CD_CW

    Set sut = ca.calcMeasOffsetOfPoint(-1 / 2, sqr3 / 2)
    Assert.IsTrue LibMath.areDoublesEqual(LibGeom.PI / 3 * r, sut.m, epsilon), "Failed measure on second quadrant, CW, left offset!"
    Assert.IsTrue LibMath.areDoublesEqual(-0.5, sut.o, epsilon), "Failed offset on second quadrant, CW, left offset!"

    ' Third quadrant, CW, PI/2 length, Pi/3 measure, left offset
    ca.initFromSCLD 0, -0.5, 0, 0, LibGeom.PI / 2 * r, CURVE_DIR.CD_CW

    Set sut = ca.calcMeasOffsetOfPoint(-sqr3 / 2, -1 / 2)
    Assert.IsTrue LibMath.areDoublesEqual(LibGeom.PI / 3 * r, sut.m, epsilon), "Failed measure on third quadrant, CW, left offset!"
    Assert.IsTrue LibMath.areDoublesEqual(-0.5, sut.o, epsilon), "Failed offset on third quadrant, CW, left offset!"

    ' Fourth quadrant, CW, PI/2 length, Pi/3 measure, left offset
    ca.initFromSCLD 0.5, 0, 0, 0, LibGeom.PI / 2 * r, CURVE_DIR.CD_CW

    Set sut = ca.calcMeasOffsetOfPoint(1 / 2, -sqr3 / 2)
    Assert.IsTrue LibMath.areDoublesEqual(LibGeom.PI / 3 * r, sut.m, epsilon), "Failed measure on fourth quadrant, CW, left offset!"
    Assert.IsTrue LibMath.areDoublesEqual(-0.5, sut.o, epsilon), "Failed offset on fourth quadrant, CW, left offset!"

    ' First quadrant, CCW, PI/2 length, Pi/3 measure, left offset
    r = 1.5
    ca.initFromSCLD 1.5, 0, 0, 0, LibGeom.PI / 2 * r, CURVE_DIR.CD_CCW

    Set sut = ca.calcMeasOffsetOfPoint(1 / 2, sqr3 / 2)
    Assert.IsTrue LibMath.areDoublesEqual(LibGeom.PI / 3 * r, sut.m, epsilon), "Failed measure on first quadrant, CCW, left offset!"
    Assert.IsTrue LibMath.areDoublesEqual(-0.5, sut.o, epsilon), "Failed offset on first quadrant, CCW, left offset!"

    ' Second quadrant, CCW, PI/2 length, Pi/3 measure, left offset
    ca.initFromSCLD 0, 1.5, 0, 0, LibGeom.PI / 2 * r, CURVE_DIR.CD_CCW

    Set sut = ca.calcMeasOffsetOfPoint(-sqr3 / 2, 1 / 2)
    Assert.IsTrue LibMath.areDoublesEqual(LibGeom.PI / 3 * r, sut.m, epsilon), "Failed measure on second quadrant, CCW, left offset!"
    Assert.IsTrue LibMath.areDoublesEqual(-0.5, sut.o, epsilon), "Failed offset on second quadrant, CCW, left offset!"

    ' Third quadrant, CCW, PI/2 length, Pi/3 measure, left offset
    ca.initFromSCLD -1.5, 0, 0, 0, LibGeom.PI / 2 * r, CURVE_DIR.CD_CCW

    Set sut = ca.calcMeasOffsetOfPoint(-1 / 2, -sqr3 / 2)
    Assert.IsTrue LibMath.areDoublesEqual(LibGeom.PI / 3 * r, sut.m, epsilon), "Failed measure on third quadrant, CCW, left offset!"
    Assert.IsTrue LibMath.areDoublesEqual(-0.5, sut.o, epsilon), "Failed offset on third quadrant, CCW, left offset!"

    ' Fourth quadrant, CCW, PI/2 length, Pi/3 measure, left offset
    ca.initFromSCLD 0, -1.5, 0, 0, LibGeom.PI / 2 * r, CURVE_DIR.CD_CCW

    Set sut = ca.calcMeasOffsetOfPoint(sqr3 / 2, -1 / 2)
    Assert.IsTrue LibMath.areDoublesEqual(LibGeom.PI / 3 * r, sut.m, epsilon), "Failed measure on fourth quadrant, CCW, left offset!"
    Assert.IsTrue LibMath.areDoublesEqual(-0.5, sut.o, epsilon), "Failed offset on fourth quadrant, CCW, left offset!"

    ' First quadrant, CW, PI/2 length, Pi/3 measure, right offset
    r = 1.5
    ca.initFromSCLD 0, 1.5, 0, 0, LibGeom.PI / 2 * r, CURVE_DIR.CD_CW

    Set sut = ca.calcMeasOffsetOfPoint(sqr3 / 2, 1 / 2)
    Assert.IsTrue LibMath.areDoublesEqual(LibGeom.PI / 3 * r, sut.m, epsilon), "Failed measure on first quadrant, CW, right offset!"
    Assert.IsTrue LibMath.areDoublesEqual(0.5, sut.o, epsilon), "Failed offset on first quadrant, CW, right offset!"

    ' Second quadrant, CW, PI/2 length, Pi/3 measure, right offset
    ca.initFromSCLD -1.5, 0, 0, 0, LibGeom.PI / 2 * r, CURVE_DIR.CD_CW

    Set sut = ca.calcMeasOffsetOfPoint(-1 / 2, sqr3 / 2)
    Assert.IsTrue LibMath.areDoublesEqual(LibGeom.PI / 3 * r, sut.m, epsilon), "Failed measure on second quadrant, CW, right offset!"
    Assert.IsTrue LibMath.areDoublesEqual(0.5, sut.o, epsilon), "Failed offset on second quadrant, CW, right offset!"

    ' Third quadrant, CW, PI/2 length, Pi/3 measure, right offset
    ca.initFromSCLD 0, -1.5, 0, 0, LibGeom.PI / 2 * r, CURVE_DIR.CD_CW

    Set sut = ca.calcMeasOffsetOfPoint(-sqr3 / 2, -1 / 2)
    Assert.IsTrue LibMath.areDoublesEqual(LibGeom.PI / 3 * r, sut.m, epsilon), "Failed measure on third quadrant, CW, right offset!"
    Assert.IsTrue LibMath.areDoublesEqual(0.5, sut.o, epsilon), "Failed offset on third quadrant, CW, right offset!"

    ' Fourth quadrant, CW, PI/2 length, Pi/3 measure, right offset
    ca.initFromSCLD 1.5, 0, 0, 0, LibGeom.PI / 2 * r, CURVE_DIR.CD_CW

    Set sut = ca.calcMeasOffsetOfPoint(1 / 2, -sqr3 / 2)
    Assert.IsTrue LibMath.areDoublesEqual(LibGeom.PI / 3 * r, sut.m, epsilon), "Failed measure on fourth quadrant, CW, right offset!"
    Assert.IsTrue LibMath.areDoublesEqual(0.5, sut.o, epsilon), "Failed offset on fourth quadrant, CW, right offset!"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub TestGetMeasOffsetOfPointInCenter()
    On Error GoTo TestFail
    
    'Arrange:
    Dim ca As New CircularArc  ' circular arc used for testing
    Dim sut As Point
    
    'Act:
    ca.initFromSCLD 1, 0, 0, 0, LibGeom.PI / 2, CURVE_DIR.CD_CCW

    'Assert:
    Assert.IsNothing ca.calcMeasOffsetOfPoint(0, 0), "Nothing must be returned if input point is in the center of the circular arc!"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub TestGetMeasOffsetOfPointBeyondCenter()
    On Error GoTo TestFail
    
    'Arrange:
    Dim ca As New CircularArc  ' circular arc used for testing
    Dim sut As MeasOffset
    Dim r As Double
    Dim epsilon As Double
    Dim sqr2 As Double
    Dim sqr3 As Double

    'Act:
    epsilon = 0.00000000000001       '1E-14
    sqr3 = Sqr(3)
    sqr2 = Sqr(2)
    
    'Assert:
    
    ' First quadrant, CCW, PI/2 length, Pi/3 measure, left offset
    r = 1
    ca.initFromSCLD 1, 0, 0, 0, LibGeom.PI / 2 * r, CURVE_DIR.CD_CCW

    Set sut = ca.calcMeasOffsetOfPoint(-1 / 2, -sqr3 / 2)
    Assert.IsTrue LibMath.areDoublesEqual(LibGeom.PI / 3 * r, sut.m, epsilon), "Failed measure on first quadrant, CCW, left offset!"
    Assert.IsTrue LibMath.areDoublesEqual(-2, sut.o, epsilon), "Failed offset on first quadrant, CCW, left offset!"

    ' Second quadrant, CCW, PI/2 length, Pi/3 measure, left offset
    ca.initFromSCLD 0, 1, 0, 0, LibGeom.PI / 2 * r, CURVE_DIR.CD_CCW

    Set sut = ca.calcMeasOffsetOfPoint(sqr3 / 2, -1 / 2)
    Assert.IsTrue LibMath.areDoublesEqual(LibGeom.PI / 3 * r, sut.m, epsilon), "Failed measure on second quadrant, CCW, left offset!"
    Assert.IsTrue LibMath.areDoublesEqual(-2, sut.o, epsilon), "Failed offset on second quadrant, CCW, left offset!"

    ' Third quadrant, CCW, PI/2 length, Pi/3 measure, left offset
    ca.initFromSCLD -1, 0, 0, 0, LibGeom.PI / 2 * r, CURVE_DIR.CD_CCW

    Set sut = ca.calcMeasOffsetOfPoint(1 / 2, sqr3 / 2)
    Assert.IsTrue LibMath.areDoublesEqual(LibGeom.PI / 3 * r, sut.m, epsilon), "Failed measure on third quadrant, CCW, left offset!"
    Assert.IsTrue LibMath.areDoublesEqual(-2, sut.o, epsilon), "Failed offset on third quadrant, CCW, left offset!"

    ' Fourth quadrant, CCW, PI/2 length, Pi/3 measure, left offset
    ca.initFromSCLD 0, -1, 0, 0, LibGeom.PI / 2 * r, CURVE_DIR.CD_CCW

    Set sut = ca.calcMeasOffsetOfPoint(-sqr3 / 2, 1 / 2)
    Assert.IsTrue LibMath.areDoublesEqual(LibGeom.PI / 3 * r, sut.m, epsilon), "Failed measure on fourth quadrant, CCW, left offset!"
    Assert.IsTrue LibMath.areDoublesEqual(-2, sut.o, epsilon), "Failed offset on fourth quadrant, CCW, left offset!"

    ' First quadrant, CW, PI/2 length, Pi/3 measure, right offset
    r = 1
    ca.initFromSCLD 0, 1, 0, 0, LibGeom.PI / 2 * r, CURVE_DIR.CD_CW

    Set sut = ca.calcMeasOffsetOfPoint(-sqr3 / 2, -1 / 2)
    Assert.IsTrue LibMath.areDoublesEqual(LibGeom.PI / 3 * r, sut.m, epsilon), "Failed measure on first quadrant, CW, right offset!"
    Assert.IsTrue LibMath.areDoublesEqual(2, sut.o, epsilon), "Failed offset on first quadrant, CW, right offset!"

    ' Second quadrant, CW, PI/2 length, Pi/3 measure, right offset
    ca.initFromSCLD -1, 0, 0, 0, LibGeom.PI / 2 * r, CURVE_DIR.CD_CW

    Set sut = ca.calcMeasOffsetOfPoint(1 / 2, -sqr3 / 2)
    Assert.IsTrue LibMath.areDoublesEqual(LibGeom.PI / 3 * r, sut.m, epsilon), "Failed measure on second quadrant, CW, right offset!"
    Assert.IsTrue LibMath.areDoublesEqual(2, sut.o, epsilon), "Failed offset on second quadrant, CW, right offset!"

    ' Third quadrant, CW, PI/2 length, Pi/3 measure, right offset
    ca.initFromSCLD 0, -1, 0, 0, LibGeom.PI / 2 * r, CURVE_DIR.CD_CW

    Set sut = ca.calcMeasOffsetOfPoint(sqr3 / 2, 1 / 2)
    Assert.IsTrue LibMath.areDoublesEqual(LibGeom.PI / 3 * r, sut.m, epsilon), "Failed measure on third quadrant, CW, right offset!"
    Assert.IsTrue LibMath.areDoublesEqual(2, sut.o, epsilon), "Failed offset on third quadrant, CW, right offset!"

    ' Fourth quadrant, CW, PI/2 length, Pi/3 measure, right offset
    ca.initFromSCLD 1, 0, 0, 0, LibGeom.PI / 2 * r, CURVE_DIR.CD_CW

    Set sut = ca.calcMeasOffsetOfPoint(-1 / 2, sqr3 / 2)
    Assert.IsTrue LibMath.areDoublesEqual(LibGeom.PI / 3 * r, sut.m, epsilon), "Failed measure on fourth quadrant, CW, right offset!"
    Assert.IsTrue LibMath.areDoublesEqual(2, sut.o, epsilon), "Failed offset on fourth quadrant, CW, right offset!"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub TestGetMeasOffsetOfPointInRnd()
    On Error GoTo TestFail
    
    'Arrange:
    Dim ca As New CircularArc  ' circular arc used for testing
    Dim sut As MeasOffset
    Dim epsilon As Double
    
    'Act:
    epsilon = 0.000000001
    
    ca.initFromSCLD -87818.3779652715, 771356.94026947, -88186.2640380859, 771356.94026947, 5.00231883943564, CURVE_DIR.CD_CW
    Set sut = ca.calcMeasOffsetOfPoint(-88373.0500207291, 771357.986905455)
    Assert.IsTrue LibMath.areDoublesEqual(2.06139008360862, sut.m, epsilon)
    Assert.IsTrue LibMath.areDoublesEqual(554.674987792969, sut.o, epsilon)

    ca.initFromSCLD 918308.416069448, 185580.134391785, 917782.306671143, 185580.134391785, 4.41651344775256, CURVE_DIR.CD_CW
    Set sut = ca.calcMeasOffsetOfPoint(917350.17029878, 185580.834560072)
    Assert.IsTrue LibMath.areDoublesEqual(0.852427190745306, sut.m, epsilon)
    Assert.IsTrue LibMath.areDoublesEqual(958.246337890625, sut.o, epsilon)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub




