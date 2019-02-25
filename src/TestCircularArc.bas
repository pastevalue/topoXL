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
Public Sub TestInitFromScenLenDir()
    On Error GoTo TestFail
    
    'Arrange:
    Dim sut As New CircularArc
    Dim epsilon As Double
    
    'Act:
    epsilon = 0.000000000000001                  '1E-15
    

    ' First quadrant, CW, PI/6 length
    sut.initFromScenLenDir 0, 1, 0, 0, PI / 6, CURVE_DIR.CD_CW
    
    'Assert:
    Assert.AreEqual 1#, sut.rad, "Radius must be 1!"
    Assert.IsTrue MathLib.AreDoublesEqual(PI / 2, sut.cToStheta, epsilon), "Theta of center to start point line must be PI/2!"
    Assert.IsTrue MathLib.AreDoublesEqual(PI / 3, sut.cToEtheta, epsilon), "Theta of center to end point line must be PI/3!"
    Assert.IsTrue MathLib.AreDoublesEqual(Cos(PI / 3), sut.e.x, epsilon), "X of end point must be" & Cos(PI / 3) & "!"
    Assert.IsTrue MathLib.AreDoublesEqual(Sin(PI / 3), sut.e.y, epsilon), "Y of end point must be" & Sin(PI / 3) & "!"
    
    ' Second quadrant, CW, PI/6 length
    sut.initFromScenLenDir -1, 0, 0, 0, PI / 6, CURVE_DIR.CD_CW

    'Assert:
    Assert.AreEqual 1#, sut.rad, "Radius must be 1!"
    Assert.IsTrue MathLib.AreDoublesEqual(PI, sut.cToStheta, epsilon), "Theta of center to start point line must be PI!"
    Assert.IsTrue MathLib.AreDoublesEqual(5 * PI / 6, sut.cToEtheta, epsilon), "Theta of center to end point line must be 5*PI/6!"
    Assert.IsTrue MathLib.AreDoublesEqual(Cos(5 * PI / 6), sut.e.x, epsilon), "X of end point must be" & Cos(5 * PI / 6) & "!"
    Assert.IsTrue MathLib.AreDoublesEqual(Sin(5 * PI / 6), sut.e.y, epsilon), "Y of end point must be" & Sin(5 * PI / 6) & "!"
    
    ' Third quadrant, CW, PI/6 length
    sut.initFromScenLenDir 0, -1, 0, 0, PI / 6, CURVE_DIR.CD_CW
    
    'Assert:
    Assert.AreEqual 1#, sut.rad, "Radius must be 1!"
    Assert.IsTrue MathLib.AreDoublesEqual(-PI / 2, sut.cToStheta, epsilon), "Theta of center to start point line must be -PI/2!"
    Assert.IsTrue MathLib.AreDoublesEqual(-2 * PI / 3, sut.cToEtheta, epsilon), "Theta of center to end point line must be -2*PI/3!"
    Assert.IsTrue MathLib.AreDoublesEqual(Cos(-2 * PI / 3), sut.e.x, epsilon), "X of end point must be" & Cos(-2 * PI / 3) & "!"
    Assert.IsTrue MathLib.AreDoublesEqual(Sin(-2 * PI / 3), sut.e.y, epsilon), "Y of end point must be" & Sin(-2 * PI / 3) & "!"
    
    ' Fourth quadrant, CW, PI/6 length
    sut.initFromScenLenDir 1, 0, 0, 0, PI / 6, CURVE_DIR.CD_CW
    
    'Assert:
    Assert.AreEqual 1#, sut.rad, "Radius must be 1!"
    Assert.IsTrue MathLib.AreDoublesEqual(0, sut.cToStheta, epsilon), "Theta of center to start point line must be 0.0!"
    Assert.IsTrue MathLib.AreDoublesEqual(-PI / 6, sut.cToEtheta, epsilon), "Theta of center to end point line must be -PI/6!"
    Assert.IsTrue MathLib.AreDoublesEqual(Cos(-PI / 6), sut.e.x, epsilon), "X of end point must be" & Cos(-PI / 6) & "!"
    Assert.IsTrue MathLib.AreDoublesEqual(Sin(-PI / 6), sut.e.y, epsilon), "Y of end point must be" & Sin(-PI / 6) & "!"
    
    
    ' First quadrant, CCW, PI/6 length
    sut.initFromScenLenDir 1, 0, 0, 0, PI / 6, CURVE_DIR.CD_CCW
    
    'Assert:
    Assert.AreEqual 1#, sut.rad, "Radius must be 1!"
    Assert.IsTrue MathLib.AreDoublesEqual(0, sut.cToStheta, epsilon), "Theta of center to start point line must be 0!"
    Assert.IsTrue MathLib.AreDoublesEqual(PI / 6, sut.cToEtheta, epsilon), "Theta of center to end point line must be PI/6!"
    Assert.IsTrue MathLib.AreDoublesEqual(Cos(PI / 6), sut.e.x, epsilon), "X of end point must be" & Cos(PI / 6) & "!"
    Assert.IsTrue MathLib.AreDoublesEqual(Sin(PI / 6), sut.e.y, epsilon), "Y of end point must be" & Sin(PI / 6) & "!"
    
    ' Second quadrant, CCW, PI/6 length
    sut.initFromScenLenDir 0, 1, 0, 0, PI / 6, CURVE_DIR.CD_CCW
    
    'Assert:
    Assert.AreEqual 1#, sut.rad, "Radius must be 1!"
    Assert.IsTrue MathLib.AreDoublesEqual(PI / 2, sut.cToStheta, epsilon), "Theta of center to start point line must be PI/2!"
    Assert.IsTrue MathLib.AreDoublesEqual(2 * PI / 3, sut.cToEtheta, epsilon), "Theta of center to end point line must be 2*PI/3!"
    Assert.IsTrue MathLib.AreDoublesEqual(Cos(2 * PI / 3), sut.e.x, epsilon), "X of end point must be" & Cos(2 * PI / 3) & "!"
    Assert.IsTrue MathLib.AreDoublesEqual(Sin(2 * PI / 3), sut.e.y, epsilon), "Y of end point must be" & Sin(2 * PI / 3) & "!"
    
    ' Third quadrant, CCW, PI/6 length
    sut.initFromScenLenDir -1, 0, 0, 0, PI / 6, CURVE_DIR.CD_CCW
    
    'Assert:
    Assert.AreEqual 1#, sut.rad, "Radius must be 1!"
    Assert.IsTrue MathLib.AreDoublesEqual(PI, sut.cToStheta, epsilon), "Theta of center to start point line must be PI!"
    Assert.IsTrue MathLib.AreDoublesEqual(-5 * PI / 6, sut.cToEtheta, epsilon), "Theta of center to end point line must be -5*PI/6!"
    Assert.IsTrue MathLib.AreDoublesEqual(Cos(-5 * PI / 6), sut.e.x, epsilon), "X of end point must be" & Cos(-5 * PI / 6) & "!"
    Assert.IsTrue MathLib.AreDoublesEqual(Sin(-5 * PI / 6), sut.e.y, epsilon), "Y of end point must be" & Sin(-5 * PI / 6) & "!"
    
    ' Fourth quadrant, CCW, PI/6 length
    sut.initFromScenLenDir 0, -1, 0, 0, PI / 6, CURVE_DIR.CD_CCW
    
    'Assert:
    Assert.AreEqual 1#, sut.rad, "Radius must be 1!"
    Assert.IsTrue MathLib.AreDoublesEqual(-PI / 2, sut.cToStheta, epsilon), "Theta of center to start point line must be -PI/2!"
    Assert.IsTrue MathLib.AreDoublesEqual(-PI / 3, sut.cToEtheta, epsilon), "Theta of center to end point line must be -PI/3!"
    Assert.IsTrue MathLib.AreDoublesEqual(Cos(-PI / 3), sut.e.x, epsilon), "X of end point must be" & Cos(-PI / 3) & "!"
    Assert.IsTrue MathLib.AreDoublesEqual(Sin(-PI / 3), sut.e.y, epsilon), "Y of end point must be" & Sin(-PI / 3) & "!"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub TestInitFromSEradDir()
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
    sut.initFromSEradDir 0, 1, 1 / 2, sqr3 / 2, 1, CURVE_DIR.CD_CW
    
    'Assert:
    Assert.IsTrue MathLib.AreDoublesEqual(0, sut.c.x, epsilon), "X of end point must be 0!"
    Assert.IsTrue MathLib.AreDoublesEqual(0, sut.c.y, epsilon), "Y of end point must be 0!"
    Assert.IsTrue MathLib.AreDoublesEqual(PI / 6, sut.length, epsilon), "Length of circular arc must be be PI/6!"
    Assert.IsTrue MathLib.AreDoublesEqual(PI / 2, sut.cToStheta, epsilon), "Theta of center to start point line must be PI/2!"
    Assert.IsTrue MathLib.AreDoublesEqual(PI / 3, sut.cToEtheta, epsilon), "Theta of center to end point line must be PI/3!"
    
    ' Second quadrant, CW, PI/6 length
    sut.initFromSEradDir -1, 0, -sqr3 / 2, 1 / 2, 1, CURVE_DIR.CD_CW

    'Assert:
    Assert.IsTrue MathLib.AreDoublesEqual(0, sut.c.x, epsilon), "X of end point must be 0!"
    Assert.IsTrue MathLib.AreDoublesEqual(0, sut.c.y, epsilon), "Y of end point must be 0!"
    Assert.IsTrue MathLib.AreDoublesEqual(PI / 6, sut.length, epsilon), "Length of circular arc must be be PI/6!"
    Assert.IsTrue MathLib.AreDoublesEqual(PI, sut.cToStheta, epsilon), "Theta of center to start point line must be PI!"
    Assert.IsTrue MathLib.AreDoublesEqual(5 * PI / 6, sut.cToEtheta, epsilon), "Theta of center to end point line must be 5*PI/6!"

    ' Third quadrant, CW, PI/6 length
    sut.initFromSEradDir 0, -1, -1 / 2, -sqr3 / 2, 1, CURVE_DIR.CD_CW

    'Assert:
    Assert.IsTrue MathLib.AreDoublesEqual(0, sut.c.x, epsilon), "X of end point must be 0!"
    Assert.IsTrue MathLib.AreDoublesEqual(0, sut.c.y, epsilon), "Y of end point must be 0!"
    Assert.IsTrue MathLib.AreDoublesEqual(PI / 6, sut.length, epsilon), "Length of circular arc must be be PI/6!"
    Assert.IsTrue MathLib.AreDoublesEqual(-PI / 2, sut.cToStheta, epsilon), "Theta of center to start point line must be -PI/2!"
    Assert.IsTrue MathLib.AreDoublesEqual(-2 * PI / 3, sut.cToEtheta, epsilon), "Theta of center to end point line must be -2*PI/3!"

    ' Fourth quadrant, CW, PI/6 length
    sut.initFromSEradDir 1, 0, sqr3 / 2, -1 / 2, 1, CURVE_DIR.CD_CW

    'Assert:
    Assert.IsTrue MathLib.AreDoublesEqual(0, sut.c.x, epsilon), "X of end point must be 0!"
    Assert.IsTrue MathLib.AreDoublesEqual(0, sut.c.y, epsilon), "Y of end point must be 0!"
    Assert.IsTrue MathLib.AreDoublesEqual(PI / 6, sut.length, epsilon), "Length of circular arc must be be PI/6!"
    Assert.IsTrue MathLib.AreDoublesEqual(0#, sut.cToStheta, epsilon), "Theta of center to start point line must be 0.0!"
    Assert.IsTrue MathLib.AreDoublesEqual(-PI / 6, sut.cToEtheta, epsilon), "Theta of center to end point line must be -PI/6!"


    ' First quadrant, CCW, PI/6 length
    sut.initFromSEradDir 1, 0, sqr3 / 2, 1 / 2, 1, CURVE_DIR.CD_CCW

    'Assert:
    Assert.IsTrue MathLib.AreDoublesEqual(0, sut.c.x, epsilon), "X of end point must be 0!"
    Assert.IsTrue MathLib.AreDoublesEqual(0, sut.c.y, epsilon), "Y of end point must be 0!"
    Assert.IsTrue MathLib.AreDoublesEqual(PI / 6, sut.length, epsilon), "Length of circular arc must be be PI/6!"
    Assert.IsTrue MathLib.AreDoublesEqual(0#, sut.cToStheta, epsilon), "Theta of center to start point line must be 0.0!"
    Assert.IsTrue MathLib.AreDoublesEqual(PI / 6, sut.cToEtheta, epsilon), "Theta of center to end point line must be PI/6!"

    ' Second quadrant, CCW, PI/6 length
    sut.initFromSEradDir 0, 1, -1 / 2, sqr3 / 2, 1, CURVE_DIR.CD_CCW

    'Assert:
    Assert.IsTrue MathLib.AreDoublesEqual(0, sut.c.x, epsilon), "X of end point must be 0!"
    Assert.IsTrue MathLib.AreDoublesEqual(0, sut.c.y, epsilon), "Y of end point must be 0!"
    Assert.IsTrue MathLib.AreDoublesEqual(PI / 6, sut.length, epsilon), "Length of circular arc must be be PI/6!"
    Assert.IsTrue MathLib.AreDoublesEqual(PI / 2, sut.cToStheta, epsilon), "Theta of center to start point line must be PI/2!"
    Assert.IsTrue MathLib.AreDoublesEqual(2 * PI / 3, sut.cToEtheta, epsilon), "Theta of center to end point line must be 2*PI/3!"

    ' Third quadrant, CCW, PI/6 length
    sut.initFromSEradDir -1, 0, -sqr3 / 2, -1 / 2, 1, CURVE_DIR.CD_CCW

    'Assert:
    Assert.IsTrue MathLib.AreDoublesEqual(0, sut.c.x, epsilon), "X of end point must be 0!"
    Assert.IsTrue MathLib.AreDoublesEqual(0, sut.c.y, epsilon), "Y of end point must be 0!"
    Assert.IsTrue MathLib.AreDoublesEqual(PI / 6, sut.length, epsilon), "Length of circular arc must be be PI/6!"
    Assert.IsTrue MathLib.AreDoublesEqual(-PI, sut.cToStheta, epsilon), "Theta of center to start point line must be -PI!"
    Assert.IsTrue MathLib.AreDoublesEqual(-5 * PI / 6, sut.cToEtheta, epsilon), "Theta of center to end point line must be -5*PI/6!"

    ' Fourth quadrant, CCW, PI/6 length
    sut.initFromSEradDir 0, -1, 1 / 2, -sqr3 / 2, 1, CURVE_DIR.CD_CCW

    'Assert:
    Assert.IsTrue MathLib.AreDoublesEqual(0, sut.c.x, epsilon), "X of end point must be 0!"
    Assert.IsTrue MathLib.AreDoublesEqual(0, sut.c.y, epsilon), "Y of end point must be 0!"
    Assert.IsTrue MathLib.AreDoublesEqual(PI / 6, sut.length, epsilon), "Length of circular arc must be be PI/6!"
    Assert.IsTrue MathLib.AreDoublesEqual(-PI / 2, sut.cToStheta, epsilon), "Theta of center to start point line must be -PI/2!"
    Assert.IsTrue MathLib.AreDoublesEqual(-PI / 3, sut.cToEtheta, epsilon), "Theta of center to end point line must be -PI/3!"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

