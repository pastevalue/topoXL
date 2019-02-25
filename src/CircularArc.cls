VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CircularArc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

''========================================================================
'' Description
'' Represents a circular arc.
''------------------------------------------------------------------------
'' Called by:
''    Modules: None
''    Classes: None
'' Calls:
''    Modules: None
''    Classes: None
''========================================================================

'@Folder("TopoXL.geom")

Option Explicit

Private m_s As Point                             ' start Point of the circular arc
Private m_c As Point                             ' center Point of the circular arc
Private m_e As Point                             ' end Point of the circular arc
Private m_length As Double                       ' length of the circular arc
Private m_curveDir As CURVE_DIR                  ' circular arc curve direction
Private m_rad As Double                          ' radius of the circular arc
Private m_cToStheta As Double                    ' theta of the center to start line
Private m_cToEtheta As Double                    ' theta of the center to end line

' Class constructor with default values
Private Sub Class_Initialize()
    Set m_s = New Point
    Set m_c = New Point
    Set m_e = New Point
    m_length = 0
    m_curveDir = CURVE_DIR.CD_NONE
    m_rad = 0
    m_cToStheta = 0
    m_cToEtheta = 0
End Sub

' Initialize by: start Point, center Point, length and curve direction
Public Sub initFromScenLenDir(ByVal xs As Double, ByVal ys As Double, _
                              ByVal xc As Double, ByVal yc As Double, _
                              ByVal length As Double, ByVal curveDir As CURVE_DIR)
    m_s.Init xs, ys
    m_c.Init xc, yc
    m_length = length
    m_curveDir = curveDir
    
    ' Compute other members (m_e, m_rad, m_cToStheta, m_cToEtheta)
    m_rad = GeomLib.Dist2D(xs, ys, xc, yc)
    m_cToStheta = GeomLib.Atn2(xs - xc, ys - yc)
    m_cToEtheta = GeomLib.NormalizeAngle(m_cToStheta, GeomLib.PI) - length / m_rad * curveDir
    m_cToEtheta = GeomLib.NormalizeAngle(m_cToEtheta, 0#) ' normalize theta in 0.0
    m_e.x = m_c.x + m_rad * Math.Cos(m_cToEtheta)
    m_e.y = m_c.y + m_rad * Math.Sin(m_cToEtheta)
End Sub

Public Sub initFromSEradDir(ByVal xs As Double, ByVal ys As Double, _
                            ByVal xe As Double, ByVal ye As Double, _
                            ByVal rad As Double, ByVal curveDir As CURVE_DIR)
    m_s.Init xs, ys
    m_c.Init xe, ye
    m_rad = rad
    m_curveDir = curveDir
    
    ' Compute other members (m_c, m_length, m_cToStheta, m_cToEtheta)
    Dim q As Double  ' Distance between s (start) point and e (end) point
    Dim x As Double  ' X of q mid point
    Dim y As Double  ' Y of q mid point
    
    q = GeomLib.Dist2D(xs, ys, xe, ye)
    x = (xs + xe) / 2
    y = (ys + ye) / 2
    
    m_c.x = x - curveDir * Math.Sqr(rad ^ 2 - (q / 2) ^ 2) * ((ys - ye) / q)
    m_c.y = y - curveDir * Math.Sqr(rad ^ 2 - (q / 2) ^ 2) * ((xe - xs) / q)
    
    m_length = rad * GeomLib.ACos(1 - q * q / (2 * rad * rad))
    
    m_cToStheta = GeomLib.Atn2(xs - m_c.x, ys - m_c.y)
    m_cToEtheta = GeomLib.Atn2(xe - m_c.x, ye - m_c.y)
End Sub

Public Function GetPointByMeasOffset(ByVal m As Double, ByVal o As Double) As Point

    If Abs(m) > m_length Then
        Set GetPointByMeasOffset = Nothing
        Exit Function
    End If
    
    Dim pTheta As Double                         ' Theta from the arc center to the result point
    Dim pDist As Double                          ' Distance fromt the arc center to the result point
        
    pTheta = GeomLib.NormalizeAngle(m_cToStheta, GeomLib.PI) - m / m_rad * curveDir
    pTheta = GeomLib.NormalizeAngle(pTheta, 0#)  ' normalize theta in 0.0
    pDist = m_rad - o * m_curveDir
    
    Set GetPointByMeasOffset = New Point
    GetPointByMeasOffset.x = m_c.x + pDist * Math.Sin(pTheta)
    GetPointByMeasOffset.y = m_c.y + pDist * Math.Cos(pTheta)
End Function

Public Function GetMeasOffsetOfPoint(ByVal x As Double, ByVal y As Double) As MeasOffset

End Function

' Start Point of circular arc property
Public Property Get s() As Point
    Set s = m_s
End Property

' Start Center of circular arc property
Public Property Get c() As Point
    Set c = m_c
End Property

' End Point of circular arc property
Public Property Get e() As Point
    Set e = m_e
End Property

' Length of circular arc property
Public Property Get length() As Double
    length = m_length
End Property

' Length of circular arc property
Public Property Get curveDir() As CURVE_DIR
    curveDir = m_curveDir
End Property

' Radius of circular arc property
Public Property Get rad() As Double
    rad = m_rad
End Property

' Theta of center to start point line
Public Property Get cToStheta() As Double
    cToStheta = m_cToStheta
End Property

' Theta of center to start point line
Public Property Get cToEtheta() As Double
    cToEtheta = m_cToEtheta
End Property


