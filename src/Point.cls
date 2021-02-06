VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Point"
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
'' Represents a location defined by a set of 3D grid coordinates
''========================================================================

'@Folder("TopoXL.CL")

Option Explicit

Private m_X As Double                            'Abscissa: distance measured parallel to the horizontal axis
Private m_Y As Double                            'Ordinate: distance measured parallel to the vertical axis
Private m_Z As Double                            'Ordinate: distance measured from XY plan

' Class constructor
Private Sub Class_Initialize()
    m_X = 0
    m_Y = 0
    m_Z = 0
End Sub

' Class initialization
Public Sub init(ByVal x As Double, ByVal y As Double, Optional ByVal z As Double = 0#)
    m_X = x
    m_Y = y
    m_Z = z
End Sub

' Test if this Point is equivalent to another
' Returns TRUE if all members are equal
Public Function isEqual(p As Point) As Boolean
    isEqual = m_X = p.x And m_Y = p.y And m_Z = p.z
End Function

' Returns a string representation of the Point class (2D - X and Y only)
Public Function toStringXY() As String
    toStringXY = "(" & m_X & "," & m_Y & ")"
End Function

' Returns a string representation of the Point class (3D - X,Y and Z)
Public Function toStringXYZ() As String
    toStringXYZ = "(" & m_X & "," & m_Y & "," & m_Z & ")"
End Function

' Returns an array of the X and Y coordinates
Public Function toArrayXY() As Variant
    toArrayXY = Array(m_X, m_Y)
End Function

' X property
Public Property Get x() As Double
    x = m_X
End Property

Public Property Let x(ByVal x As Double)
    m_X = x
End Property

' Y property
Public Property Get y() As Double
    y = m_Y
End Property

Public Property Let y(ByVal y As Double)
    m_Y = y
End Property

' Z property
Public Property Get z() As Double
    z = m_Z
End Property

Public Property Let z(ByVal z As Double)
    m_Z = z
End Property


