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

''===================================================================
'' Description
'' Represents a location defined by a set of 2D grid
'' coordinates
''-------------------------------------------------------
'' Called by:
''    Modules: None
''    Classes: None
'' Calls:
''    Modules: None
''    Classes: None
''=======================================================

'@Folder("TopoXL.CL")
Option Explicit

Private m_X As Double                            'Abscissa: distance measured parallel to the horizontal axis
Private m_Y As Double                            'Ordinate: distance measured parallel to the vertical axis

' Class constructor
Private Sub Class_Initialize()
    m_X = 0
    m_Y = 0
End Sub

' Class initialization
Public Sub init(ByVal x As Double, ByVal y As Double)
    m_X = x
    m_Y = y
End Sub

' Test if this Point is equivalent to another
' Returns TRUE if all members are equal
Public Function isEqual(p As Point) As Boolean
    isEqual = m_X = p.x And m_Y = p.y
End Function

' Returns a string representation of the Point class
Public Function toString() As String
    toString = "(" & m_X & "," & m_Y & ")"
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

