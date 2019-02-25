VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MeasOffset"
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

''=======================================================
'' Description
'' Represents a location defined by linear reference. It
'' is described by:
''    - Measure: a distance, applied along a geometrical element,
''      starting from one of its ends (+/-, start/end)
''    - Offset: a perpendicular distance on the geometrical
''      element applied at the end of the "Measure"
''      (-/+, left/right)
''-------------------------------------------------------
'' Called by:
''    Modules: None
''    Classes: None
'' Calls:
''    Modules: None
''    Classes: None
''=======================================================

'@Folder("TopoXL.geom")
Option Explicit

Private m_m As Double                            'Measure: the measure of a distance on a geometric element
Private m_o As Double                            'Offset: the perpendicular distance between a point and a geometric element

'Class constructor
Private Sub Class_Initialize()
    m_m = 0
    m_o = 0
End Sub

'Class initialization
Public Sub Init(ByVal m As Double, ByVal o As Double)
    m_m = m
    m_o = o
End Sub

' Test if this MeasOffset is equivalent to another
' Returns TRUE if all members are equal
Public Function IsEqual(mo As MeasOffset) As Boolean
    IsEqual = Me.m = mo.m And Me.o = mo.o
End Function

' Returns a string representation of the Point class
Public Function ToString() As String
    ToString = "(" & m_m & "," & m_o & ")"
End Function

' Measure property
Public Property Get m() As Double
    m = m_m
End Property

Public Property Let m(m As Double)
    m_m = m
End Property

' Offset property
Public Property Get o() As Double
    o = m_o
End Property

Public Property Let o(o As Double)
    m_o = o
End Property

