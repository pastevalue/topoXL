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

''=======================================================
''Called by:
''    Modules: None
''    Classes: None
''Calls:
''    Modules: None
''    Classes: None
''=======================================================

'@Folder("TopoXL.geom")
Option Explicit

Private m_X As Double                            'Abscissa: distance measured parallel to the horizontal axis
Private m_Y As Double                            'Ordinate: distance measured parallel to the vertical axis

'Class constructor
Private Sub Class_Initialize()
    m_X = 0
    m_Y = 0
End Sub

'Class initialization
Public Sub Init(ByVal X As Double, ByVal Y As Double)
    m_X = X
    m_Y = Y
End Sub

'Returns a string representation of the Point class
Public Function ToString() As String
    ToString = "(" & m_X & "," & m_Y & ")"
End Function

'X property
Public Property Get X() As Double
    X = m_X
End Property

Public Property Let X(X As Double)
    m_X = X
End Property

'Y property
Public Property Get Y() As Double
    Y = m_Y
End Property

Public Property Let Y(Y As Double)
    m_Y = Y
End Property
