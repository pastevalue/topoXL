VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IGeom"
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
'' Represents an interface to be implemented by all geom classes which
'' are used to definea CL (Centerline) object
''========================================================================

'@Folder("TopoXL.CL.geom")

Option Explicit

Public Function calcPointByMeasOffset(ByVal m As Double, ByVal o As Double) As Point
End Function

Public Function calcMeasOffsetOfPoint(ByVal x As Double, ByVal y As Double) As MeasOffset
End Function

Public Function calcXatY(ByVal y As Double) As Variant
End Function

Public Function calcYatX(ByVal x As Double) As Variant
End Function

Public Property Get length() As Double
End Property

Public Property Get geomTypeName() As String
End Property

