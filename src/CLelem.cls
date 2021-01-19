VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CLelem"
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
'' Represents a wrapper for geom classes
'' Has members for start and end measure. This allows for geometry
'' measures to have a defined starting position by convention:
''  - start measure = 0 of the geom element
''  - end measure = start measure + geom length
'' Reversed geometry means that start measure is applied at the end
'' of the geometry (m_geom). This is to deal with ClothoidArc geometries
'' which are designed to exit a curve
''========================================================================

'@Folder("TopoXL.CL")

Option Explicit

Private m_geom As IGeom
Private m_startM As Double                       ' Start measure
Private m_endM As Double                         ' End measure
Private m_reversed As Boolean                    ' Reversed geometry

' Class constructor with default values
Private Sub Class_Initialize()
    Set m_geom = Nothing
    m_startM = 0
    m_endM = 0
    m_reversed = False
End Sub

' Initialize by geometry object and start measure
' Raise error if geom parameter is Nothing
Public Sub init(geom As IGeom, ByVal startM As Double, Optional reversed As Boolean = False)

    ' Raise error if start and center coordinates are equal
    If geom Is Nothing Then
        Err.Raise 5, TypeName(Me), "Can't initiate CLelem with Nothing geom!"
    End If
    Set m_geom = geom
    m_startM = startM
    m_endM = startM + geom.length
    m_reversed = reversed
End Sub

' Sets a new value to start measure
Public Sub setMeasures(ByVal startM As Double)
    Me.init m_geom, startM
End Sub

' Returns calcPointByMeasOffset of geom member for the assigned start measure
Public Function calcPointByMeasOffset(ByVal m As Double, ByVal o As Double) As Point
    ' If geometry is reversed then call calcPointByMeasOffset of the geom obj with adjusted values
    If m_reversed Then
        Set calcPointByMeasOffset = m_geom.calcPointByMeasOffset(m_geom.length - (m - m_startM), -o)
    Else
        Set calcPointByMeasOffset = m_geom.calcPointByMeasOffset(m - m_startM, o)
    End If
End Function

' Returns calcMeasOffsetOfPoint of geom member for the assigned start measure
Public Function calcMeasOffsetOfPoint(ByVal x As Double, ByVal y As Double) As MeasOffset
    Set calcMeasOffsetOfPoint = m_geom.calcMeasOffsetOfPoint(x, y)
    If Not calcMeasOffsetOfPoint Is Nothing Then
        If m_reversed Then
            calcMeasOffsetOfPoint.m = m_startM + m_geom.length - calcMeasOffsetOfPoint.m
            calcMeasOffsetOfPoint.o = -calcMeasOffsetOfPoint.o
        Else
            calcMeasOffsetOfPoint.m = calcMeasOffsetOfPoint.m + m_startM
        End If
    End If
End Function

' Start measure of the element
Public Property Get startM() As Double
    startM = m_startM
End Property

' End measure of the element
Public Property Get endM() As Double
    endM = m_endM
End Property

' Length of the element
Public Property Get length() As Double
    length = m_geom.length
End Property

' Geometry type name
Public Property Get geomTypeName() As String
    If m_geom Is Nothing Then
        geomTypeName = "None"
    Else
        geomTypeName = m_geom.geomTypeName
    End If
End Property
