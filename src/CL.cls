VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CL"
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
'' Represents a centerline (axis) object which contains a "geom" object
'' collection.
''========================================================================

'@Folder("TopoXL.CL")
Option Explicit

Private m_name As String                         ' Axis name
Private m_elems As Collection                    ' Axis elements

' Class constructor with default values
Private Sub Class_Initialize()
    m_name = vbNullString
    Set m_elems = New Collection
End Sub

Public Sub init(ByVal name As String)
    m_name = name
End Sub

' Add a new element to this Axis
Public Sub addElem(elem As CLelem)
    m_elems.Add elem
End Sub

Public Function calcPointByMeasOffset(ByVal m As Double, ByVal o As Double) As Point
    Dim elem As CLelem
    Set calcPointByMeasOffset = Nothing          ' set result to Nothing
    For Each elem In m_elems
        Set calcPointByMeasOffset = elem.calcPointByMeasOffset(m, o)
        If Not calcPointByMeasOffset Is Nothing Then Exit Function
    Next elem
End Function

Public Function calcMeasOffsetOfPoint(ByVal x As Double, ByVal y As Double) As MeasOffset
    Dim elem As CLelem
    Dim currentMO As MeasOffset
    
    Set calcMeasOffsetOfPoint = Nothing
    ' Loop all elements and calc MeasOffset. Assign result to smallest offset found
    For Each elem In m_elems
        Set currentMO = elem.calcMeasOffsetOfPoint(x, y)
        ' If MeasureOffset valid
        If Not currentMO Is Nothing Then
            If calcMeasOffsetOfPoint Is Nothing Then ' Set result if first valid obj found
                Set calcMeasOffsetOfPoint = currentMO
            Else
                ' Set result if offset of current is less than offset of result
                If Abs(currentMO.o) < Abs(calcMeasOffsetOfPoint.o) Then Set calcMeasOffsetOfPoint = currentMO
            End If
        End If
    Next elem
End Function

Public Property Get name() As String
    name = m_name
End Property

Public Property Let name(ByVal n As String)
    m_name = n
End Property

