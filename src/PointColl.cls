VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PointColl"
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
'' Represents a collection of Point class instances
''========================================================================

'@Folder("TopoXL.CL.geom")

Option Explicit

Private m_pnts As Collection

' Class constructor with default values
Private Sub Class_Initialize()
    Set m_pnts = New Collection
End Sub

' Adds a new Point to the collection
Public Sub add(p As Variant)
    If IsObject(p) And TypeOf p Is Point Then
        m_pnts.add p
    Else
        Err.Raise 5, VBA.TypeName(Me), "Can't add item to Point Collection. TypeName must be Point!"

    End If
End Sub

' Test if this PointColl is equivalent to another
' Returns TRUE if all Points are equal
Public Function isEqual(pc As PointColl) As Boolean
    ' Return False if count is not the same
    If m_pnts.count <> pc.count Then
        isEqual = False
        Exit Function
    End If
    
    Dim i As Long
    For i = 1 To m_pnts.count
        If Not m_pnts.item(i).isEqual(pc.getItem(i)) Then
            isEqual = False
            Exit Function
        End If
    Next i
    isEqual = True                               ' Return True if core reaches this point
End Function

' Returns an array (1 based index) of the X and Y coordinate of all points in the collection
' Parameters:
'   - colWise: specifies if array is returned as column wise: each point has its own row
Public Function toArrayXY(Optional ByVal colWise As Boolean = False) As Variant
    Dim r() As Double
    Dim i As Long
    If colWise Then
        ReDim r(1 To m_pnts.count, 1 To 2)
        For i = 1 To m_pnts.count
            r(i, 1) = m_pnts.item(i).x
            r(i, 2) = m_pnts.item(i).y
        Next i
    Else
        ReDim r(1 To 2 * m_pnts.count)
        For i = 1 To m_pnts.count
            r(i * 2 - 1) = m_pnts.item(i).x
            r(i * 2) = m_pnts.item(i).y
        Next i
    End If
    toArrayXY = r
End Function

' Item property
Public Property Get getItem(ByVal idx As Long) As Point
    Set getItem = m_pnts.item(idx)
End Property

' Count property
Public Property Get count() As Long
    count = m_pnts.count
End Property



