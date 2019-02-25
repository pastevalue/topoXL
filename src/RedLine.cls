VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RedLine"
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
''    Modules: RedLinesInit
''    Classes: RedLines
''Calls:
''    Modules: None
''    Classes: RLelement, Coo2D
''=======================================================
Option Explicit
Private m_name As String
Private m_elements As Collection
Private m_startSt As Double
Private m_endSt As Double

'constructor
Private Sub Class_Initialize()
    m_name = ""
    m_startSt = 0
    m_endSt = 0
    Set m_elements = New Collection
End Sub



'add new element to Red Line
Public Sub addElem(elem As RLelement)
    'add element
    m_elements.Add elem
    
    'update Red Line station
    If m_elements.count = 1 Then
        m_startSt = elem.startStation
        m_endSt = elem.endStation
    Else
        Call updateStations
    End If
End Sub

Private Sub updateStations()
    Dim tempElem As RLelement
    
    For Each tempElem In m_elements
        If tempElem.startStation < m_startSt Then m_startSt = tempElem.startStation
        If tempElem.endStation > m_endSt Then m_endSt = tempElem.endStation
    Next

End Sub


Public Sub setPosition(startPosition As Double)
    m_startSt = startPosition
    
    Dim rlElem As RLelement
    Dim length As Double
    length = 0
    For Each rlElem In m_elements
        length = length + rlElem.getLength
    Next rlElem
    m_endSt = m_startSt + length
End Sub

Public Function getCoo(station As Double, offset As Double) As Coo2D
    Dim rlElem As RLelement
    For Each rlElem In m_elements
        If station >= rlElem.startStation And station <= rlElem.endStation Then
            Set getCoo = rlElem.getCoo(station, offset)
            Exit Function
        End If
    Next rlElem
    
    'return error if no match for station has been found
    Set getCoo = Nothing
End Function

'Gets abscissa at a specified ordinate
Public Function getX(y As Double) As Double
    Dim rlElem As RLelement
    
    For Each rlElem In m_elements
            getX = rlElem.getX(y)
            If getX > 0 Then Exit Function
    Next rlElem
    getX = Null
End Function

'Gets ordinate at a specified abscissa
Public Function getY(x As Double) As Double
    Dim rlElem As RLelement

    For Each rlElem In m_elements
            getY = rlElem.getY(x)
            If getY > 0 Then Exit Function
    Next rlElem
    getY = Null
End Function

Public Function getPerpOrRadFromPoint(px As Double, py As Double, station As Double, distFromElem As Double) As Boolean
    Dim rlElem As RLelement
    Dim tempPosition As Double
    Dim tempDistFromElem As Double
    Dim firstPositionFound As Boolean
    
    firstPositionFound = False
    For Each rlElem In m_elements
        If rlElem.getPerpOrRadFromPoint(px, py, station, distFromElem) Then
            If firstPositionFound = False Or Abs(tempDistFromElem) > Abs(distFromElem) Then
                tempDistFromElem = distFromElem
                tempPosition = station
            End If
            firstPositionFound = True
        End If
    Next rlElem
    If firstPositionFound Then
        getPerpOrRadFromPoint = True
        distFromElem = tempDistFromElem
        station = tempPosition
    Else
        'return error if no match for station has been found
        station = 0
        distFromElem = 0
        getPerpOrRadFromPoint = False
    End If
End Function

'get first element from elements collections of red line
Public Function getFirstElement() As RLelement
    Set getFirstElement = m_elements.item(1)
End Function

'get last element from elements collections of red line
Public Function getLastElement() As RLelement
    Set getLastElement = m_elements.item(m_elements.count)
End Function


Public Function getLength() As Double
    Dim i As Integer
    getLength = 0
    For i = 1 To m_elements.count
        getLength = getLength + m_elements.item(i).getLength
    Next i
End Function

Public Function getElem(index As Integer) As RLelement
    Set getElem = m_elements.item(index)
End Function

Public Function getElements() As Collection
    Set getElements = m_elements
End Function

Public Function getElemCount() As Integer
    getElemCount = m_elements.count
End Function

Public Sub clear()
    Dim i As Integer
    For i = 1 To m_elements.count
        m_elements.remove i
    Next i
    m_startSt = 0
    m_endSt = 0
End Sub

Public Property Get name() As String
    name = m_name
End Property

Public Property Let name(n As String)
    m_name = n
End Property

Public Property Get startStation() As Double
    startStation = m_startSt
End Property

Public Property Get endStation() As Double
    endStation = m_endSt
End Property







