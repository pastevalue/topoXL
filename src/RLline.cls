VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RLline"
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
''    Classes: None
''Calls:
''    Modules: RLenums
''    Classes: Coo2D, Line2D
''=======================================================
Option Explicit
Implements RLelement

Private m_Type As RL_ELEM_TYPES
Private m_Line2D As Line2D
Private m_startSt As Double 'start station
Private m_endSt As Double ' end station

'Class constructor
Private Sub Class_Initialize()
    m_Type = RL_ELEM_TYPES.ELEM_LINE
    Set m_Line2D = New Line2D
    m_startSt = 0
    m_endSt = 0
End Sub

'Initialize class
Public Sub init(l As Line2D, startStation As Double)
    Set m_Line2D = l
    RLelement_setStations startStation
End Sub


Public Function initFromCollection(coll As Collection) As Boolean
    Dim tempVal As String
    Dim tempElemType As RL_ELEM_TYPES
    Dim tempInputType As RL_INPUT_TYPE
    
    On Error GoTo FailInit
    
    tempElemType = RLenums.rlElemTypeFromString(coll.item(RedLinesInit.TBL_ELEM_TYPE_COL))
    If tempElemType <> ELEM_LINE Then GoTo FailInit
    
    tempInputType = RLenums.rlElemInputTypeFromString(coll.item(RedLinesInit.TBL_ELEM_INPUT_TYPE_COL))
    
    Select Case tempInputType
        Case RL_INPUT_TYPE.INPUT_IN_OUT_ST
            'validate Coo In
            Dim tempCooIn As New Coo2D
            If Not tempCooIn.initFromStrings(coll.item(RLenums.rlElemPartToString(RL_ELEM_PARTS.PART_COO_IN_X)), coll.item(RLenums.rlElemPartToString(RL_ELEM_PARTS.PART_COO_IN_Y))) _
            Then GoTo FailInit
            
            'validate Coo Center
            Dim tempCooOut As New Coo2D
            If Not tempCooOut.initFromStrings(coll.item(RLenums.rlElemPartToString(RL_ELEM_PARTS.PART_COO_OUT_X)), coll.item(RLenums.rlElemPartToString(RL_ELEM_PARTS.PART_COO_OUT_Y))) _
            Then GoTo FailInit
            
            'validate start station
            Dim tempStartStation As Double
            tempVal = coll.item(RLenums.rlElemPartToString(RL_ELEM_PARTS.PART_STATION))
            If IsNumeric(tempVal) Then
                tempStartStation = tempVal
             Else
                Debug.Print "Can't read start station! " & tempVal & " must be numeric values!"
                GoTo FailInit
            End If
            
            
            'init tempLine2D element
            Dim tempLine2D As New Line2D
            tempLine2D.init tempCooIn, tempCooOut
            
            'initialize class
            Call init(tempLine2D, tempStartStation)
            initFromCollection = True
            
        Case Else
              GoTo FailInit
    End Select

Exit Function
FailInit:
    initFromCollection = False
End Function


Public Function getLineAzimuth() As Double
    getLineAzimuth = m_Line2D.getAzimuth
End Function

'Implement RLelement interface
'Gets coordinates at a specified station and offset from line
Private Function RLelement_getCoo(station As Double, offset As Double) As Coo2D
    Set RLelement_getCoo = m_Line2D.getCoo(station - m_startSt, offset)
End Function

'Gets abscissa at a specified ordinate
Private Function RLelement_getX(y As Double) As Double
    RLelement_getX = m_Line2D.getX(y)
End Function

'Gets ordinate at a specified abscissa
Private Function RLelement_getY(x As Double) As Double
    RLelement_getY = m_Line2D.getY(x)
End Function

'Gets station and perpendicular distance from a specified set of coordinates to the RLline
Public Function RLelement_getPerpOrRadFromPoint(px As Double, py As Double, station As Double, distFromElem As Double) As Boolean
    RLelement_getPerpOrRadFromPoint = m_Line2D.getPerpDistanceFromPoint(px, py, station, distFromElem)
    station = station + m_startSt
End Function

Private Property Get RLelement_getElemType() As Double
    RLelement_getElemType = m_Type
End Property

Private Property Get RLelement_getlength() As Double
    RLelement_getlength = m_Line2D.getLength
End Property

Private Sub RLelement_setStations(startStation As Double)
    m_startSt = startStation
    m_endSt = m_startSt + RLelement_getlength
End Sub

Private Property Get RLelement_startStation() As Double
    RLelement_startStation = m_startSt
End Property

Private Property Get RLelement_endStation() As Double
    RLelement_endStation = m_endSt
End Property

Private Property Get RLelement_startX() As Double
    RLelement_startX = m_Line2D.startCoo.x
End Property

Private Property Get RLelement_startY() As Double
    RLelement_startY = m_Line2D.startCoo.y
End Property

Private Property Get RLelement_endX() As Double
    RLelement_endX = m_Line2D.endCoo.x
End Property

Private Property Get RLelement_endY() As Double
    RLelement_endY = m_Line2D.endCoo.y
End Property


