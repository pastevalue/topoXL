VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RLarcCircle"
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
''    Modules: RLenums, UDF_COGO
''    Classes: Coo2D, ArcCircle
''=======================================================
Option Explicit
Implements RLelement
Private m_Type As RL_ELEM_TYPES
Private m_arcCircle As ArcCircle
Private m_startSt As Double 'start station
Private m_endSt As Double 'end station


'constructor
Private Sub Class_Initialize()
    m_Type = ELEM_ARC_CIRCLE
    Set m_arcCircle = New ArcCircle
    m_startSt = 0
    m_endSt = 0
End Sub

'initialize class
Public Sub Init(ac As ArcCircle, startStation As Double)
    Set m_arcCircle = ac
    RLelement_setStations startStation
End Sub


Private Function RLelement_initFromSepValues(inputType As RL_INPUT_TYPE, separator As String, values As String) As Boolean
    
End Function

Public Function initFromCollection(coll As Collection) As Boolean
    Dim tempVal As String
    Dim tempElemType As RL_ELEM_TYPES
    Dim tempInputType As RL_INPUT_TYPE
    
    Dim tempCooIn As Coo2D
    Dim tempCooOut As Coo2D
    Dim tempCooCenter As Coo2D
    Dim tempLength As Double
    Dim tempCurveDirection As CURVE_DIRECTION
    Dim tempStartStation As Double
    
    Dim tempArcCircle As ArcCircle
    
    On Error GoTo FailInit
    
    tempElemType = RLenums.rlElemTypeFromString(coll.item(RedLinesInit.TBL_ELEM_TYPE_COL))
    If tempElemType <> ELEM_ARC_CIRCLE Then GoTo FailInit
    
    tempInputType = RLenums.rlElemInputTypeFromString(coll.item(RedLinesInit.TBL_ELEM_INPUT_TYPE_COL))
    
    Select Case tempInputType
        Case RL_INPUT_TYPE.INPUT_IN_CEN_LEN_DIR_ST

            'validate Coo In
            Set tempCooIn = New Coo2D
            If Not tempCooIn.initFromStrings(coll.item(RLenums.rlElemPartToString(RL_ELEM_PARTS.PART_COO_IN_X)), coll.item(RLenums.rlElemPartToString(RL_ELEM_PARTS.PART_COO_IN_Y))) _
            Then GoTo FailInit
            
            'validate Coo Center
            Set tempCooCenter = New Coo2D
            If Not tempCooCenter.initFromStrings(coll.item(RLenums.rlElemPartToString(RL_ELEM_PARTS.PART_COO_CENTER_X)), coll.item(RLenums.rlElemPartToString(RL_ELEM_PARTS.PART_COO_CENTER_Y))) _
            Then GoTo FailInit
    
            'validate length of arc circle
            tempVal = coll.item(RLenums.rlElemPartToString(RL_ELEM_PARTS.PART_LENGTH))
            If IsNumeric(tempVal) Then
                tempLength = tempVal
             Else
                Debug.Print "Can't read arc circle length! " & tempVal & " must be numeric value!"
                GoTo FailInit
            End If
            If tempLength > 2 * PI * cogoDistance2D(tempCooIn.x, tempCooIn.y, tempCooCenter.x, tempCooCenter.y) Then
                Debug.Print "Arc length connot be higher than 2*pi()*radius = " & 2 * PI * cogoDistance2D(tempCooIn.x, tempCooIn.y, tempCooCenter.x, tempCooCenter.y)
                GoTo FailInit
            End If
            
            'validate Curve Direction of arc circle
            tempVal = coll.item(RLenums.rlElemPartToString(RL_ELEM_PARTS.PART_CURVE_DIRECTION))
            tempCurveDirection = RLenums.curveDirectionFromString(tempVal)
            If tempCurveDirection = CURVE_DIRECTION.NONE Then
                Debug.Print "Can't read arc circle curve direction! " & tempVal & " must be " & RLenums.curveDirectionToString(CLOCKWISE) & " or " & RLenums.curveDirectionToString(COUNTERCLOCKWISE)
                GoTo FailInit
            End If
            
            'validate start station
            tempVal = coll.item(RLenums.rlElemPartToString(RL_ELEM_PARTS.PART_STATION))
            If IsNumeric(tempVal) Then
                tempStartStation = tempVal
             Else
                Debug.Print "Can't read start station! " & tempVal & " must be numeric values!"
                GoTo FailInit
            End If
            
            'init tempArcCircle
            Set tempArcCircle = New ArcCircle
            tempArcCircle.initFromInCenLenDir tempCooIn, tempCooCenter, tempLength, tempCurveDirection
            
            'initialize class
            Call Init(tempArcCircle, tempStartStation)
            initFromCollection = True
            
        Case RL_INPUT_TYPE.INPUT_IN_OUT_RAD_DIR_ST
            
            'validate Coo In
            Set tempCooIn = New Coo2D
            If Not tempCooIn.initFromStrings(coll.item(RLenums.rlElemPartToString(RL_ELEM_PARTS.PART_COO_IN_X)), coll.item(RLenums.rlElemPartToString(RL_ELEM_PARTS.PART_COO_IN_Y))) _
            Then GoTo FailInit
                
            'validate Coo Out
            Set tempCooOut = New Coo2D
            If Not tempCooOut.initFromStrings(coll.item(RLenums.rlElemPartToString(RL_ELEM_PARTS.PART_COO_OUT_X)), coll.item(RLenums.rlElemPartToString(RL_ELEM_PARTS.PART_COO_OUT_Y))) _
            Then GoTo FailInit
                
            'validate Radius of arc circle
            Dim tempRadius As Double
            tempVal = coll.item(RLenums.rlElemPartToString(RL_ELEM_PARTS.PART_RADIUS))
            If IsNumeric(tempVal) Then
                tempRadius = tempVal
             Else
                Debug.Print "Can't read arc circle radius! " & tempVal & " must be numeric value!"
                GoTo FailInit
            End If
                
            'validate Curve Direction of arc circle
            tempVal = coll.item(RLenums.rlElemPartToString(RL_ELEM_PARTS.PART_CURVE_DIRECTION))
            tempCurveDirection = RLenums.curveDirectionFromString(tempVal)
            If tempCurveDirection = CURVE_DIRECTION.NONE Then
                Debug.Print "Can't read arc circle curve direction! " & tempVal & " must be " & RLenums.curveDirectionToString(CLOCKWISE) & " or " & RLenums.curveDirectionToString(COUNTERCLOCKWISE)
                GoTo FailInit
            End If
    
            'validate start station
            tempVal = coll.item(RLenums.rlElemPartToString(RL_ELEM_PARTS.PART_STATION))
            If IsNumeric(tempVal) Then
                tempStartStation = tempVal
             Else
                Debug.Print "Can't read start station! " & tempVal & " must be numeric values!"
                GoTo FailInit
            End If
            
            'init tempArcCircle
            Set tempArcCircle = New ArcCircle
            tempArcCircle.initFromInOutRadDir tempCooIn, tempCooOut, tempRadius, tempCurveDirection
            
            'initialize class
            Call Init(tempArcCircle, tempStartStation)
            initFromCollection = True
            
        Case Else
              GoTo FailInit
    End Select

Exit Function
FailInit:
    initFromCollection = False
End Function

'sets stations of the arc circle
Private Sub RLelement_setStations(startStation As Double)
    m_startSt = startStation
    m_endSt = m_startSt + RLelement_getlength
End Sub

'gets coordinates at a specified station and offset from arc circle
Private Function RLelement_getCoo(station As Double, offset As Double) As Coo2D
    Set RLelement_getCoo = m_arcCircle.getCoo(station - m_startSt, offset)
End Function

'gets abscissa at a specified ordinate
Private Function RLelement_getX(y As Double) As Double
    RLelement_getX = m_arcCircle.getX(y)
End Function

'gets ordinate at a specified abscissa
Private Function RLelement_getY(x As Double) As Double
    RLelement_getY = m_arcCircle.getY(x)
End Function

Public Function RLelement_getPerpOrRadFromPoint(px As Double, py As Double, station As Double, distFromElem As Double) As Boolean
    RLelement_getPerpOrRadFromPoint = m_arcCircle.getRadialDistFromPoint(px, py, station, distFromElem)
    station = station + m_startSt
End Function

'getting and setting class members
Private Property Get RLelement_getElemType() As Double
    RLelement_getElemType = m_Type
End Property


Private Property Get RLelement_getlength() As Double
    RLelement_getlength = m_arcCircle.getLength
End Property

Private Property Get RLelement_startStation() As Double
    RLelement_startStation = m_startSt
End Property

Private Property Get RLelement_endStation() As Double
    RLelement_endStation = m_endSt
End Property

Private Property Get RLelement_startX() As Double
    RLelement_startX = m_arcCircle.startCoo.x
End Property

Private Property Get RLelement_startY() As Double
    RLelement_startY = m_arcCircle.startCoo.y
End Property

Private Property Get RLelement_endX() As Double
    RLelement_endX = m_arcCircle.endCoo.x
End Property

Private Property Get RLelement_endY() As Double
    RLelement_endY = m_arcCircle.endCoo.y
End Property
