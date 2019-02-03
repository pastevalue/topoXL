VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RLarcClothoid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''=======================================================
''Called by:
''    Modules: RedLinesInit
''    Classes: None
''Calls:
''    Modules: RLenums
''    Classes: Coo2D, ArcClothoid
''=======================================================
Option Explicit
Implements RLelement

Private m_Type As RL_ELEM_TYPES
Private m_clothoid  As ArcClothoid
Private m_spiralType As SPIRAL_TYPE
Private m_startPos As Double 'start station of the arc clothoid element
Private m_endPos As Double 'end station of the arc clothoid element

'Constructor
Private Sub Class_Initialize()
    m_Type = ELEM_ARC_CLOTHOID
    Set m_clothoid = New ArcClothoid
    m_spiralType = SPIRAL_TYPE.NONE
    m_startPos = 0
    m_endPos = 0
End Sub

'Initialize class
Public Sub init(c As ArcClothoid, spiralType As SPIRAL_TYPE, startStation As Double)
    Set m_clothoid = c
    m_spiralType = spiralType
    RLelement_setStations startStation
End Sub

Public Function initFromCollection(coll As Collection) As Boolean
    Dim tempVal As String
    Dim tempElemType As RL_ELEM_TYPES
    Dim tempInputType As RL_INPUT_TYPE
    On Error GoTo FailInit
    
    tempElemType = RLenums.rlElemTypeFromString(coll.item(RedLinesInit.TBL_ELEM_TYPE_COL))
    If tempElemType <> ELEM_ARC_CLOTHOID Then GoTo FailInit
    
    tempInputType = RLenums.rlElemInputTypeFromString(coll.item(RedLinesInit.TBL_ELEM_INPUT_TYPE_COL))
    
    Select Case tempInputType
        Case RL_INPUT_TYPE.INPUT_IN_PEAK_LEN_RAD_DIR_TYPE_ST

            'validate Coo In of arc clothoid
            Dim tempCooIn As New Coo2D
            If Not tempCooIn.initFromStrings(coll.item(RLenums.rlElemPartToString(RL_ELEM_PARTS.PART_COO_IN_X)), coll.item(RLenums.rlElemPartToString(RL_ELEM_PARTS.PART_COO_IN_Y))) _
            Then GoTo FailInit
            
            'validate Coo Peak of arc clothoid
            Dim tempCooPeak As New Coo2D
            If Not tempCooPeak.initFromStrings(coll.item(RLenums.rlElemPartToString(RL_ELEM_PARTS.PART_COO_PEAK_X)), coll.item(RLenums.rlElemPartToString(RL_ELEM_PARTS.PART_COO_PEAK_Y))) _
            Then GoTo FailInit
            
            'validate Length of arc clothoid
            Dim tempLength As Double
            tempVal = coll.item(RLenums.rlElemPartToString(RL_ELEM_PARTS.PART_LENGTH))
            If IsNumeric(tempVal) Then
                tempLength = tempVal
             Else
                Debug.Print "Can't read arc clothoid length! " & tempVal & " must be numeric value!"
                GoTo FailInit
            End If
            
            'validate Radius of arc clothoid
            Dim tempRadius As Double
            tempVal = coll.item(RLenums.rlElemPartToString(RL_ELEM_PARTS.PART_RADIUS))
            If IsNumeric(tempVal) Then
                tempRadius = tempVal
             Else
                Debug.Print "Can't read arc clothoid radius! " & tempVal & " must be numeric value!"
                GoTo FailInit
            End If
            
            'validate Curve Direction of arc clothoid
            Dim tempCurveDirection As CURVE_DIRECTION
            tempVal = coll.item(RLenums.rlElemPartToString(RL_ELEM_PARTS.PART_CURVE_DIRECTION))
            tempCurveDirection = RLenums.curveDirectionFromString(tempVal)
            If tempCurveDirection = CURVE_DIRECTION.NONE Then
                Debug.Print "Can't read arc clothoid Curve Direction! " & tempVal & " must be " & RLenums.curveDirectionToString(CLOCKWISE) & " or " & RLenums.curveDirectionToString(COUNTERCLOCKWISE)
                GoTo FailInit
            End If

            'validate Spiral Type of arc clothoid
            Dim tempSpiralType As SPIRAL_TYPE
            tempVal = coll.item(RLenums.rlElemPartToString(RL_ELEM_PARTS.PART_SPYRAL_TYPE))
            tempSpiralType = RLenums.spiralTypeFromString(tempVal)
            If tempSpiralType = SPIRAL_TYPE.NONE Then
                Debug.Print "Can't read arc clothoid Spiral Type! " & tempVal & " must be " & RLenums.spiralTypeToString(IN_CURVE) & " or " & RLenums.spiralTypeToString(OUT_CURVE)
                GoTo FailInit
            End If
            
            'validate Start Station of arc clothoid
            Dim tempStartPosition As Double
            tempVal = coll.item(RLenums.rlElemPartToString(RL_ELEM_PARTS.PART_STATION))
            If IsNumeric(tempVal) Then
                tempStartPosition = tempVal
             Else
                Debug.Print "Can't read Start Station! " & tempVal & " must be numeric values!"
                GoTo FailInit
            End If
             
            'init tempArcClothoid
            Dim tempArcClothoid As ArcClothoid
            Set tempArcClothoid = New ArcClothoid
            tempArcClothoid.init tempCooIn, tempCooPeak, tempLength, tempRadius, tempCurveDirection
            
            
            'initialize class
            Call init(tempArcClothoid, tempSpiralType, tempStartPosition)
            initFromCollection = True
            
        Case Else
              GoTo FailInit
    End Select

Exit Function
FailInit:
    initFromCollection = False
End Function

'sets stations of the arc clothoid
Private Sub RLelement_setStations(startStation As Double)
    m_startPos = startStation
    m_endPos = m_startPos + RLelement_getlength
End Sub

'gets coordinates at a specified station and offset from clothoid
Private Function RLelement_getCoo(station As Double, offset As Double) As Coo2D
    If m_spiralType = SPIRAL_TYPE.IN_CURVE Then
        Set RLelement_getCoo = m_clothoid.getCoo(station - m_startPos, offset)
    End If
    
    If m_spiralType = SPIRAL_TYPE.OUT_CURVE Then
        Set RLelement_getCoo = m_clothoid.getCoo(m_clothoid.getLength - (station - m_startPos), -offset)
    End If
End Function

'gets abscissa at a specified ordinate
Private Function RLelement_getX(y As Double) As Double
        RLelement_getX = m_clothoid.getX(y)
End Function

'gets ordinate at a specified abscissa
Private Function RLelement_getY(x As Double) As Double
        RLelement_getY = m_clothoid.getY(x)
End Function

Public Function RLelement_getPerpOrRadFromPoint(px As Double, py As Double, station As Double, distFromElem As Double) As Boolean
    If m_spiralType = SPIRAL_TYPE.IN_CURVE Then
        RLelement_getPerpOrRadFromPoint = m_clothoid.getRadialDistFromPoint(px, py, station, distFromElem)
        station = station + m_startPos
    End If
    
    If m_spiralType = SPIRAL_TYPE.OUT_CURVE Then
        RLelement_getPerpOrRadFromPoint = m_clothoid.getRadialDistFromPoint(px, py, station, distFromElem)
        station = m_startPos + (m_clothoid.getLength - station)
        distFromElem = -distFromElem
    End If
End Function

'getting and setting class members
Private Property Get RLelement_getElemType() As Double
    RLelement_getElemType = m_Type
End Property

Private Property Get RLelement_getlength() As Double
    RLelement_getlength = m_clothoid.getLength
End Property


'Start station property
Private Property Get RLelement_startStation() As Double
    RLelement_startStation = m_startPos
End Property

'End station property
Private Property Get RLelement_endStation() As Double
    RLelement_endStation = m_endPos
End Property

Private Property Get RLelement_startX() As Double
    RLelement_startX = m_clothoid.startCoo.x
End Property

Private Property Get RLelement_startY() As Double
    RLelement_startY = m_clothoid.startCoo.y
End Property

Private Property Get RLelement_endX() As Double
    RLelement_endX = m_clothoid.endCoo.x
End Property

Private Property Get RLelement_endY() As Double
    RLelement_endY = m_clothoid.endCoo.y
End Property

