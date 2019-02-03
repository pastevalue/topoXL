Attribute VB_Name = "RLenums"
''=======================================================
''Called by:
''    Modules: None
''    Classes: RLline, RLarcCircle, RLarcClothoid, RedLinesInit
''Calls:
''    Modules: IC_ENUMS
''    Classes: None
''=======================================================
Option Explicit
Option Private Module

'Red Line element type strings
Private Const STR_ELEM_NONE As String = "None"
Private Const STR_ELEM_LINE As String = "Line"
Private Const STR_ELEM_ARC_CIRCLE As String = "ArcCircle"
Private Const STR_ELEM_ARC_CLOTHOID As String = "ArcClothoid"

'Red Line element types enum
Public Enum RL_ELEM_TYPES
    ELEM_NONE = 0
    ELEM_LINE = 1
    ELEM_ARC_CIRCLE = 2
    ELEM_ARC_CLOTHOID = 3
    ELEM_COUNT = 4
End Enum


'Red Line input type strings
Private Const STR_INPUT_NONE As String = "None"
Private Const STR_INPUT_IN_OUT_ST As String = "InOutSt"
Private Const STR_INPUT_IN_CEN_LEN_DIR_ST As String = "InCenLenDirSt"
Private Const STR_INPUT_IN_OUT_RAD_DIR_ST As String = "InOutRadDirSt"
Private Const STR_INPUT_IN_PEAK_LEN_RAD_DIR_TYPE_ST As String = "InPeakLenRadDirTypeSt"

'Red Line input types enum
Public Enum RL_INPUT_TYPE
    INPUT_NONE = 0
    INPUT_IN_OUT_ST = 1
    INPUT_IN_CEN_LEN_DIR_ST = 2
    INPUT_IN_OUT_RAD_DIR_ST = 3
    INPUT_IN_PEAK_LEN_RAD_DIR_TYPE_ST = 4
    INPUT_COUNT = 5
End Enum


'Red Line Element part strings
Private Const STR_PART_NONE As String = "None"
Private Const STR_PART_COO_IN_X As String = "Coo In X"
Private Const STR_PART_COO_IN_Y As String = "Coo In Y"
Private Const STR_PART_COO_OUT_X As String = "Coo Out X"
Private Const STR_PART_COO_OUT_Y As String = "Coo Out Y"
Private Const STR_PART_COO_CENTER_X As String = "Coo Center X"
Private Const STR_PART_COO_CENTER_Y As String = "Coo Center Y"
Private Const STR_PART_COO_PEAK_X As String = "Coo Peak X"
Private Const STR_PART_COO_PEAK_Y As String = "Coo Peak Y"
Private Const STR_PART_LENGTH As String = "Length"
Private Const STR_PART_RADIUS As String = "Radius"
Private Const STR_PART_CURVE_DIRECTION As String = "Curve Direction"
Private Const STR_PART_SPYRAL_TYPE As String = "Spiral Type"
Private Const STR_PART_STATION As String = "Station"

'Red Line Element parts enum
Public Enum RL_ELEM_PARTS
    PART_NONE = 0
    PART_COO_IN_X = 1
    PART_COO_IN_Y = 2
    PART_COO_OUT_X = 3
    PART_COO_OUT_Y = 4
    PART_COO_CENTER_X = 5
    PART_COO_CENTER_Y = 6
    PART_COO_PEAK_X = 7
    PART_COO_PEAK_Y = 8
    PART_LENGTH = 9
    PART_RADIUS = 10
    PART_CURVE_DIRECTION = 11
    PART_SPYRAL_TYPE = 12
    PART_STATION = 13
    PART_COUNT = 14
End Enum

Public Enum CURVE_DIRECTION
    CLOCKWISE = 1
    NONE = 0
    COUNTERCLOCKWISE = -1
End Enum



Public Enum SPIRAL_TYPE
    IN_CURVE = 1
    NONE = 0
    OUT_CURVE = -1
End Enum

'get RL_ELEM_TYPES string
Public Function rlElemTypeToString(elemType As RL_ELEM_TYPES) As String
    Select Case elemType
        Case ELEM_LINE
            rlElemTypeToString = STR_ELEM_LINE
        Case ELEM_ARC_CIRCLE
            rlElemTypeToString = STR_ELEM_ARC_CIRCLE
        Case ELEM_ARC_CLOTHOID
            rlElemTypeToString = STR_ELEM_ARC_CLOTHOID
        Case Else
            rlElemTypeToString = STR_ELEM_NONE
    End Select
End Function

'returns an array of Red Line Element Type strings
Public Function getElemTypeStringArray() As Variant
    Dim i As RL_ELEM_TYPES
    Dim outArray(0 To RL_ELEM_TYPES.ELEM_COUNT - 2) As String
    For i = 1 To RL_ELEM_TYPES.ELEM_COUNT - 1
        outArray(i - 1) = rlElemTypeToString(i)
    Next i
    getElemTypeStringArray = outArray
End Function

'get RL_ELEM_TYPES from string
Public Function rlElemTypeFromString(s As String) As RL_ELEM_TYPES
    Dim arr As Variant
    Dim i As Integer
    
    arr = getElemTypeStringArray
    For i = 0 To UBound(arr)
        If arr(i) = s Then
            rlElemTypeFromString = i + 1
            Exit Function
        End If
    Next i
    rlElemTypeFromString = INPUT_NONE
End Function

'get RL_INPUT_TYPE string
Public Function rlElemInputTypeToString(elemPart As RL_ELEM_PARTS) As String
    Select Case elemPart
        Case INPUT_IN_OUT_ST
            rlElemInputTypeToString = STR_INPUT_IN_OUT_ST
        Case INPUT_IN_CEN_LEN_DIR_ST
            rlElemInputTypeToString = STR_INPUT_IN_CEN_LEN_DIR_ST
        Case INPUT_IN_OUT_RAD_DIR_ST
            rlElemInputTypeToString = STR_INPUT_IN_OUT_RAD_DIR_ST
        Case INPUT_IN_PEAK_LEN_RAD_DIR_TYPE_ST
            rlElemInputTypeToString = STR_INPUT_IN_PEAK_LEN_RAD_DIR_TYPE_ST
        Case Else
            rlElemInputTypeToString = STR_INPUT_NONE
    End Select
End Function

'get RL_INPUT_TYPE from string
Public Function rlElemInputTypeFromString(s As String) As RL_INPUT_TYPE
    Dim arr As Variant
    Dim i As Integer
    
    arr = getInputStringArray
    For i = 0 To UBound(arr)
        If arr(i) = s Then
            rlElemInputTypeFromString = i + 1
            Exit Function
        End If
    Next i
    rlElemInputTypeFromString = INPUT_NONE
End Function

'get RL_ELEM_PARTS string
Public Function rlElemPartToString(inputType As RL_INPUT_TYPE) As String
    Select Case inputType
        Case PART_COO_IN_X
            rlElemPartToString = STR_PART_COO_IN_X
        Case PART_COO_IN_Y
            rlElemPartToString = STR_PART_COO_IN_Y
        Case PART_COO_OUT_X
            rlElemPartToString = STR_PART_COO_OUT_X
        Case PART_COO_OUT_Y
            rlElemPartToString = STR_PART_COO_OUT_Y
        Case PART_COO_CENTER_X
            rlElemPartToString = STR_PART_COO_CENTER_X
        Case PART_COO_CENTER_Y
            rlElemPartToString = STR_PART_COO_CENTER_Y
        Case PART_COO_PEAK_X
            rlElemPartToString = STR_PART_COO_PEAK_X
        Case PART_COO_PEAK_Y
            rlElemPartToString = STR_PART_COO_PEAK_Y
        Case PART_LENGTH
            rlElemPartToString = STR_PART_LENGTH
        Case PART_RADIUS
            rlElemPartToString = STR_PART_RADIUS
        Case PART_CURVE_DIRECTION
            rlElemPartToString = STR_PART_CURVE_DIRECTION
        Case PART_SPYRAL_TYPE
            rlElemPartToString = STR_PART_SPYRAL_TYPE
         Case PART_STATION
            rlElemPartToString = STR_PART_STATION
        Case Else
            rlElemPartToString = STR_PART_NONE
    End Select
End Function
'returns an array of Red Line Element Parts strings
Public Function getElemPartStringArray() As Variant
    Dim i As RL_ELEM_PARTS
    Dim outArray(0 To RL_ELEM_PARTS.PART_COUNT - 2) As String
    For i = 1 To RL_ELEM_PARTS.PART_COUNT - 1
        outArray(i - 1) = rlElemPartToString(i)
    Next i
    getElemPartStringArray = outArray
End Function

'returns an array of Red Line Element Input Type strings
Public Function getInputStringArray() As Variant
    Dim i As RL_INPUT_TYPE
    Dim outArray(0 To RL_INPUT_TYPE.INPUT_COUNT - 2) As String
    For i = 1 To RL_INPUT_TYPE.INPUT_COUNT - 1
        outArray(i - 1) = rlElemInputTypeToString(i)
    Next i
    getInputStringArray = outArray
End Function

Public Function curveDirectionToString(curveDir As CURVE_DIRECTION) As String
    Select Case curveDir
        Case CLOCKWISE
            curveDirectionToString = "Clockwise"
        Case COUNTERCLOCKWISE
            curveDirectionToString = "Counterclockwise"
        Case Else
            curveDirectionToString = "None"
    End Select
End Function

Public Function curveDirectionFromString(s As String) As CURVE_DIRECTION
    Select Case s
        Case "Clockwise"
            curveDirectionFromString = CURVE_DIRECTION.CLOCKWISE
        Case "Counterclockwise"
            curveDirectionFromString = CURVE_DIRECTION.COUNTERCLOCKWISE
        Case Else
            curveDirectionFromString = CURVE_DIRECTION.NONE
    End Select
End Function

Public Function spiralTypeToString(spiralType As SPIRAL_TYPE) As String
    Select Case spiralType
        Case IN_CURVE
            spiralTypeToString = "InCurve"
        Case OUT_CURVE
            spiralTypeToString = "OutCurve"
        Case Else
            spiralTypeToString = "None"
    End Select
End Function

Public Function spiralTypeFromString(s As String) As SPIRAL_TYPE
    Select Case s
    Case "InCurve"
        spiralTypeFromString = SPIRAL_TYPE.IN_CURVE
    Case "OutCurve"
        spiralTypeFromString = SPIRAL_TYPE.OUT_CURVE
    Case Else
        spiralTypeFromString = SPIRAL_TYPE.NONE
    End Select
End Function

