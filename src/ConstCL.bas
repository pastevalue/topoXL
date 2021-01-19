Attribute VB_Name = "ConstCL"
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
'' Stores constants, enumerations relevant for CL (Centerline) classes
'' This includes various functions to work with the constant values like:
'' toString, fromString, enumContains
'' Classes related to CL include geom classes
''========================================================================

'@Folder("TopoXL.CL")

Option Explicit
Option Private Module

' CL classes relevant constants ------------------------------------------
Public Const CL_MEASURE As String = "Measure"
Public Const CL_REVERSED As String = "Reversed"
' ------------------------------------------------------------------------

' Geometry classes relevant constants ------------------------------------
Public Const GEOM_TYPE As String = "GeomType"
Public Const GEOM_INIT_TYPE As String = "InitType"
' ------------------------------------------------------------------------

' LineSegment relevant constants -----------------------------------------
Public Const LS_NAME As String = "LineSegment"
Public Const LS_INIT_SE As String = "SE"

Public Const LS_M_START_X As String = "StartX"
Public Const LS_M_START_Y As String = "StartY"
Public Const LS_M_END_X As String = "EndX"
Public Const LS_M_END_Y As String = "EndY"
Public Const LS_M_DX As String = "dX"
Public Const LS_M_DY As String = "dY"
Public Const LS_M_LENGTH As String = "Length"
Public Const LS_M_THETA As String = "Theta"
' ------------------------------------------------------------------------

' CircularArc relevant constants -----------------------------------------
Public Const CA_NAME As String = "CircularArc"
Public Const CA_INIT_SCLD As String = "SCLD"
Public Const CA_INIT_SERD As String = "SERD"

Public Const CA_M_START_X As String = "StartX"
Public Const CA_M_START_Y As String = "StartY"
Public Const CA_M_CENTER_X As String = "CenterX"
Public Const CA_M_CENTER_Y As String = "CenterY"
Public Const CA_M_END_X As String = "EndX"
Public Const CA_M_END_Y As String = "EndY"
Public Const CA_M_LENGTH As String = "Length"
Public Const CA_M_RADIUS As String = "Radius"
Public Const CA_M_CURVE_DIR As String = "CurveDirection"
Public Const CA_M_START_T As String = "StartTheta"
Public Const CA_M_END_T As String = "EndTheta"
' ------------------------------------------------------------------------

' ClothoidArc relevant constants -----------------------------------------
Public Const CLA_NAME As String = "ClothoidArc"
Public Const CLA_INIT_SLRDT As String = "SLRDT"

Public Const CLA_M_START_X As String = "StartX"
Public Const CLA_M_START_Y As String = "StartY"
Public Const CLA_M_LENGTH As String = "Length"
Public Const CLA_M_END_RADIUS As String = "Radius"
Public Const CLA_M_CURVE_DIR As String = "CurveDirection"
Public Const CLA_M_START_T As String = "StartTheta"
' ------------------------------------------------------------------------

' Curve direction enumeration --------------------------------------------
Private Const STR_CD_CW As String = "CW"
Private Const STR_CD_NONE As String = "none"
Private Const STR_CD_CCW As String = "CCW"

Public Enum CURVE_DIR
    [_FIRST] = -2                                ' First index
    CD_CCW = -1                                  ' Counter-clockwise
    CD_NONE = 0                                  ' None
    CD_CW = 1                                    ' Clockwise
    [_LAST] = 2                                  ' Last index
End Enum

Public Function curveDirToString(ByVal curveDir As CURVE_DIR) As String
    Select Case curveDir
    Case CD_CW
        curveDirToString = STR_CD_CW
    Case CD_CCW
        curveDirToString = STR_CD_CCW
    Case Else
        curveDirToString = STR_CD_NONE
    End Select
End Function

' Returns
'   - CD_CCW or CD_CW enum element if the paramater is matched with
'     CD_CCW/CD_CW or their corresponding strings
'   - CD_None if no match is found
Public Function curveDirFromVariant(ByVal v As Variant) As CURVE_DIR
    On Error GoTo curveDirFromVariant
    If curveDirContains(v) Then                  ' try if value matches
        curveDirFromVariant = v
    Else                                         ' try if string matches
        Select Case LCase(v)
        Case LCase(STR_CD_CW)
            curveDirFromVariant = CURVE_DIR.CD_CW
        Case LCase(STR_CD_CCW)
            curveDirFromVariant = CURVE_DIR.CD_CCW
        Case Else
            curveDirFromVariant = CURVE_DIR.CD_NONE
        End Select
    End If
    Exit Function
curveDirFromVariant:
    curveDirFromVariant = CURVE_DIR.CD_NONE
End Function

Private Function curveDirContains(ByVal item As Variant) As Boolean
    Dim i As Integer

    On Error GoTo FailContains:
    curveDirContains = False
    For i = CURVE_DIR.[_FIRST] + 1 To CURVE_DIR.[_LAST] - 1
        If i = item Then
            curveDirContains = True
            Exit Function
        End If
    Next i
    Exit Function
FailContains:
    curveDirContains = False
End Function

'-------------------------------------------------------------------------




