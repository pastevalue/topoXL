VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RedLineRange"
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
''    Modules: None
''    Classes: None
''Calls:
''    Modules: RLenums
''    Classes: None
''=======================================================
Option Explicit
Private Const NAME_PREFIX As String = "tblRL" 'use this constant to validate, by name prefix, that a table contains Red Line information
Private Const TYPE_STR As String = "Element Type"
Private Const COO_IN_X_STR As String = "Coo In X"
Private Const COO_IN_Y_STR As String = "Coo In Y"
Private Const COO_OUT_X_STR As String = "Coo Out X"
Private Const COO_OUT_Y_STR As String = "Coo Out Y"
Private Const COO_CENTER_X_STR As String = "Coo Center X"
Private Const COO_CENTER_Y_STR As String = "Coo Center Y"
Private Const COO_PEAK_X_STR As String = "Coo Peak X"
Private Const COO_PEAK_Y_STR As String = "Coo Peak Y"
Private Const LENGTH_STR As String = "Lengh"
Private Const RADIUS_STR As String = "Radius"
Private Const CURVE_DIRECTION_STR As String = "Curve Direction"
Private Const SPYRAL_TYPE_STR As String = "Spyral Type"

Private m_name As String

Private Enum RL_ELEM_PARTS
    PART_COO_IN_X
    PART_COO_IN_Y
    PART_COO_OUT_X
    PART_COO_OUT_Y
    PART_COO_CENTER_X
    PART_COO_CENTER_Y
    PART_COO_PEAK_X
    PART_COO_PEAK_Y
    PART_LENGTH
    PART_RADIUS
    PART_CURVE_DIRECTION
    PART_SPYRAL_TYPE
    PART_STATION
End Enum


Private Function getElemPartFromString(part As String) As RL_ELEM_PARTS
Select Case part
    Case COO_IN_X_STR
        getElemPartFromString = RL_ELEM_PARTS.PART_COO_IN_X
    Case COO_IN_Y_STR
        getElemPartFromString = RL_ELEM_PARTS.PART_COO_IN_Y
    Case COO_OUT_X_STR
        getElemPartFromString = RL_ELEM_PARTS.PART_COO_OUT_X
    Case COO_OUT_Y_STR
        getElemPartFromString = RL_ELEM_PARTS.PART_COO_OUT_Y
    Case COO_CENTER_X_STR
        getElemPartFromString = RL_ELEM_PARTS.PART_COO_CENTER_X
    Case COO_CENTER_Y_STR
        getElemPartFromString = RL_ELEM_PARTS.PART_COO_CENTER_Y
    Case COO_PEAK_X_STR
        getElemPartFromString = RL_ELEM_PARTS.PART_COO_PEAK_X
    Case COO_PEAK_Y_STR
        getElemPartFromString = RL_ELEM_PARTS.PART_COO_PEAK_Y
    Case LENGTH_STR
        getElemPartFromString = RL_ELEM_PARTS.PART_LENGTH
    Case RADIUS_STR
        getElemPartFromString = RL_ELEM_PARTS.PART_RADIUS
    Case CURVE_DIRECTION_STR
        getElemPartFromString = RL_ELEM_PARTS.PART_CURVE_DIRECTION
    Case SPYRAL_TYPE_STR
        getElemPartFromString = RL_ELEM_PARTS.PART_SPYRAL_TYPE
End Select
End Function


Private Sub Class_Initialize()
    m_name = vbNullString
End Sub

Public Function initFromRange(r As Range) As Boolean
    m_name = r.name
End Function

Private Function isRangeValid(r As Range) As Boolean
    isRangeValid = False
    
    'check if Range name is valid
    If Left(r.name, 5) <> NAME_PREFIX Then
        isRangeValid = False
        Exit Function
    End If
End Function


