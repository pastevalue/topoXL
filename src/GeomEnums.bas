Attribute VB_Name = "GeomEnums"
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
'' Stores enumerations related to the geometry elements.
'' Stores functionality for working with the strings associated to
'' an enumeration
''========================================================================

'@Folder("TopoXL.geom")

Option Explicit
Option Private Module

' Curve direction enumeration --------------------------------------------

Private Const STR_CD_CW As String = "clockwise"
Private Const STR_CD_NONE As String = "none"
Private Const STR_CD_CCW As String = "counter-clockwise"

Public Enum CURVE_DIR
    CD_CW = 1                                    ' Clockwise
    CD_NONE = 0                                  ' None
    CD_CCW = -1                                  ' Counter-clockwise
End Enum

Public Function curveDirToString(curveDir As CURVE_DIR) As String
    Select Case curveDir
    Case CD_CW
        curveDirToString = STR_CD_CW
    Case CD_CCW
        curveDirToString = STR_CD_CCW
    Case CD_NONE
        curveDirToString = STR_CD_NONE
    Case Else
        Err.Raise 5, "Curve Direction to String function", _
                  "Curve direction enumeration doesn't include member: " & curveDir & "!"
    End Select
End Function

Public Function curveDirFromString(s As String) As CURVE_DIR
    Select Case LCase(s)
    Case STR_CD_CW
        curveDirFromString = CURVE_DIR.CD_CW
    Case STR_CD_CCW
        curveDirFromString = CURVE_DIR.CD_CCW
    Case STR_CD_NONE
        curveDirFromString = CURVE_DIR.CD_NONE
    Case Else
        Err.Raise 5, "Curve direction enumeration doesn't include a member named: " & s & "!"
    End Select
End Function

'-------------------------------------------------------------------------



