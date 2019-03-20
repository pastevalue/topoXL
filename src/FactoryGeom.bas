Attribute VB_Name = "FactoryGeom"
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
'' Factory module used to create geometry related
'' classes instances (Point, LineSegment, CircularArc)
''========================================================================

'@Folder("TopoXL.CL.geom")

Option Explicit
Option Private Module

' Creates a Point from a pair of grid coordinates
Public Function newPnt(ByVal x As Double, ByVal y As Double) As Point
    Set newPnt = New Point
    newPnt.init x, y
End Function

' Creates a Point from Variant values
' Returns:
'   - a new Point object
'   - Nothing if conversion of Variant to Double fails
Public Function newPntVar(ByVal x As Variant, ByVal y As Variant) As Point
    On Error GoTo ErrHandler
    Set newPntVar = newPnt(CDbl(x), CDbl(y))
    Exit Function
ErrHandler:
    Set newPntVar = Nothing
End Function

' Creates a MeasOffset from a measure distance and an offset
Public Function newMO(ByVal m As Double, ByVal o As Double) As MeasOffset
    Set newMO = New MeasOffset
    newMO.init m, o
End Function

' Creates a MeasOffset from Variant values
' Returns:
'  - a new MeasOffset object
'  - Nothing if conversion of Variant to Double fails
Public Function NewMOvar(ByVal m As Variant, ByVal o As Variant) As MeasOffset
    On Error GoTo ErrHandler
    Set NewMOvar = newMO(CDbl(m), CDbl(o))
    Exit Function
ErrHandler:
    Set NewMOvar = Nothing
End Function

' Creates a LineSegment from two sets of grid coordinates
Public Function newLnSeg(ByVal x1 As Double, ByVal y1 As Double, _
                         ByVal x2 As Double, ByVal y2 As Double) As LineSegment
    Set newLnSeg = New LineSegment
    newLnSeg.init x1, y1, x2, y2
End Function

' Creates a LineSegment from two sets of grid coordinates defined
' as Variant values
' Returns:
'   - a new LineSegment object
'   - Nothing if conversion of Variant to Double fails
Public Function newLnSegVar(ByVal x1 As Variant, ByVal y1 As Variant, _
                            ByVal x2 As Variant, ByVal y2 As Variant) As LineSegment
    On Error GoTo ErrHandler
    Set newLnSegVar = newLnSeg(CDbl(x1), CDbl(y1), CDbl(x2), CDbl(y2))
    Exit Function
ErrHandler:
    Set newLnSegVar = Nothing
End Function

' Creates a LineSegment from a key: value list (collection)
' Keys are defined in ConstCL
' Returns:
'   - a new LineSegment object
'   - Nothing if creation fails: wrong values or missing keys
Public Function newLnSegColl(coll As Collection) As LineSegment
    On Error GoTo FailNewLS
    
    If coll.item(ConstCL.GEOM_TYPE) <> ConstCL.LS_NAME Then GoTo FailNewLS
    
    ' Select init type
    Select Case coll.item(ConstCL.GEOM_INIT_TYPE)
        ' Case Start and End point
    Case ConstCL.LS_INIT_SE
        Set newLnSegColl = newLnSegVar(coll.item(ConstCL.LS_M_START_X), _
                                       coll.item(ConstCL.LS_M_START_Y), _
                                       coll.item(ConstCL.LS_M_END_X), _
                                       coll.item(ConstCL.LS_M_END_Y))
    Case Else
        GoTo FailNewLS
    End Select

    Exit Function
FailNewLS:
    Set newLnSegColl = Nothing
End Function

' Creates a CircularArc from Start and Center point coordinates, length and curve direction
' Returns:
'   - a new CircularArc object
'   - Nothing if wrong input parameters are passed
Public Function newCircArcSCLD(ByVal sX As Double, ByVal sY As Double, _
                               ByVal cX As Double, ByVal cY As Double, _
                               ByVal length As Double, ByVal curveDir As CURVE_DIR) As CircularArc
    On Error GoTo ErrHandler
    Set newCircArcSCLD = New CircularArc
    newCircArcSCLD.initFromSCLD sX, sY, cX, cY, length, curveDir
    Exit Function
ErrHandler:
    Set newCircArcSCLD = Nothing
End Function

' Creates a CircularArc from Start and Center point coordinates, length and curve direction
' Parameters are defined as Variant values
' Returns:
'   - a new CircularArc object
'   - Nothing if conversion of Variant to relevant type fails or wrong input parameters are passed
Public Function newCircArcSCLDvar(ByVal sX As Variant, ByVal sY As Variant, _
                                  ByVal cX As Variant, ByVal cY As Variant, _
                                  ByVal length As Variant, ByVal curveDir As Variant) As CircularArc
    Dim tmpCurveDir As CURVE_DIR
    tmpCurveDir = ConstCL.curveDirFromVariant(curveDir)
    On Error GoTo ErrHandler
    If tmpCurveDir = CD_NONE Then
        Set newCircArcSCLDvar = Nothing
        Exit Function
    End If
    Set newCircArcSCLDvar = newCircArcSCLD(CDbl(sX), CDbl(sY), CDbl(cX), CDbl(cY), CDbl(length), tmpCurveDir)
    Exit Function
ErrHandler:
    Set newCircArcSCLDvar = Nothing
End Function

' Creates a CircularArc from Start and End point coordinates, radius and curve direction
' Returns:
'   - a new CircularArc object
'   - Nothing if wrong input parameters are passed
Public Function newCircArcSERD(ByVal sX As Double, ByVal sY As Double, _
                               ByVal eX As Double, ByVal eY As Double, _
                               ByVal rad As Double, ByVal curveDir As CURVE_DIR) As CircularArc
    On Error GoTo ErrHandler
    Set newCircArcSERD = New CircularArc
    newCircArcSERD.initFromSERD sX, sY, eX, eY, rad, curveDir
    Exit Function
ErrHandler:
    Set newCircArcSERD = Nothing
End Function

' Creates a CircularArc from Start and End point coordinates, radius and curve direction
' Parameters are defined as Variant values
' Returns:
'   - a new CircularArc object
'   - Nothing if conversion of Variant to relevant type fails or wrong input parameters are passed
Public Function newCircArcSERDvar(ByVal sX As Variant, ByVal sY As Variant, _
                                  ByVal eX As Variant, ByVal eY As Variant, _
                                  ByVal rad As Variant, ByVal curveDir As Variant) As CircularArc
    Dim tmpCurveDir As CURVE_DIR
    tmpCurveDir = ConstCL.curveDirFromVariant(curveDir)
    On Error GoTo ErrHandler
    If tmpCurveDir = CD_NONE Then
        Set newCircArcSERDvar = Nothing
        Exit Function
    End If
    Set newCircArcSERDvar = newCircArcSERD(CDbl(sX), CDbl(sY), CDbl(eX), CDbl(eY), CDbl(rad), tmpCurveDir)
    Exit Function
ErrHandler:
    Set newCircArcSERDvar = Nothing
End Function

' Creates a CircularArc from a key: value list (collection)
' Keys are defined in ConstCL
' Returns:
'   - a new CircularArc object
'   - Nothing if creation fails: wrong values or missing keys
Public Function newCircArcColl(coll As Collection) As CircularArc
    On Error GoTo FailNewCA
    
    If coll.item(ConstCL.GEOM_TYPE) <> ConstCL.CA_NAME Then GoTo FailNewCA
    
    ' Select init type
    Select Case coll.item(ConstCL.GEOM_INIT_TYPE)
        ' Case Start, Center, Length and Curve Direction
    Case ConstCL.CA_INIT_SCLD
        Set newCircArcColl = newCircArcSCLDvar(coll.item(ConstCL.CA_M_START_X), _
                                               coll.item(ConstCL.CA_M_START_Y), _
                                               coll.item(ConstCL.CA_M_CENTER_X), _
                                               coll.item(ConstCL.CA_M_CENTER_Y), _
                                               coll.item(ConstCL.CA_M_LENGTH), _
                                               coll.item(ConstCL.CA_M_CURVE_DIR))
        ' Case Start, End, Radius and Curve Direction
    Case ConstCL.CA_INIT_SERD
        Set newCircArcColl = newCircArcSERDvar(coll.item(ConstCL.CA_M_START_X), _
                                               coll.item(ConstCL.CA_M_START_Y), _
                                               coll.item(ConstCL.CA_M_END_X), _
                                               coll.item(ConstCL.CA_M_END_Y), _
                                               coll.item(ConstCL.CA_M_RADIUS), _
                                               coll.item(ConstCL.CA_M_CURVE_DIR))
    
    Case Else
        GoTo FailNewCA
    End Select

    Exit Function
FailNewCA:
    Set newCircArcColl = Nothing
End Function





