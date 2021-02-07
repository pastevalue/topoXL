Attribute VB_Name = "FactoryCL"
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
'' Factory module used to create CL (Centerline) related classes
''========================================================================

'@Folder("TopoXL.CL")

Option Explicit
Option Private Module

Private Const CL_TBL_NAME_PREFIX As String = "tblCL"

' Creates a CLelem instance from a geom object and a start measure value
' Returns:
'   - a new CLelem object
'   - Nothing if geom is Nothing
Public Function newCLelem(geom As IGeom, ByVal startM As Double, Optional reversed As Boolean = False) As CLelem
    Set newCLelem = New CLelem
    newCLelem.init geom, startM, reversed
End Function

' Creates a CLelem from a geom object and a start measure value as Variant
' Returns:
'   - a new CLelem object
'   - Nothing if conversion of Variant to Double fails
Public Function newCLelemVar(geom As IGeom, ByVal startM As Variant, Optional reversed As Variant = False) As CLelem
    On Error GoTo ErrHandler
    Set newCLelemVar = newCLelem(geom, CDbl(startM), CBool(reversed))
    Exit Function
ErrHandler:
    Set newCLelemVar = Nothing
End Function

' Creates a CLelem from a key: value list (collection)
' Keys are defined in ConstCL
' Returns:
'   - a new CLelem object. Its type is defined by ConstCL.GEOM_TYPE key
'   - Nothing if creation fails: wrong values or missing keys
Public Function newCLelemColl(coll As Collection) As CLelem
    Dim startM As Variant
    Dim reversed As Variant
    On Error GoTo FailNewCLelem
    
    startM = coll.item(ConstCL.CL_MEASURE)
    reversed = coll.item(ConstCL.CL_REVERSED)
    ' Select init type
    Select Case coll.item(ConstCL.GEOM_TYPE)
    Case ConstCL.LS_NAME
        Set newCLelemColl = newCLelemVar(FactoryGeom.newLnSegColl(coll), startM, reversed)
    Case ConstCL.CA_NAME
        Set newCLelemColl = newCLelemVar(FactoryGeom.newCircArcColl(coll), startM, reversed)
    Case ConstCL.CLA_NAME
        Set newCLelemColl = newCLelemVar(FactoryGeom.newClothArcColl(coll), startM, reversed)
    Case Else
        GoTo FailNewCLelem
    End Select

    Exit Function
FailNewCLelem:
    Set newCLelemColl = Nothing
End Function

' Creates a CL instance with the specified name
Public Function newCL(ByVal name As String) As CL
    Set newCL = New CL
    newCL.init name
End Function

' Creates a CL instance based on an Excel table
' Name of the CL is assigned as the name of the Excel table object (ListObject name)
Public Function newCLtbl(tbl As ListObject) As CL
    Dim elem As CLelem
    Dim lr As ListRow
    Dim inputColl As Collection
    
    If isValidCLtable(tbl) Then
        Set newCLtbl = New CL
        For Each lr In tbl.ListRows
            Set inputColl = tblRowToColl(lr, tbl)
            Set elem = newCLelemColl(inputColl)
            If Not elem Is Nothing Then
                newCLtbl.addElem elem
             Else
                GoTo FailNewCenterLine
            End If
        Next lr
    Else
        GoTo FailNewCenterLine
    End If
    
    newCLtbl.name = tbl.name
    Exit Function
FailNewCenterLine:
    Set newCLtbl = Nothing
End Function

' Returns True if a table is valid for Center Line initialization, False otherwise
Private Function isValidCLtable(tbl As ListObject) As Boolean
        
    'check if Table name is valid
    If VBA.Left(tbl.name, 5) <> CL_TBL_NAME_PREFIX Then
        isValidCLtable = False
        Exit Function
    End If
    
    ' TODO: instead of the below table header must be checked as following:
    '   - there must be columns for AXIS_TBL_ELEM_TYPE_COL and AXIS_TBL_ELEM_INPUT_TYPE_COL
    '   - for each input type present relevant column header must be present
    
    'if code executes till this line then table format is valid
    isValidCLtable = True
End Function

' Returns a table row as a collection (key = table header value, value = table value)
Private Function tblRowToColl(ByVal lr As ListRow, tbl As ListObject) As Collection
    Set tblRowToColl = New Collection
    
    Dim c As Range
    Dim tempVal As Variant
    
    For Each c In tbl.HeaderRowRange
        tempVal = lr.Range.Cells(1, c.Column).value
        If tempVal <> vbNullString Then
            tblRowToColl.add tempVal, c.value
        End If
    Next c
End Function






