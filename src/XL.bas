Attribute VB_Name = "XL"
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
'' "getUserCLs" returns all CLs found in this book and refreshes all
'   formulas in order to forcefully calculate non-volatile UDFs
''========================================================================

'@Folder("TopoXL.Init")

Option Explicit
Option Private Module

Private m_bookCalcCallback As BookCalculateCallback

' Searches all worksheets for tables containing Center Line data
' If errors are found in the input table (like missing columns/values)
'  then the center line is not initialized
' Initialized center lines are returned and also stored for reuse
Public Function getUserCLs(Optional forceRefresh As Boolean = False) As CLs
    Static userCLs As CLs
    
    If userCLs Is Nothing Or forceRefresh Then
        Dim ws As Worksheet
        Dim tbl As ListObject
        Dim tmpCL As CL
        
        Set userCLs = New CLs
        For Each ws In ThisWorkbook.Worksheets
            For Each tbl In ws.ListObjects
                Set tmpCL = FactoryCL.newCLtbl(tbl)
                If Not tmpCL Is Nothing Then
                    userCLs.addCL tmpCL
                End If
            Next tbl
        Next ws
        
        refreshFormulas
    End If
    Set getUserCLs = userCLs
End Function

' refreshes all formulas on the spot or defers the call if code is UDF mode
Private Function refreshFormulas()
    If UDFMode() Then 'Defer
        If Not Application.EnableEvents Then
            Application.EnableEvents = True 'Works in UDF mode as well
        End If

        ' Try both Application.OnTime and SheetCalculate event to make sure
        '   one of them will trigger the Refresh
        Application.OnTime Now(), "'" & ThisWorkbook.name & "'!refreshCallback"
        If m_bookCalcCallback Is Nothing Then
            Set m_bookCalcCallback = New BookCalculateCallback
            m_bookCalcCallback.init ThisWorkbook, "refreshCallback", True
        End If
    Else
        refreshCallback
    End If
End Function

'Returns a boolean indicating if code was called from a UDF
Public Function UDFMode() As Boolean
    Dim dispAlerts As Boolean: dispAlerts = Application.DisplayAlerts
    
    On Error Resume Next
    Application.DisplayAlerts = Not dispAlerts 'Cannot be changed in UDF mode
    On Error GoTo 0
    
    UDFMode = (Application.DisplayAlerts = dispAlerts)
    If Not UDFMode Then Application.DisplayAlerts = dispAlerts 'Revert
End Function

' Callback method
' Do not change 'Function' to 'Sub' to avoid Application.Run disconnection!
' Do not delete the 'dummy' parameter (caller can be a SheetCalculate event)!
Private Function refreshCallback(Optional ByVal dummy As Variant)
    Static waitRefresh As Boolean
    
    If Not waitRefresh Then
        waitRefresh = True
        If ThisWorkbook.IsAddin Then
            Dim book As Workbook
            For Each book In Application.Workbooks
                forceRefreshBookFormulas book
            Next book
        Else
            forceRefreshBookFormulas ThisWorkbook
        End If
        Set m_bookCalcCallback = Nothing
        waitRefresh = False
    End If
End Function

' Refreshes all formulas in a Workbook
Private Sub forceRefreshBookFormulas(ByVal book As Workbook)
    Dim sht As Worksheet
    Dim scrUpdate As Boolean: scrUpdate = Application.ScreenUpdating
    Dim dispAlerts As Boolean: dispAlerts = Application.DisplayAlerts
    Dim calcMode As XlCalculation: calcMode = Application.Calculation
    '
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    '
    For Each sht In book.Worksheets
        If sht.EnableCalculation Then
            sht.EnableCalculation = False
            sht.EnableCalculation = True
        End If
    Next sht
RestoreState:
    Application.ScreenUpdating = scrUpdate
    Application.DisplayAlerts = dispAlerts
    Application.Calculation = calcMode
End Sub

' Refreshes all CLs
Public Sub refreshCLs()
    getUserCLs forceRefresh:=True
End Sub
