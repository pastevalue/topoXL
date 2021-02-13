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
'' "getUserCLs" exposes all initialized CLs stored in "userCLs" variable
'' If "userCLs" is not initialized then "getUserCLs" (most likely called
''  from a UDF) deffers the call to "initCLsCallback which in turn
''  initializes the CLs and refreshes all formulas in order to forcefully
''  calculate non-volatile UDFs
''========================================================================

'@Folder("TopoXL.Init")

Option Explicit
Option Private Module

Private userCLs As CLs

' Calculate Event Callback and wait indicator used for assuring the persistance
'  of the userCLs data
Private m_bookCalcCallback As BookCalculateCallback
Private m_waitRefresh As Boolean

' Returns the persistant userCLs Object
Public Function getUserCLs() As CLs
    If userCLs Is Nothing Then
        ' Try both Application.OnTime and SheetCalculate event to make sure
        '   one of them will trigger the Refresh
        Application.OnTime Now(), "'" & ThisWorkbook.name & "'!initCLsCallback"
        If m_bookCalcCallback Is Nothing Then
            Set m_bookCalcCallback = New BookCalculateCallback
            m_bookCalcCallback.init ThisWorkbook, "initCLsCallback", True
        End If
    Else
        Set getUserCLs = userCLs
    End If
End Function

' Callback method. See 'GetUserCLs' method above
' Do not change 'Function' to 'Sub' to avoid Application.Run disconnection!
' Do not delete the 'dummy' parameter (caller can be a SheetCalculate event)!
'*******************************************************************************
Private Function initCLsCallback(Optional ByVal dummy As Variant)
    If Not m_waitRefresh Then
        m_waitRefresh = True
        initCLs
        If ThisWorkbook.IsAddin Then
            Dim book As Workbook
            For Each book In Application.Workbooks
                forceRefreshBookFormulas book
            Next book
        Else
            forceRefreshBookFormulas ThisWorkbook
        End If
        Set m_bookCalcCallback = Nothing
        m_waitRefresh = False
    End If
End Function

' Searches all worksheets for tables which can be used to initialize a Center Line
' If errors are found in the input table (missing columns, missing values,...),
' center line is not initialized
' Initilized center lines are stored in the userCLs variable of this module
Private Sub initCLs()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim tmpCL As CL
    
    Set userCLs = New CLs
    'loop all worksheets
    For Each ws In ThisWorkbook.Worksheets
        'loop all tables
        For Each tbl In ws.ListObjects
            Set tmpCL = FactoryCL.newCLtbl(tbl)
            If Not tmpCL Is Nothing Then
                userCLs.addCL tmpCL
            End If
        Next tbl
    Next ws
End Sub

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
    Set userCLs = Nothing
    initCLsCallback
End Sub
