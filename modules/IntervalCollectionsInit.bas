Attribute VB_Name = "IntervalCollectionsInit"
''=======================================================
''Called by:
''    Modules: None
''    Classes: None
''Calls:
''    Modules: None
''    Classes: Interval, IntvlColl, IntvlColls
''=======================================================
Option Private Module
Option Explicit
Public INTVL_COLLS As New IntvlColls

'table name prefix
Private Const TBL_NAME_PREFIX As String = "tblIC"
Public Const TBL_INTVL_INPUT_TYPE_COL As String = "Input Type"


Public Sub initIntvlColls()
    Dim WS As Worksheet
    Dim tbl As ListObject
    Dim tempIC As IntvlColl
    
    'loop all worksheets
    For Each WS In ThisWorkbook.Worksheets
        'loop all tables
        For Each tbl In WS.ListObjects
            Set tempIC = getIntvlCollFromTable(tbl)
            If Not tempIC Is Nothing Then
                INTVL_COLLS.addIntvlColl tempIC
            End If
        Next tbl
    Next WS
End Sub

Private Function isTableValid(tbl As ListObject) As Boolean
    
    isTableValid = False
    
    'check if Table name is valid
    If Left(tbl.name, 5) <> TBL_NAME_PREFIX Then
        isTableValid = False
        Exit Function
    End If
    
    'check if Table Header is valid
    Dim i As Integer
    Dim c As Range
    Dim foundColName As Boolean
    Dim inpValidArray() As String
    inpValidArray = getIntvlPartStringArray
    
    ReDim Preserve inpValidArray(UBound(inpValidArray) + 1)
    inpValidArray(UBound(inpValidArray)) = TBL_INTVL_INPUT_TYPE_COL
    
    For i = 0 To UBound(inpValidArray)
        foundColName = False
        For Each c In tbl.HeaderRowRange
            If c.value = inpValidArray(i) Then
                foundColName = True
                Exit For
            End If
        Next c
        If foundColName = False Then
            isTableValid = False
            Exit Function
        End If
    Next i
    
    'if code executes till this line then table format is valid
    isTableValid = True
End Function

Private Function getIntvlCollFromTable(tbl As ListObject) As IntvlColl
    Dim outIC As IntvlColl
    Dim lr As ListRow
    Dim inputColl As Collection
    Dim tempInterval As New Interval
    If isTableValid(tbl) Then
        Set outIC = New IntvlColl
        For Each lr In tbl.ListRows
            Set inputColl = tableRowToCollection(lr.index, tbl)
            Set tempInterval = Nothing
            If tempInterval.initFromCollection(inputColl) Then
                If Not outIC.addIntvl(tempInterval) Then GoTo FailGetIntvlColl
            Else
                GoTo FailGetIntvlColl
            End If
        Next lr
    Else
        GoTo FailGetIntvlColl
    End If
    
    outIC.name = tbl.name
    Set getIntvlCollFromTable = outIC
    
    Exit Function
FailGetIntvlColl:
    Set outIC = Nothing
    Set getIntvlCollFromTable = outIC
    Exit Function
End Function

Private Function tableRowToCollection(rowIndex As Integer, tbl As ListObject) As Collection
    Dim result As New Collection
    Dim c As Range
    Dim tempVal As Variant
    
    For Each c In tbl.HeaderRowRange
        tempVal = tbl.ListRows(rowIndex).Range.Cells(1, c.Column).value
        If tempVal <> vbNullString Then
            result.Add tempVal, c.value
        End If
    Next c
    Set tableRowToCollection = result
End Function


