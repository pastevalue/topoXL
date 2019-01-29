Attribute VB_Name = "UDF_ToAcad"
Option Explicit
Private Const ACAD_CMD_LAYER = "-layer"
Private Const ACAD_CMD_POINT = "point"
Private Const ACAD_CMD_PLINE = "pline"
Private Const ACAD_CMD_3DPOLY = "3dpoly"
Private Const ACAD_CMD_TEXT = "-text"
Private Const ACAD_CMD_INSERT = "-insert"
Private Const ACAD_CMD_HATCH = "-hatch"

Public Function toAcadScrPoint(axisOrder As String, ParamArray values() As Variant) As Variant
    Application.Volatile False
    Dim tempColl As Collection
    Dim tempArray() As Variant
    Dim i As Long
    Dim t As Variant
    Dim axisOrderVar As AXIS_ORDER
    Dim cooCount As Integer
    Dim xInd As Integer, yInd As Integer, zInd As Integer
    
    axisOrderVar = ENUMS.getAxisOrderFromString(axisOrder)
    cooCount = Len(ENUMS.axisOrderToString(axisOrderVar))
    'IF "axisOrder" parameter is invalid THEN return error
    If axisOrderVar = AXIS_NONE Then GoTo failInput
    
    t = values
    Set tempColl = rangeFunctions.valuesToCollection(t)
    
    'IF the number of values is not divisible by number of coordinates THEN return error
    If tempColl.count Mod cooCount <> 0 Then GoTo failInput
    
    tempArray = rangeFunctions.collectionTo2Darray(cooCount, tempColl)
    
    'get coordintes positions
    xInd = ENUMS.getAxisPosition("X", axisOrderVar)
    yInd = ENUMS.getAxisPosition("Y", axisOrderVar)
    zInd = ENUMS.getAxisPosition("Z", axisOrderVar)
    If zInd = 0 Then
        'build return string for 2D
        For i = 0 To UBound(tempArray)
            If WorksheetFunction.IsNumber(tempArray(i, xInd - 1)) And WorksheetFunction.IsNumber(tempArray(i, yInd - 1)) Then
                toAcadScrPoint = toAcadScrPoint & ACAD_CMD_POINT & " " & tempArray(i, xInd - 1) & "," & tempArray(i, yInd - 1) & vbNewLine
            Else
               GoTo failInput
            End If
        Next
    Else
        'build return string for 3D
        For i = 0 To UBound(tempArray)
            If WorksheetFunction.IsNumber(tempArray(i, xInd - 1)) And WorksheetFunction.IsNumber(tempArray(i, yInd - 1)) And WorksheetFunction.IsNumber(tempArray(i, zInd - 1)) Then
                toAcadScrPoint = toAcadScrPoint & ACAD_CMD_POINT & " " & tempArray(i, xInd - 1) & "," & tempArray(i, yInd - 1) & "," & tempArray(i, zInd - 1) & vbNewLine
            Else
                GoTo failInput
            End If
        Next
    End If
     
    Exit Function
failInput:
    toAcadScrPoint = CVErr(xlErrNum)
    Exit Function
End Function

Public Function toAcadScrPLine(axisOrder As String, ParamArray values() As Variant) As Variant
    Application.Volatile False
    Dim tempColl As Collection
    Dim tempArray() As Variant
    Dim i As Long
    Dim t As Variant
    Dim axisOrderVar As AXIS_ORDER
    Dim cooCount As Integer
    Dim xInd As Integer, yInd As Integer, zInd As Integer
 
    axisOrderVar = ENUMS.getAxisOrderFromString(axisOrder)
    cooCount = Len(ENUMS.axisOrderToString(axisOrderVar))
    'IF  "axisOrder" parameter is invalid THEN return error
    If axisOrderVar = AXIS_NONE Then GoTo failInput
    
    t = values
    Set tempColl = rangeFunctions.valuesToCollection(t)
    
    'IF the number of values is not divisible by number of coordinates THEN return error
    If tempColl.count Mod cooCount <> 0 Then GoTo failInput
    
    tempArray = rangeFunctions.collectionTo2Darray(cooCount, tempColl)
    
    'get coordintes positions
    xInd = ENUMS.getAxisPosition("X", axisOrderVar) - 1
    yInd = ENUMS.getAxisPosition("Y", axisOrderVar) - 1
    zInd = ENUMS.getAxisPosition("Z", axisOrderVar) - 1
    
    
    Select Case cooCount
        'two dimension coordinates
        Case 2
            toAcadScrPLine = ACAD_CMD_PLINE & " "
            If tempColl.count < 4 Then GoTo failInput 'return error if not enought values are provided
            For i = 0 To UBound(tempArray)
                If WorksheetFunction.IsNumber(tempArray(i, xInd)) And WorksheetFunction.IsNumber(tempArray(i, yInd)) Then
                    toAcadScrPLine = toAcadScrPLine & tempArray(i, xInd) & "," & tempArray(i, yInd) & vbNewLine
                Else
                    GoTo failInput
                End If
            Next
        'three dimension coordinates
        Case 3
            toAcadScrPLine = ACAD_CMD_3DPOLY & " "
            If tempColl.count < 6 Then GoTo failInput 'return error if not enought values are provided
            For i = 0 To UBound(tempArray)
                If WorksheetFunction.IsNumber(tempArray(i, xInd)) And WorksheetFunction.IsNumber(tempArray(i, yInd)) And WorksheetFunction.IsNumber(tempArray(i, zInd)) Then
                    toAcadScrPLine = toAcadScrPLine & tempArray(i, xInd) & "," & tempArray(i, yInd) & "," & tempArray(i, zInd) & vbNewLine
                Else
                    GoTo failInput
                End If
            Next
        Case Else
            GoTo failInput
    End Select
    
    Exit Function
failInput:
    toAcadScrPLine = CVErr(xlErrNum)
    Exit Function
End Function
'Returns an Autocad script that can be used as a script file to automaticaly insert multiple polylines into the drawing
'Parameters:
'   -axisOrder: a string that specifies the order in which X, Y and Z values will be parsed. Valid entries for this function are combinations of X, Y and Z characters;
'   -separator: a string that separates individual polylines. The string must be placed like a pair of _
    coordinates but instead of numbers just the separator (2 or 3 separators depending if 2D or 3D points are used). Single point polylines will give failInput error.
'   -values: a variant that contains the value(s) that will be used as Acad Mpline command parameters.
'            This can be an Excel Range Object, simple values (numbers or strings) or an (multidimensional) array of these types.
'            Other types than those specified will be ignored.
'            The values must reprezent the coordinates (2 or 3 dimensions) for Autocad Mpline command;
'Result is a string that is enclosed between " character. Must replace it with nothing before running Autocad script file.
Public Function toAcadScrMPline(axisOrder As String, separator As String, ParamArray values() As Variant) As Variant
    Application.Volatile False
    Dim tempColl As Collection
    Dim tempArray() As Variant
    Dim i As Long
    Dim t As Variant
    Dim axisOrderVar As AXIS_ORDER
    Dim cooCount As Integer
    Dim xInd As Integer, yInd As Integer, zInd As Integer
    Dim firstScr As Boolean
 
    axisOrderVar = ENUMS.getAxisOrderFromString(axisOrder)
    cooCount = Len(ENUMS.axisOrderToString(axisOrderVar))
    'IF  "axisOrder" parameter is invalid THEN return error
    If axisOrderVar = AXIS_NONE Then GoTo failInput
    
    t = values
    Set tempColl = rangeFunctions.valuesToCollection(t)
    
    'IF the number of values is not divisible by number of coordinates THEN return error
    If tempColl.count Mod cooCount <> 0 Then GoTo failInput
    
    tempArray = rangeFunctions.collectionTo2Darray(cooCount, tempColl)
    
    'get coordintes positions
    xInd = ENUMS.getAxisPosition("X", axisOrderVar) - 1
    yInd = ENUMS.getAxisPosition("Y", axisOrderVar) - 1
    zInd = ENUMS.getAxisPosition("Z", axisOrderVar) - 1
    
    toAcadScrMPline = ""
    Dim endPline As Boolean
    Dim cooPairCount As Long
    endPline = True
    cooPairCount = 0
    firstScr = True
    Select Case cooCount
        'two dimension coordinates
        Case 2
            For i = 0 To UBound(tempArray)
                If tempArray(i, xInd) = separator And tempArray(i, yInd) = separator Then
                    If cooPairCount = 1 Then GoTo failInput
                    endPline = True
                    cooPairCount = 0
                Else
                    If WorksheetFunction.IsNumber(tempArray(i, xInd)) And WorksheetFunction.IsNumber(tempArray(i, yInd)) Then
                        If endPline Then
                            If firstScr Then
                                toAcadScrMPline = toAcadScrMPline & ACAD_CMD_PLINE & " "
                                firstScr = False
                            Else
                                toAcadScrMPline = toAcadScrMPline & vbNewLine & ACAD_CMD_PLINE & " "
                            End If
                            endPline = False
                        End If
                        cooPairCount = cooPairCount + 1
                        toAcadScrMPline = toAcadScrMPline & tempArray(i, xInd) & "," & tempArray(i, yInd) & vbNewLine
                    Else
                        GoTo failInput
                    End If
                End If
            Next i
        'three dimension coordinates
        Case 3
            For i = 0 To UBound(tempArray)
                If tempArray(i, xInd) = separator And tempArray(i, yInd) = separator And tempArray(i, zInd) = separator Then
                    If cooPairCount = 1 Then GoTo failInput
                    endPline = True
                    cooPairCount = 0
                 Else
                    If WorksheetFunction.IsNumber(tempArray(i, xInd)) And WorksheetFunction.IsNumber(tempArray(i, yInd)) And WorksheetFunction.IsNumber(tempArray(i, zInd)) Then
                        If endPline Then
                            If firstScr Then
                                toAcadScrMPline = toAcadScrMPline & ACAD_CMD_3DPOLY & " "
                                firstScr = False
                            Else
                                toAcadScrMPline = toAcadScrMPline & vbNewLine & ACAD_CMD_3DPOLY & " "
                            End If
                            endPline = False
                        End If
                        cooPairCount = cooPairCount + 1
                        toAcadScrMPline = toAcadScrMPline & tempArray(i, xInd) & "," & tempArray(i, yInd) & "," & tempArray(i, zInd) & vbNewLine
                     Else
                        GoTo failInput
                    End If
                End If
            Next
        Case Else
            GoTo failInput
    End Select
    Exit Function
failInput:
    toAcadScrMPline = CVErr(xlErrNum)
    Exit Function
End Function

'Returns an Autocad script that can be used as a script file to automaticaly insert blocks into the drawing
'Parameters:
'   -axisOrder: a string that specifies the order in which X, Y and Z values will be parsed. Valid entries for this function are combinations of X, Y and Z characters;
'   -values: a variant that contains the value(s) that will be used as Acad Insert command parameters.
'            This can be an Excel Range Object, simple values (numbers or strings) or an (multidimensional) array of these types.
'            Other types than those specified will be ignored.
'            In order, values must reprezent the following parameters for Autocad Insert command;
'               - block name;
'               - coordinates (2 or 3 dimensions);
'               - x scale;
'               - y scale;
'               - rotation
'Result is a string that is enclosed between " character. Must replace it with nothing before running Autocad script file.
Public Function toAcadScrInsert(axisOrder As String, ParamArray values() As Variant) As Variant
    Application.Volatile False
    Dim tempColl As Collection
    Dim tempArray() As Variant
    Dim i As Long
    Dim t As Variant
    Dim axisOrderVar As AXIS_ORDER
    Dim cooCount As Integer
    Dim xInd As Integer, yInd As Integer, zInd As Integer
 
    axisOrderVar = ENUMS.getAxisOrderFromString(axisOrder)
    cooCount = Len(ENUMS.axisOrderToString(axisOrderVar))
    'IF  "axisOrder" parameter is invalid THEN return error
    If axisOrderVar = AXIS_NONE Then GoTo failInput
    
    t = values
    Set tempColl = rangeFunctions.valuesToCollection(t)
    
    'IF the number of values is not divisible by number of coordinates THEN return error
    If tempColl.count Mod (cooCount + 4) <> 0 Then GoTo failInput
    
    tempArray = rangeFunctions.collectionTo2Darray(cooCount + 4, tempColl)
    
    'get coordintes positions
    xInd = ENUMS.getAxisPosition("X", axisOrderVar)
    yInd = ENUMS.getAxisPosition("Y", axisOrderVar)
    zInd = ENUMS.getAxisPosition("Z", axisOrderVar)
    
    toAcadScrInsert = ACAD_CMD_INSERT & " "
    Select Case cooCount
        'two dimension coordinates
        Case 2
            For i = 0 To UBound(tempArray)
                'IF input is valid THEN create Acad Script ELSE return input error
                If WorksheetFunction.IsNumber(tempArray(i, xInd)) And WorksheetFunction.IsNumber(tempArray(i, yInd)) And _
                   WorksheetFunction.IsNumber(tempArray(i, 3)) And WorksheetFunction.IsNumber(tempArray(i, 4)) And WorksheetFunction.IsNumber(tempArray(i, 5)) Then
                    'add name and coo (x and y) to Acad script
                    toAcadScrInsert = toAcadScrInsert & tempArray(i, 0) & vbNewLine & tempArray(i, xInd) & "," & tempArray(i, yInd) & vbNewLine
                    'add x and y scale to Acad script
                    toAcadScrInsert = toAcadScrInsert & tempArray(i, 3) & vbNewLine & tempArray(i, 4) & vbNewLine
                    'add rotation to Acad script
                    toAcadScrInsert = toAcadScrInsert & tempArray(i, 5) '& vbNewLine
                Else
                    GoTo failInput
                End If
            Next
        'three dimension coordinates
        Case 3
            For i = 0 To UBound(tempArray)
                'IF input is valid THEN create Acad Script ELSE return input error
                If WorksheetFunction.IsNumber(tempArray(i, xInd)) And WorksheetFunction.IsNumber(tempArray(i, yInd)) And WorksheetFunction.IsNumber(tempArray(i, zInd)) And _
                   WorksheetFunction.IsNumber(tempArray(i, 4)) And WorksheetFunction.IsNumber(tempArray(i, 5)) And WorksheetFunction.IsNumber(tempArray(i, 6)) Then
                    'add name and coo (x, y and z) to Acad script
                    toAcadScrInsert = toAcadScrInsert & tempArray(i, 0) & vbNewLine & tempArray(i, xInd) & "," & tempArray(i, yInd) & "," & tempArray(i, zInd) & vbNewLine
                    'add x and y scale to Acad script
                    toAcadScrInsert = toAcadScrInsert & tempArray(i, 4) & vbNewLine & tempArray(i, 5) & vbNewLine
                    'add rotation to Acad script
                    toAcadScrInsert = toAcadScrInsert & tempArray(i, 6) '& vbNewLine
                Else
                    GoTo failInput
                End If
            Next
        Case Else
            GoTo failInput
    End Select
    
    Exit Function
failInput:
    toAcadScrInsert = CVErr(xlErrNum)
    Exit Function
End Function

'Returns an Autocad script that can be used as a script file to automaticaly draw texts into the drawing
'Parameters:
'   -axisOrder: a string that specifies the order in wich X, Y and Z values will be parsed. Valid entries for this function are combinations of X, Y and Z characters;
'   -values: a variant that contains the value(s) that will be used as Acad Text command parameters.
'            This can be an Excel Range Object, simple values (numbers or strings) or an (multidimensional) array of these types.
'            Other types than those specified will be ignored.
'            In order, values must reprezent the following parameters for Autocad Insert command:
'               - text height;
'               - text rotation;
'               - text string;
'Result is a string that is enclosed between " character. Must replace it with nothing before running Autocad script file.
Public Function toAcadScrText(axisOrder As String, ParamArray values() As Variant) As Variant
    Application.Volatile False
    Application.Volatile False
    Const VALUES_COUNT As Integer = 3
    Dim tempColl As Collection
    Dim tempArray() As Variant
    Dim i As Long
    Dim t As Variant
    Dim axisOrderVar As AXIS_ORDER
    Dim cooCount As Integer
    Dim xPos As Integer
    Dim yPos As Integer
    Dim zPos As Integer
    
    axisOrderVar = ENUMS.getAxisOrderFromString(axisOrder)
    cooCount = Len(ENUMS.axisOrderToString(axisOrderVar))
    'IF "axisOrder" parameter is invalid THEN return error
    If axisOrderVar = AXIS_NONE Then GoTo failInput
        
    t = values
    Set tempColl = rangeFunctions.valuesToCollection(t)
    
    'IF the number of values is not divisible by number of values needed for input THEN return error
    If tempColl.count Mod (cooCount + VALUES_COUNT) <> 0 Then GoTo failInput
    
    tempArray = rangeFunctions.collectionTo2Darray(cooCount + VALUES_COUNT, tempColl)
        
    'get coordintes positions
    xPos = ENUMS.getAxisPosition("X", axisOrderVar)
    yPos = ENUMS.getAxisPosition("Y", axisOrderVar)
    zPos = ENUMS.getAxisPosition("Z", axisOrderVar)
    
    Select Case cooCount
        Case 2 '2D coordinates
            'build return string
            For i = 0 To UBound(tempArray)
                'validate coordinates
                If WorksheetFunction.IsNumber(tempArray(i, xPos - 1)) And WorksheetFunction.IsNumber(tempArray(i, yPos - 1)) Then
                    'validate text height and text rotation values
                    If WorksheetFunction.IsNumber(tempArray(i, 2)) And WorksheetFunction.IsNumber(tempArray(i, 3)) Then
                        toAcadScrText = toAcadScrText & ACAD_CMD_TEXT & " " & tempArray(i, xPos - 1) & "," & tempArray(i, yPos - 1) & _
                        " " & tempArray(i, 2) & " " & tempArray(i, 3) & " " & tempArray(i, 4)
                        If i < UBound(tempArray) Then toAcadScrText = toAcadScrText & vbNewLine 'add new line if there are more commands to follow
                    Else
                        GoTo failInput
                    End If
                Else
                    GoTo failInput
                End If
            Next i
        Case 3 '3D coordinates
            'build return string
            For i = 0 To UBound(tempArray)
                'validate coordinates
                If WorksheetFunction.IsNumber(tempArray(i, xPos - 1)) And WorksheetFunction.IsNumber(tempArray(i, yPos - 1)) And WorksheetFunction.IsNumber(tempArray(i, zPos - 1)) Then
                    'validate text height and text rotation values
                    If WorksheetFunction.IsNumber(tempArray(i, 3)) And WorksheetFunction.IsNumber(tempArray(i, 4)) Then
                        toAcadScrText = toAcadScrText & ACAD_CMD_TEXT & " " & tempArray(i, xPos - 1) & "," & tempArray(i, yPos - 1) & "," & tempArray(i, zPos - 1) & _
                        " " & tempArray(i, 3) & " " & tempArray(i, 4) & " " & tempArray(i, 5)
                        If i < UBound(tempArray) Then toAcadScrText = toAcadScrText & vbNewLine 'add new line if there are more commands to follow
                    Else
                        GoTo failInput
                    End If
                Else
                    GoTo failInput
                End If
            Next i
        Case Else 'else return error
            GoTo failInput
    End Select
    Exit Function
failInput:
    toAcadScrText = CVErr(xlErrNum)
    Exit Function
End Function

Public Function toAcadScrChangeLayer(layerName As String) As Variant
    Application.Volatile False
    If Len(layerName) > 0 Then
        toAcadScrChangeLayer = ACAD_CMD_LAYER & " m " & layerName & vbNewLine
    Else
        GoTo failInput
    End If
Exit Function
failInput:
    toAcadScrChangeLayer = CVErr(xlErrNum)
End Function

'Returns an Autocad script that can be used to insert a hatch
'Parameters:
'   -patternName: a string that specifies the name of an existing pattern in Autocad (ex. ANSI31);
'   -patternScale: a double that specifies the scale of the pattern;
'   -patternAngle: a double that specifies the angle of the pattern;
'   -drawBoundary: an integer that specifies if the boundary of the hatch is drawn (0 - No, 1 - Yes)
'   -values: a variant that contains the coordinates that will be used as boundary for the hatch;
'            This can be an Excel Range Object, simple values (numbers) or an (multidimensional) array of these types.
'            Other types than those specified will be ignored.
'Result is a string that is enclosed between " character. Must replace it with nothing before running Autocad script file.
Public Function toAcadScrHatch(patternName As String, patternScale As Double, patternAngle As Double, _
                               drawBoundary As Integer, ParamArray values() As Variant) As Variant
    Application.Volatile False
    Dim tempColl As Collection
    Dim tempArray() As Variant
    Dim i As Long
    Dim t As Variant
    Dim drawBoundaryString As String
 
    t = values
    Set tempColl = rangeFunctions.valuesToCollection(t)
    
    'IF the drawBoundary parameter is different than 0 or 1 then exit function
    If drawBoundary < 0 Or drawBoundary > 1 Then GoTo failInput
    'IF the number of values is not divisible by 2 (2D coordinates) and minimum 6 (minimum 3 points to describe an area) then exit function
    If tempColl.count Mod 2 <> 0 Or tempColl.count < 6 Then GoTo failInput
    
    tempArray = rangeFunctions.collectionTo2Darray(2, tempColl)
    
    If drawBoundary = 0 Then
        drawBoundaryString = "N"
    Else
        drawBoundaryString = "Y"
    End If
    toAcadScrHatch = ACAD_CMD_HATCH & " p " & patternName & " " & patternScale & " " & patternAngle & " w " & drawBoundaryString & " "
            For i = 0 To UBound(tempArray)
                'IF input is valid THEN create Acad Script ELSE return input error
                If WorksheetFunction.IsNumber(tempArray(i, 0)) And WorksheetFunction.IsNumber(tempArray(i, 1)) Then
                    'add coo (x and y) to Acad script
                    toAcadScrHatch = toAcadScrHatch & tempArray(i, 0) & "," & tempArray(i, 1) & vbNewLine
                Else
                    GoTo failInput
                End If
            Next i
            toAcadScrHatch = toAcadScrHatch & tempArray(0, 0) & "," & tempArray(0, 1) & vbNewLine & "c " & vbNewLine & vbNewLine
    Exit Function
failInput:
    toAcadScrHatch = CVErr(xlErrNum)
    Exit Function
End Function




