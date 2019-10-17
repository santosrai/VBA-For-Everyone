''
' @Purpose:  Get Corresponding sheet
' @Param  :  {Workbook} Book
'            {String}   sheetname　
' @Return :　Worksheet
''
Public Function GetSheet(ByVal Book As Workbook, ByVal SheetName As String) As Worksheet
    Dim sheet As Object
    For Each sheet In Book.Worksheets
        If sheet.Name = SheetName Then
            Set GetSheet = sheet
            Exit Function
        End If
    Next
    Set GetSheet = Nothing
    
End Function

''
' @Purpose:  Return the Last Row Or Column Number
' @Param  :  {Workbook} Workbook
'            {String}   RowColumn
' @Return :　{Long} RowColumn
''
Public Function GetLastRowColumn(ws As Worksheet, RowColumn As String) As Long
    Dim LastRowColumn As Long
    Select Case LCase(Left(RowColumn, 1)) 'If they put in 'row' or column instead of 'r' or 'c'.
        Case "c"
            LastRowColumn = ws.Cells.Find("*", LookIn:=xlFormulas, SearchOrder:=xlByColumns, _
            SearchDirection:=xlPrevious).Column
        Case "r"
            LastRowColumn = ws.Cells.Find("*", LookIn:=xlFormulas, SearchOrder:=xlByRows, _
            SearchDirection:=xlPrevious).Row
        Case Else
            LastRowColumn = 1
        End Select
    'Return
    GetLastRowColumn = LastRowColumn
End Function
