<div align="center">
<h1> VBA for Everyone </h1>

<sub>Author: Santosh Rai
<small> January, 2018</small>
</sub>

</div>

- [Introduction](#introduction)
- [Requirements](#requirements)
- [Setup](#setup)
- [Variables](#variables)
- [Comments](#comments)
- [Coding Style](#codingstyle)
- [Data types](#data-types)
- [Checking Data Type and casting](#data-types-casting)
- [Conditionals](#conditionals)
  - [if](#if)
  - [if else](#if-else)
  - [if else if else](#if-else-if-else)
  - [switch](#switch)
  - [Ternary Operators](#ternary-operators)
  - [While loop](#while-loop)
  - [Do while loop](#do-while-loop)
- [Arrays](#arrays)
- [Functions](#functions)
- [Classes](#classes)
  - [Defining a classes](#defining-a-classes)
  - [Class Instantiation](#class-instantiation)
  - [Class Constructor](#class-constructor)
  - [Default values with constructor](#default-values-with-constructor)
  - [Class methods](#class-methods)
  - [Properties with initial value](#properties-with-initial-value)
  - [getter](#getter)
  - [setter](#setter)
  - [Static method](#static-method)
  - [CommonVBA.Utility](#CommonVBA-Utility)


## Introduction
<!-- TODO: add -->
## Requirments
<!-- TODO: add -->
## Setup
<!-- TODO: add -->
## CodingStyle
<!-- TODO: add -->
```vb
''
' @Purpose:  Get Corresponding sheet
' @Param  :  {Workbook} Book
'            {String}   sheetname　
' @Return :　Worksheet
''
```

## CommonVBA-Utility
This is utility module with lots of reuseable functions for VBA
* Function to get worksheet object

```vb
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
```
* Function to get last row or column count number
```vb
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
```

* Function to get workbook
```vb
''
' @Purpose:  Get Corresponding workbook
' @Param  :  {String}    Name of workbook
' @Return :　{Workbook}  Corresponding workbook if it find the workbook otherwise Nothing
''
Public Function GetWorkbook(ByVal WorkBookName As String) As Workbook
    Dim EachWorkbook As Object
    
    If Not Trim(WorkBookName) = vbNullString Then
        For Each EachWorkbook In Excel.Workbooks
            If EachWorkbook.Name = WorkBookName Then
                  Set GetWorkbook = EachWorkbook
                  Exit Function
            End If
        Next EachWorkbook
    End If
    
    Set GetWorkbook = Nothing
    
End Function 
```

* Function to find dynamic array is empty or not
([source from cpearson](http://www.cpearson.com/excel/IsArrayAllocated.aspx))
```vb
''
' @Purpose:  Find out dynamci array allocated or not
' @Param  :  {Varaint}  Arr
' @Return :　{Boolean}  Return True if Arr is a valid and allocted array
''
Function IsArrayAllocated(Arr As Variant) As Boolean
        On Error Resume Next
        IsArrayAllocated = IsArray(Arr) And _
                           Not IsError(LBound(Arr, 1)) And _
                           LBound(Arr, 1) <= UBound(Arr, 1)
End Function                           
```

* Function to check whether given file exist or not
```vb
''
' @Purpose:  Check whether given file exist or not
' @Param  :  {String} FilePath
' @Return :　{Boolean} True if successful
''
Public Function FileExist(FilePath As String) As Boolean
Dim GetFile As String
Dim FileExistResult As Boolean
    
     GetFile = Dir(FilePath)
    
     If GetFile <> "" Then
            FileExistResult = True
     End If
    
    'return
    FileExist = FileExistResult
    
End Function
```

* Function to check whether input is valid or not

```vb
''
' @Purpose:  Check whether TextBox is valid or not
' @Param  :  {String} ctrlName *Optional
' @Return :　Nothing
''
Public Sub ValidateForm(Optional ctrlName As String)
Dim ctrl As Object
Dim ControlName As String

    For Each ctrl In Me.Controls
        If TypeName(ctrl) = "TextBox" Then
            ControlName = IIf(ctrlName <> "", ctrlName, ctrl.ControlName)
            If IsNull(ctrl.Value) Then
                ctrl.SetFocus
                MsgBox ControlName & "に入力してから実行してください。"
                End
            End If
        End If
    Next ctrl
End Sub

```

