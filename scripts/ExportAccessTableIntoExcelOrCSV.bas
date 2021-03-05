

''
' @Purpose:  Export tables from access db to excel or csv
''
Public Sub Export_Table_TO_Excel_OR_CSV()

  Dim strOut As String
  Dim tbl As AccessObject
  Dim isExportToExcel As Boolean

  With Application.FileDialog(4) ' msoFileDialogFolderPicker

    .Title = "Please select the target folder"
    
    'exit
    If Not .Show Then
          MsgBox "You didn't select a target folder.", vbExclamation
          Exit Sub
    End If
    
    strOut = .SelectedItems(1)
    If Not Right(strOut, 1) = "\" Then
      strOut = strOut & "\"
    End If

  End With
 
  'Prompt user to select excel or CSV
  isExportToExcel = (MsgBox("Do you want to export all tables to Excel (No = CSV)?", vbQuestion + vbYesNo) = vbYes)

  'Looping All table from access db
  For Each tbl In CurrentData.AllTables
            'donot export "MSys* and "~"
            If Not tbl.Name Like "MSys*" And Not tbl.Name Like "~" Then
                    If isExportToExcel Then
                      'Export to Excel file
                      DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, _
                                  tbl.Name, strOut & tbl.Name & ".xlsx", True
                    Else
                      'Export to CSV file
                      DoCmd.TransferText acExportDelim, , _
                               tbl.Name, strOut & tbl.Name & ".csv", True
                    End If  '//isExportToExcel
            End If  '//tbl.Name Like "MSys*"

  Next tbl

End Sub
