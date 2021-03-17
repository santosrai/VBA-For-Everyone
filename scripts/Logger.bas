'/**
' * VBA-Logger v1.0.0
' * (c) Santosh Rai
' * @Purpose: Export all logs
' * @Pre-condition: 
'            Public logPath as String --> Name log folder
'            Public logFilename as String  --> Application Name
' * @Param:  {String} sStatus
'            {String} sTitleName   
'            {String} sMessage
' * @Ref:    Microsoft Scripting Runtime
' * @Use:    1.Download logger.bas or copy this whole code
'            2.change folder name with replacing XXXXX (line no.18)
'            3.insert or call LogFile_Write function inside your code where you want to log
' * @Note:  sStatus can be {Info,Trace,Warn,Error,Debug}
' */
Public Sub LogFile_Write(ByVal sStatus As String, ByVal sTitleName As String, _
                         ByVal sMessage As String)
                             
On Error GoTo ErrorHandler
   Dim sText As String
   Dim g_objFSO As Scripting.FileSystemObject
   Dim g_scrText As Scripting.TextStream
   Dim logFile As String

   'log file
   logFile = getLog_FolderPath & "\" & getLog_FileName

   ' initialise filesystem object
   If (g_objFSO Is Nothing) Then
      Set g_objFSO = New FileSystemObject
   End If

   ' IO mode for file
   If (g_scrText Is Nothing) Then
      If (g_objFSO.FileExists(logFile) = False) Then
         Set g_scrText = g_objFSO.OpenTextFile(logFile, IOMode.ForWriting, True)
      Else
         Set g_scrText = g_objFSO.OpenTextFile(logFile, IOMode.ForAppending)
      End If
   End If

   ' writing status and message
   sText = sText
   sText = sText & Format(Date, "yyyy/MM/dd") & "-" & Time()
   sText = sText & " " & sStatus
   sText = sText & " {""" & "title"":""" & sTitleName & """"
   sText = sText & "," & """message"":""" & sMessage & """}"
   g_scrText.WriteLine sText
   g_scrText.Close

   Set g_scrText = Nothing
   Exit sub
ErrorHandler:
   Set g_scrText = Nothing
   Debug.Print "unable to write in log file"
End Sub

'/**
' * @Purposes: Open the log file
' */
Public Sub LogFile_Open()
   Dim objFSO As New Scripting.FileSystemObject
   Dim logFile As String
   
   'log file
    logFile = getLog_FolderPath & "\" & getLog_FileName
 
    If objFSO.FileExists(logFile) = True Then
       Shell "notepad.exe """ & logFile & """", vbNormalFocus
    Else
       MsgBox "Log file not found"
    End If
    
End Sub

'/**
' * @Purposes: Provide log file Name
' * @Return  : {string} file name
' */
Public Function getLog_FileName()
      getLog_FileName = Format(Now(), "yyyy-MM-dd") & "_ExecutionLog.txt"
End Function

'/**
' * @Purposes: Provide log folder path
' * @Return  : {string} folder path
' */
Public function getLog_FolderPath()
      Dim folderPath as String
      Dim folderName as String
      
      'change folder name only
      folderName = "XXXXX_log"
      folderPath ="C:\Users\" & VBA.Environ$("username") & "\AppData\Roaming\" & folderName & "\"
      'create folder
      Create_Folder folderPath

      ' Return
      getLog_FolderPath = folderPath
End Function
                                    
'/**
' * @Purposes: Create folder on given path
' * @Param:    {String}  Folder path
' * @Return:    {Boolean} True if it successfully create folder
' */
Public Function Create_Folder(folderPath As String) As Boolean
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
                        
    'if folder path already exist
    If FSO.FolderExists(folderPath) Then Exit Function
    
    'Create folder
    FSO.CreateFolder folderPath
    ' Return
    Create_Folder = True
 
End Function
