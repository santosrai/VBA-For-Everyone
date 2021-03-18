Option Compare Database
Option Explicit

  
'/**
' * VBA-Logger v1.0.0
' * (c) Santosh Rai
' * @Purpose: Export all logs
' * @Pre-condition:
'            Public logPath as String --> Name log folder
'            Public logFilename as String  --> Application Name
' * @Param:  {String} sStatus
'                   {String} sTitleName
'                   {String} sMessage
' * @Ref:    Microsoft Scripting Runtime
' * @Use:    1.Download logger.bas or copy this whole code
'                  2.Change folder name with replacing XXXXX (line no.18)
'                  3.Insert or call Log with level function inside your code where you want to log
'                       - LogTrace "Start logging", user-defined function name
'                  4. Call Log file with LogFile_Open function
'           ---Huge thanks to timhall----
' */


' Levels
' 0: Off
' 1: Trace/All
' 2: Debug
' 3: Info
' 4: Warn
' 5: Error

Public LogThreshold As Long    'To enble or disable log
Public pLogged As Variant       ' For unit-testing


'/**
' * @property LogEnabled
' * @type Boolean
' * @default False
' */
Public Property Get LogEnabled() As Boolean
    If LogThreshold = 0 Then
        LogEnabled = False
    Else
        LogEnabled = True
    End If
End Property

Public Property Let LogEnabled(Value As Boolean)
    If Value Then
        LogThreshold = 1
    Else
        LogThreshold = 0
    End If
End Property

'/**
' * @Purpose:  Log
' * @Param: {Long} Level
' * @Param: {String} [sTitleName = ""]
' * @Param: {String} Message
' */
Public Sub log(Level As Long, sMessage As String, sTitleName As String)
Dim log_LevelValue As String
    If LogThreshold = 0 Or Level < LogThreshold Then
        pLogged = Array(Level, sMessage, sTitleName)
        Exit Sub
    End If

     Select Case Level
        Case 1
            log_LevelValue = "Trace"
        Case 2
            log_LevelValue = "Debug"
        Case 3
            log_LevelValue = "Info "
        Case 4
            log_LevelValue = "WARN "
        Case 5
            log_LevelValue = "ERROR"
    End Select
        
    ' Write to log file
    LogFile_Write log_LevelValue, sTitleName, sMessage
    
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
Public Function getLog_FolderPath()
      Dim FolderPath As String
      Dim folderName As String
      
      'change folder name only
      folderName = "XXXXX_log"
      FolderPath = "C:\Users\" & VBA.Environ$("username") & "\AppData\Roaming\" & folderName & "\"

      ' Return
      getLog_FolderPath = FolderPath
End Function

'/**
' * @Purpose: Export all logs
' * @Pre-condition:
'            Public logPath as String --> Name log folder
'            Public logFilename as String  --> Application Name
' * @Param:  {String} sStatus
'            {String} sTitleName
'            {String} sMessage
' * @Ref:    Microsoft Scripting Runtime
' */
Public Sub LogFile_Write(ByVal sStatus As String, ByVal sTitleName As String, _
                             ByVal sMessage As String)
                             
On Error GoTo ErrorHandler
   Dim sText As String
   Dim objFSO As Scripting.FileSystemObject
   Dim scrText As Scripting.TextStream
   Dim logFile As String
   Dim FileName As String
   Dim FolderPath As String
   
    FolderPath = getLog_FolderPath
    FileName = getLog_FileName
     
   'log full file
   logFile = getLog_FolderPath & "\" & getLog_FileName
   
    'create folder
    Create_Folder FolderPath

   ' initialise filesystem object
   If (objFSO Is Nothing) Then
      Set objFSO = New FileSystemObject
   End If

   ' IO mode for file
   If (scrText Is Nothing) Then
      If (objFSO.FileExists(logFile) = False) Then
         Set scrText = objFSO.OpenTextFile(logFile, 2, True) 'IOMode.ForWriting
      Else
         Set scrText = objFSO.OpenTextFile(logFile, 8) 'IOMode.ForAppending
      End If
   End If

   ' writing status and message
   sText = sText
   sText = sText & Format(Date, "yyyy/MM/dd") & "-" & Time()
   sText = sText & " " & sStatus
   sText = sText & " {""" & "title"":""" & sTitleName & """"
   sText = sText & "," & """message"":""" & sMessage & """}"
   scrText.WriteLine sText
   scrText.Close

   Set scrText = Nothing
   Exit Sub
ErrorHandler:
   Set scrText = Nothing
   Debug.Print Err.Number & " unable to write in log file" & "  : " & Err.Description
End Sub

'/**
' * @Purposes: Open the log file with notepad
' */
Public Sub LogFile_Open()
   Dim objFSO As Scripting.FileSystemObject
   Dim logFile As String
   
    ' initialise filesystem object
   If (objFSO Is Nothing) Then
      Set objFSO = New FileSystemObject
   End If
   
   ' log file
    logFile = getLog_FolderPath & "\" & getLog_FileName
    
    ' open notepad
    If objFSO.FileExists(logFile) = True Then
       Shell "notepad.exe """ & logFile & """", vbNormalFocus
    Else
       MsgBox "Log file not found"
    End If
    
End Sub
                                    
                                    
'/**
' * @Purposes: Create folder on given path
' * @Param:    {String}  Folder path
' * @Return:    {Boolean} True if it successfully create folder
' */
Public Function Create_Folder(FolderPath As String) As Boolean
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
                        
    'if folder path already exist
    If FSO.FolderExists(FolderPath) Then Exit Function
    
    'Create folder
    FSO.CreateFolder FolderPath
    ' Return
    Create_Folder = True

End Function

'/**
' * @Purposes: LogTrace
' * @Param: {String} Message
' * @Param: {String} From
' */
Public Sub LogTrace(Message As String, Optional sTitleName As String = "")
    log 1, Message, sTitleName
End Sub

'/**
' * @Purposes: LogDebug
' * @Param: {String} Message
' * @Param: {String} sTitleName
' */
Public Sub LogDebug(Message As String, Optional sTitleName As String = "")
    log 2, Message, sTitleName
End Sub

'/**
' * @Purposes: LogInfo
' * @Param: {String} Message
' * @Param: {String} sTitleName
' */
Public Sub LogInfo(Message As String, Optional sTitleName As String = "")
    log 3, Message, sTitleName
End Sub

'/**
' * @Purposes: LogWarning
' * @Param: {String} Message
' * @Param: {String} sTitleName
' */
Public Sub LogWarn(Message As String, Optional sTitleName As String = "")
    log 4, Message, sTitleName
End Sub

'/**
' * @Purposes: LogError
' * @Param:  {String} Message
' * @Param:  {String} [sTitleName = ""]
' * @Param:  {Long} [ErrNumber = 0]
' * @Note:     got from timhall
' */
Public Sub LogError(Message As String, Optional sTitleName As String = "", Optional ErrNumber As Long = 0)
    Dim log_ErrorValue As String
    If ErrNumber <> 0 Then
        log_ErrorValue = ErrNumber
    
        ' For object errors, extract from vbObjectError and get Hex value
        If ErrNumber < 0 Then
            log_ErrorValue = log_ErrorValue & " (" & (ErrNumber - vbObjectError) & " / " & VBA.LCase$(VBA.Hex$(ErrNumber)) & ")"
        End If
        
        log_ErrorValue = log_ErrorValue & ", "
    End If

    log 5, log_ErrorValue & Message, sTitleName
End Sub
