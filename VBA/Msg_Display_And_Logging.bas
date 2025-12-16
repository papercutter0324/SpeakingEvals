Option Explicit

#Const Windows = (Mac = 0)

Public Function DisplayMessage(ByVal msgKey As String, ParamArray args() As Variant) As Variant
#If Mac Then
    Dim dialogParameters As String
    Dim dialogResult As Variant
    Dim messageDisplayed As Boolean
    Dim msgWidth As Long
#End If
    Dim msgText As String
    Dim msgTitle As String
    Dim msgType As Long
    Dim dialogType As String
    Dim iconType As String
    Dim tempArray() As Variant
    Dim isArrayEmpty As Boolean
    Dim i As Long
    
    msgText = GetMsg(msgKey & ".Message")
    msgTitle = GetMsg(msgKey & ".Title")
    dialogType = GetMsg(msgKey & ".Buttons")
    iconType = GetMsg(msgKey & ".Icon")
    
#If Mac Then
    msgWidth = CLng(GetMsg(msgKey & ".Width"))
#End If
    
    If IsMissing(args) Then
        msgText = UpdateMsgPlaceholders(msgText)
    Else
        ReDim tempArray(LBound(args) To UBound(args))
        For i = LBound(args) To UBound(args)
            tempArray(i) = args(i)
        Next i
        
        msgText = UpdateMsgPlaceholders(msgText, tempArray)
    End If
    
#If Mac Then
    If AreEnhancedDialogsEnabled Then
        If g_UserOptions.EnableLogging Then
            DebugAndLogging GetMsg("Debug.DialogToolKitPlus.AttemptToUse", msgText)
        End If
                
        ' Update SpeakingEvals.scpt to supoport:
        ' dialogParameters = msgText & APPLE_SCRIPT_SPLIT_KEY & dialogType & APPLE_SCRIPT_SPLIT_KEY & iconType & APPLE_SCRIPT_SPLIT_KEY & msgTitle & APPLE_SCRIPT_SPLIT_KEY & msgWidth
        dialogParameters = msgText & APPLE_SCRIPT_SPLIT_KEY & dialogType & APPLE_SCRIPT_SPLIT_KEY & msgTitle & APPLE_SCRIPT_SPLIT_KEY & msgWidth
        
        On Error Resume Next
        dialogResult = AppleScriptTask("DialogDisplay.scpt", "DisplayDialog", dialogParameters)
        On Error GoTo 0
        
        If dialogResult = vbNullString Then
            If g_UserOptions.EnableLogging Then
                DebugAndLogging GetMsg("Debug.DialogToolKitPlus.FailedToDisplay", Err.Number & " - " & Err.Description)
            End If
        Else
            DisplayMessage = dialogResult
            Exit Function
        End If
    End If
#End If
    
    msgType = GetVbButtonCode(dialogType) + GetVbIconCode(iconType)
    DisplayMessage = MsgBox(msgText, msgType, msgTitle)
End Function

Private Function GetVbButtonCode(ByVal msgType As String) As Long
    Static buttonMap As New Dictionary
    
    If Not buttonMap.Exists("Test") Then
        buttonMap("Test") = "Okay"
        buttonMap("OKOnly") = vbOKOnly
        buttonMap("OKCancel") = vbOKCancel
        buttonMap("YesNo") = vbYesNo
        buttonMap("YesNoCancel") = vbYesNoCancel
        buttonMap("RetryCancel") = vbRetryCancel
        buttonMap("AbortRetryIgnore") = vbAbortRetryIgnore
    End If
    
    If buttonMap.Exists(msgType) Then
        GetVbButtonCode = buttonMap(msgType)
    Else
        GetVbButtonCode = vbOKOnly ' default fallback
    End If
End Function

Private Function GetVbIconCode(ByVal msgIcon As String) As Long
    Static iconMap As New Dictionary
    
    If Not iconMap.Exists("Test") Then
        iconMap("Test") = "Okay"
        iconMap("Critical") = vbCritical
        iconMap("Question") = vbQuestion
        iconMap("Exclamation") = vbExclamation
        iconMap("Information") = vbInformation
    End If
    
    If iconMap.Exists(msgIcon) Then
        GetVbIconCode = iconMap(msgIcon)
    Else
        GetVbIconCode = vbInformation ' default fallback
    End If
End Function

Public Sub DebugAndLogging(ByVal debugMsg As String, Optional ByVal startNewLogFile As Boolean = False, Optional ByVal WriteLog As Boolean = False)
    Const DATE_FORMAT           As String = "dd-mmm-yy"
    Const DATE_AND_TIME_FORMAT  As String = "yyyy-mm-dd hh:nn:ss"
    Const ENTRY_DIVIDER         As String = "--------------------------------------------------------------------------" & vbNewLine
    Const LOG_FILENAME_PREFIX   As String = "SpeakingEvalsLog_"
    Const LOG_FILENAME_FILETYPE As String = ".txt"
    
    Static logData     As String
    Static logFolder   As String
    Static logFileName As String
    
    Debug.Print debugMsg
    
    If logFolder = vbNullString Then
        logFolder = GetDefaultFolderPaths("Logs")
        
        If Not CheckForAndAttemptToCreateFolder(logFolder) Then
            logFolder = vbNullString
            Exit Sub
        End If
    End If
    
    If logFileName = vbNullString Then
        logFileName = LOG_FILENAME_PREFIX & Format$(Date, DATE_FORMAT) & LOG_FILENAME_FILETYPE
        
        If Not DoesFileExist(logFolder & logFileName, True) Then
            If Not CreateNewLogFile(logFolder, logFileName) Then
                Exit Sub
            End If
        End If
    End If
    
    If startNewLogFile Then
        logData = ENTRY_DIVIDER & Format$(Now, DATE_AND_TIME_FORMAT) & vbNewLine & vbNewLine
    End If
    
    logData = logData & vbNewLine & debugMsg
    
    If WriteLog Then
        WriteLogToFile logData & vbNewLine, logFolder & logFileName
        logData = vbNullString
    End If
End Sub

Private Function CreateNewLogFile(ByVal logFolder As String, ByVal logFileName As String) As Boolean
    On Error GoTo FileCreationFailed
#If Mac Then
    CreateNewLogFile = AppleScriptTask(APPLE_SCRIPT_FILE, "CreateTextFile", logFolder & APPLE_SCRIPT_SPLIT_KEY & logFileName)
#Else
    Dim LogFile As Object: Set LogFile = CreateObject("ADODB.Stream")
    
    With LogFile
        .Type = 2 ' adTypeText
        .Charset = "UTF-8"
        .Open
        .SaveToFile logFolder & logFileName
        .Close
    End With
#End If
    On Error GoTo 0

    CreateNewLogFile = True
    Exit Function
FileCreationFailed:
    CreateNewLogFile = False
End Function

Private Sub WriteLogToFile(ByVal logData As String, ByVal logPath As String)
#If Mac Then
    Dim writeSuccessful As Boolean
    writeSuccessful = AppleScriptTask(APPLE_SCRIPT_FILE, "WriteToLog", logPath & APPLE_SCRIPT_SPLIT_KEY & logData)
#Else
    Dim LogFile As Object: Set LogFile = CreateObject("ADODB.Stream")
    
    With LogFile
        .Type = 2 ' adTypeText
        .Charset = "UTF-8"
        .Open
        On Error Resume Next
        .LoadFromFile logPath
        On Error GoTo 0
        .Position = .Size
        .WriteText logData
        .SaveToFile logPath, 2 ' adSaveCreateOverWrite
        .Close
    End With
#End If
End Sub

Public Sub CleanUpOldLogs(ByVal logFolder As String)
    Const LOG_NAMING_CONVENTION As String = "SpeakingEvalsLog_??-???-??*.txt"
    Const MAX_LOGS              As Long = 5
    
    Dim logNamesMacOS() As String
    Dim logFiles()      As Variant
    Dim logCount        As Long
    Dim i               As Long
    Dim errEncountered  As Boolean
    
    logCount = GetLogCount(logFolder, LOG_NAMING_CONVENTION, logNamesMacOS(), errEncountered)

    If logCount <= MAX_LOGS Or errEncountered Then
        Exit Sub
    End If
    
    ReDim logFiles(1 To logCount)
    logCount = CreateListOfLogFiles(logFiles(), logNamesMacOS(), logFolder, LOG_NAMING_CONVENTION)
    
    BubbleSortOne2DArray logFiles(), logCount
    
    For i = MAX_LOGS + 1 To logCount
        DeleteFile logFiles(i)(1)
    Next i
End Sub

Private Sub BubbleSortOne2DArray(ByRef unsortedArray() As Variant, Optional ByVal upperBound As Long = 0, Optional ByVal lowerBound As Long = 0)
    Dim i As Long
    Dim j As Long
    Dim swapped As Boolean
    
    If lowerBound < 1 Then
        lowerBound = LBound(unsortedArray)
    End If
    
    If upperBound < 1 Then
        upperBound = UBound(unsortedArray)
    End If
    
    For i = lowerBound To upperBound - 1
        swapped = False
        For j = lowerBound To upperBound - (i - lowerBound) - 1
            If unsortedArray(j)(0) < unsortedArray(j + 1)(0) Then
                SwapPlaces unsortedArray(j), unsortedArray(j + 1)
                swapped = True
            End If
        Next j
        If Not swapped Then
            Exit For
        End If
    Next i
End Sub

Private Function GetLogCount(ByVal logFolder As String, ByVal logFileNamingConvention As String, ByRef logNamesMacOS() As String, ByRef errEncountered As Boolean) As Long
#If Mac Then
    Dim appleScriptResult As String
    Dim i        As Long
#Else
    Dim fso         As Object
    Dim fsoFolder   As Object
    Dim fsoFile     As Object
#End If
    
    Dim logCount As Long
    
#If Mac Then
    On Error Resume Next
    appleScriptResult = AppleScriptTask(APPLE_SCRIPT_FILE, "ListFolderContents", logFolder & APPLE_SCRIPT_SPLIT_KEY & "txt")

    If Err.Number <> 0 Then
        ' Display message?
        Exit Function
    End If
    On Error GoTo 0

    logNamesMacOS() = Split(appleScriptResult, APPLE_SCRIPT_SPLIT_KEY)

    For i = LBound(logNamesMacOS) To UBound(logNamesMacOS)
        If logNamesMacOS(i) Like logFileNamingConvention Then
            logCount = logCount + 1
        End If
    Next i
#Else
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fsoFolder = fso.GetFolder(logFolder)

    For Each fsoFile In fsoFolder.Files
        If fsoFile.Name Like logFileNamingConvention Then
            logCount = logCount + 1
        End If
    Next fsoFile
#End If

    GetLogCount = logCount
End Function

Private Function CreateListOfLogFiles(ByRef logFiles() As Variant, ByRef logNamesMacOS() As String, ByVal logFolder As String, ByVal logFileNamingConvention As String) As Long
    Dim datePart As String
    Dim fileName As String
    Dim filePath As String
    Dim fileDate As Date
    Dim logCount As Long
    
#If Mac Then
    Dim i        As Long
#Else
    Dim fso         As Object
    Dim fsoFolder   As Object
    Dim fsoFile     As Object
#End If

#If Mac Then
    For i = LBound(logNamesMacOS) To UBound(logNamesMacOS)
        fileName = logNamesMacOS(i)
        
        If fileName Like logFileNamingConvention Then
            filePath = logFolder & fileName
            
            If Len(fileName) >= 27 Then
                datePart = Mid$(fileName, 18, 9)
                On Error Resume Next
                fileDate = DateValue(datePart)
                On Error GoTo 0
                
                If fileDate > 0 Then
                    logCount = logCount + 1
                    logFiles(logCount) = Array(fileDate, filePath)
                End If
            End If
        End If
    Next i
#Else
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fsoFolder = fso.GetFolder(logFolder)
    
    For Each fsoFile In fsoFolder.Files
        fileName = fsoFile.Name
        
        If fileName Like logFileNamingConvention Then
            filePath = logFolder & fileName
            
           If Len(fileName) >= 27 Then
                datePart = Mid$(fileName, 18, 9)
                On Error Resume Next
                fileDate = DateValue(datePart)
                On Error GoTo 0
                
                If fileDate > 0 Then
                    logCount = logCount + 1
                    logFiles(logCount) = Array(fileDate, filePath)
                End If
            End If
        End If
    Next fsoFile
#End If

    CreateListOfLogFiles = logCount
End Function