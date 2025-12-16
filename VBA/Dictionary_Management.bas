Option Explicit

#Const Windows = (Mac = 0)

' ================= Update all calls to the following two subs to use AddValueToDictionary instead =================
' Private Sub AddToFileNamesHashesAndUrlsDictionary(ByVal dicKey As String, ByVal dicValue As String)
' Public Sub AddToMessagesDictionary(ByVal dicKey As String, ByVal dicValue As String)
' Public Sub AddToShadingDictionary(ByRef shadingDictionary As Dictionary, ByVal cellToShade As String, ByVal shadingValue As Long)

' ================= Update these calls to be ReadValueFromDictionary(g_dictFileData, fontName, "englishdisplayname") =================
' Public Function GetFontDisplayName(ByVal fontName As String) As String

' ================= Update these calles to be ReadValueFromDictionary(g_dictFileData, key, suffix) instead =================
' Public Function GetDictionaryValue(ByVal key As String, ByVal keySuffix As String) As String

' ================= End of update =================

Public Sub AddValueToDictionary(ByRef dictName As Dictionary, ByVal dictKey As String, ByVal dictValue As Variant)
    If Not dictName.Exists(dictKey) Then
        dictName.Add dictKey, dictValue
    Else
        dictName(dictKey) = dictValue
    End If
End Sub

Public Function ReadValueFromDictionary(ByRef dictName As Dictionary, ByVal Key As String, Optional ByVal suffix As String = vbNullString) As Variant
    If suffix <> vbNullString Then
        Key = Key & "." & suffix
    End If
    
    If dictName.Exists(Key) Then
        ReadValueFromDictionary = dictName(Key)
    Else
        ReadValueFromDictionary = "Entry not found: " & Key & IIf(suffix <> vbNullString, vbNullString, "." & suffix)
    End If
End Function

Public Function GetDownloadUrl(ByVal fileType As String, Optional ByVal fileName As String = vbNullString) As String
    Const BASE_JSON_URL As String = "https://raw.githubusercontent.com/papercutter0324/SpeakingEvals/main/JSON/"
    
    If fileType = "JSON" Then
        GetDownloadUrl = BASE_JSON_URL & fileName
    Else
        GetDownloadUrl = ReadValueFromDictionary(g_dictFileData, fileType, "url")
    End If
End Function

Public Function GetMsg(ByVal Key As String, ParamArray args() As Variant) As String
    Dim tempArray() As Variant
    Dim i           As Long
    
    KeyHasValidSuffix Key
    
    If g_dictMessages.Exists(Key) Then
        If IsMissing(args) Then
            GetMsg = CStr(UpdateMsgPlaceholders(g_dictMessages(Key)))
        Else
            ReDim tempArray(LBound(args) To UBound(args))
            For i = LBound(args) To UBound(args)
                tempArray(i) = args(i)
            Next i
            
            GetMsg = UpdateMsgPlaceholders(g_dictMessages(Key), tempArray)
        End If
    Else
        GetMsg = "Entry not found: " & Key
    End If
End Function

Private Function KeyHasValidSuffix(ByRef Key As String) As Boolean
    Dim suffixList   As Variant
    Dim suffixOption As Variant
    
    suffixList = Array(".Message", ".Buttons", ".Icon", ".Title", ".Width", ".Messages", ".FileUrls", ".FileHashes", ".Number", ".Text")
    
    For Each suffixOption In suffixList
        If Right$(Key, Len(suffixOption)) = suffixOption Then
            KeyHasValidSuffix = True
            Exit Function
        End If
    Next suffixOption
    
    Key = Key & ".Message"
End Function

Public Function UpdateMsgPlaceholders(ByRef msgText As String, Optional ByVal args As Variant) As String
    Dim i As Long

    If Not IsMissing(args) Then
        For i = LBound(args) To UBound(args)
            msgText = Replace(msgText, "{" & i & "}", CStr(args(i)))
        Next i
    End If
    
    msgText = Replace(msgText, "{INDENT1}", INDENT_LEVEL_1)
    msgText = Replace(msgText, "{INDENT2}", INDENT_LEVEL_2)
    msgText = Replace(msgText, "{INDENT3}", INDENT_LEVEL_3)
    msgText = Replace(msgText, "{NEWLINE1}", vbNewLine)
    msgText = Replace(msgText, "{NEWLINE2}", vbNewLine & vbNewLine)
    
    UpdateMsgPlaceholders = msgText
End Function

Public Sub InitializeDictionaries(Optional ByVal jsonFolder As String = vbNullString, Optional ByVal requestFilePermissions As Boolean = False)
    Const FILE_RECORDS_JSON As String = "dictFileNamesHashesAndUrls.json"
    Const MSGS_RECORDS_JSON As String = "dictMsgRecords.json"

    Dim notDefaultPath   As Boolean
    Dim msgsJsonPath     As String
    Dim msgsJsonExists   As Boolean
    Dim fileJsonPath     As String
    Dim fileJsonExists   As Boolean
    Dim jsonFilesUpdated As Boolean
    
    If jsonFolder = vbNullString Then jsonFolder = GetDefaultFolderPaths("JSON")

    If Not CheckForAndAttemptToCreateFolder(jsonFolder) Then
        jsonFolder = Left$(jsonFolder, Len(jsonFolder) - 10)
        notDefaultPath = True
    End If
    
    msgsJsonPath = jsonFolder & MSGS_RECORDS_JSON
    msgsJsonExists = DoesFileExist(msgsJsonPath, requestFilePermissions)
    
    fileJsonPath = jsonFolder & FILE_RECORDS_JSON
    fileJsonExists = DoesFileExist(fileJsonPath, requestFilePermissions)
    
    If msgsJsonExists And fileJsonExists Then
        VerifyDictionariesAreLoaded
        jsonFilesUpdated = ProcessJsonUpdate(MSGS_RECORDS_JSON, msgsJsonPath, CLng(GetMsg("Version.Number"))) Or _
                           ProcessJsonUpdate(FILE_RECORDS_JSON, fileJsonPath, CLng(ReadValueFromDictionary(g_dictFileData, "Version", "Number")))
    Else
        If Not msgsJsonExists Then DownloadFile MSGS_RECORDS_JSON, "JSON", msgsJsonPath
        If Not fileJsonExists Then DownloadFile FILE_RECORDS_JSON, "JSON", fileJsonPath
        jsonFilesUpdated = True
    End If
    
    If jsonFilesUpdated Then
        Set g_dictMessages = Nothing
        Set g_dictFileData = Nothing
        VerifyDictionariesAreLoaded
    End If

    If notDefaultPath Then
        DeleteFile msgsJsonPath
        DeleteFile fileJsonPath
    End If
End Sub

Public Function VerifyDictionariesAreLoaded() As Boolean
    Const FILE_RECORDS_JSON   As String = "dictFileNamesHashesAndUrls.json"
    Const MSGS_RECORDS_JSON   As String = "dictMsgRecords.json"
    Const MSGS_TEST_KEY       As String = "EntryTests.Messages"
    Const URL_ENTRY_NOT_FOUND As String = "Entry not found: EntryTests.FileUrls"
    Const MSG_ENTRY_NOT_FOUND As String = "Entry not found: EntryTests.Messages"
    
    Static jsonFolder As String
    
    Dim fileRecordsJsonIsLoaded As Boolean
    Dim msgsRecordsJsonIsLoaded As Boolean
    
    If jsonFolder = vbNullString Then
        jsonFolder = GetDefaultFolderPaths("JSON")
    End If

    fileRecordsJsonIsLoaded = VerifyFileRecordsDictionaryIsLoaded("EntryTests", "FileUrls", URL_ENTRY_NOT_FOUND)
    If Not fileRecordsJsonIsLoaded Then
        fileRecordsJsonIsLoaded = LoadDictionary(jsonFolder, FILE_RECORDS_JSON, "FileUrls")
    End If

    msgsRecordsJsonIsLoaded = Not (GetMsg(MSGS_TEST_KEY) = MSG_ENTRY_NOT_FOUND)
    If Not msgsRecordsJsonIsLoaded Then
        msgsRecordsJsonIsLoaded = LoadDictionary(jsonFolder, MSGS_RECORDS_JSON, "Messages")
    End If
    
    VerifyDictionariesAreLoaded = fileRecordsJsonIsLoaded And msgsRecordsJsonIsLoaded
End Function

Private Function LoadDictionary(ByVal dictPath As String, ByVal jsonFileName As String, ByVal dictCheckType As String) As Boolean
    Dim jsonExists As Boolean
    Dim dictionaryLoaded As Boolean
    
    jsonExists = DoesFileExist(dictPath & jsonFileName)
    
    If Not jsonExists Then
        jsonExists = DownloadFile(jsonFileName, vbNullString, dictPath & jsonFileName)
    End If

    Select Case dictCheckType
        Case "FileUrls", "FileHashes"
            dictionaryLoaded = LoadValuesFromJson(LoadDataFromJson(dictPath & jsonFileName), vbNullString, g_dictFileData)
        Case "Messages"
            dictionaryLoaded = LoadValuesFromJson(LoadDataFromJson(dictPath & jsonFileName), vbNullString, g_dictMessages)
    End Select
    
    LoadDictionary = dictionaryLoaded
End Function

Public Function LoadDataFromJson(ByVal jsonFilePath As String) As Object
    Dim fileNum  As Long
    Dim jsonText As String
    
    fileNum = FreeFile
    Open jsonFilePath For Input As #fileNum
        jsonText = Input$(LOF(fileNum), fileNum)
    Close #fileNum
    
    Set LoadDataFromJson = Parse(jsonText).Value
End Function

Public Function LoadValuesFromJson(ByRef obj As Object, Optional ByVal prefix As String, Optional ByRef dict As Dictionary) As Boolean
    Dim Key       As Variant
    Dim newPrefix As String
    
    On Error GoTo LoadError
    
    If dict Is Nothing Then
        Err.Raise vbObjectError + 1000, "LoadValuesFromJson", "Dictionary object is required."
    End If
    
    For Each Key In obj.Keys
        newPrefix = IIf(prefix = vbNullString, Key, prefix & "." & Key)
        
        If IsObject(obj(Key)) Then
            ' Recursive call a¢æ¡± if it fails, bubble up the False result
            If Not LoadValuesFromJson(obj(Key), newPrefix, dict) Then
                LoadValuesFromJson = False
                Exit Function
            End If
            ' Old code in case this update doesn't work
            ' LoadValuesFromJson obj(key), newPrefix, dict
        Else
            dict(newPrefix) = obj(Key)
        End If
    Next Key

    LoadValuesFromJson = True
    Exit Function

LoadError:
    LoadValuesFromJson = False
End Function

Public Function VerifyFileRecordsDictionaryIsLoaded(ByVal Key As String, ByVal suffix As String, ByVal dictNotLoadedMsg As String) As Boolean
    Dim dictionaryNotLoaded As Boolean

    dictionaryNotLoaded = (ReadValueFromDictionary(g_dictFileData, Key, suffix) = dictNotLoadedMsg)
    VerifyFileRecordsDictionaryIsLoaded = Not dictionaryNotLoaded
End Function