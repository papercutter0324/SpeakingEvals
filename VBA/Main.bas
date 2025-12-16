Option Explicit

#Const Windows = (Mac = 0)

#If Mac Then
    Public Const APPLE_SCRIPT_FILE      As String = "SpeakingEvals.scpt"
    Public Const APPLE_SCRIPT_SPLIT_KEY As String = "-,-"
#End If

Public g_UserOptions As ConfigSettings

Public g_dictMessages           As New Dictionary
Public g_dictFileData           As New Dictionary
' Public ScoreSheetCellTypeMap    As New Dictionary

Public Const INDENT_LEVEL_1     As String = vbTab
Public Const INDENT_LEVEL_2     As String = vbTab & vbTab
Public Const INDENT_LEVEL_3     As String = vbTab & vbTab & vbTab

Public Const g_NATIVE_TEACHER                   As String = "C1"
Public Const g_KOREAN_TEACHER                   As String = "C2"
Public Const g_CLASS_LEVEL                      As String = "C3"
Public Const g_CLASS_DAYS                       As String = "C4"
Public Const g_CLASS_TIME                       As String = "C5"
Public Const g_EVALUATION_DATE                  As String = "C6"
Public Const g_CLASS_INFO                       As String = "C1:C6"
Public Const g_ZIP_FILENAME                     As String = "C2:C5"
Public Const g_ENGLISH_NAMES                    As String = "B8:B32"
Public Const g_KOREAN_NAMES                     As String = "C8:C32"
Public Const g_FULL_NAMES                       As String = "B8:C32"
Public Const g_STUDENT_GRADES                   As String = "D8:I32"
Public Const g_COMMENTS                         As String = "J8:J32"
Public Const g_TEACHER_NOTES                    As String = "K8:M32"
Public Const g_WINNER_NAMES                     As String = "L2:L4"
Public Const g_FIRST_PLACE_WINNER               As String = "L2"
Public Const g_VALIDATION_LIST                  As String = "O8:O32"
Public Const g_ALL_MONITORED                    As String = "C1:C2,C6,B8:J32,L2:L4"
Public Const g_FONT_INSTALLATION_STATUS         As String = "K6"
Public Const g_7ZIP_SUPPORT_STATUS              As String = "K8"
Public Const g_VALID_HASHES_STATUS              As String = "K9"
Public Const g_OPTIONS_SETTINGS                 As String = "K2:K9"
Public Const g_CERTIFICATE_OPTIONS              As String = "K11:K16"
Public Const g_CERTIFICATE_SETTINGS_AND_OPTIONS As String = "J11:K16"

Public Const g_STUDENT_INDEX_OFFSET  As Long = 7
Public Const g_FIRST_STUDENT_ROW     As Long = 8

Public Type ConfigSettings
    OpenSavePathWhenDone    As Boolean
    DisplayEntryTips        As Boolean
    EnableLogging           As Boolean
    DisplayInitialWarning   As Boolean
    AllFontsAreInstalled    As Boolean
    ZipSupportEnabled       As Boolean
    ValidFileHashes         As Boolean
End Type

Public Type ValidationSettings
    TypeOfValidation As Long
    AlertStyle       As Long
    InputTitle       As String
    InputMessage     As String
    Formula          As String
    Formula1         As String
    Formula2         As String
    Operator         As Long
    IgnoreBlank      As Boolean
    InCellDropdown   As Boolean
    ShowInput        As Boolean
    ShowError        As Boolean
End Type

Public Sub Main()
    Dim ws As Worksheet
    Dim wsName As String
    Dim clickedButtonName As String
    Dim startTime As Date
    
    GetRunTime "Start", startTime
    If Not VerifyDictionariesAreLoaded Then
        InitializeDictionaries
    End If
    ReadOptionsValues
    ToggleApplicationFeatures "Disable"
    
    If Not VerifyKeySheetsExist Then
        ' Display Error message
        GoTo ReenableEvents
    End If
    
    Set ws = ActiveSheet
    wsName = ws.Name
    clickedButtonName = Application.Caller
    
    If g_UserOptions.EnableLogging Then
        DebugAndLogging GetMsg("Debug.CodeExecution.EntryPointButton", wsName, clickedButtonName), True
        DebugAndLogging GetMsg("Debug.CodeExecution.BeginningTasks", Format$(startTime, "hh:mm:ss"), wsName, clickedButtonName)
    End If
    
    If IsFileLoadedFromTempDir Then
        #If Mac Then
            DisplayMessage "Display.Workbook.LoadedFromTempMac"
        #Else
            DisplayMessage "Display.Workbook.LoadedFromTempWindows"
        #End If
        GoTo ReenableEvents
    End If
    
    Select Case clickedButtonName
        Case "Button_SignatureEmbedded", "Button_SignatureMissing"
            ToggleEmbeddedSignature clickedButtonName
            GoTo ReenableEvents
        Case "Button_CreateNewClassSheet"
            CreateNewClassRecordsSheet
            GoTo ReenableEvents
        Case "Button_RepairLayout"
            RepairLayouts ws
            GoTo ReenableEvents
        Case "Button_AutoSelectWinners"
            ' This could use a better name in the future
            UpdateWinnersLists ws, True
            GoTo ReenableEvents
#If Mac Then
        Case "Button_EnhancedDialogs_Enable", "Button_EnhancedDialogs_Disable"
            ToggleMacSettingsButtons ws, clickedButtonName
            GoTo ReenableEvents
#End If
    End Select
    
#If Mac Then
    If Not AreAppleScriptsInstalled(, , True) Then
        RemindUserToInstallSpeakingEvalsScpt
        GoTo ReenableEvents
    End If
#End If
    
    Select Case clickedButtonName
        Case "Button_GenerateReports", "Button_GenerateProofs", "Button_GenerateCertificates"
            If g_UserOptions.DisplayInitialWarning Then
                DisplayMessage "Display.PowerPoint.WarnAboutClosingDelay"
            End If
                
            CreateReportsAndCertificates ws, clickedButtonName
            ws.Activate
    End Select
    
ReenableEvents:
    ToggleApplicationFeatures "Enable"
    GetRunTime "End", startTime
End Sub

Public Sub GetRunTime(ByVal timerToggle As String, ByRef startTime As Date)
    Dim endTime As Date
    Dim elapsedTime As Double
    
    Select Case timerToggle
        Case "Start"
            startTime = Now
        Case "End"
            endTime = Now
            elapsedTime = Now - startTime
            
            If g_UserOptions.EnableLogging Then
                DebugAndLogging GetMsg("Debug.CodeExecution.ExectutionTimeStats", Format$(startTime, "hh:mm:ss"), Format$(endTime, "hh:mm:ss"), Format$(elapsedTime * 86400, "0.00")), , True
            End If
    End Select
End Sub

Public Sub SetDefaultSheetVisibility()
    Dim ws As Worksheet

    For Each ws In ThisWorkbook.Worksheets
        With ws
            Select Case .CodeName
                Case "MacOS_Users"
                    #If Mac Then
                        If .Visible <> xlSheetVisible Then .Visible = xlSheetVisible
                    #Else
                        If .Visible <> xlSheetHidden Then .Visible = xlSheetHidden
                    #End If
                Case "Class_"
                    If .Visible <> xlSheetVeryHidden Then .Visible = xlSheetVeryHidden
                Case Else
                    If .Visible <> xlSheetVisible Then .Visible = xlSheetVisible
            End Select
        End With
    Next ws
End Sub

Public Sub ReadOptionsValues()
    Dim optionsValues As Variant: optionsValues = Options.Range(g_OPTIONS_SETTINGS).Value
    
    With g_UserOptions
        .DisplayEntryTips = (optionsValues(1, 1) = "Yes")
        .OpenSavePathWhenDone = (optionsValues(2, 1) = "Yes")
        .EnableLogging = (optionsValues(3, 1) = "Yes")
        .DisplayInitialWarning = (optionsValues(4, 1) = "No")
        .AllFontsAreInstalled = (optionsValues(5, 1) = "Yes")
        .ZipSupportEnabled = (optionsValues(7, 1) = "Yes")
        .ValidFileHashes = (optionsValues(8, 1) = "Yes")
    End With
End Sub

Public Sub ToggleApplicationFeatures(ByVal enabledStatus As String, Optional ByVal cleanLogFolder As Boolean = False)
    With Application
        Select Case enabledStatus
            Case "Enable"
                .Calculation = xlCalculationAutomatic
                .EnableAnimations = True
                .EnableEvents = True
                .ScreenUpdating = True
            Case "Disable"
                .Calculation = xlCalculationManual
                .EnableAnimations = False
                .EnableEvents = False
                .ScreenUpdating = False
        End Select
    End With
End Sub

Public Sub ToggleSheetProtection(ByRef ws As Worksheet, ByVal protectionStatus As Boolean)
    With ws
        If protectionStatus Then
            .Protect
            If .EnableSelection <> xlUnlockedCells Then .EnableSelection = xlUnlockedCells
        Else
            .Unprotect
        End If
    End With
End Sub

Public Sub ValidateOptionsValues(Optional ByVal startupCheck As Boolean = False)
    Dim optionsValues As Variant: optionsValues = Options.Range(g_OPTIONS_SETTINGS)
    
    If startupCheck Then
        g_UserOptions.AllFontsAreInstalled = VerifyFontInstallation(startupCheck)
        g_UserOptions.ZipSupportEnabled = CheckFor7Zip(startupCheck)
        
        ToggleSheetProtection Options, False
        
        With Options
            WriteNewRangeValue .Range(g_FONT_INSTALLATION_STATUS), IIf(g_UserOptions.AllFontsAreInstalled, "Yes", "No")
            WriteNewRangeValue .Range(g_7ZIP_SUPPORT_STATUS), IIf(g_UserOptions.ZipSupportEnabled, "Yes", "No")
        End With
        
        ToggleSheetProtection Options, True
    Else
        With g_UserOptions
            .DisplayEntryTips = (optionsValues(1, 1) = "Yes")
            .OpenSavePathWhenDone = (optionsValues(2, 1) = "Yes")
            .EnableLogging = (optionsValues(3, 1) = "Yes")
            .DisplayInitialWarning = (optionsValues(4, 1) = "No")
            .AllFontsAreInstalled = (optionsValues(5, 1) = "Yes")
            .ZipSupportEnabled = (optionsValues(7, 1) = "Yes")
            .ValidFileHashes = (optionsValues(8, 1) = "Yes")
        End With
    End If
End Sub

Public Function VerifyKeySheetsExist() As Boolean
    Dim ws As Worksheet
    Dim wsCodeName As String
    Dim keyWorksheetCount As Long
    
    For Each ws In ThisWorkbook.Worksheets
        wsCodeName = ws.CodeName
        Select Case wsCodeName
            Case "Instructions"
                keyWorksheetCount = keyWorksheetCount + 1
            Case "Options"
                keyWorksheetCount = keyWorksheetCount + 1
            Case "MacOS_Users"
                keyWorksheetCount = keyWorksheetCount + 1
            Case "Class_"
                keyWorksheetCount = keyWorksheetCount + 1
        End Select
    Next ws
    
    VerifyKeySheetsExist = (keyWorksheetCount = 4)
End Function

Private Sub CreateNewClassRecordsSheet()
    Const BASE_NAME  As String = "New Class"
    
    Dim newSheetName As String
    Dim sheetSuffix  As Long
    Dim wbSheetCount As Long
    
    newSheetName = BASE_NAME
    sheetSuffix = 1
    wbSheetCount = ThisWorkbook.Sheets.Count
    
    Do While SheetExists(newSheetName)
        sheetSuffix = sheetSuffix + 1
        newSheetName = BASE_NAME & " (" & sheetSuffix & ")"
    Loop
    
    With Class_
        .Visible = xlSheetVisible
        .Copy After:=ThisWorkbook.Sheets(wbSheetCount)
        .Visible = xlSheetVeryHidden
    End With
    
    With ThisWorkbook.Sheets(wbSheetCount + 1)
        .Unprotect
        .Name = newSheetName
        .Protect
    End With
End Sub