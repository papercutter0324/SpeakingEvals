(*
Helper Scripts for the DYB Speaking Evaluations Excel spreadsheet

Version: 1.0.0
Build:   20250113
Warren Feltmate
© 2025
*)

-- Requirements and Dependencies
use AppleScript version "2.4" -- Yosemite (10.10) or later
use scripting additions
use script "Dialog Toolkit Plus" version "1.1.3"

-- Environment Variables

on GetScriptVersionNumber(paramString)
    -- Use build number to determine if an update is available
    return 20250113
end GetScriptVersionNumber

on GetMacOSVersion(paramString)
    -- Not currently used, but could be helpful if there are issues with older versions of MacOS
    try
        set osVersion to do shell script "sw_vers -productVersion"
        return osVersion
    end try
end GetMacOSVersion

-- Parameter Manipulation

on SplitString(passedParamString, parameterSeparator)
    -- Excel can only pass on parameter to this file. This makes it possible to split one into many.
    tell AppleScript
        set oldTextItemsDelimiters to text item delimiters
        set text item delimiters to parameterSeparator
        set separatedParameters to text items of passedParamString
        set text item delimiters to oldTextItemsDelimiters
    end tell
    return separatedParameters
end SplitString

-- Application Manipulations

on LoadApplication(appName)
    -- A simple function to tell the needed program to open.
    try
        tell application appName to activate
        return ""
    on error errMsg number errNum
        return "Error loading " & appName & ": " & errNum & " - " & errMsg
    end try
end LoadApplication

on IsAppLoaded(appName)
    -- This lets Excel check that the other program is open before continuing.
    try
        tell application "System Events"
            if (name of every process) contains appName then
                set loadResult to appName & " is now running."
            else
                set loadResult to "Error opening " & appName
            end if
        end tell
        return loadResult
    on error errMsg number errNum
        return "Error loading " & appName & ": " & errNum & " - " & errMsg
    end try
end IsAppLoaded

on CloseWord(paramString)
    -- This will completely close MS Word, even from the Dock. This reduces the chances of errors on subsequent runs.
    try
        tell application "System Events"
            if (name of every process) contains "Microsoft Word" then
                tell application "Microsoft Word" to quit
                set closeResult to "Word has successfully been closed."
            else
                set closeResult to "Word is not currently running."
            end if
            return closeResult
        end tell
    on error
        return "There was an error trying to close Word."
    end try
end CloseWord

-- File Manipulation

on CompareMD5Hashes(paramString)
    -- This will check the file integrity of the downloaded template against the known good value.
    set {filePath, validHash} to SplitString(paramString, "-,-")

    if not DoesFileExist(filePath) then
        return false
    end if

    try
        set checkResult to (do shell script "md5 -q " & quoted form of filePath)
        return checkResult is validHash
    on error
        return false
    end try
end CompareMD5Hashes

on CopyFile(paramString)
    -- Self-explanatory. Copy file from place A to place B. The original file will still exist.
    set {tempTemplatePath, finalTemplatePath} to SplitString(paramString, "-,-")
    try
        do shell script "cp " & (quoted form of tempTemplatePath) & " " & (quoted form of finalTemplatePath)
        return true
    on error
        return false
    end try
end CopyFile

on CreateZipFile(paramString)
    -- Create a ZIP file of all the PDFs in the target folder. Makes it simpler for you to send them to your KTs.
    set {savePath, zipPath} to SplitString(paramString, "-,-")
    try
        do shell script "cd " & quoted form of savePath & " && /usr/bin/zip -j " & quoted form of zipPath & " *.pdf"
        return "Success"
    on error
        return errMsg
    end try
end CreateZipFile

on DeleteFile(filePath)
    --Self-explanatory. This will delete the target file, skipping the Trash.
    (* The value of filePath passed to this function is always carefully considered
    (and limited), but at a future point, I will likely add in some safety checks for extra security
    to prevent a dangerous value accidentally being sent to this function.
    *)
    try
        do shell script "rm -f " & (quoted form of filePath)
        return true
    on error
    return false
    end try
end DeleteFile

on DoesBundleExist(bundlePath)
    -- Used to check if the Dialog Toolkit Plus script bundle exists
    tell application "System Events" to return (exists disk item bundlePath)
end DoesBundleExist

on DoesFileExist(filePath)
    -- Self-explanatory
    tell application "System Events" to return (exists disk item filePath) and class of disk item filePath = file
end DoesFileExist

on DownloadFile(paramString)
    -- Self-explanatory. The value of fileURL is the internet address to the desired file.
    set {destinationPath, fileURL} to SplitString(paramString, "-,-")
    try
        do shell script "curl -L -o " & (quoted form of destinationPath) & " " & (quoted form of fileURL)
        return true
    on error
        display dialog "Error downloading file: " & fileURL
        return false
    end try
end DownloadFile

on FindSignature(signaturePath)
    -- If your signature isn't embedded in the Excel file, it will try to find an external JPG or PNG version
    try
        if DoesFileExist(signaturePath & "mySignature.png") then
            return signaturePath & "mySignature.png"
        else if DoesFileExist(signaturePath & "mySignature.jpg") then
            return signaturePath & "mySignature.png"
        else
            return ""
        end if
    on error
        return ""
    end try
end FindSignature

on RenameFile(paramString)
    -- This pulls double duty for renaming a file or moving it to a new location. (It's the same process to the computer.)
    set {targetFile, newFilename} to SplitString(paramString, "-,-")
    set targetFile to quoted form of POSIX path of targetFile
    set newFilename to quoted form of POSIX path of newFilename
    try
        do shell script "mv " & targetFile & space & newFilename
        return true
    on error
        return false
    end try
end RenameFile

-- Folder Manipulation

on ClearFolder(folderToEmpty)
    -- Empties the target folder, but only of PDF and ZIP files. This folder will not be deleted.
    try
        do shell script "find " & (quoted form of folderToEmpty) & " -type f -name '*.pdf' -delete"
        do shell script "find " & (quoted form of folderToEmpty) & " -type f -name '*.zip' -delete"
        -- It then checks for a Proofs folder and clears it of DOCX files.
        set folderToEmpty to folderToEmpty & "Proofs/"
        if DoesFolderExist(folderToEmpty) then
            do shell script "find " & (quoted form of folderToEmpty) & " -type f -name '*.docx' -delete"
            set folderContents to list folder folderToEmpty without invisibles
            -- If found and empty, it then deletes the Proofs folder
            if (count of folderContents) is 0 then DeleteFolder(folderToEmpty)
        end if
        return true
    on error
        return false
    end try
end ClearFolder

on CreateFolder(folderPath)
    -- Self-explanatory. Needed for creating the folder for where the reports will be saved.
    try
        do shell script "mkdir -p " & (quoted form of folderPath)
        return true
    on error
        return false
    end try
end CreateFolder

on DeleteFolder(folderPath)
    -- Self-explanatory. Same as with DeleteFile, extra security checks will likely be added later.
    try
        do shell script "rm -rf " & (quoted form of folderPath)
        return true
    on error
        return false
    end try
end DeleteFolder

on DoesFolderExist(folderPath)
    -- Self-explanatory
    tell application "System Events" to return (exists disk item folderPath) and class of disk item folderPath = folder
end DoesFolderExist

-- Dialog Boxes

on DisplayDialog(messageString)
    -- This will display the nicer looking messages. A great improvement over the default version.
    set {dialogMessage, dialogType, dialogTitle, accViewWidth} to SplitString(messageString, "-,-")

    -- Select button type
    set defaultButton to 1 -- Set default unless overridden below
    if dialogType is "OkOnly" then
        set displayedButtons to {"OK"}
        set buttonKeys to {"", "1", ""}
    else if dialogType is "OkCancel" then
        set displayedButtons to {"Cancel", "OK"}
        set buttonKeys to {"", "2", "1", ""}
    else if dialogType is "YesNo" then
        set displayedButtons to {"No", "Yes"}
        set buttonKeys to {"", "2", "1", ""}
    else if dialogType is "RetryCancel" then
        set displayedButtons to {"Cancel", "Retry"}
        set buttonKeys to {"", "2", "1", ""}
    else if dialogType is "YesNoCancel" then
        set displayedButtons to {"Cancel", "No", "Yes"}
        set buttonKeys to {"", "3", "2", "1", ""}
    else
        set displayedButtons to {"OK"}
        set buttonKeys to {"", "1", ""}
    end if
	
    -- Create a Dialog Toolkit dialog window
    -- set accViewWidth to 250 -- This is passed in by the VBA code but kept here in case of future updates.
    set theTop to 10
    set {theButtons, minWidth} to create buttons displayedButtons button keys buttonKeys default button defaultButton
    if minWidth > accViewWidth then set accViewWidth to minWidth

    -- Create label for the message
    set {messageLabel, theTop} to create label dialogMessage bottom 0 max width accViewWidth control size regular size

    -- Display the dialog window
    set {buttonName, controlsResults} to display enhanced window dialogTitle ¬
        acc view width accViewWidth ¬
        acc view height theTop ¬
        acc view controls {messageLabel} ¬
        buttons theButtons
	
    -- Set return values
    if buttonName is "Ok" then
        return 1
    else if buttonName is "Cancel" then
        return 2
    else if buttonName is "Abort" then
        return 3
    else if buttonName is "Retry" then
        return 4
    else if buttonName is "Ignore" then
        return 5
    else if buttonName is "Yes" then
        return 6
    else if buttonName is "No" then
        return 7
    end if
end DisplayDialog

on InstallDialogToolkitPlus(paramString)
    set scriptLibrariesFolder to POSIX path of (path to home folder) & "Library/Script Libraries"
    set dialogToolkitPlusBundle to scriptLibrariesFolder & "/Dialog Toolkit Plus.scptd"
    set downloadDestination to POSIX path of (path to downloads folder)
    set zipFilePath to downloadDestination & "Dialog_Toolkit.zip"
    set zipExtractionPath to downloadDestination & "dialogToolkitTemp"
    set downloadURL to "https://raw.githubusercontent.com/papercutter0324/SpeakingEvals/main/Dialog_Toolkit.zip"

    -- Check if Dialog Toolkit is already installed
    if DoesBundleExist(dialogToolkitPlusBundle) then
        return true
        -- If not installed, ensure the required folder exists
    else if not DoesFolderExist(scriptLibrariesFolder) then
        try
            -- ~/Library is typically a read-only folder, so I need to requst your password to create the need folder
            do shell script "mkdir -p " & quoted form of scriptLibrariesFolder with administrator privileges
        on error
            -- If the folder cannot be created, tell the VBA script to use the default MsgBox command
            return false
        end try
    end if

    -- Ensure old versions of the file are not present in the Downloads folder
    if DoesFileExist(zipFilePath) then
        DeleteFile(zipFilePath)
    end if
	
    if DoesFolderExist(zipExtractionPath) then
        DeleteFolder(zipExtractionPath)
    end if

    -- Download, extract, and copy the script bundle to the required folder
    if DownloadFile(zipFilePath & "-,-" & downloadURL) then
        try
            do shell script "unzip -o " & quoted form of zipFilePath & " -d " & quoted form of (downloadDestination & "dialogToolkitTemp")
            RenameFile(zipExtractionPath & "/Dialog_Toolkit/Dialog Toolkit Plus.scptd" & "-,-" & dialogToolkitPlusBundle)
        end try
    end if

    -- Remove unneeded files and folders created during this process
    try
        DeleteFile(zipFilePath)
        DeleteFolder(zipExtractionPath)
    end try

    -- One final check to verify installation was successful and return true if it was
    return DoesBundleExist(dialogToolkitPlusBundle)
end InstallDialogToolkitPlus
