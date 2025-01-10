(*
Helper Scripts for the DYB Speaking Evaluations Excel spreadsheet

Version: 20250107
Warren Feltmate
© 2025
*)

-- Dialog Toolkit Plus.scptd should be in ~/Library/Script Libraries
use AppleScript version "2.4" -- Yosemite (10.10) or later
use scripting additions
use script "Dialog Toolkit Plus" version "1.1.3"
-- Add code to install this automatically

-- Environment Variables

on GetScriptVersionNumber(paramString)
	return 20250107
end GetScriptVersionNumber

-- Parameter Manipulation

on SplitString(passedParamString, parameterSeparator)
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
	try
		tell application appName to activate
		return ""
	on error errMsg number errNum
		return "Error loading " & appName & ": " & errNum & " - " & errMsg
	end try
end LoadApplication

on IsAppLoaded(appName)
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
	set {filePath, validHash} to SplitString(paramString, ",")
	
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
	set {tempTemplatePath, finalTemplatePath} to SplitString(paramString, ",")
	try
		do shell script "cp " & (quoted form of tempTemplatePath) & " " & (quoted form of finalTemplatePath)
		return true
	on error
		return false
	end try
end CopyFile

on CreateZipFile(paramString)
	set {savePath, zipPath} to SplitString(paramString, ",")
	try
		do shell script "cd " & quoted form of savePath & " && /usr/bin/zip -j " & quoted form of zipPath & " *.pdf"
		return "Success"
	on error
		return errMsg
	end try
end CreateZipFile

on DeleteFile(filePath)
	try
		do shell script "rm -f " & (quoted form of filePath)
		return true
	on error
		return false
	end try
end DeleteFile

on DoesFileExist(filePath)
	tell application "System Events" to return (exists disk item filePath) and class of disk item filePath = file
end DoesFileExist

on DownloadFile(paramString)
	set {destinationPath, fileURL} to SplitString(paramString, ",")
	try
		do shell script "curl -L -o " & (quoted form of destinationPath) & " " & (quoted form of fileURL)
		return true
	on error
		display dialog "Error downloading file: " & fileURL
		return false
	end try
end DownloadFile

on FindSignature(signaturePath)
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
	set {targetFile, newFilename} to SplitString(paramString, ",")
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
	try
		do shell script "find " & (quoted form of folderToEmpty) & " -type f -name '*.pdf' -delete"
		do shell script "find " & (quoted form of folderToEmpty) & " -type f -name '*.zip' -delete"
		set folderToEmpty to folderToEmpty & "Proofs/"
		if DoesFolderExist(folderToEmpty) then
			do shell script "find " & (quoted form of folderToEmpty) & " -type f -name '*.docx' -delete"
			set folderContents to list folder folderToEmpty without invisibles
			if (count of folderContents) is 0 then DeleteFolder(folderToEmpty)
		end if
		return true
	on error
		return false
	end try
end ClearFolder

on CreateFolder(folderPath)
	try
		do shell script "mkdir -p " & (quoted form of folderPath)
		return true
	on error
		return false
	end try
end CreateFolder

on DeleteFolder(folderPath)
	try
		do shell script "rm -rf " & (quoted form of folderPath)
		return true
	on error
		return false
	end try
end DeleteFolder

on DoesFolderExist(folderPath)
	tell application "System Events" to return (exists disk item folderPath) and class of disk item folderPath = folder
end DoesFolderExist

(*
The following are not yet used by the Speaking Evals spreadsheet
but are here in anticipation of future improvements.
*)

-- Dialog Boxes

DisplayDialog("Hello,Test,YesNo")

on DisplayDialog(messageString)
	set {dialogMessage, dialogTitle, dialogType} to SplitString(messageString, ",")
	
	-- Select button type
	if dialogType is "OkCancel" then
		set displayedButtons to {"Cancel", "OK"}
		set buttonKeys to {"", "2", "1", ""}
		set defaultButton to 2
	else if dialogType is "YesNo" then
		set displayedButtons to {"No", "Yes"}
		set buttonKeys to {"", "2", "1", ""}
		set defaultButton to 2
	else if dialogType is "OkOnly" then
		set displayedButtons to {"OK"}
		set buttonKeys to {"", "1", ""}
		set defaultButton to 1
	end if
	
	-- Create a Dialog Toolkit dialog window
	set accViewWidth to 300
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
	else if buttonName is "Yes" then
		return 6
	else if buttonName is "No" then
		return 7
	end if
end DisplayDialog

-- Environment Variables

on GetMacOSVersion(paramString)
	try
		set osVersion to do shell script "sw_vers -productVersion"
		return osVersion
	end try
end GetMacOSVersion

on SetConfigDirectory(paramString)
	try
		set osVersion to GetMacOSVersion("")
		
		if osVersion starts with "10.1" or osVersion starts with "11." or osVersion starts with "12." then
			set configFolder to POSIX path of (path to documents folder) & "DYB"
		else if osVersion starts with "13." or osVersion starts with "14." or osVersion starts with "15." then
			set configFolder to POSIX path of (path to home folder) & "Documents/DYB"
		else
			return "unsupported"
		end if
		
		if not ExistsFolder(configFolder) then
			do shell script "mkdir -p " & quoted form of configFolder
		end if
		
		set configFolder to configFolder & "/AngryBirdsTrivia"
		if not ExistsFolder(configFolder) then
			do shell script "mkdir -p " & quoted form of configFolder
		end if
		
		return configFolder
	on error
		display dialog "Error: " & errMsg buttons {"OK"} default button "OK"
		return ""
	end try
end SetConfigDirectory

on SetTempDirectory(paramString)
	return POSIX path of (path to temporary items)
end SetTempDirectory
