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

on CopyFile(paramString)
	set {tempTemplatePath, finalTemplatePath} to SplitString(paramString, ",")
	try
		do shell script "cp " & (quoted form of tempTemplatePath) & " " & (quoted form of finalTemplatePath)
		return true
	on error
		return false
	end try
end CopyFile

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
	set {savePath, fileURL} to SplitString(paramString, ",")
	try
		do shell script "curl -L -o " & (quoted form of savePath) & " " & (quoted form of fileURL)
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

on ClearFolder(folderToEmpty)
	try
		do shell script "find " & (quoted form of folderToEmpty) & " -type f -name '*.pdf' -delete"
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

on CompareMD5Hashes(paramString)
	set {filePath, validHash} to SplitString(paramString, ",")
	
	if not DoesFileExist(filePath) then
		display dialog "File does not exist: " & filePath
		return false
	end if
	
	try
		if (do shell script "md5 -q " & quoted form of filePath) is validHash then
			return true
		else
			return false
		end if
	on error
		display dialog "Error generating hash of " & filePath
		return false
	end try
end CompareMD5Hashes

on SplitString(passedParamString, parameterSeparator)
	tell AppleScript
		set oldTextItemsDelimiters to text item delimiters
		set text item delimiters to parameterSeparator
		set separatedParameters to text items of passedParamString
		set text item delimiters to oldTextItemsDelimiters
	end tell
	return separatedParameters
end SplitString

on YesNoDialog(messageString)
	if button returned of (display dialog messageString buttons {"Yes", "No"} default button "No") is "Yes" then
		return 6
	else
		return 7
	end if
end YesNoDialog

on OkCancelDialog(messageString)
	if button returned of (display dialog messageString buttons {"Ok", "Cancel"} default button "Cancel") is "Ok" then
		return 1
	else
		return 2
	end if
end OkCancelDialog

on OKDialog(messageString)
	display dialog messageString buttons {"OK"} default button "OK"
	return 0
end OKDialog
