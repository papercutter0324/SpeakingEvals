(*
Helper Scripts for the DYB Speaking Evaluations Excel spreadsheet

Version: 1.9.1
Build:   20260120
Warren Feltmate
© 2025
*)

-- Environment Variables

on GetScriptVersionNumber(paramString)
	--- Use build number to determine if an update is available
	return 20260120
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
		set oldTextItemsDelimiters to AppleScript's text item delimiters
		set AppleScript's text item delimiters to parameterSeparator
		set separatedParameters to text items of passedParamString
		set AppleScript's text item delimiters to oldTextItemsDelimiters
	end tell
	return separatedParameters
end SplitString

on JoinString(passedParamArray, parameterSeparator)
	tell AppleScript
		set oldTextItemsDelimiters to AppleScript's text item delimiters
		set AppleScript's text item delimiters to parameterSeparator
		set joinedParameters to passedParamArray as string
		set AppleScript's text item delimiters to oldTextItemsDelimiters
	end tell
	return joinedParameters
end JoinString

-- Application Manipulations

on LoadApplication(appName)
	-- A simple function to tell the needed program to open.
	try
		tell application appName to activate
		return ""
	on error errMsg number errNum
		return "Error loading" & space & appName & ": " & errNum & " - " & errMsg
	end try
end LoadApplication

on IsAppLoaded(appName)
	-- This lets Excel check that the other program is open before continuing.
	try
		tell application "System Events"
			if (name of every process) contains appName then
				set loadResult to appName & space & "is now running."
			else
				set loadResult to "Error opening" & space & appName
			end if
		end tell
		return loadResult
	on error errMsg number errNum
		return "Error loading " & appName & ": " & errNum & " - " & errMsg
	end try
end IsAppLoaded

on ClosePowerPoint(paramString)
	-- This will completely close MS PowerPoint, even from the Dock. This reduces the chances of errors on subsequent runs.
	try
		tell application "System Events"
			if (name of every process) contains "Microsoft PowerPoint" then
				tell application "Microsoft PowerPoint" to quit
				set closeResult to "PowerPoint has successfully been closed."
			else
				set closeResult to "PowerPoint is not currently running."
			end if
			return closeResult
		end tell
	on error
		return "There was an error trying to close PowerPoint."
	end try
end ClosePowerPoint

-- File Manipulation

on ChangeFilePermissions(paramString)
	set {newPermissions, filePath} to SplitString(paramString, "-,-")
	
	try
		do shell script "xattr -d com.apple.quarantine " & quoted form of filePath
	end try
	
	try
		do shell script "chmod " & newPermissions & space & quoted form of filePath
		return true
	on error
		return false
	end try
end ChangeFilePermissions

on CompareMD5Hashes(paramString)
	-- This will check the file integrity of the downloaded template against the known good value.
	set {filePath, validHash} to SplitString(paramString, "-,-")
	
	if not DoesFileExist(filePath) then
		return false
	end if
	
	try
		set checkResult to (do shell script "md5 -q" & space & quoted form of filePath)
		return checkResult is validHash
	on error
		return false
	end try
end CompareMD5Hashes

on CopyFile(filePaths)
	-- Self-explanatory. Copy file from place A to place B. The original file will still exist.
	set {targetFile, destinationFile} to SplitString(filePaths, "-,-")
	try
		do shell script "cp" & space & (quoted form of targetFile) & space & (quoted form of destinationFile)
		return true
	on error
		return false
	end try
end CopyFile

on CreateTextFile(paramString)
	set {folderPath, fileName} to SplitString(paramString, "-,-")
	
	try
		do shell script "touch " & quoted form of (folderPath & fileName)
		return true
	on error
		return false
	end try
end CreateTextFile

on CreateZipWithLocal7Zip(zipCommand)
	try
		do shell script zipCommand
		return "Success"
	on error
		return errMsg
	end try
end CreateZipWithLocal7Zip

on CreateZipWithDefaultArchiver(paramString)
	-- Create a ZIP file of all the PDFs in the target folder. Makes it simpler for you to send them to your KTs.
	set {savePath, zipPath} to SplitString(paramString, "-,-")
	try
		do shell script "cd" & space & quoted form of savePath & " && /usr/bin/zip -j " & quoted form of zipPath & space & "*.pdf"
		return "Success"
	on error errMsg
		return errMsg
	end try
end CreateZipWithDefaultArchiver

on DeleteFile(filePath)
	try
		do shell script "rm -f " & quoted form of filePath & " && test ! -e " & quoted form of filePath
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
		do shell script "curl -fL -o " & (quoted form of destinationPath) & " " & (quoted form of fileURL)
		return true
	on error
		return false
	end try
end DownloadFile

on FileSelectDialog(paramString)
	set theFile to choose file with prompt "Select the workbook to import from"
	return POSIX path of theFile
end FileSelectDialog

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

on GetMD5Hash(filePath)
	try
		return do shell script "md5 -q " & quoted form of filePath
	on error errMsg
		return errMsg
	end try
end GetMD5Hash

on InstallFonts(paramString)
	set {fontName, fontURL} to SplitString(paramString, "-,-")
	set userFontPath to POSIX path of (path to home folder) & "Library/Fonts/" & fontName
	
	-- Check if the font is already installed in user or system-wide font directories
	if IsFontInstalled(fontName) then
		return true
	end if
	
	-- If not, download a copy to the fonts folder
	return DownloadFile(userFontPath & "-,-" & fontURL)
end InstallFonts

on IsFileEmpty(filePath)
	try
		do shell script "test -s " & quoted form of filePath
		-- file exists AND has size > 0
		return false
	on error
		-- file is missing OR size == 0
		return true
	end try
end IsFileEmpty

on IsFontInstalled(fontName)
	set userFontPath to POSIX path of (path to home folder) & "Library/Fonts/" & fontName
	set systemFontPath to "/Library/Fonts/" & fontName
	
	if DoesFileExist(userFontPath) or DoesFileExist(systemFontPath) then
		return true
	else
		return false
	end if
end IsFontInstalled

on RenameFile(paramString)
	-- This pulls double duty for renaming a file or moving it to a new location. (It's the same process to the computer.)
	set {targetFile, newFilename} to SplitString(paramString, "-,-")
	try
		do shell script "mv -f " & quoted form of targetFile & " " & quoted form of newFilename & " && test -f " & quoted form of newFilename
		return true
	on error
		return false
	end try
end RenameFile

on SavePptAsPdf(tempSavePath)
	try
		tell application "Microsoft PowerPoint"
			set thisDocument to active presentation
			save thisDocument in (POSIX file tempSavePath) as save as PDF
		end tell
		return true
	on error
		return false
	end try
end SavePptAsPdf

on TriggerPermission(paramString)
	try
		tell application "System Events" to get name of every process
		return true
	on error
		return false
	end try
end TriggerPermission

on WriteToLog(paramString)
	set {logPath, logData} to SplitString(paramString, "-,-")
	
	try
		do shell script "printf '%s\\n' " & quoted form of logData & " | iconv -t UTF-8 >> " & quoted form of logPath
		return true
	on error errMsg number errNum
		display dialog ("WriteToLog shell error: " & errMsg & " (" & errNum & ")")
		return false
	end try
end WriteToLog

-- Folder Manipulation

on ClearFolder(folderToEmpty)
	-- Empties the target folder, but only of DOCX, PDF, and ZIP files. This folder will not be deleted.
	try
		do shell script "find" & space & (quoted form of folderToEmpty) & space & "-type f -name '*.pdf' -delete"
		do shell script "find" & space & (quoted form of folderToEmpty) & space & "-type f -name '*.zip' -delete"
		do shell script "find" & space & (quoted form of folderToEmpty) & space & "-type f -name '*.pptx' -delete"
		return true
	on error
		return false
	end try
end ClearFolder

on ClearPDFsAfterZipping(folderToEmpty)
	try
		do shell script "find" & space & (quoted form of folderToEmpty) & space & "-type f -name '*.pdf' -delete"
		return true
	on error
		return false
	end try
end ClearPDFsAfterZipping

on CopyFolder(folderPath)
	-- Self-explanatory. Copy a folder (or bundle) from place A to place B. The original file will still exist.
	set {targetFolder, destinationFolder} to SplitString(folderPath, "-,-")
	try
		do shell script "cp -Rf" & space & (quoted form of targetFolder) & space & (quoted form of destinationFolder)
		return true
	on error
		return false
	end try
end CopyFolder

on CreateFolder(folderPath)
	try
		do shell script "mkdir -p " & quoted form of folderPath
		tell application "System Events"
			return exists folder folderPath
		end tell
	on error
		return false
	end try
end CreateFolder

on DeleteFolder(folderPath)
	-- Self-explanatory. Same as with DeleteFile, extra security checks will likely be added later.
	try
		do shell script "rm -rf" & space & (quoted form of folderPath)
		return true
	on error
		return false
	end try
end DeleteFolder

on DoesFolderExist(folderPath)
	try
		do shell script "test -d " & quoted form of folderPath
		return true
	on error
		return false
	end try
end DoesFolderExist

on ListFolderContents(paramString)
	set {folderPath, fileExtension} to SplitString(paramString, "-,-")
	
	tell application "System Events"
		try
			set fileList to name of files of folder folderPath whose name extension is fileExtension
			
			if fileList is {} then
				return ""
			end if
			
			set oldTextItemsDelimiters to AppleScript's text item delimiters
			set AppleScript's text item delimiters to "-,-"
			
			set joinedFileList to fileList as string
			set AppleScript's text item delimiters to oldTextItemsDelimiters
			
			return joinedFileList
		on error errMsg
			return "Error: " & errMsg
		end try
	end tell
end ListFolderContents

on OpenFolder(folderPath)
	try
		set pathAlias to POSIX file folderPath as alias
		tell application "Finder"
			open pathAlias
			return true
		end tell
	on error
		return false
	end try
end OpenFolder

-- Dialog Boxes

on InstallDialogDisplayScript(paramString)
	set scriptPath to POSIX path of (path to home folder) & "Library/Application Scripts/com.microsoft.Excel/DialogDisplay.scpt"
	set downloadURL to "https://raw.githubusercontent.com/papercutter0324/SpeakingEvals/main/AppleScript/DialogDisplay.scpt"
	
	return DownloadFile(scriptPath & "-,-" & downloadURL)
end InstallDialogDisplayScript

on CheckForScriptLibrariesFolder(paramString)
	set scriptLibrariesFolder to (POSIX path of (path to home folder)) & "Library/Script Libraries"
	
	if DoesFolderExist(scriptLibrariesFolder) then
		return scriptLibrariesFolder
	else
		try
			do shell script "mkdir -p " & quoted form of scriptLibrariesFolder -- with administrator privileges
			-- do shell script "chown " & quoted form of (short user name of (system info)) & space & quoted form of scriptLibrariesFolder with administrator privileges
			do shell script "chmod u+rwx " & quoted form of scriptLibrariesFolder -- with administrator privileges
			return scriptLibrariesFolder
		on error
			return ""
		end try
	end if
end CheckForScriptLibrariesFolder

on InstallDialogToolkitPlus(resourcesFolder)
	set downloadURL to "https://raw.githubusercontent.com/papercutter0324/SpeakingEvals/main/AppleScript/Dialog_Toolkit.zip"
	set scriptLibrariesFolder to POSIX path of (path to home folder) & "Library/Script Libraries"
	set dialogBundleName to "/Dialog Toolkit Plus.scptd"
	set dialogToolkitPlusBundle to scriptLibrariesFolder & dialogBundleName
	set zipFilePath to resourcesFolder & "/Dialog_Toolkit.zip"
	set zipExtractionPath to resourcesFolder & "/dialogToolkitTemp"
	
	if DoesBundleExist(dialogToolkitPlusBundle) then return true
	
	if not DoesFolderExist(resourcesFolder) then
		try
			CreateFolder(resourcesFolder)
		on error
			return false
		end try
	end if
	
	if DoesBundleExist(resourcesFolder & dialogBundleName) then
		if CopyFolder(resourcesFolder & dialogBundleName & "-,-" & dialogToolkitPlusBundle) then
			return true
		end if
	end if
	
	if DownloadFile(zipFilePath & "-,-" & downloadURL) then
		try
			do shell script "unzip -o " & (quoted form of zipFilePath) & " -d " & (quoted form of zipExtractionPath)
			if not CopyFolder(zipExtractionPath & "/Dialog_Toolkit" & dialogBundleName & "-,-" & resourcesFolder & dialogBundleName) then error "Copy failed"
			if not CopyFolder(zipExtractionPath & "/Dialog_Toolkit" & dialogBundleName & "-,-" & dialogToolkitPlusBundle) then error "Copy failed"
		on error
			return false
		end try
	end if
	
	DeleteFile(zipFilePath)
	DeleteFolder(zipExtractionPath)
	
	return DoesBundleExist(dialogToolkitPlusBundle)
end InstallDialogToolkitPlus

on UninstallDialogToolkitPlus(resourcesFolder)
	set dialogToolkitPlusBundle to POSIX path of (path to home folder) & "Library/Script Libraries/Dialog Toolkit Plus.scptd"
	set localCopy to resourcesFolder & "/Dialog Toolkit Plus.scptd"
	
	if DoesBundleExist(dialogToolkitPlusBundle) then
		try
			if not DoesBundleExist(localCopy) then CopyFolder(dialogToolkitPlusBundle & "-,-" & localCopy)
			DeleteFolder(dialogToolkitPlusBundle)
			set removalResult to true
		on error
			set removalResult to false
		end try
	else
		set removalResult to true
	end if
	
	return removalResult
end UninstallDialogToolkitPlus

on VerifyFileIsAvailableLocally(filePath)
	set posixFile to POSIX file filePath
	
	try
		tell application "Finder"
			set theItem to item posixFile
			set logicalSize to size of theItem
			set physicalSize to physical size of theItem
		end tell
		
		if physicalSize > 0 then
			set checkResult to "Ok"
		else
			set checkResult to "Error: Local copy not available."
		end if
	on error errMsg
		set checkResult to "Error: " & errMsg
	end try
	
	return checkResult
end VerifyFileIsAvailableLocally
