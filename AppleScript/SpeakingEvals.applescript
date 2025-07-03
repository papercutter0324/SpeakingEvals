(*
Helper Scripts for the DYB Speaking Evaluations Excel spreadsheet

Version: 1.5.0
Build:   20250704
Warren Feltmate
© 2025
*)

-- Environment Variables

on GetScriptVersionNumber(paramString)
	--- Use build number to determine if an update is available
	return 20250704
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
	
	-- Check if quarantine status is set; remove if necessary
	try
		set quarantineStatus to do shell script "xattr -p com.apple.quarantine" & space & quoted form of filePath
		if quarantineStatus is not "" then
			do shell script "xattr -d com.apple.quarantine" & space & quoted form of filePath
		end if
	end try
	
	-- Change file permissions
	try
		do shell script "chmod" & space & newPermissions & space & quoted form of filePath
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
	on error
		return errMsg
	end try
end CreateZipWithDefaultArchiver

on DeleteFile(filePath)
	--Self-explanatory. This will delete the target file, skipping the Trash.
	(* The value of filePath passed to this function is always carefully considered
	(and limited), but at a future point, I will likely add in some safety checks for extra security
	to prevent a dangerous value accidentally being sent to this function.
	*)
	try
		do shell script "rm -f" & space & (quoted form of filePath)
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

on InstallFonts(paramString)
	set {fontName, fontURL} to SplitString(paramString, "-,-")
	set userFontPath to POSIX path of (path to home folder) & "Library/Fonts/" & fontName
	set systemFontPath to "/Library/Fonts/" & fontName
	
	-- Check if the font is already installed in user or system-wide font directories
	if DoesFileExist(userFontPath) or DoesFileExist(systemFontPath) then
		return true
	end if
	
	-- If not, download a copy to the fonts folder
	return DownloadFile(userFontPath & "-,-" & fontURL)
end InstallFonts

on RenameFile(paramString)
	-- This pulls double duty for renaming a file or moving it to a new location. (It's the same process to the computer.)
	set {targetFile, newFilename} to SplitString(paramString, "-,-")
	set targetFile to quoted form of POSIX path of targetFile
	set newFilename to quoted form of POSIX path of newFilename
	try
		do shell script "mv -f" & space & targetFile & space & newFilename
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
	-- Self-explanatory. Needed for creating the folder for where the reports will be saved.
	try
		do shell script "mkdir -p" & space & (quoted form of folderPath)
		return true
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
	-- Self-explanatory
	tell application "System Events" to return (exists disk item folderPath) and class of disk item folderPath = folder
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
	set downloadURL to "https://raw.githubusercontent.com/papercutter0324/SpeakingEvals/main/DialogDisplay.scpt"
	
	-- If an existing version is not found, download a fresh copy
	-- Skip this first check until a full update function can be designed. For now, install each time
	-- if DoesFileExist(scriptPath) then return true
	return DownloadFile(scriptPath & "-,-" & downloadURL)
end InstallDialogDisplayScript

on CheckForScriptLibrariesFolder(paramString)
	set scriptLibrariesFolder to POSIX path of (path to home folder) & "Library/Script Libraries"
	
	if DoesFolderExist(scriptLibrariesFolder) then
		return scriptLibrariesFolder
	else
		try
			-- ~/Library is typically a read-only folder, so I need to requst your password to create the need folder
			do shell script "mkdir -p" & space & quoted form of scriptLibrariesFolder with administrator privileges
			-- Set your username as the owner
			do shell script "chown " & quoted form of (short user name of (system info)) & space & quoted form of scriptLibrariesFolder with administrator privileges
			-- Give your username READ and WRITE permissions.
			do shell script "chmod u+rw " & quoted form of scriptLibrariesFolder with administrator privileges
			return scriptLibrariesFolder
		on error
			return ""
		end try
	end if
end CheckForScriptLibrariesFolder

on InstallDialogToolkitPlus(resourcesFolder)
	set downloadURL to "https://raw.githubusercontent.com/papercutter0324/SpeakingEvals/main/Dialog_Toolkit.zip"
	set scriptLibrariesFolder to POSIX path of (path to home folder) & "Library/Script Libraries"
	set dialogBundleName to "/Dialog Toolkit Plus.scptd"
	set dialogToolkitPlusBundle to scriptLibrariesFolder & dialogBundleName
	set zipFilePath to resourcesFolder & "/Dialog_Toolkit.zip"
	set zipExtractionPath to resourcesFolder & "/dialogToolkitTemp"
	
	-- Initial check to see if already installed
	if DoesBundleExist(dialogToolkitPlusBundle) then return true
	
	-- Ensure resources folder exists for later use
	if not DoesFolderExist(resourcesFolder) then
		try
			CreateFolder(resourcesFolder)
		on error
			return false
		end try
	end if
	
	-- Check for a local copy and move it to the needed folder if found
	if DoesBundleExist(resourcesFolder & dialogBundleName) then
		if CopyFolder(resourcesFolder & dialogBundleName & "-,-" & dialogToolkitPlusBundle) then
			return true
		end if
	end if
	
	-- Otherwise, download and...
	if DownloadFile(zipFilePath & "-,-" & downloadURL) then
		try
			-- ...extract the files...
			do shell script "unzip -o" & space & (quoted form of zipFilePath) & " -d " & (quoted form of zipExtractionPath)
			-- ...keep a local copy in the resources folder...
			CopyFolder(zipExtractionPath & "/Dialog_Toolkit" & dialogBundleName & "-,-" & resourcesFolder & dialogBundleName)
			-- ...and copy the script bundle to the required folder
			CopyFolder(zipExtractionPath & "/Dialog_Toolkit" & dialogBundleName & "-,-" & dialogToolkitPlusBundle)
		end try
	end if
	
	-- Remove unneeded files and folders created during this process
	DeleteFile(zipFilePath)
	DeleteFolder(zipExtractionPath)
	
	-- One final check to verify installation was successful and return true if it was
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
