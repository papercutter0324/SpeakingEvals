(*
Helper Scripts for the DYB Speaking Evaluations Excel spreadsheet

Version: 2.4.0
Build:   20260306
Warren Feltmate
© 2025
*)

-- Propertites
property DELIM : "-,-"

-- System Manipulation

on GetScriptVersionNumber(paramString)
	--- Use build number to determine if an update is available
	return 20260306
end GetScriptVersionNumber

on GetMacOSVersion(paramString)
	-- Not currently used, but could be helpful if there are issues with older versions of MacOS
	try
		set osVersion to do shell script "sw_vers -productVersion"
		return osVersion
	end try
end GetMacOSVersion

on IsScriptPresent(paramString)
	return true
end IsScriptPresent

on ListOfPermittedFolders()
	set homeFolder to NormalizeFolderPath(POSIX path of (path to home folder))
	set scriptLibrariesFolder to NormalizeFolderPath(homeFolder & "Library/Script Libraries")
	set excelScriptsFolder to NormalizeFolderPath(homeFolder & "Library/Application Scripts/com.microsoft.Excel")
	
	set pathList to {scriptLibrariesFolder, excelScriptsFolder}
	return pathList
end ListOfPermittedFolders

on GetListOfInstalledPrinters(paramString)
	set printerList to paragraphs of (do shell script "lpstat -p | awk '{print $2}'")
	
	if printerList is {} then return ""
	
	set joinedPrinterList to JoinString(printerList, DELIM)
	return joinedPrinterList
end GetListOfInstalledPrinters

on PrintUsingPDFPrintTool(paramString)
	set partsList to SplitString(paramString, DELIM)
	
	if (count of partsList) < 5 then
		return "Error: Incomplete paramString"
	end if
	
	set printToolPath to item -1 of partsList
	set printScaling to item -2 of partsList
	set paperSize to item -3 of partsList
	set printerName to item -4 of partsList
	set printQueue to JoinString(items 1 thru -5 of partsList, ",")
	
	if printScaling is "" then
		set printScaling to "fit"
	end if
	
	if paperSize is "" then
		set paperSize to "pdf"
	end if
	
	try
		do shell script "nohup" & space & ¬
			quoted form of printToolPath & space & ¬
			"-f " & quoted form of printQueue & space & ¬
			"-d " & quoted form of printerName & space & ¬
			"-s " & quoted form of printScaling & space & ¬
			"-p " & paperSize
		return "Success"
	on error errMsg
		return "Error: " & errMsg
	end try
end PrintUsingPDFPrintTool

-- Debuggin Helpers
on displayDebugMessage(msgToDisplay)
	display dialog msgToDisplay
end displayDebugMessage

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

on CloseApplication(appName)
	-- Not working as intended
	try
		tell application "System Events"
			if exists (process appName) then
				tell application appName to quit
				return appName & " has been told to quit."
			else
				return appName & " is not currently running."
			end if
		end tell
	on error errMsg number errNum
		return "Error closing " & appName & ": " & errNum & " - " & errMsg
	end try
end CloseApplication

on LoadApplication(appName)
	try
		tell application appName to activate
		return "Success"
	on error errMsg number errNum
		if errNum = -10814 then
			set rtnMsg to "Error: " & appName & " not found."
		else
			set rtnMsg to "Error loading " & appName & ": " & errNum & " - " & errMsg
		end if
		
		return rtnMsg
	end try
end LoadApplication

on IsAppLoaded(appName)
	-- This lets Excel check that the other program is open before continuing.
	try
		tell application "System Events"
			-- if (name of every process) contains appName then
			if exists (process appName) then
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

on ClosePowerPoint(paramString)
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
	on error errMsg number errNum
		return "Error closing Microsoft PowerPoint: " & errNum & " - " & errMsg
	end try
end ClosePowerPoint

on ClosePowerPointV2(paramString)
	try
		if application "Microsoft PowerPoint" is running then
			tell application "Microsoft PowerPoint"
				quit saving no
			end tell
			return "PowerPoint has successfully been closed."
		else
			return "PowerPoint is not currently running."
		end if
	on error errMsg number errNum
		return "Error closing Microsoft PowerPoint: " & errNum & " - " & errMsg
	end try
end ClosePowerPointV2

-- File Manipulation

on ChangeFilePermissions(paramString)
	set {newPermissions, filePath} to SplitString(paramString, DELIM)
	
	try
		do shell script "xattr -dr com.apple.quarantine " & quoted form of filePath
	end try
	
	try
		do shell script "chmod " & quoted form of newPermissions & space & quoted form of filePath
		return true
	on error
		return false
	end try
end ChangeFilePermissions

on CompareMD5Hashes(paramString)
	-- This will check the file integrity of the downloaded template against the known good value.
	set {filePath, validHash} to SplitString(paramString, DELIM)
	
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

on CopyBundle(sourcePath, destPath)
	try
		do shell script "rm -rf " & quoted form of destPath
		do shell script "ditto " & quoted form of sourcePath & " " & quoted form of destPath
		return true
	on error
		return false
	end try
end CopyBundle

on CopyFile(filePaths)
	set {targetFile, destinationFile} to SplitString(filePaths, DELIM)
	try
		do shell script "cp -f " & quoted form of targetFile & space & quoted form of destinationFile & " && test -f " & quoted form of destinationFile
		return true
	on error
		return false
	end try
end CopyFile

on CopyAllFiles(paramString)
	set partsList to SplitString(paramString, DELIM)
	
	if (count of partsList) < 3 then
		return false
	end if
	
	set sourceFolder to item 1 of partsList
	set destinationFolder to item 2 of partsList
	set fileExtension to item 3 of partsList
	
	if (count of partsList) ≥ 4 then
		set filesToSkip to items 4 thru -1 of partsList
	else
		set filesToSkip to {}
	end if
	
	set excludeRules to ""
	
	repeat with f in filesToSkip
		set excludeRules to excludeRules & " --exclude=" & quoted form of (contents of f)
	end repeat
	
	set cmd to "rsync -a" & space & ¬
		excludeRules & space & ¬
		"--include=" & quoted form of ("*." & fileExtension) & space & ¬
		"--exclude='*'" & space & ¬
		quoted form of sourceFolder & space & ¬
		quoted form of destinationFolder
	
	try
		do shell script "nohup" & space & cmd
		return true
	on error
		return false
	end try
end CopyAllFiles

on CreateTextFile(paramString)
	set {folderPath, fileName} to SplitString(paramString, DELIM)
	
	if fileName starts with "/" then
		set fileName to text 2 thru -1 of fileName
	end if
	
	try
		set folderPath to POSIX path of POSIX file folderPath
	on error
		return false
	end try
	
	if folderPath does not end with "/" then
		set folderPath to folderPath & "/"
	end if
	
	set filePath to folderPath & fileName
	
	try
		do shell script "touch " & quoted form of filePath
		return true
	on error
		return false
	end try
end CreateTextFile

on CreateZipWithLocal7Zip(zipCommand)
	try
		do shell script zipCommand
		return true
	on error
		return false
	end try
end CreateZipWithLocal7Zip

on CreateZipWithDefaultArchiver(paramString)
	-- Create a ZIP file of all the PDFs in the target folder. Makes it simpler for you to send them to your KTs.
	set {savePath, zipPath} to SplitString(paramString, DELIM)
	
	try
		set savePath to POSIX path of POSIX file savePath
	on error
		return "Error 1: Invalid savePath value."
	end try
	
	if savePath does not end with "/" then
		set savePath to savePath & "/"
	end if
	
	try
		set pdfCount to do shell script "ls " & quoted form of (savePath & "*.pdf") & " 2>/dev/null | wc -l"
	on error
		return "Error 2: Unable to read savePath"
	end try
	
	if pdfCount is "0" then
		return "Error 3: No PDF files found in savePath."
	end if
	
	try
		do shell script "/usr/bin/zip -j " & quoted form of zipPath & space & quoted form of (savePath & "*.pdf")
		return "Success"
	on error errMsg number errNum
		return "Error " & errNum & ": " & errMsg
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
	try
		do shell script "test -d " & quoted form of bundlePath
		return true
	on error
		return false
	end try
end DoesBundleExist

on DoesFileExist(filePath)
	try
		do shell script "test -f " & quoted form of filePath
		return true
	on error
		return false
	end try
end DoesFileExist

on DownloadFile(paramString)
	set {destinationPath, fileURL} to SplitString(paramString, DELIM)
	
	try
		do shell script "curl -fL --proto '=https' --create-dirs -o " & quoted form of destinationPath & space & quoted form of fileURL
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
			return signaturePath & "mySignature.jpg"
		else
			return "Error: mySignature not found."
		end if
	on error errMsg number errNum
		return "Error " & errNum & ": " & errMsg
	end try
end FindSignature

on GetMD5Hash(filePath)
	try
		return do shell script "md5 -q " & quoted form of filePath
	on error errMsg number errNum
		return "Error " & errNum & ": " & errMsg
	end try
end GetMD5Hash

on InstallFonts(paramString)
	set {fontName, fontURL} to SplitString(paramString, DELIM)
	set userFontPath to POSIX path of (path to home folder) & "Library/Fonts/" & fontName
	-- set systemFontPath to "/Library/Fonts/" & fontName
	
	-- Check if the font is already installed in user or system-wide font directories
	if IsFontInstalled(fontName) then
		return true
	end if
	
	-- If not, download a copy to the fonts folder
	if DownloadFile(userFontPath & DELIM & fontURL) then
		return DoesFileExist(userFontPath)
	else
		return false
	end if
end InstallFonts

on IsFileEmpty(filePath)
	try
		do shell script "test -s " & quoted form of filePath
		return false -- non-empty file
	on error
		return true -- empty or missing
	end try
end IsFileEmpty

on IsFontInstalled(fontName)
	set userFontPath to POSIX path of (path to home folder) & "Library/Fonts/" & fontName
	set systemFontPath to "/Library/Fonts/" & fontName
	
	return DoesFileExist(userFontPath) or DoesFileExist(systemFontPath)
end IsFontInstalled

on RenameFile(paramString)
	set {targetFile, newFilename} to SplitString(paramString, DELIM)
	try
		do shell script "mv -f " & quoted form of targetFile & space & quoted form of newFilename & " && test -e " & quoted form of newFilename
		return true
	on error
		return false
	end try
end RenameFile

on SavePptAsPdf(tempSavePath)
	try
		tell application "Microsoft PowerPoint"
			save active presentation in (POSIX file tempSavePath) as save as PDF
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
	set {logPath, logData} to SplitString(paramString, DELIM)
	
	try
		do shell script "printf '%s\\n' " & quoted form of logData & " >> " & quoted form of logPath
		return true
	on error errMsg number errNum
		display dialog ("WriteToLog shell error: " & errMsg & " (" & errNum & ")")
		return false
	end try
end WriteToLog

-- Folder Manipulation

on ClearFolder(paramString)
	set partsList to SplitString(paramString, DELIM)
	
	set folderToEmpty to item 1 of partsList
	
	set excludeClause to ""
	
	if (count of partsList) > 1 then
		set filesToKeep to items 2 thru -1 of partsList
		
		repeat with f in filesToKeep
			set excludeClause to excludeClause & " ! -name " & quoted form of (contents of f)
		end repeat
	end if
	
	set shellCommand to "find " & quoted form of folderToEmpty & ¬
		" -maxdepth 1 -type f \\( -iname '*.pdf' -o -iname '*.zip' -o -iname '*.pptx' \\)" & ¬
		excludeClause & " -delete"
	
	try
		do shell script shellCommand
		return true
	on error errMsg number errNum
		return false
	end try
end ClearFolder

on ClearPDFsAfterZipping(folderToEmpty)
	try
		-- do shell script "find" & space & (quoted form of folderToEmpty) & space & "-type f -name '*.pdf' -delete"
		do shell script "if [ -d " & quoted form of folderToEmpty & " ]; then " & ¬
			"find " & quoted form of folderToEmpty & " -type f -name '*.pdf' -exec rm -f {} +; fi"
		return true
	on error
		return false
	end try
end ClearPDFsAfterZipping

on CopyFolder(folderPath)
	-- Self-explanatory. Copy a folder (or bundle) from place A to place B. The original file will still exist.
	set {sourceFolder, destPath} to SplitString(folderPath, DELIM)
	try
		-- do shell script "cp -Rf "& quoted form of sourceFolder & space & quoted form of destPath
		-- Guard against nesting
		do shell script "test -e " & quoted form of destPath & " && rm -rf " & quoted form of destPath
		
		do shell script "cp -R " & quoted form of sourceFolder & space & quoted form of destPath
		return true
	on error
		return false
	end try
end CopyFolder

on CreateFolder(folderPath)
	try
		do shell script "mkdir -p " & quoted form of folderPath
		return true
	on error
		return false
	end try
end CreateFolder

on DeleteFolder(paramString)
	set {targetPath, resourcesFolder} to SplitString(paramString, DELIM)
	
	set targetPath to NormalizeFolderPath(targetPath)
	set resourcesFolder to NormalizeFolderPath(resourcesFolder)
	
	set permittedFolders to ListOfPermittedFolders()
	set end of permittedFolders to resourcesFolder
	
	if targetPath is "/" or targetPath is "" then
		return false
	end if
	
	repeat with baseFolder in permittedFolders
		set baseFolder to NormalizeFolderPath(contents of baseFolder)
		
		if targetPath starts with (contents of baseFolder) then
			try
				do shell script "rm -rf " & quoted form of targetPath
				return true
			on error
				return false
			end try
		end if
	end repeat
	
	return false
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
	set {folderPath, fileExtension} to SplitString(paramString, DELIM)
	
	if fileExtension starts with "." then
		set fileExtension to text 2 thru -1 of fileExtension
	end if
	
	try
		set cmd to "find " & quoted form of folderPath & ¬
			" -maxdepth 1 -type f -iname '*." & fileExtension & "' -print"
		
		set shellResult to do shell script cmd
		
		if shellResult is "" then
			return ""
		end if
		
		set fileList to paragraphs of shellResult
		
		set oldTID to AppleScript's text item delimiters
		set AppleScript's text item delimiters to DELIM
		set joinedFileList to fileList as string
		set AppleScript's text item delimiters to oldTID
		
		return joinedFileList
	on error errMsg
		return "Error: " & errMsg
	end try
end ListFolderContents

on ModifyFolderPermissions(paramString)
	set {targetFolder, desiredPermissions} to SplitString(paramString, DELIM)
	
	try
		do shell script "chmod " & quoted form of desiredPermissions & space & quoted form of targetFolder
		return true
	on error
		return false
	end try
end ModifyFolderPermissions

on NormalizeFolderPath(folderPath)
	if folderPath does not end with "/" then
		return folderPath & "/"
	end if
	
	return folderPath
end NormalizeFolderPath

on OpenFolder(folderPath)
	if folderPath ends with "/" then
		set folderPath to text 1 thru -2 of folderPath
	end if
	
	try
		tell application "Finder"
			repeat with w in windows
				try
					set windowPath to POSIX path of (target of w)
					if windowPath ends with "/" then
						set windowPath to text 1 thru -2 of windowPath
					end if
					
					if windowPath = folderPath then
						set index of w to 1
						activate
						return true
					end if
				end try
			end repeat
			
			set pathAlias to POSIX file folderPath as alias
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
	
	if DoesFileExist(scriptPath) then
		return true
	end if
	
	return DownloadFile(scriptPath & DELIM & downloadURL)
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

on IsDialogToolkitInstalled(resourcesFolder)
	set homeFolder to NormalizeFolderPath(POSIX path of (path to home folder))
	set scriptLibrariesFolder to NormalizeFolderPath(homeFolder & "Library/Script Libraries")
	set resourcesFolder to NormalizeFolderPath(resourcesFolder)
	
	set dialogBundleName to "Dialog Toolkit Plus.scptd"
	
	set dialogToolkitPlusBundle to scriptLibrariesFolder & dialogBundleName
	set resourcesBundle to resourcesFolder & dialogBundleName
	
	if DoesBundleExist(dialogToolkitPlusBundle) then
		return true
	else if DoesBundleExist(resourcesBundle) then
		return CopyFolder(resourcesBundle & DELIM & dialogToolkitPlusBundle)
	else
		return false
	end if
end IsDialogToolkitInstalled

on InstallDialogToolkitPlus(resourcesFolder)
	set downloadURL to "https://raw.githubusercontent.com/papercutter0324/SpeakingEvals/main/AppleScript/Dialog_Toolkit.zip"
	
	set dialogBundleName to "Dialog Toolkit Plus.scptd"
	
	set resourcesFolder to NormalizeFolderPath(resourcesFolder)
	set homeFolder to NormalizeFolderPath(POSIX path of (path to home folder))
	set scriptLibrariesFolder to NormalizeFolderPath(homeFolder & "Library/Script Libraries")
	
	set dialogToolkitPlusBundle to scriptLibrariesFolder & dialogBundleName
	set zipFilePath to resourcesFolder & "Dialog_Toolkit.zip"
	set zipExtractionPath to resourcesFolder & "dialogToolkitTemp"
	set extractedBundlePath to zipExtractionPath & "/Dialog_Toolkit/" & dialogBundleName
	
	if DoesBundleExist(dialogToolkitPlusBundle) then
		return true
	end if
	
	if DoesBundleExist(resourcesFolder & dialogBundleName) then
		if CopyBundle(resourcesFolder & dialogBundleName, dialogToolkitPlusBundle) then
			return true
		end if
	end if
	
	if DownloadFile(zipFilePath & DELIM & downloadURL) then
		try
			do shell script "mkdir -p " & quoted form of zipExtractionPath
			do shell script "unzip -oq " & quoted form of zipFilePath & " -d " & quoted form of zipExtractionPath
			
			CopyBundle(extractedBundlePath, resourcesFolder & dialogBundleName)
			CopyBundle(extractedBundlePath, dialogToolkitPlusBundle)
			
			DeleteFile(zipFilePath)
			DeleteFolder(zipExtractionPath & DELIM & resourcesFolder)
			
			return DoesBundleExist(dialogToolkitPlusBundle)
		on error errMsg number errNum
			display alert "Error" message ("(" & errNum & ") " & errMsg) as critical
			return false
		end try
	end if
end InstallDialogToolkitPlus

on UninstallDialogToolkitPlus(resourcesFolder)
	set dialogToolkitPlusBundle to POSIX path of (path to home folder) & "Library/Script Libraries/Dialog Toolkit Plus.scptd"
	set localCopy to resourcesFolder & "/Dialog Toolkit Plus.scptd"
	
	if DoesBundleExist(dialogToolkitPlusBundle) then
		try
			if not DoesBundleExist(localCopy) then CopyBundle(dialogToolkitPlusBundle, localCopy)
			DeleteFolder(dialogToolkitPlusBundle & DELIM & resourcesFolder)
			return true
		on error
			return false
		end try
	else
		return true
	end if
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
			return "Ok"
		else
			return "Error: Local copy not available."
		end if
	end try
end VerifyFileIsAvailableLocally
