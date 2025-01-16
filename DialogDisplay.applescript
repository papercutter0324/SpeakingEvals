use AppleScript version "2.4" -- Yosemite (10.10) or later
use scripting additions
use script "Dialog Toolkit Plus" version "1.1.3"

on GetScriptVersionNumber(paramString)
	--- Use build number to determine if an update is available
	return 20250113
end GetScriptVersionNumber

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
	set {buttonName, controlsResults} to display enhanced window dialogTitle Â
		acc view width accViewWidth Â
		acc view height theTop Â
		acc view controls {messageLabel} Â
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

