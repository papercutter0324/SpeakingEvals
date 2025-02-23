#If Mac Then
#Else
    #If VBA7 Then
        Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
            ByVal hwnd As LongPtr, ByVal lpOperation As String, ByVal lpFile As String, _
            ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As LongPtr
    #Else
        Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
            ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
            ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
    #End If
#End If
