 On Error Resume Next
Dim myNameSpace
Dim ofChosenFolder
Dim myOlApp
Dim myItem
Dim objItem
Dim myFolder
Dim strSubject
Dim strName
Dim strFile
Dim strReceived
Dim strSavePath
 
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set myOlApp = CreateObject("Outlook.Application")
Set myNameSpace = myOlApp.GetNamespace("MAPI")
Set ofChosenFolder = myNameSpace.PickFolder
 
 
'get path to My Docs
Dim szDocsFolder, g_shell
Set g_shell = CreateObject("WScript.Shell")
szTempFolder = g_shell.SpecialFolders("MyDocuments")
 
'Get the current Username
Set WshNetwork = WScript.CreateObject("WScript.Network")
strUser = WshNetwork.UserName
 
strSavePath = InputBox("Please enter the path to save to and be sure to end with a backslash at the end of your path. You can enter a new folder name if you like and it will be created", "Save Emails To:", szTempFolder & "\Saved Emails\" & ofChosenFolder & "\")
 
If not right(strSavePath,1) = "\" then
strSavePath = strSavePath & "\"
wscript.echo "You forgot a backslash at the end of your path." & vbcrlf & "But don't worry, I added one for you."
End If
 
' strSavePath = strSavePath & ofChosenFolder & "\"
 
strSaveFolder = Left(strSavePath, Len(strSavePath)-1)
 
If Not objFSO.FolderExists(strSaveFolder) then
if MsgBox("The folder you specified does not exist." & vbcrlf & "Would you like one created?", VBYesNo, "Folder Not Found") = 7 then
wscript.echo "Exiting script. Try again."
Else
objFSO.CreateFolder(strSaveFolder)
wscript.echo strSaveFolder & " - Created"
End if
End if
 
 
i = 1
For each Item in ofChosenFolder.Items
Set myItem = ofChosenFolder.Items(i)
strReceived = ArrangedDate(myitem.ReceivedTime)
' strSubject = myItem.Subject
strSubject = myitem.SenderName & "_" & myitem.Subject
strName = StripIllegalChar(strSubject)
strFile = strSavePath & strReceived & "_" & strName & ".msg"
If Not Len(strfile) > 256 then
myItem.SaveAs strfile, 3
Else
wscript.echo strfile & vbcrlf & "Path and filename too long."
End If
i = i + 1
next
 
 
 
Function StripIllegalChar(strInput)
 
'***************************************************
'Simple function that removes illegal file system
'characters.
'***************************************************
 
Set RegX = New RegExp
 
RegX.pattern = "[\" & chr(34) & "\!\@\#\$\%\^\&\*\(\)\=\+\|\[\]\{\}\`\'\;\:\<\>\?\/\,]"
RegX.IgnoreCase = True
RegX.Global = True
 
StripIllegalChar = RegX.Replace(strInput, "")
Set RegX = nothing
 
End Function
 
 
Function ArrangedDate(strDateInput)
 
'***************************************************
'This function re-arranges the date data in order
'for it to display in chronilogical order in a
'sorted list in the file system. It also removes
'illegal file system characters and replaces them
'with dashes.
'Example:
'Input: 2/26/2004 7:07:33 AM
'Output: 2004-02-26_AM-07-07-33
'***************************************************
 
Dim strFullDate
Dim strFullTime
Dim strAMPM
Dim strTime
Dim strYear
Dim strMonthDay
Dim strMonth
Dim strDay
Dim strDate
Dim strDateTime
Dim RegX
 
If not Left(strDateInput, 2) = "10" Then
If not Left(strDateInput, 2) = "11" Then
If not Left(strDateInput, 2) = "12" Then
strDateInput = "0" & strDateInput
End If
End If
End If
 
strFullDate = Left(strDateInput, 10)
 
If Right(strFullDate, 1) = " " Then
strFullDate = Left(strDateInput, 9)
End If
 
strFullTime = Replace(strDateInput,strFullDate & " ","")
 
If Len(strFullTime) = 10 Then
strFullTime = "0" & strFullTime
End If
 
strAMPM = Right(strFullTime, 2)
 
strTime = strAMPM & "-" & Left(strFullTime, 8)
 
strYear = Right(strFullDate,4)
 
strMonthDay = Replace(strFullDate,"/" & strYear,"")
 
strMonth = Left(strMonthDay, 2)
 
strDay = Right(strMonthDay,len(strMonthDay)-3)
 
If len(strDay) = 1 Then
strDay = "0" & strDay
End If
 
strDate = strYear & "-" & strMonth & "-" & strDay
 
'strDateTime = strDate & "_" & strTime
strDateTime = strDate
 
Set RegX = New RegExp
 
RegX.pattern = "[\:\/\ ]"
RegX.IgnoreCase = True
RegX.Global = True
 
ArrangedDate = RegX.Replace(strDateTime, "-")
 
Set RegX = nothing
 
End Function