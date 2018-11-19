Dim SavePath
Dim Sender
Dim FileExtension
Dim DateToday

SavePath = "<Save_Path_Var>\"
Sender = "'<Sender_Name_Var>'"
FileExtension = "<File_Extension_Var>"
DateToday = CreateObject("system.text.stringbuilder").AppendFormat("{0:MM}-{0:dd}-{0:yyyy}", now).ToString()

Set objFileToWrite = CreateObject("Scripting.FileSystemObject").OpenTextFile("<LogPath_Var>\Log.txt",8,true)
objFileToWrite.WriteLine(vbcrlf)

Set objOutlook = CreateObject("Outlook.Application")
Set objNamespace = objOutlook.GetNamespace("MAPI")
Set objFolder = objNamespace.GetDefaultFolder(6) 'Inbox

Set colItems = objFolder.Items
Set colFilteredItems = colItems.Restrict("[ReceivedTime] >=' "  & DateToday & "'")
Set colFilteredItems = colFilteredItems.Restrict("[SenderName] = " & Sender)

For Each objMessage In colFilteredItems
strSubject = objMessage.Subject
    intCount = objMessage.Attachments.Count
    If intCount > 0 Then
        For i = 1 To intCount
                objMessage.Attachments.Item(i).SaveAsFile SavePath &  _
                    objMessage.Attachments.Item(i).FileName
                objFileToWrite.WriteLine("($Day$-$Month$-$Year$ $Hour$:$Minute$:$Second$) Downloaded " & objMessage.Attachments.Item(i).FileName & " for " & strSubject)
        Next
        objMessage.Unread = False
    End If
Next