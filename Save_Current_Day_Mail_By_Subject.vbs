Dim SavePath
Dim Subject
Dim FileExtension
Dim DateToday


SavePath = "<Save_Path_Var>\"
Subject = "'<Subject_Var>'"
FileExtension = "<FileExtension_Var>"
DateToday = CreateObject("system.text.stringbuilder").AppendFormat("{0:MM}-{0:dd}-{0:yyyy}", now).ToString()

Set objOutlook = CreateObject("Outlook.Application")
Set objNamespace = objOutlook.GetNamespace("MAPI")
Set objFolder = objNamespace.GetDefaultFolder(6) 'Inbox

Set colItems = objFolder.Items
Set colFilteredItems = colItems.Restrict("[ReceivedTime] >=' "  & DateToday & "'")
Set colFilteredItems = colFilteredItems.Restrict("[Subject] = " & Subject)

For Each objMessage In colFilteredItems
    intCount = objMessage.Attachments.Count
    If intCount > 0 Then
        For i = 1 To intCount
                objMessage.Attachments.Item(i).SaveAsFile SavePath &  _
                    objMessage.Attachments.Item(i).FileName
        Next
        objMessage.Unread = False
    End If
Next