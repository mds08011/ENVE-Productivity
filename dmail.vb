' VBA Macro to Save Selected Email and Attachments to a Dated Folder with Sender Name

Sub SaveEmailAndAttachments()
    ' --- USER: UPDATE THIS VARIABLE ---
    ' Set the base path for your project. Make sure it ends with a backslash \
    Const baseFolderPath As String = "K:\Path\To\Your\ProjectFolder\"
    ' ---------------------------------

    Dim olItem As Object
    Dim olMail As Outlook.MailItem
    Dim olAttachment As Outlook.Attachment
    Dim subjectStr As String
    Dim dateStr As String
    Dim folderPath As String
    Dim filePath As String
    Dim fs As Object ' FileSystemObject
    
    ' *** NEW: Added variable for the sender's name ***
    Dim senderStr As String

    ' Get the currently selected item in Outlook
    Set olItem = Application.ActiveExplorer.Selection.Item(1)

    ' Ensure the selected item is an email
    If TypeName(olItem) <> "MailItem" Then
        MsgBox "Please select an email before running this macro.", vbExclamation
        Exit Sub
    End If

    Set olMail = olItem
    Set fs = CreateObject("Scripting.FileSystemObject")

    ' --- 1. Prepare Folder and File Names ---
    
    ' *** NEW: Get and sanitize the sender's name ***
    senderStr = olMail.SenderName
    senderStr = Replace(senderStr, "/", "-")
    senderStr = Replace(senderStr, "\", "-")
    senderStr = Replace(senderStr, ":", "-")
    
    ' Sanitize the subject line to remove characters invalid for folder names
    subjectStr = olMail.Subject
    subjectStr = Replace(subjectStr, "/", "-")
    subjectStr = Replace(subjectStr, "\", "-")
    subjectStr = Replace(subjectStr, ":", "-")
    subjectStr = Replace(subjectStr, "*", "-")
    subjectStr = Replace(subjectStr, "?", "")
    subjectStr = Replace(subjectStr, """", "")
    subjectStr = Replace(subjectStr, "<", "")
    subjectStr = Replace(subjectStr, ">", "")
    subjectStr = Replace(subjectStr, "|", "-")

    ' Format the received date as yyyy.mm.dd
    dateStr = Format(olMail.ReceivedTime, "yyyy.mm.dd")

    ' --- 2. Create the Destination Folder ---
    ' *** UPDATED: Added the senderStr variable to the folder path ***
    folderPath = baseFolderPath & dateStr & " - " & senderStr & " - " & subjectStr
    
    If Not fs.FolderExists(folderPath) Then
        fs.CreateFolder folderPath
    End If

    ' --- 3. Save the Email Message ---
    ' Use the sanitized subject for the file name as well
    filePath = folderPath & "\" & subjectStr & ".msg"
    olMail.SaveAs filePath, olMSG

    ' --- 4. Save All Attachments ---
    If olMail.Attachments.Count > 0 Then
        For Each olAttachment In olMail.Attachments
            olAttachment.SaveAsFile folderPath & "\" & olAttachment.FileName
        Next olAttachment
    End If
    
    MsgBox "Email and attachments saved successfully to:" & vbCrLf & folderPath, vbInformation

End Sub
