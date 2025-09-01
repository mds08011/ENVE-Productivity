' VBA Macro that Prompts for a Project Folder

Sub SaveEmailAndAttachments_PromptForProject()
    ' --- USER: UPDATE THIS VARIABLE ---
    ' Set the root path where all your project folders are located.
    Const projectRootPath As String = ""
    ' ---------------------------------

    Dim olItem As Object
    Dim olMail As Outlook.MailItem
    Dim olAttachment As Outlook.Attachment
    Dim subjectStr, dateStr, folderPath, filePath, senderStr As String
    Dim projectFolderName As String
    Dim finalBasePath As String
    Dim fs As Object ' FileSystemObject

    ' --- 1. Get Project Name from User ---
    projectFolderName = InputBox("Enter the name of the project folder:", "Select Project")
    
    ' Exit if the user cancels or enters nothing
    If Trim(projectFolderName) = "" Then
        Exit Sub
    End If
    
    ' Construct the final base path for this project
    finalBasePath = projectRootPath & projectFolderName & "\"

    ' Get the currently selected item in Outlook
    Set olItem = Application.ActiveExplorer.Selection.Item(1)
    If TypeName(olItem) <> "MailItem" Then
        MsgBox "Please select an email before running this macro.", vbExclamation
        Exit Sub
    End If
    Set olMail = olItem
    Set fs = CreateObject("Scripting.FileSystemObject")

    ' --- 2. Check if the Project Folder Exists ---
    If Not fs.FolderExists(finalBasePath) Then
        MsgBox "The specified project folder does not exist:" & vbCrLf & finalBasePath, vbCritical
        Exit Sub
    End If
    
    ' --- 3. Prepare Subfolder and File Names ---
    senderStr = olMail.SenderName
    senderStr = Replace(Replace(senderStr, "/", "-"), "\", "-") ' Sanitize

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

    dateStr = Format(olMail.ReceivedTime, "yyyy.mm.dd")

    ' --- 4. Create the Destination Subfolder ---
    folderPath = finalBasePath & dateStr & " - " & subjectStr & " [from " & senderStr & "]"
    If Not fs.FolderExists(folderPath) Then
        fs.CreateFolder folderPath
    End If

    ' --- 5. Save the Email and Attachments ---
    filePath = folderPath & "\" & subjectStr & ".msg"
    olMail.SaveAs filePath, olMSG

    If olMail.Attachments.Count > 0 Then
        For Each olAttachment In olMail.Attachments
            olAttachment.SaveAsFile folderPath & "\" & olAttachment.FileName
        Next olAttachment
    End If
    
    MsgBox "Email and attachments saved successfully to:" & vbCrLf & folderPath, vbInformation
End Sub
