' VBA Macro that Prompts for a Project Folder and saves as .MSG and .PDF

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
    
    '--- Variables for PDF Export ---
    Dim olInspector As Outlook.Inspector
    Dim wdDoc As Word.Document
    Dim pdfPath As String

    ' --- 1. Get Project Name from User ---
    projectFolderName = InputBox("Enter the name of the project folder:", "Select Project")
    
    If Trim(projectFolderName) = "" Then
        Exit Sub
    End If
    
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
    senderStr = Replace(Replace(senderStr, "/", "-"), "\", "-")

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
    ' *** UPDATED FOLDER NAME ORDER AS REQUESTED ***
    folderPath = finalBasePath & dateStr & " - " & subjectStr & " [from " & senderStr & "]"
    
    If Not fs.FolderExists(folderPath) Then
        fs.CreateFolder folderPath
    End If

    ' --- 5. Save the Email as .MSG and Attachments ---
    filePath = folderPath & "\" & subjectStr & ".msg"
    olMail.SaveAs filePath, olMSG

    If olMail.Attachments.Count > 0 Then
        For Each olAttachment In olMail.Attachments
            olAttachment.SaveAsFile folderPath & "\" & olAttachment.FileName
        Next olAttachment
    End If
    
    ' --- 6. Save the Email Body as a PDF ---
    On Error Resume Next ' In case Word automation fails
    
    pdfPath = folderPath & "\" & subjectStr & ".pdf"
    
    Set olInspector = olMail.GetInspector
    Set wdDoc = olInspector.WordEditor
    
    If Not wdDoc Is Nothing Then
        wdDoc.ExportAsFixedFormat OutputFileName:=pdfPath, _
                                  ExportFormat:=wdExportFormatPDF
    End If
    
    olInspector.Close olDiscard
    
    Set wdDoc = Nothing
    Set olInspector = Nothing
    On Error GoTo 0 ' Resume normal error handling
    
    MsgBox "Email, PDF, and attachments saved successfully to:" & vbCrLf & folderPath, vbInformation
End Sub
