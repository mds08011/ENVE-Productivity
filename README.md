# ENVE-Productivity
Small scripts to aid common tasks

## VBA Macro in Outlook
This is the most integrated solution. You can add a button to your Outlook Quick Access Toolbar that runs a macro on the currently selected email. It will extract the date and subject, create the formatted folder on your K: drive, and save both the email and its attachments into it.

How It Works:
Select an important email in your Outlook inbox.

Click a custom button on your toolbar.

The macro automatically creates a folder like K:\YourProject\2025.09.01 - Project Update and saves the email and any PDFs inside.

The new folder structure will be: yyyy.mm.dd - Sender Name - Subject
For example: K:\YourProject\2025.09.01 - John Smith - Submittal RFI Response



Setup and Code:

You'll need to enable the "Developer" tab in Outlook and paste this code into a new module.

Enable Developer Tab: Go to File > Options > Customize Ribbon and check the box for Developer.

Open VBA Editor: Click the new Developer tab, then click Visual Basic.

Insert a Module: In the new window, go to Insert > Module.

Paste the Code: Copy and paste the code below into the white space.

Customize and Save: You must change the baseFolderPath variable to your project's root folder path. Save your work (Ctrl + S).
