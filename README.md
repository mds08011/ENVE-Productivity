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






This version will pop up a simple input box each time you run the macro, asking for the specific project folder name. You just type the name (e.g., "Project 123 - City Hall") and it does the rest.
This method is great because it requires no pre-configuration.

How it Works:

Select an email and run the macro.
An input box appears asking for the project folder name.
You type Project 123 - City Hall.
The macro saves the files to K:\Master\Project List\Project 123 - City Hall\2025.09.01 - Sender


Outlook's VBA doesn't have a direct "Save as PDF" command, so the best way to do this is to use the Microsoft Word engine that Outlook uses to edit emails. The macro will programmatically open the email, use Word's "Export to PDF" function, and then close it.
​Step 1: Enable the Microsoft Word Library
​First, you need to tell Excel's VBA editor that it's allowed to use Microsoft Word commands. This is a one-time setup.
​In the VBA Editor, go to the menu and click Tools -> References....
​A new window will pop up. Scroll down the list until you find "Microsoft Word 16.0 Object Library". The version number (16.0) might be different on your computer, but the name will be the same.
​Check the box next to it and click OK.
