# Save Excel Attachments from Outlook to Folder (Outlook Macro)

## Author
This script was authored by [aviado1](https://github.com/aviado1).

## Description
This script enables users to select an Outlook folder and save all Excel file attachments (.xlsx and .xls) from emails in that folder to a specified directory on the computer.

## Usage
1. Open Outlook.
2. Press `Alt` + `F11` to open the Visual Basic for Applications (VBA) editor.
3. In the VBA editor, click `Insert` > `Module` to insert a new module.
4. Copy and paste the provided script into the module window.
5. Modify the `SaveFolderPath` variable to specify the folder path where you want to save the attachments.
6. Run the `saveexcel1` subroutine by clicking `Run` > `Run Sub/UserForm` or by pressing `F5`.
7. Select the Outlook folder containing the emails with Excel attachments when prompted.
8. Once the script completes, Excel attachments will be saved to the specified folder.

## Important Note
Ensure macro settings in Outlook are configured to allow the execution of macros for this script to work properly.

## Script
```vba
' Script Title: Save Excel Attachments from Outlook to Folder
' Description: This script allows the user to select an Outlook folder and then saves 
' all Excel file attachments (.xlsx and .xls) from emails in that folder to a specified 
' directory on the computer.

Sub saveexcel1()
    Dim OutlookApp As Outlook.Application
    Dim Namespace As Outlook.Namespace
    Dim Folder As Outlook.Folder
    Dim MailItem As Object
    Dim Attachments As Outlook.Attachments
    Dim attachment As Outlook.attachment
    Dim SaveFolderPath As String
    Dim FldrPicker As Outlook.Folder

    ' Specify the folder path where you want to save the attachments
    SaveFolderPath = "C:\SampleFolder"

    Set OutlookApp = New Outlook.Application
    Set Namespace = OutlookApp.GetNamespace("MAPI")

    ' Display a dialog to select a folder
    Set FldrPicker = Namespace.PickFolder

    ' Exit if no folder is selected
    If FldrPicker Is Nothing Then
        MsgBox "No Folder Selected. Exiting.", vbExclamation
        Exit Sub
    End If

    Set Folder = FldrPicker

    ' Loop through each mail item in the folder
    For Each MailItem In Folder.Items
        Set Attachments = MailItem.Attachments

        ' Loop through each attachment in the mail item
        For Each attachment In Attachments
            ' Check if the attachment is an Excel file
            If LCase(Right(attachment.FileName, 5)) = ".xlsx" Or LCase(Right(attachment.FileName, 4)) = ".xls" Then
                ' Save the Excel attachment to the specified folder
                attachment.SaveAsFile SaveFolderPath & "\" & attachment.FileName
            End If
        Next attachment
    Next MailItem

    ' Clean up
    Set OutlookApp = Nothing
    Set Namespace = Nothing
    Set Folder = Nothing
    Set Attachments = Nothing
    Set attachment = Nothing
    Set FldrPicker = Nothing

    ' Confirmation message
    MsgBox "Excel attachments have been saved to " & SaveFolderPath, vbInformation
End Sub
