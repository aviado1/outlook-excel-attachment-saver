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
    Set Namespace = OutlookApp.GetNamespace("SampleNamespace")

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
