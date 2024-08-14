Attribute VB_Name = "module1"
Sub SimplePrintPDFsFromSelectedEmails()
    Dim olItem As Object
    Dim olAttachment As Attachment
    Dim tempFolderPath As String
    Dim filePath As String
    Dim pdfCount As Integer
    Dim shellObj As Object
    Dim tempFiles As Collection
    Dim tempFile As Variant

    ' Initialize PDF counter and collection for temp files
    pdfCount = 0
    Set tempFiles = New Collection
    
    ' Get the Temp folder path
    tempFolderPath = Environ("Temp") & "\OutlookTempPDFs\"

    ' Create the temp folder if it doesn't exist
    If Dir(tempFolderPath, vbDirectory) = "" Then
        MkDir tempFolderPath
    End If
    
    ' Loop through the selected items
    For Each olItem In Application.ActiveExplorer.Selection
        ' Check if the item is a MailItem
        If TypeOf olItem Is MailItem Then
            ' Loop through the attachments
            For Each olAttachment In olItem.Attachments
                ' Check if the attachment is a PDF
                If LCase(Right(olAttachment.FileName, 4)) = ".pdf" Then
                    ' Save the attachment to the temp folder
                    filePath = tempFolderPath & olAttachment.FileName
                    olAttachment.SaveAsFile filePath
                    pdfCount = pdfCount + 1
                    tempFiles.Add filePath
                End If
            Next olAttachment
        End If
    Next olItem

    ' Notify the user how many PDFs were found
    MsgBox "Found " & pdfCount & " PDF(s) to print.", vbInformation, "PDF Count"
    
    ' Only proceed if PDFs were found
    If pdfCount > 0 Then
        ' Confirm if the user wants to print the PDFs
        If MsgBox("Do you want to print these PDFs?", vbYesNo + vbQuestion, "Print Confirmation") = vbYes Then
            ' Print each PDF file
            Set shellObj = CreateObject("Shell.Application")
            For Each tempFile In tempFiles
                shellObj.ShellExecute tempFile, "", "", "print", 0
            Next tempFile
            
            MsgBox "PDFs have been sent to the printer. Run the cleanup macro to remove the temporary files.", vbInformation, "Print Complete"
        Else
            MsgBox "Printing canceled.", vbExclamation, "Process Canceled"
        End If
    Else
        MsgBox "No PDF attachments found.", vbExclamation, "No PDFs Found"
    End If
End Sub

