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

    pdfCount = 0
    Set tempFiles = New Collection

    tempFolderPath = Environ("Temp") & "\OutlookTempPDFs\"

    If Dir(tempFolderPath, vbDirectory) = "" Then
        MkDir tempFolderPath
    End If

    For Each olItem In Application.ActiveExplorer.Selection
        If TypeOf olItem Is MailItem Then
            For Each olAttachment In olItem.Attachments
                If LCase(Right(olAttachment.FileName, 4)) = ".pdf" Then
                    filePath = tempFolderPath & olAttachment.FileName
                    olAttachment.SaveAsFile filePath
                    pdfCount = pdfCount + 1
                    tempFiles.Add filePath
                End If
            Next olAttachment
        End If
    Next olItem

    MsgBox "Found " & pdfCount & " PDF(s) to print.", vbInformation, "PDF Count"
    
    If pdfCount > 0 Then
        If MsgBox("Do you want to print these PDFs?", vbYesNo + vbQuestion, "Print Confirmation") = vbYes Then
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

