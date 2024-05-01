Attribute VB_Name = "Module1"
Private Declare Function ShellExecute Lib "shell32.dll" Alias _
  "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, _
  ByVal lpFile As String, ByVal lpParameters As String, _
  ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Sub PrintSelectedAttachments()
    Dim objFileSystem As Object
    Dim Exp As Outlook.Explorer
    Dim Sel As Outlook.Selection
    Dim obj As Object
    Dim proceed As VbMsgBoxResult
    
    ' Ask user for confirmation
    proceed = MsgBox("Are you sure you want to print all selected attachments?", vbYesNo + vbQuestion, "Print Attachments Confirmation")
    
    If proceed = vbNo Then
        Exit Sub
    End If
    
    Set Exp = Application.ActiveExplorer
    Set Sel = Exp.Selection
    For Each obj In Sel
        If TypeOf obj Is Outlook.MailItem Then
            PrintAttachments obj
        End If
    Next
End Sub


Private Sub PrintAttachments(oMail As Outlook.MailItem)
    On Error Resume Next
    Dim colAtts As Outlook.Attachments
    Dim oAtt As Outlook.Attachment
    Dim sFile As String
    Dim sDirectory As String
    Dim sFileType As String
    Dim folderCounter As Integer
    Dim objFileSystem As Object
    
    Set objFileSystem = CreateObject("Scripting.FileSystemObject")
    sDirectory = Environ("TEMP") & "\" ' Get the path to the system's temporary folder
    
    ' Loop through existing folders to find the next available folder name
    folderCounter = 1
    Do While objFileSystem.FolderExists(sDirectory & "Temp for Attachments_" & folderCounter)
        folderCounter = folderCounter + 1
    Loop
    
    sDirectory = sDirectory & "Temp for Attachments_" & folderCounter & "\"
    objFileSystem.CreateFolder sDirectory ' Create the folder
    
    Set colAtts = oMail.Attachments

    If colAtts.Count Then
        For Each oAtt In colAtts
            sFileType = LCase$(Right$(oAtt.FileName, 4))

            Select Case sFileType
                Case ".xls", ".doc", ".pdf"
                    sFile = sDirectory & oAtt.FileName
                    oAtt.SaveAsFile sFile
                    ShellExecute 0, "print", sFile, vbNullString, vbNullString, 0
            End Select
        Next
        MsgBox "  Printing All Selected Attachments", vbInformation, "Success!"
    End If
End Sub

Sub ClearTempFolders()
    Dim objFileSystem As Object
    Dim strTempFolder As String
    Dim folderPath As String
    Dim folderName As String
    
    Set objFileSystem = CreateObject("Scripting.FileSystemObject")
    strTempFolder = objFileSystem.GetSpecialFolder(2).Path ' Get the path to the temp folder
    
    ' Loop through all subfolders in the temp folder
    For Each Folder In objFileSystem.GetFolder(strTempFolder).SubFolders
        folderName = Folder.Name
        ' Check if the folder name contains "Temp for Attachments"
        If InStr(folderName, "Temp for Attachments") > 0 Then
            folderPath = Folder.Path
            ' Delete the folder and all its contents
            objFileSystem.DeleteFolder folderPath, True
        End If
    Next Folder
    MsgBox "  Temporary Files Cleared.", vbOKOnly, "Success!"
End Sub

' Created By: Trey McBride


