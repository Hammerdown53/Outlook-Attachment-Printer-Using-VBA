Attribute VB_Name = "Module2"
Sub CleanUpTempFiles()
    Dim tempFolderPath As String
    Dim filePath As String
    
    ' Get the Temp folder path
    tempFolderPath = Environ("Temp") & "\OutlookTempPDFs\"

    ' Clean up the temp files
    On Error Resume Next
    filePath = Dir(tempFolderPath & "*.pdf")
    Do While filePath <> ""
        Kill tempFolderPath & filePath
        filePath = Dir
    Loop
    On Error GoTo 0
    
    ' Remove the temp folder if empty
    On Error Resume Next
    RmDir tempFolderPath
    On Error GoTo 0
    
    MsgBox "Temporary files cleaned up.", vbInformation, "Cleanup Complete"
End Sub
