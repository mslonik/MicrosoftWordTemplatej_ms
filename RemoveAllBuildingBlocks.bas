Attribute VB_Name = "RemoveAllBuildingBlocks"
Sub RemoveAllBuildingBlocksFromBuiltInTemplate()
    Dim template As template
    Dim bb As buildingBlock
    Dim j As Integer
    Dim userName As String
    Dim langID As String
    Dim filePath As String
    Dim response As VbMsgBoxResult

    ' Get the currently logged user name
    userName = Environ("USERNAME")

    ' Get the Office authoring languages and proofing language ID
    authoringLanguageID = Application.LanguageSettings.LanguageID(msoLanguageIDInstall)
    
    ' Construct the file path
    filePath = "C:\Users\" & userName & "\AppData\Roaming\Microsoft\Document Building Blocks\" & authoringLanguageID & "\16\Built-In Building Blocks.dotx"

    ' Ask the user to confirm the file path
    response = MsgBox("The following file will be used: " & filePath & vbCrLf & "Is this correct?", vbYesNo + vbQuestion, "Confirm File Path")

    ' If the user confirms, proceed to load the template and delete building blocks
    If response = vbYes Then
        On Error Resume Next
        Set template = Templates(filePath)
        On Error GoTo 0

        ' Check if the template is loaded successfully
        If template Is Nothing Then
            MsgBox "The 'Built-In Building Blocks.dotx' template could not be loaded."
            Exit Sub
        End If

        ' Loop through all building blocks and delete them
        For j = template.BuildingBlockEntries.Count To 1 Step -1
            Set bb = template.BuildingBlockEntries(j)
            bb.Delete
        Next j

        MsgBox "All Building Blocks have been deleted from the 'Built-In Building Blocks.dotx' template." _
             & vbCrLf & vbCrLf & _
             "To save the changes in the 'Built-In Building Blocks.dotx' template" & vbCrLf & _
             "exit currently edited file, exit the Microsoft Word and save the changes to the" & vbCrLf & _
             "'Built-In Building Blocks.dotx' template"
    Else
        MsgBox "Operation canceled by the user."
    End If
End Sub

