Attribute VB_Name = "RefToHyperlinks"
' RefToHyperlinks macro
' 2025-02-10 by ms
' - fixed indentation
' - text strings translated into English
' - code rework

Sub RefToHyperlinks()
    Dim oSource As Document
    Set oSource = ActiveDocument
    Dim i As Long, j As Long
    
    If Application.Version <> "14.0" And Application.Version <> "16.0" Then
        MsgBox "This macro couldn't run with your version of Microsoft Office!", vbCritical
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
   ' When Application.ScreenUpdating is set to False, it turns off screen updating, which can significantly speed up the execution of a macro by preventing the screen from refreshing until the macro has finished running. This is particularly useful for macros that perform a lot of operations, as it reduces the time spent on rendering the screen.
        
    ' Adds a bookmark in place where cursor is present
    If ActiveWindow.ActivePane.view.SeekView = wdSeekMainDocument Then _
        Selection.Bookmarks.Add ("LastCursorPosition")
    
    i = ActiveDocument.Fields.Count
    j = 0
    ProgressHyperlinks.progresslabel = "Finished: " & j & " out of " & i
    ProgressHyperlinks.Show
    For Each aField In oSource.Fields
        If aField.Type = wdFieldRef Then
            If InStr(aField.Code, "_Ref") > 0 And aField.Code Like "*\h*" = 0 Then
                aField.Select
                aField.Code.InsertAfter (" \h")
            End If
            If (InStr(aField.Code, "\h") Or InStr(aField.Code, "\H")) Then
                If (aField.Code Like "*\* MERGEFORMAT*" = -1) Then
                    aField.Code.Text = Replace(aField.Code, " \* MERGEFORMAT ", "", 1, -1, vbTextCompare)
                    aField.Update
                End If
                If (aField.Code Like "*\* CHARFORMAT*" = 0) Then
                    ' adds tag \*Charformat
                    aField.Code.InsertAfter ("\* CHARFORMAT ")
                    aField.Update
                End If
                aField.Select
                Selection.Font.Underline = wdUnderlineSingle
                Selection.Font.ColorIndex = wdBlue  ' this color should be set to current theme color, but I don't know how to do that
            End If
        End If
        ' only the numbering in ToC is shown as a hyperlink (wdFieldPageRef); alternative: wdFieldHyperlink
        If (aField.Type = wdFieldPageRef) Then
            aField.Select
            Selection.Font.Underline = wdUnderlineSingle
            Selection.Font.ColorIndex = wdBlue  ' this color should be set to current theme color, but I don't know how to do that
        End If
        
        j = j + 1
        ProgressHyperlinks.progresslabel = "Finished: " & j & " out of " & i
        ' MsgBox "" ' for debugging only
        DoEvents
        ' The DoEvents function in Visual Basic for Applications (VBA) for Microsoft Word is used to yield execution so that the operating system can process other events. This function allows the operating system to handle other tasks, such as updating the screen, responding to user inputs, or processing other events in the queue, while your macro is running
    Next aField
       
    ActiveWindow.view.Type = wdPrintView
    Application.ScreenUpdating = True

    ' Goes to a place where temporary bookmark was located and removes it afterwards
    If ActiveDocument.Bookmarks.Exists("LastCursorPosition") Then
        Selection.GoTo What:=wdGoToBookmark, Name:="LastCursorPosition"
        ActiveDocument.Bookmarks("LastCursorPosition").Delete
    Else
        Selection.HomeKey wdStory
    End If

    Unload ProgressHyperlinks
End Sub
