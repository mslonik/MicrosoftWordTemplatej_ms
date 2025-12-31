Attribute VB_Name = "Validation"
' VBA Module name: Validation.bas
' https://github.com/mslonik/Microsoft-Word-Configuration
'
'   License: MIT License.
'
'
'+----+----------------------------+-------------+----------------------------+--------------------------+
'| No | Sub name                   | Ribbon name | Ribbon section             | Ribbon button name       |
'+----+----------------------------+-------------+----------------------------+--------------------------+
'| 1  | Tables_Format              | Validation  | Tables                     | Tables_Format            |
'| 2  | Tables_CheckNestingLevel   | Validation  | Tables                     | Tables_CheckNestingLevel |
'| 3  | UpdateAllFields            | Validation  | Custom (no name)           | UpdateAllFields          |
'| 4  | InsertNoBrakeSpace         | Validation  | Custom (no name)           | InsertNoBrakeSpace       |
'| 5  | ModifyReferencesToPicTab   | Validation  | ModifyReferencesToPicTab   |                          |
'| 6  | ReplaceUnwantedTextstrings | Validation  | ReplaceUnwantedTextstrings |                          |
'| 7  | FindParagraphStyling       | Validation  | FindParagraphStyling       |                          |
'| 8  | FindCharacterStyling       | Validation  | FindCharacterStyling       |                          |
'+----+----------------------------+-------------+----------------------------+--------------------------+
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
' Used to enforce the explicit declaration of all variables in a module. When you include Option Explicit at the beginning of a module, it ensures that you must
' declare all variables using the Dim, Private, Public, ReDim, or Static statements before using them. This helps prevent errors caused by typos or
' undeclared variables.
Option Explicit

' Declare the Sleep function at the top of your module
Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)


' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
' Checks if there are in the document body any reference displayed as "0". If it does, macro stops there.
' This is part of validation routines.
' Reworked by ms on 2025-02-11
' Changed to private function by ms on 2025-03-15
Private Function CheckRefZero() As Boolean
    Dim fld As Field

    Dim FileName As String:     FileName = C_F_Macros
    Dim ModuleName As String:   ModuleName = C_M_Validation
    Dim MacroName As String:    MacroName = "CheckRefZero"
    Dim MsgBoxTitle As String:  MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    For Each fld In ActiveDocument.Fields
        With fld
            If .Type = wdFieldRef Then
                If InStr(.result, ".0.") > 0 Or .result Like "0.*" Or .result Like "*.0" Or .result = "0" Then
                    .Select
                    CheckRefZero = False
                    Exit Function
                End If
            End If
        End With
    Next

    MsgBox _
        Prompt:="No 0 references were found.", _
        Buttons:=vbInformation, _
        Title:=MsgBoxTitle
    CheckRefZero = True
End Function

' Check if document are error references. Mark them in yellow.
' 2025-03-15 by ms
Private Function CheckRefError() As Boolean
    Dim fld As Field
    Dim rng As Range
    
    Dim FileName As String:     FileName = C_F_Macros
    Dim ModuleName As String:   ModuleName = C_M_Validation
    Dim MacroName As String:    MacroName = "CheckRefError"
    Dim MsgBoxTitle As String:  MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    Dim ErrorString As String
    ErrorString = "Error! Reference source not found."
    
    ' Loop through all fields in the document
    For Each fld In ActiveDocument.Fields
        ' Check if the field result contains the error text
        If InStr(fld.result.Text, ErrorString) > 0 Then
            Set rng = fld.result
            ' Jump to the field location
            rng.Select
            ' Highlight the field
            rng.HighlightColorIndex = wdYellow
            MsgBox _
                Prompt:="Error reference field was found. Aborting.", _
                Buttons:=vbCritical, _
                Title:=MsgBoxTitle
            CheckRefError = False
            Exit Function
        Else
           ' Switch off highlighting if condition isn't met
            fld.result.HighlightColorIndex = wdNoHighlight
        End If
    Next fld
    MsgBox _
        Prompt:="No reference error was found.", _
        Buttons:=vbInformation, _
        Title:=MsgBoxTitle
    CheckRefError = True
    
    ' Clear object variables
    Set rng = Nothing
End Function

' Search the current document body and delete the following text string sequences:
'   - double space: "  "
'   - triple space: "   "
'   - space + dot: " ."
'   - space + comma: " ,"
' 2025-02-11 by ms
Sub ReplaceUnwantedTextstrings()
    Dim findText As String
    Dim replaceText As String
    Dim Summary As String
    Dim counter As String
    
    Dim FileName As String:     FileName = C_F_Macros
    Dim ModuleName As String:   ModuleName = C_M_Validation
    Dim MacroName As String:    MacroName = "ReplaceUnwantedTextstrings"
    Dim MsgBoxTitle As String:  MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    Application.ScreenUpdating = True
    Call AddLastCursorPositionBookmark

    Summary = "Finished processing." _
        & vbNewLine
    
    ' Selection.HomeKey unit:=wdStory
    ' Section with regular expressions oriented to "Find and replace" syntax.
    ' The following syntax is similar to the "Find and Replace" dialog in Microsoft Word, which uses a different syntax for wildcards compared to regular expressions known from other applications. This is because Word's wildcard search is designed to be more user-friendly and accessible to those who may not be familiar with the more complex syntax of regular expressions.
    ' More details on that subject: https://wordmvp.com/FAQs/General/UsingWildcards.htm
    ' In other words the "Find and replace" syntax must not be confused with regular expressions syntax.
    ' Square Brackets [ ]: In regex, square brackets are used to define a character class. In this case, [ ] represents a single space character. It means that the regex will match any single space character.
    ' Curly Braces {2,}: Curly braces are used to specify the number of occurrences of the preceding element (in this case, the space character). The {2,} means "two or more" occurrences. So, [ ]{2,} will match any sequence of two or more consecutive space characters.
    
    findText = " {2,}"    ' "Find and replace" syntax to find two or more spaces
    replaceText = " "
    ReplacePattern findText, replaceText, Summary, MsgBoxTitle
            
    findText = " {1,}."      ' "Find and replace" syntax to find one or more spaces followed by a dot
    replaceText = "."
    ReplacePattern findText, replaceText, Summary, MsgBoxTitle
    
    findText = " {1,},"         ' "Find and replace" syntax to find one or more spaces followed by a comma
    replaceText = ","
    ReplacePattern findText, replaceText, Summary, MsgBoxTitle
    
    findText = " {1,}\!"         ' "Find and replace" syntax to find one or more spaces followed by a comma
    replaceText = "!"
    ReplacePattern findText, replaceText, Summary, MsgBoxTitle
    
    findText = " {1,}\?"         ' "Find and replace" syntax to find one or more spaces followed by a comma
    replaceText = "?"
    ReplacePattern findText, replaceText, Summary, MsgBoxTitle
    
    Call DeleteEmptyParagraphs(Summary, MsgBoxTitle)
    Call RemoveLastCursorPositionBookmark
    Call Logging(Summary, MsgBoxTitle)
    MsgBox _
        Prompt:=Summary, _
        Buttons:=vbInformation, _
        Title:=MsgBoxTitle
End Sub

Private Sub DeleteEmptyParagraphs(ByRef Summary As String, _
            MsgBoxHeader As String)
    
    Dim para As Paragraph
    Dim yesCounter As Integer
    Dim paraText As String
    Dim isEmptyParagraph As Boolean
    Dim i As Integer
    
    ' Initialize the counter
    yesCounter = 0
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    ' Loop through all paragraphs in the document
    For Each para In ActiveDocument.Paragraphs
        
        ' Get the text of the paragraph
        paraText = para.Range.Text
        
        ' Check if the paragraph contains only empty paragraphs (^13 characters)
        isEmptyParagraph = True
        For i = 1 To Len(paraText)
            If Mid(paraText, i, 1) <> Chr(13) Then
            ' If Mid(paraText, i, 1) <> Chr(13) And Mid(paraText, i, 1) <> Chr(10) And Mid(paraText, i, 1) <> " " Then
                isEmptyParagraph = False
                Exit For
            End If
        Next i
    
        ' If the paragraph is empty, set shading to yellow and ask user for action
        If isEmptyParagraph Then
            DoEvents    ' Force a screen refresh
            para.Range.shading.BackgroundPatternColor = wdColorYellow
            DoEvents    ' Force a screen refresh
            para.Range.Select
            Selection.Collapse Direction:=wdCollapseStart
            
            ' Display a message box to the user
            Dim UserDecision As VbMsgBoxResult
            Beep
            UserDecision = MsgBox( _
                Prompt:="An empty paragraph was found. Do you want to delete it?", _
                Buttons:=vbYesNoCancel + vbQuestion, _
                Title:=MsgBoxHeader)
            
            ' If the user chooses Yes, delete the empty paragraph
            If UserDecision = vbYes Then
                Selection.TypeBackspace
                ' Increment the counter
                yesCounter = yesCounter + 1
            
            ' If the user chooses No, change shading of the empty paragraph to "no color"
            ElseIf UserDecision = vbNo Then
                para.Range.shading.BackgroundPatternColor = wdColorAutomatic
                DoEvents    ' Force a screen refresh
            
            ' If the user chooses Cancel, exit the loop and remove highlighting and shading
            ElseIf UserDecision = vbCancel Then
                para.Range.shading.BackgroundPatternColor = wdColorAutomatic
                DoEvents    ' Force a screen refresh
                Summary = Summary & vbNewLine & "Number of removed empty paragraphs: " & yesCounter
                Exit Sub
            
            End If
        End If
    Next para
    
    Summary = Summary & vbNewLine & "Number of removed empty paragraphs: " & yesCounter
End Sub

Private Sub Logging(Summary As String, MsgBoxHeader As String)
    Dim FilePath As String
    Dim DocName As String
    Dim TemplatePath As String
    Dim filenum As Integer
    Dim CurrentDate As String
    Dim CurrentTime As String

    FilePath = Options.DefaultFilePath(wdDocumentsPath) & "\" & "Validation" & ".txt"
    DocName = ActiveDocument.Name
    TemplatePath = ActiveDocument.AttachedTemplate.FullName
    CurrentDate = Format(Date, "yyyy-mm-dd")
    CurrentTime = Format(Time, "hh:mm:ss")

    filenum = FreeFile
    If Dir(FilePath) <> "" Then
        Open FilePath For Append As filenum
    Else
        Open FilePath For Output As filenum
    End If
    Print #filenum, "Document Name: " & DocName
    Print #filenum, "Template Name: " & TemplatePath
    Print #filenum, "Macro Name: " & MsgBoxHeader
    Print #filenum, "Date: " & CurrentDate
    Print #filenum, "Time: " & CurrentTime
    Print #filenum, Summary
    Close filenum

End Sub

Private Sub ReplacePattern(findText As String, _
            replaceText As String, _
            ByRef Summary As String, _
            MsgBoxHeader As String)
            
    Dim count As Integer
    
    ' Initialize the counter
    count = 0
        
    ' Set parameters of the "Find and Replace"
    With Selection.Find
        .Text = findText
        .Replacement.Text = replaceText
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchWildcards = True
    End With
    
    Do While Selection.Find.Execute(Replace:=wdReplaceNone)
        Selection.Range.Select
        DoEvents    ' Force a screen refresh
        Selection.Range.HighlightColorIndex = wdYellow  ' Highlight the selection in yellow
        DoEvents    ' Force a screen refresh
        Dim UserDecision As VbMsgBoxResult
        Beep
        UserDecision = MsgBox( _
            Prompt:="Do you want to replace '" & findText & "' with '" & replaceText & "' here?", _
            Buttons:=vbYesNoCancel + vbQuestion, _
            Title:=MsgBoxHeader)
        
        If UserDecision = vbYes Then
            Selection.Find.Execute Replace:=wdReplaceOne
            DoEvents    ' Force a screen refresh
            Selection.Range.HighlightColorIndex = wdNoHighlight ' Remove the highlight if not replaced
            count = count + 1
        
        ElseIf UserDecision = vbNo Then
            DoEvents    ' Force a screen refresh
            Selection.Range.HighlightColorIndex = wdNoHighlight ' Remove the highlight if not replaced
            ' Remove shading if not replaced and it was set to yellow
            If Len(Trim(Selection.Range.Text)) = 0 Then
                Selection.Range.shading.BackgroundPatternColor = wdColorAutomatic
            End If
        
        ElseIf UserDecision = vbCancel Then
            ' Remove shading if canceled and it was set to yellow
            If Len(Trim(Selection.Range.Text)) = 0 Then
                Selection.Range.shading.BackgroundPatternColor = wdColorAutomatic
            End If
            Exit Do     ' Exit the loop if Cancel is pressed
        End If
    
    Loop

    Summary = Summary & vbNewLine & "Number of replacements """ & findText & """: " & count
End Sub

' Update all fields in the document, including headers / footers / Table of Contents (ToCs)
' 2025-02-11 by ms
Sub UpdateAllFields()
    Dim aStory As Range
    Dim aField As Field
    Dim toC As TableOfContents
    Dim Summary As String   ' argument for the sub "Logging"

    Dim FileName As String:     FileName = C_F_Macros
    Dim ModuleName As String:   ModuleName = C_M_Validation
    Dim MacroName As String:    MacroName = "UpdateAllFields"
    Dim MsgBoxTitle As String:  MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName

    Call MacroBeginning                     ' in module Validation
    Call AddLastCursorPositionBookmark      ' in module Validation

    UpdateAllFields_Form.ProgressLabel = "Macro UpdateAllFields is running..."
    UpdateAllFields_Form.ProgressLabel.font = "Consolas"
    UpdateAllFields_Form.Show vbModeless ' sets ShowModal to False in the corresponding Form
    ' The DoEvents function in Visual Basic for Applications (VBA) for Microsoft Word is used to yield execution so that the operating system can process other events. This function allows the operating system to handle other tasks, such as updating the screen, responding to user inputs, or processing other events in the queue, while your macro is running
    DoEvents
   
    For Each aStory In ActiveDocument.StoryRanges
        aStory.Fields.Update
        UpdateAllFields_Form.ProgressLabel = "Document fields content update..."
        ' The DoEvents function in Visual Basic for Applications (VBA) for Microsoft Word is used to yield execution so that the operating system can process other events. This function allows the operating system to handle other tasks, such as updating the screen, responding to user inputs, or processing other events in the queue, while your macro is running
        DoEvents
    Next aStory
    
    ' surprisingly the following loop do not update the fields in headers and footers
    For Each toC In ActiveDocument.TablesOfContents
        toC.Update
        UpdateAllFields_Form.ProgressLabel = "Table of Contents (ToCs) content update"
        ' The DoEvents function in Visual Basic for Applications (VBA) for Microsoft Word is used to yield execution so that the operating system can process other events. This function allows the operating system to handle other tasks, such as updating the screen, responding to user inputs, or processing other events in the queue, while your macro is running
        DoEvents
    Next toC
    
    ' end of macro
    Call MacroFinish                        ' in module Validation
    Call RemoveLastCursorPositionBookmark   ' in module Validation
    Call UpdateHeadersFootersSub            ' in module Validation
    Unload UpdateAllFields_Form
        
    MsgBox _
        Prompt:="Finished processing.", _
        Buttons:=vbInformation, _
        Title:=MsgBoxTitle
End Sub

' 2025-10-05 by ms
' Checks content against 'reference zero' or 'reference error'. Loggs information into the dedicated file. If error was found it exits. Thanks to that user may fix an issue.
' This function is called in the Scenarios module
Function CheckFieldsAgainstErrors() As Boolean
    CheckFieldsAgainstErrors = True         ' by default everything is fine
    Dim Summary As String
    Dim Flag_CheckRefZero As Boolean
    Flag_CheckRefZero = CheckRefZero        ' in module Validation
    Dim Flag_CheckRefError As Boolean
    Flag_CheckRefError = CheckRefError      ' in module Validation
    
    Dim FileName As String:     FileName = C_F_Macros
    Dim ModuleName As String:   ModuleName = C_M_Validation
    Dim MacroName As String:    MacroName = "CheckFieldsAgainstErrors"
    Dim MsgBoxTitle As String:  MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    If Flag_CheckRefZero = False Then
        Summary = "Found error: zero reference." & vbNewLine & vbNewLine & "It is highligted with yellow color in the content."
        MsgBox _
            Prompt:=Summary, _
            Buttons:=vbCritical, _
            Title:=MsgBoxTitle
        Summary = Summary & vbNewLine & "Finished processing with error: zero reference."
        Call Logging(Summary, MsgBoxTitle)      ' in module Validation
        CheckFieldsAgainstErrors = False
        Exit Function
    End If
    
    If Flag_CheckRefError = False Then
        Summary = "Found error: error reference." & vbNewLine & vbNewLine & "It is highligted with yellow color in the content."
        MsgBox _
            Prompt:=Summary, _
            Buttons:=vbCritical, _
            Title:=MsgBoxTitle
        Summary = Summary & vbNewLine & "Finished processing with error: error reference."
        Call Logging(Summary, MsgBoxTitle)      ' in module Validation
        CheckFieldsAgainstErrors = False
        Exit Function
    End If
    
    Summary = "Either no zero reference, either no reference error was found."
    MsgBox _
        Prompt:=Summary, _
        Buttons:=vbInformation, _
        Title:=MsgBoxTitle
    Call Logging(Summary, MsgBoxTitle)      ' in module Validation
End Function

Private Sub RemoveLastCursorPositionBookmark()
    If ActiveDocument.Bookmarks.Exists(C_BM_LastCursorPosition) Then
        Selection.GoTo What:=wdGoToBookmark, Name:=C_BM_LastCursorPosition
        ActiveDocument.Bookmarks(C_BM_LastCursorPosition).Delete
    Else
        ActiveDocument.GoTo wdStory ' it moves the selection (or cursor) to the very beginning of the document.
    End If
End Sub

Private Sub AddLastCursorPositionBookmark()
    ' Adds a bookmark in place where cursor is present
    If ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument Then _
        Selection.Bookmarks.Add (C_BM_LastCursorPosition)
End Sub

Sub MacroBeginning()
    ' Switch to Print Layout view
    ActiveWindow.View.Type = wdPrintView
    Application.ScreenUpdating = False
    ' When Application.ScreenUpdating is set to False, it turns off screen updating, which can significantly speed up the execution of a macro by preventing the screen from refreshing until the macro has finished running. This is particularly useful for macros that perform a lot of operations, as it reduces the time spent on rendering the screen.
    Application.DisplayAlerts = wdAlertsNone
End Sub

Sub MacroFinish()
    ActiveWindow.View.Type = wdPrintView
    Application.DisplayAlerts = wdAlertsAll
    Application.ScreenUpdating = True
End Sub

' Update only fields in headers and footers
' 2025-02-11 by ms
Private Sub UpdateHeadersFootersSub()
    Dim i As Integer

    For i = 1 To ActiveDocument.Sections.count
        With ActiveDocument.Sections(i)
            .Headers(wdHeaderFooterPrimary).Range.Fields.Update
            .Headers(wdHeaderFooterFirstPage).Range.Fields.Update
            .Footers(wdHeaderFooterPrimary).Range.Fields.Update
            .Footers(wdHeaderFooterFirstPage).Range.Fields.Update
        End With
    Next
End Sub

' Each reference is changed to style hyperlink.
' Future: apply style from Theme to hyperlinks.
' 2025-02-10 by ms
Private Sub RefToHyperlinks()
    Dim oSource As Document
    Set oSource = ActiveDocument
    Dim i As Long, j As Long
    Dim aField As Field
    
    Dim FileName As String:     FileName = C_F_Macros
    Dim ModuleName As String:   ModuleName = C_M_Validation
    Dim MacroName As String:    MacroName = "RefToHyperlinks"
    Dim MsgBoxTitle As String:  MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    Call CheckMicrosoftWordVersion(MacroName)
    
    Call MacroBeginning
    Call AddLastCursorPositionBookmark
    
    i = ActiveDocument.Fields.count
    j = 0
    RefToHyperlinks_Form.ProgressLabel.font = "Consolas"
    RefToHyperlinks_Form.ProgressLabel = "Finished: " & j & " out of " & i
    RefToHyperlinks_Form.Show vbModeless ' this means ShowModal is set to False in the corresponding Form
    For Each aField In oSource.Fields
        Call RefFormatToHyperlink(aField)   ' in module Tools
        ' only the numbering in ToC is shown as a hyperlink (wdFieldPageRef); alternative: wdFieldHyperlink
        If (aField.Type = wdFieldPageRef) Then
            aField.Select
            Selection.font.Underline = wdUnderlineSingle
            Selection.font.color = RGB(0, 130, 180) ' Surprisingly this doesn't work: Selection.font.color = wdThemeColorHyperlink
        End If
        
        j = j + 1
        RefToHyperlinks_Form.ProgressLabel = "Finished: " & j & " out of " & i
        ' The DoEvents function in Visual Basic for Applications (VBA) for Microsoft Word is used to yield execution so that the operating system can process other events. This function allows the operating system to handle other tasks, such as updating the screen, responding to user inputs, or processing other events in the queue, while your macro is running
        DoEvents
    Next aField
    
    Unload RefToHyperlinks_Form
    Call MacroFinish
    Call RemoveLastCursorPositionBookmark
    
    Call Logging("Finished processing.", MsgBoxTitle)
    
    MsgBox _
        Prompt:="Finished processing.", _
        Buttons:=vbInformation, _
        Title:=MsgBoxTitle
        
    ' Clear object variables
    Set oSource = Nothing
End Sub

Sub CheckMicrosoftWordVersion(MacroName As String)
    If Application.Version <> "14.0" And Application.Version <> "16.0" Then
        MsgBox _
            Prompt:="This macro is not compatible to this version of Office!", _
            Buttons:=vbCritical, _
            Title:=MacroName
        Exit Sub ' Exit the subroutine
    End If
End Sub

Sub Tables_CheckNestingLevel()
    Dim oTbl As Table
    Dim oCell As Cell
    Dim tblCount As Long
    Dim nestedTblCount As Long
    Dim highlightedTblCount As Long

    Dim FileName As String:     FileName = C_F_Macros
    Dim ModuleName As String:   ModuleName = C_M_Validation
    Dim MacroName As String:    MacroName = "Tables_CheckNestingLevel"
    Dim MsgBoxTitle As String:  MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName

    tblCount = 0
    nestedTblCount = 0
    highlightedTblCount = 0

    For Each oTbl In ActiveDocument.Tables
        tblCount = tblCount + 1
        If oTbl.NestingLevel > 1 Then
            nestedTblCount = nestedTblCount + 1
            oTbl.Range.HighlightColorIndex = wdRed
            Dim UserDecision As VbMsgBoxResult
            Beep
            UserDecision = MsgBox( _
                Prompt:="Table at index " & tblCount & " has a nesting level greater than 1. Do you want to leave it highlighted?", _
                Buttons:=vbYesNo, _
                Title:=MsgBoxTitle)
            If UserDecision = vbYes Then
                highlightedTblCount = highlightedTblCount + 1
            Else
                oTbl.Range.HighlightColorIndex = wdNoHighlight
            End If
        End If
    Next oTbl

    MsgBox _
        Prompt:="Summary:" & vbCrLf & _
           "Total tables: " & tblCount & vbCrLf & _
           "Tables with nesting level > 1: " & nestedTblCount & vbCrLf & _
           "Tables left highlighted: " & highlightedTblCount, _
        Buttons:=vbInformation, _
        Title:=MsgBoxTitle
End Sub

' Format all tables in the current document upon asking user for decision:
' 1. set the first row to repeat as header row,
' 2. fit between available text borders,
' 3. text in a table rows are not allowed to split across a page break
' 4. center cells vertically
' future: check if in this document there are tables in the tables.
'
' 2025-02-27 by ms and AI
' 2025-03-05 by ms
' 2025-12-05 by ms
Sub Tables_Format()
    Dim tbl As Table
    Dim UserDecision As VbMsgBoxResult
    Dim YesHeaderCount As Integer
    Dim YesAutoFitWindow As Integer
    Dim YesAllowBreakAcrossPages As Integer
    Dim YesCenterCellsVertically As Integer
    Dim TotalNoTables As Integer
    Dim Summary As String
    
    Dim FileName As String:     FileName = C_F_Macros
    Dim ModuleName As String:   ModuleName = C_M_Validation
    Dim MacroName As String:    MacroName = "Tables_Format"
    Dim MsgBoxTitle As String:  MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    Call CheckMicrosoftWordVersion(MacroName)

    TotalNoTables = ActiveDocument.Tables.count
    If TotalNoTables = 0 Then
        MsgBox _
            Prompt:="No tables were found in this document." & vbNewLine & vbNewLine & "Exiting.", _
            Buttons:=vbExclamation, _
            Title:=MsgBoxTitle
        Exit Sub
    End If

    ' Macro to set last position of the cursor
    Call AddLastCursorPositionBookmark
    
    ' Switch to Print View mode of operation
    ActiveWindow.View.Type = wdPrintView
    
    ' Update the screen on time when this macro runs
    Application.ScreenUpdating = True
    
    ' Move the selection to the beginning of the document
    Selection.HomeKey wdStory

    ' Initialization of counters
    YesHeaderCount = 0
    YesAutoFitWindow = 0
    YesAllowBreakAcrossPages = 0
    YesCenterCellsVertically = 0

    For Each tbl In ActiveDocument.Tables
        DoEvents    ' Force a screen refresh
        YesHeaderCount = YesHeaderCount + Table_RepeatHeaderRows(tbl:=tbl, MsgBoxHeader:=MsgBoxTitle)                           ' 1. set the first row to repeat as header row
        YesAutoFitWindow = YesAutoFitWindow + Table_AutoFitWindow(tbl:=tbl, MsgBoxHeader:=MsgBoxTitle)                          ' 2. fit between available text borders,
        YesAllowBreakAcrossPages = YesAllowBreakAcrossPages + Table_AllowBreakAcrossPages(tbl:=tbl, MsgBoxHeader:=MsgBoxTitle)  ' 3. text in a table rows are not allowed to split across a page break
        YesCenterCellsVertically = YesCenterCellsVertically + Table_CenterCellsVertically(tbl:=tbl, MsgBoxHeader:=MsgBoxTitle)  ' 4. Center cells vertically
    Next tbl
    
    Call RemoveLastCursorPositionBookmark
    
     Summary = "Total number of tables: " & TotalNoTables & vbNewLine & _
              "Number of tables with header row set: " & YesHeaderCount & vbNewLine & _
              "Number of tables with set auto fit to text borders: " & YesAutoFitWindow & vbNewLine & _
              "Number of tables with set prohibited breaking the content across pages: " & YesAllowBreakAcrossPages & vbNewLine & _
              "Number of tables with cells centered vertically: " & YesCenterCellsVertically & vbNewLine & _
              "Finished processing."
    Call Logging(Summary:=Summary, MsgBoxHeader:=MsgBoxTitle)  ' this module
    MsgBox _
        Prompt:=Summary, _
        Buttons:=vbInformation, _
        Title:=MsgBoxTitle
End Sub

' Checks and sets breaking a table content across pages.
' 2025-03-05 by ms
Private Function Table_AllowBreakAcrossPages(tbl As Table, MsgBoxHeader As String) As Integer
    
    If tbl.Rows.AllowBreakAcrossPages = True Then
        tbl.Range.HighlightColorIndex = wdYellow
        tbl.Range.Select
        DoEvents    ' Force a screen refresh

        Dim UserDecision As VbMsgBoxResult
        Beep
        UserDecision = MsgBox( _
            Prompt:="Do you want to prohibit breaking the content across pages for this table?", _
            Buttons:=vbYesNoCancel + vbQuestion, _
            Title:=MsgBoxHeader)
        
        tbl.Range.HighlightColorIndex = wdNoHighlight
        If UserDecision = vbYes Then
            tbl.Rows.AllowBreakAcrossPages = False
            Table_AllowBreakAcrossPages = 1
        ElseIf UserDecision = vbNo Then
            Table_AllowBreakAcrossPages = 0
        ElseIf UserDecision = vbCancel Then
            Table_AllowBreakAcrossPages = 0
        End If
    End If
End Function

' Sets a table autofit to window.
' It is not possible to ask first about AutoFit type of a specific table.
' 2025-03-05 by ms
Private Function Table_AutoFitWindow(tbl As Table, MsgBoxHeader As String) As Integer

    tbl.Range.HighlightColorIndex = wdYellow
    tbl.Range.Select
    DoEvents    ' Force a screen refresh
    
    Dim UserDecision As VbMsgBoxResult
    Beep
    UserDecision = MsgBox( _
        Prompt:="Do you want to autofit table to page borders?", _
        Buttons:=vbYesNoCancel + vbQuestion + vbDefaultButton1, _
        Title:=MsgBoxHeader)
        
    tbl.Range.HighlightColorIndex = wdNoHighlight
    If UserDecision = vbYes Then
        tbl.AutoFitBehavior wdAutoFitWindow ' set the "AutoFit Window" feature for tables in your VBA code
        tbl.AllowAutoFit = False
        Table_AutoFitWindow = 1
    ElseIf UserDecision = vbNo Then
        Table_AutoFitWindow = 0
    ElseIf UserDecision = vbCancel Then
        Table_AutoFitWindow = 0
    End If

End Function

' 2025-12-05 by ms
Private Function Table_CenterCellsVertically(tbl As Table, MsgBoxHeader As String) As Integer
    tbl.Range.HighlightColorIndex = wdYellow
    tbl.Range.Select
    DoEvents        ' Force a screen refresh
    
    Dim UserDecision As VbMsgBoxResult
    Beep
    UserDecision = MsgBox( _
        Prompt:="Do you want to center content of all cells vertically?", _
        Buttons:=vbYesNoCancel + vbQuestion + vbDefaultButton1, _
        Title:=MsgBoxHeader)
    tbl.Range.HighlightColorIndex = wdNoHighlight
    If UserDecision = vbYes Then
        tbl.Range.Cells.VerticalAlignment = wdCellAlignVerticalCenter
    End If
End Function

' Checks and sets the HeadingFormat to the table.
' Trick is it is enough to check the very first cell in a table and set HeadingFormat to it.
' 2025-03-05 by ms
Private Function Table_RepeatHeaderRows(tbl As Table, MsgBoxHeader As String) As Integer
    
    If tbl.Range.Cells(1).Range.Rows.HeadingFormat = False Then
        tbl.Range.HighlightColorIndex = wdYellow
        tbl.Range.Select
        DoEvents    ' Force a screen refresh
        
        Dim UserDecision As VbMsgBoxResult
        Beep
        UserDecision = MsgBox( _
            Prompt:="Do you want to set the first row to repeat as header row for this table?", _
            Buttons:=vbYesNoCancel + vbQuestion + vbDefaultButton1, _
            Title:=MsgBoxHeader)
        
        tbl.Range.HighlightColorIndex = wdNoHighlight
        If UserDecision = vbYes Then
            tbl.Range.Cells(1).Range.Rows.HeadingFormat = True
            Table_RepeatHeaderRows = 1
        ElseIf UserDecision = vbNo Then
            Table_RepeatHeaderRows = 0
        ElseIf UserDecision = vbCancel Then
            Table_RepeatHeaderRows = 0
        End If
    End If
End Function

' Replaces ordinary space with nobreakspace in labels of legends: picture(s) and table(s). Works only for specific styles, specific labels in specific template.
' Styles:"Legend table ms" and "Legend picture ms"
' Labels: Pic. X, Tab. X
' Names of styles and labels are hardcoded.
' 2025-02-11 by ms
Sub InsertNoBrakeSpace()
    Dim toC As TableOfContents

    Dim FileName As String:     FileName = C_F_Macros
    Dim ModuleName As String:   ModuleName = C_M_Validation
    Dim MacroName As String:    MacroName = "InsertNoBrakeSpace"
    Dim MsgBoxTitle As String:  MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName

    Call MacroBeginning
    Call AddLastCursorPositionBookmark

    If StyleExists(C_S_PictureLegend) Or StyleExists(C_S_TableLegend) Then
        ' ^s = non breaking space
        With Selection.Find
            .style = C_S_PictureLegend
            .Text = C_Caption_Pic & " "
            .Replacement.Text = C_Caption_Pic & "^s"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = True
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        On Error Resume Next
        Selection.Find.Execute Replace:=wdReplaceAll
    
        Selection.HomeKey Unit:=wdStory
        ' ^s = non breaking space
        With Selection.Find
            .style = C_S_TableLegend
            .Text = C_Caption_Tab & " "
            .Replacement.Text = C_Caption_Tab & "^s"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = True
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        On Error Resume Next
        Selection.Find.Execute Replace:=wdReplaceAll
    
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
    
    ' After update of labels it is worth to update the Table of Content(s)
        For Each toC In ActiveDocument.TablesOfContents
            toC.Update
        Next toC
    End If
    
    Call MacroFinish
    Call RemoveLastCursorPositionBookmark
    
    Call Logging("Finished processing.", MsgBoxTitle)

    MsgBox _
        Prompt:="Finished processing.", _
        Buttons:=vbInformation, _
        Title:=MsgBoxTitle
        
    Call UpdateAllFields
End Sub

' Determines whether or not the target style exists in the active document.
Private Function StyleExists(strStyleName As String) As Boolean
    Dim objStyle As style

    StyleExists = False
    For Each objStyle In ActiveDocument.Styles
        If objStyle.NameLocal = strStyleName Then
            StyleExists = True
            Exit For
        End If
    Next objStyle
End Function

' This is preliminary step for multistep procedure named ModifyReferencesToPicTab().
' Adds zero length bookmarks at the beginning of specified captions 'Pic.' and 'Tab.'.
' It will be used by altered references in the next steps of this procedure.
' 2025-03-01 by ms and AI
Private Sub AddBookmarksToCaptions()
    'Const C_Caption_Pic As String = "Pic."
    'Const C_Caption_Tab As String = "Tab."
    'Const C_BM_Picture As String = "ms_picture_"

    Dim doc As Document
    Dim para As Paragraph
    Dim fld As Field
    Dim rng As Range
    Dim BookmarkName As String
    Dim bookmarkIndex As Long
    Dim bookmarkCount As Long
    
    Dim FileName As String:     FileName = C_F_Macros
    Dim ModuleName As String:   ModuleName = C_M_Validation
    Dim MacroName As String:    MacroName = "AddBookmarksToCaptions"
    Dim MsgBoxTitle As String:  MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    Set doc = ActiveDocument
    bookmarkIndex = 1
    bookmarkCount = 0
    
    For Each para In doc.Paragraphs
        If Left(para.Range.Text, 4) = C_Caption_Pic Or Left(para.Range.Text, 4) = C_Caption_Tab Then
            For Each fld In para.Range.Fields
                If fld.Type = wdFieldSequence Then
                    Set rng = fld.Code
                    rng.MoveStart wdCharacter, -1
                    rng.MoveEnd wdCharacter, 1
                    BookmarkName = C_BM_Picture & bookmarkIndex
                    doc.Bookmarks.Add Name:=BookmarkName, Range:=fld.Code
                    bookmarkIndex = bookmarkIndex + 1
                    bookmarkCount = bookmarkCount + 1
                End If
            Next fld
        End If
    Next para
    
    MsgBox _
        Prompt:=bookmarkCount & " bookmarks added to all captions starting with " & C_Caption_Pic & " or " & C_Caption_Tab & ".", _
        Buttons:=vbInformation, _
        Title:=MsgBoxTitle
    
    ' Clear object variables
    Set doc = Nothing
    Set rng = Nothing
End Sub

' 2025-03-01 by ms and AI
Private Sub DeleteBookmarksFromCaptions()
    'Const C_BM_Picture As String = "ms_picture_"
    Dim doc As Document
    Dim bookmark As bookmark
    Dim BookmarkName As String
    Dim bookmarkCount As Long
    
    Dim FileName As String:     FileName = C_F_Macros
    Dim ModuleName As String:   ModuleName = C_M_Validation
    Dim MacroName As String:    MacroName = "DeleteBookmarksFromCaptions"
    Dim MsgBoxTitle As String:  MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    Set doc = ActiveDocument
    bookmarkCount = 0
    
    For Each bookmark In doc.Bookmarks
        If Left(bookmark.Name, 11) = C_BM_Picture Then
            bookmark.Delete
            bookmarkCount = bookmarkCount + 1
        End If
    Next bookmark
    
    MsgBox _
        Prompt:=bookmarkCount & " bookmarks removed.", _
        Buttons:=vbInformation, _
        Title:=MsgBoxTitle
    
    ' Clear object variables
    Set doc = Nothing
End Sub

' DoEvents    ' Force a screen refresh
' Change the reference fields to the format \# "0" \h \* CHARFORMAT
' 2025-03-01 by ms and AI
Private Sub ModifyRefFields()
    'Const C_S_CharHidden As String = "CharHidden ms"
    Dim doc As Document
    Dim fld As Field
    Dim rng As Range
    Dim foundCount As Long
    Dim alteredCount As Long
    Dim startPos As Long
    Dim endPos As Long
    Dim refPart As String
    Dim firstFiveChars As String
    
    Dim FileName As String:     FileName = C_F_Macros
    Dim ModuleName As String:   ModuleName = C_M_Validation
    Dim MacroName As String:    MacroName = "ModifyRefFields"
    Dim MsgBoxTitle As String:  MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    Set doc = ActiveDocument
    foundCount = 0
    alteredCount = 0
    
    ' Iterate through all fields in the document body
    For Each fld In doc.Fields
        If InStr(fld.Code.Text, "_Ref") > 0 Then
            foundCount = foundCount + 1
            
            ' Move selection to the field and select it
            fld.Select
            Set rng = Selection.Range
            
            DoEvents    ' Force a screen refresh
            ' Highlight the text of the document that represents the field in yellow
            rng.HighlightColorIndex = wdYellow
            DoEvents    ' Force a screen refresh
            
            ' Ask user for decision
            Dim UserDecision As VbMsgBoxResult
            Beep
            UserDecision = MsgBox( _
                Prompt:="Field found: " & fld.Code.Text & vbCrLf & _
                    "Do you want to change the formatting of this field?", _
                Buttons:=vbYesNoCancel + vbQuestion, _
                Title:=MsgBoxTitle)
            
            If UserDecision = vbYes Then
                 
                 ' Extract the part of the field code that should not be removed
                firstFiveChars = Left(fld.result.Text, 5)
                
                ' Insert the first 5 characters in front of the field text and apply style "HiddenText ms"
                ' Thanks to hidden text it will be for user easier to decide which label should be inserted in the front of field reference number during the next step.
                rng.Collapse Direction:=wdCollapseStart
                rng.InsertBefore firstFiveChars
                rng.style = C_S_CharHidden
               
                startPos = InStr(fld.Code.Text, "_Ref")
                endPos = InStr(startPos, fld.Code.Text, " ")
                If endPos = 0 Then
                    endPos = Len(fld.Code.Text)
                End If
                refPart = Mid(fld.Code.Text, startPos, endPos - startPos)
                
                ' Change field parameters while preserving the refPart
                fld.Code.Text = "REF " & refPart & " \# ""0"" \h \* CHARFORMAT"
                alteredCount = alteredCount + 1
                
                ' Remove yellow highlighting from the text of the document that represents the field
                rng.HighlightColorIndex = wdNoHighlight
                DoEvents    ' Force a screen refresh
            ElseIf UserDecision = vbNo Then
                ' Remove yellow highlighting from the text of the document that represents the field
                rng.HighlightColorIndex = wdNoHighlight
                DoEvents    ' Force a screen refresh
            ElseIf UserDecision = vbCancel Then
                ' Cancel the procedure
                DoEvents    ' Force a screen refresh
                MsgBox _
                    Prompt:="Procedure cancelled. " & foundCount & " fields found, " & alteredCount & " fields altered.", _
                    Buttons:=vbInformation, _
                    Title:=MsgBoxTitle
                Exit Sub
            End If
        End If
    Next fld
    
    ' Display summary
    DoEvents    ' Force a screen refresh
    MsgBox _
        Prompt:=foundCount & " fields found, " & alteredCount & " fields altered.", _
        Buttons:=vbInformation, _
        Title:=MsgBoxTitle
        
    ' Update all fields
    If alteredCount > 0 Then
        Beep
        UserDecision = MsgBox( _
            Prompt:="Now it is strongly recommended to run the macro ""UpdateAllFields""" & vbNewLine & _
                "Do you want to continue?", _
            Buttons:=vbOKCancel + vbQuestion, _
            Title:=MsgBoxTitle)
        If UserDecision = vbOK Then
            UpdateAllFields
        ElseIf UserDecision = vbCancel Then
            Exit Sub
        End If
    End If
    
    ' Clear object variables
    Set doc = Nothing
    Set rng = Nothing
End Sub

' Restores previously modified fields to nominal condition.
' 2025-03-01 by ms
Private Sub RestoreModifiedRefFields()
    Dim doc As Document
    Dim fld As Field
    Dim startPos As Long
    Dim endPos As Long
    Dim refPart As String
    
    Dim FileName As String:     FileName = C_F_Macros
    Dim ModuleName As String:   ModuleName = C_M_Validation
    Dim MacroName As String:    MacroName = "RestoreModifiedRefFields"
    Dim MsgBoxTitle As String:  MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    Set doc = ActiveDocument
    
    ' Iterate through all fields in the document body
    For Each fld In doc.Fields
        If InStr(fld.Code.Text, "_Ref") > 0 Then
            ' Extract the part of the field code that should be kept
            startPos = InStr(fld.Code.Text, "_Ref")
            endPos = InStr(startPos, fld.Code.Text, " ")
            refPart = Mid(fld.Code.Text, startPos, endPos - startPos)
            
            ' Restore the field to nominal state
            fld.Code.Text = "REF " & refPart
        End If
    Next fld
    
    ' Display a message indicating the process is complete
    MsgBox _
        Prompt:="All reference fields have been restored to nominal state.", _
        Buttons:=vbInformation, _
        Title:=MsgBoxTitle
        
    ' Clear object variables
    Set doc = Nothing
End Sub

' Adds a proper form caption.
' 2025-03-01 by ms
Private Sub AddModifiedRefCaption()
'    Const C_Caption_Pic As String = "Pic."
'    Const C_Caption_Tab As String = "Tab."
'    Const C_Caption_PicSmall As String = "pic."
'    Const C_Caption_TabSmall As String = "tab."
    Dim doc As Document
    Dim fld As Field
    Dim rng As Range
    Dim foundCount As Long
    Dim currentCount As Long
    Dim InsertText As String
    Dim UserForm As Object
    
    Dim FileName As String:     FileName = C_F_Macros
    Dim ModuleName As String:   ModuleName = C_M_Validation
    Dim MacroName As String:    MacroName = "AddModifiedRefCaption"
    Dim MsgBoxTitle As String:  MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    Call AddLastCursorPositionBookmark
    
    Set doc = ActiveDocument
    foundCount = 0
    currentCount = 0
    
    ' Count the total number of fields that meet the criteria
    For Each fld In doc.Fields
        If InStr(fld.Code.Text, "_Ref") > 0 Then
            foundCount = foundCount + 1
        End If
    Next fld
    
    ' Iterate through all fields in the document body
    For Each fld In doc.Fields
        If InStr(fld.Code.Text, "_Ref") > 0 Then
            currentCount = currentCount + 1
            
            ' Highlight the field in yellow and move cursor to it
            fld.Select
            Set rng = Selection.Range
            DoEvents    ' Force a screen refresh
            rng.HighlightColorIndex = wdYellow
            DoEvents    ' Force a screen refresh
            
            ' Ask user for decision
            Dim UserDecision As VbMsgBoxResult
            Beep
            UserDecision = MsgBox( _
                Prompt:="Total fields found: " & foundCount & vbCrLf & _
                    "Current field number: " & currentCount & vbCrLf & _
                    "Do you want to insert prefix in front of this field?", _
                Buttons:=vbYesNoCancel + vbQuestion + vbDefaultButton2, _
                Title:=MsgBoxTitle)

            ' Remove highlighting based on user response
            rng.HighlightColorIndex = wdNoHighlight

            If UserDecision = vbYes Then
                ' Show the UserForm
                Set UserForm = New SelectCaption_Form
                UserForm.Show
                
                ' Get the selected text from the OptionButtons
                If UserForm.OB_PicSmall.Value = True Then
                    InsertText = C_Caption_PicSmall & Chr(160)  ' Chr(160) = non-breaking space
                ElseIf UserForm.OB_PicCapital.Value = True Then
                    InsertText = C_Caption_Pic & Chr(160)  ' Chr(160) = non-breaking space
                ElseIf UserForm.OB_TabSmall.Value = True Then
                    InsertText = C_Caption_TabSmall & Chr(160)  ' Chr(160) = non-breaking space
                ElseIf UserForm.OB_TabCapital.Value = True Then
                    InsertText = C_Caption_Tab & Chr(160)  ' Chr(160) = non-breaking space
                End If
                
                ' Move selection to the field and insert the text in front of it
                fld.Select
                Set rng = Selection.Range
                rng.Collapse Direction:=wdCollapseStart
                rng.InsertBefore InsertText
                
                ' Apply default character style to inserted text
                rng.SetRange Start:=rng.Start, End:=rng.Start + Len(InsertText)
                rng.style = ActiveDocument.Styles(wdStyleDefaultParagraphFont)
                
            ElseIf UserDecision = vbCancel Then
                ' Cancel the procedure
                MsgBox _
                    Prompt:="Procedure cancelled.", _
                    Buttons:=vbInformation, _
                    Title:=MsgBoxTitle
                Exit Sub
            End If
        End If
    Next fld
    
    Call RemoveLastCursorPositionBookmark
    DoEvents    ' Force a screen refresh
    ' Display a message indicating the process is complete
    MsgBox _
        Prompt:="Captions have been inserted for all selected fields.", _
        Buttons:=vbInformation, _
        Title:=MsgBoxTitle
    
    ' Clear object variables
    Set doc = Nothing
    Set rng = Nothing
    Set UserForm = Nothing
End Sub

' The last function in the row, which replaces text preceeding a reference value with hiperlink to a bookmark in a caption.
' 2025-03-01 by ms and AI
Private Sub AddHyperlinkToModRefFields()
'Const C_Caption_Pic As String = "Pic."
'Const C_Caption_Tab As String = "Tab."
'Const C_Caption_PicSmall As String = "pic."
'Const C_Caption_TabSmall As String = "tab."
    Dim doc As Document
    Dim fld As Field
    Dim rng As Range
    Dim Bm As bookmark
    Dim bmName As String
    Dim bmNumber As String
    Dim fldNumber As String
    Dim InsertText As String
    Dim CaptionStringLength As Long
    
    Dim FileName As String:     FileName = C_F_Macros
    Dim ModuleName As String:   ModuleName = C_M_Validation
    Dim MacroName As String:    MacroName = "AddHyperlinkToModRefFields"
    Dim MsgBoxTitle As String:  MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    Set doc = ActiveDocument
    ' Determine length of the caption
    CaptionStringLength = Len(C_Caption_Pic) + 1 ' + 1 = space
    
    ' Iterate through all fields in the document body
    For Each fld In doc.Fields
        If InStr(fld.Code.Text, "_Ref") > 0 Then
            ' Check if in front of the field there is a text string "pic. " or "Pic. "
            Set rng = fld.Code
            rng.MoveStart wdCharacter, -CaptionStringLength - 1 ' Move the range to include the preceding text
            
            ' Get the 5 characters of document body text preceding the field
            InsertText = Left(rng.Text, CaptionStringLength)
            
            ' Chr(160) = non-breaking space character
            If InsertText = C_Caption_PicSmall & Chr(160) Or InsertText = C_Caption_Pic & Chr(160) _
                Or InsertText = C_Caption_TabSmall & Chr(160) Or InsertText = C_Caption_Tab & Chr(160) Then
                ' Extract the number at the end of the visible text field
                fldNumber = Right(fld.result.Text, Len(fld.result.Text) - InStrRev(fld.result.Text, "_"))
                
                ' Search the collection of available bookmarks
                For Each Bm In doc.Bookmarks
                    bmName = Bm.Name
                    
                    ' Extract the number at the end of the bookmark name
                    bmNumber = Right(bmName, Len(bmName) - InStrRev(bmName, "_"))
                    
                    ' Compare the numbers and add a hyperlink field if they match
                    If fldNumber = bmNumber Then
                        ' reduces the range to a single point at the beginning of the original range (field code)
                        rng.Collapse Direction:=wdCollapseStart
                        rng.End = rng.Start + CaptionStringLength
                        ActiveDocument.Hyperlinks.Add Anchor:=rng, Address:="", SubAddress:=bmName, TextToDisplay:=InsertText
                        Exit For
                    End If
                Next Bm
            End If
        End If
    Next fld
    
    ' Display a message indicating the process is complete
    MsgBox _
        Prompt:="Hyperlinks have been added to all matching reference fields.", _
        Buttons:=vbInformation, _
        Title:=MsgBoxTitle
    
    ' Clear object variables
    Set doc = Nothing
    Set rng = Nothing
End Sub

' Removes hyperlinks dedicated for Pic. or pic.
' 2025-03-01 by ms and AI
Private Sub DeleteHyperlinksFromModRefFields()
    'Const C_BM_Picture As String = "ms_picture_"
    Dim doc As Document
    Dim fld As Field
    Dim rng As Range
    Dim hyperlinkText As String
    
    Dim FileName As String:     FileName = C_F_Macros
    Dim ModuleName As String:   ModuleName = C_M_Validation
    Dim MacroName As String:    MacroName = "DeleteHyperlinksFromModRefFields"
    Dim MsgBoxTitle As String:  MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    Set doc = ActiveDocument
    
    ' Iterate through all fields in the document body
    For Each fld In doc.Fields
        If fld.Type = wdFieldHyperlink Then
            ' Check if the hyperlink points to a bookmark starting with "picture_"
            If InStr(fld.Code.Text, "HYPERLINK") And InStr(fld.Code.Text, "\l " & C_BM_Picture) > 0 Then
                ' Store the original text of the hyperlink field
                hyperlinkText = fld.result.Text
                
                ' Move the range to the field code and collapse it to the start
                Set rng = fld.Code
                rng.Collapse Direction:=wdCollapseStart
                
                ' Insert the original text before the field code
                rng.InsertBefore hyperlinkText
                                
                ' Delete the hyperlink field
                fld.Delete
                
                ' Insert the original text after deleting the field
                rng.InsertAfter hyperlinkText
            End If
        End If
    Next fld
    
    ' Display a message indicating the process is complete
    MsgBox _
        Prompt:="Hyperlink fields have been replaced with their original text.", _
        Buttons:=vbInformation, _
        Title:=MsgBoxTitle
    
    ' Clear object variables
    Set doc = Nothing
    Set rng = Nothing
End Sub

' Resets to default character formatting all fields containig the "_Ref" text string.
' This macro is useful to inverse formatting set by macro RefToHyperlinks.
' 2025-03-01 by ms and AI
Private Sub ResetRefToHyperlinks()
    Dim doc As Document
    Dim fld As Field
    
    Dim FileName As String:     FileName = C_F_Macros
    Dim ModuleName As String:   ModuleName = C_M_Validation
    Dim MacroName As String:    MacroName = "ResetRefToHyperlinks"
    Dim MsgBoxTitle As String:  MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    Set doc = ActiveDocument
    
    ' Iterate through all fields in the document body
    For Each fld In doc.Fields
        If InStr(fld.Code.Text, "_Ref") > 0 Then
            ' Change the formatting of the text fld.Result.Text to default character formatting
            fld.result.font.Reset
        End If
    Next fld
    
    ' Display a message indicating the process is complete
    MsgBox _
        Prompt:="Formatting has been reset for all matching fields.", _
        Buttons:=vbInformation, _
        Title:=MsgBoxTitle
    
    ' Clear object variables
    Set doc = Nothing
End Sub

' Auxiliary function to clean up / delete hiddent text inserted earlier by the sub ModifyRefFields.
' 2025-03-03 by ms
Private Sub DeleteHiddenText() ' required in Scenarios -> ModifyReferencesToPicTab
    'Const C_Caption_Pic As String = "Pic."
    'Const C_Caption_Tab As String = "Tab."
    'Const C_Caption_PicSmall As String = "pic."
    'Const C_Caption_TabSmall As String = "tab."
    'Const C_S_CharHidden As String = "CharHidden ms"
    Dim doc As Document
    Dim rng As Range
    Dim searchStrings As Variant
    Dim i As Integer
    
    Dim FileName As String:     FileName = C_F_Macros
    Dim ModuleName As String:   ModuleName = C_M_Validation
    Dim MacroName As String:    MacroName = "DeleteHiddenText"
    Dim MsgBoxTitle As String:  MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    Set doc = ActiveDocument
    searchStrings = Array(C_Caption_Pic & " ", C_Caption_PicSmall & " ", C_Caption_Tab & " ", C_Caption_TabSmall & " ")

    ' Iterate through each search string
    For i = LBound(searchStrings) To UBound(searchStrings)
        Set rng = doc.Content
        With rng.Find
            .ClearFormatting
            .Text = searchStrings(i)
            .style = C_S_CharHidden
            .Replacement.ClearFormatting
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = True
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
            .Execute Replace:=wdReplaceAll
        End With
    Next i

    MsgBox _
        Prompt:="Auxiliary hidden text strings have been removed.", _
        Buttons:=vbInformation, _
        Title:=MsgBoxTitle
    
    ' Clear object variables
    Set doc = Nothing
    Set rng = Nothing
End Sub

' Search document for specific paragraph style
' 2025-03-15 by ms and AI
Sub FindParagraphStyling()
    'Const C_BM_SearchedStyle As String = "ms_SearchedStyle_"
    Dim styleName As String
    Dim StyleExists As Boolean
    Dim para As Paragraph
    Dim NoTotalPar As Integer
    Dim NomsNotCompliantPar As Integer
    Dim NomsCompliantPar As Integer
    Dim summaryMessage As String
    Dim i As Integer
    Dim PerVal As Double
    Dim NoParInTable As Integer
    Dim bookmarkCounter As Integer
    bookmarkCounter = 0

    Dim FileName As String:     FileName = C_F_Macros
    Dim ModuleName As String:   ModuleName = C_M_Validation
    Dim MacroName As String:    MacroName = "FindParagraphStyling"
    Dim MsgBoxTitle As String:  MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName

    ' Ask user for a style name
    Do
        styleName = InputBox("Enter the style name:", MsgBoxTitle)
        If styleName = "" Then Exit Sub ' User canceled

        ' Check if the style exists
        StyleExists = False
        On Error Resume Next
        StyleExists = Not ActiveDocument.Styles(styleName) Is Nothing
        On Error GoTo 0

        If Not StyleExists Then
            Dim UserBMDecision As VbMsgBoxResult
            Beep
            UserBMDecision = MsgBox( _
                Prompt:="Style not found. Do you want to try again?", _
                Buttons:=vbYesNo + vbQuestion, _
                Title:=MsgBoxTitle)
            If UserBMDecision = vbNo Then
                Exit Sub
            End If
        End If
    Loop Until StyleExists

    ' Ask user if they want to insert a bookmark
    Beep
    UserBMDecision = MsgBox( _
        Prompt:="Do you want to add bookmarks in paragraphs with the specified style?", _
        Buttons:=vbYesNo + vbQuestion, _
        Title:=MsgBoxTitle)

    ' Initialize counters
    NoTotalPar = ActiveDocument.Paragraphs.count
    NomsCompliantPar = 0
    NomsNotCompliantPar = 0
    NoParInTable = 0

    Call AddLastCursorPositionBookmark

    ' Initialization of the dedicated Form
    TemplateStyleValidation_Form.Show vbModeless ' sets ShowModal to False in the corresponding Form

    ' Check each paragraph
    For i = 1 To NoTotalPar
        Set para = ActiveDocument.Paragraphs(i)
        If Not para.Range.Information(wdWithInTable) Then   ' All outside the tables
            NoParInTable = NoParInTable + 1
            If para.style.NameLocal = styleName Then
                DoEvents    ' Force a screen refresh
                para.Range.HighlightColorIndex = wdYellow
                DoEvents    ' Force a screen refresh
                NomsNotCompliantPar = NomsNotCompliantPar + 1
                ' Insert a bookmark "ms_SearchedStyle_x" where x is incremented number
                If UserBMDecision = vbYes Then
                    bookmarkCounter = bookmarkCounter + 1
                    para.Range.Bookmarks.Add Name:=C_BM_SearchedStyle & bookmarkCounter
                End If
            Else
                NomsCompliantPar = NomsCompliantPar + 1
            End If
        End If

        ' Update progress label
        PerVal = (i / NoTotalPar) * 100 ' Calculate percentage value
        TemplateStyleValidation_Form.ProgressLabel = "Paragraph counter: " & i & " out of " & NoTotalPar & _
            " (" & Int(PerVal) & "%)" & vbNewLine & _
            "Compliant paragraph counter: " & NomsCompliantPar & vbNewLine & _
            "Non-compliant paragraph counter: " & NomsNotCompliantPar & vbNewLine & _
            "No. paragraphs in tables: " & NoParInTable
        
        ' Allow other events to be processed
        DoEvents
    Next i

    Unload TemplateStyleValidation_Form
    Call RemoveLastCursorPositionBookmark

    ' Display summary
    summaryMessage = "Total number of paragraphs: " & NoTotalPar & vbCrLf & _
                     "Number of paragraphs with the specified style: " & NomsNotCompliantPar & vbCrLf & _
                     "Number of paragraphs without the specified style: " & NomsCompliantPar & vbNewLine & _
                     "Number of not examined paragraphs in tables: " & NoParInTable
    MsgBox _
        Prompt:=summaryMessage, _
        Buttons:=vbInformation, _
        Title:=MsgBoxTitle
    
    ' Clear object variables
    Set para = Nothing
End Sub


Sub FindCharacterStyling()
    'Const C_BM_SearchedStyle As String = "ms_SearchedStyle_"
    Dim styleName As String
    Dim StyleExists As Boolean
    Dim isCharacterStyle As Boolean
    Dim UserDecision As VbMsgBoxResult
    Dim para As Paragraph
    Dim rng As Range
    Dim NoTotalPar As Integer
    Dim NomsNotCompliantPar As Integer
    Dim NomsCompliantPar As Integer
    Dim summaryMessage As String
    Dim i As Integer
    Dim PerVal As Double
    Dim NoParInTable As Integer
    Dim bookmarkCounter As Integer
    bookmarkCounter = 0

    Dim FileName As String:     FileName = C_F_Macros
    Dim ModuleName As String:   ModuleName = C_M_Validation
    Dim MacroName As String:    MacroName = "FindCharacterStyling"
    Dim MsgBoxTitle As String:  MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName

    ' Ask user for a style name
    Do
        styleName = InputBox("Enter the character style name:", MsgBoxTitle)
        If styleName = "" Then Exit Sub ' User canceled

        ' Check if the style exists
        StyleExists = False
        On Error Resume Next
        StyleExists = Not ActiveDocument.Styles(styleName) Is Nothing
        On Error GoTo 0

        ' Check if the style is a character type
        isCharacterStyle = False
        If StyleExists Then
            isCharacterStyle = _
            ActiveDocument.Styles(styleName).Type = wdStyleTypeCharacter _
            Or _
            ActiveDocument.Styles(styleName).Type = wdStyleTypeLinked
        End If

        If Not StyleExists Or Not isCharacterStyle Then
            Beep
            UserDecision = MsgBox( _
                Prompt:="Style not found or not a character style. Do you want to try again?" & vbNewLine & _
                    "Perhaps style name you are entering contains suffix 'Char' or 'Znak'?", _
                Buttons:=vbYesNo + vbQuestion, _
                Title:=MsgBoxTitle)
            If UserDecision = vbNo Then
                Exit Sub
            End If
        End If
    Loop Until StyleExists And isCharacterStyle

    ' Ask user if they want to insert a bookmark
    Beep
    UserDecision = MsgBox( _
        Prompt:="Do you want to add bookmarks in paragraphs with the specified character style?", _
        Buttons:=vbYesNo + vbQuestion, _
        Title:=MsgBoxTitle)

    ' Initialize counters
    NoTotalPar = ActiveDocument.Paragraphs.count
    NomsCompliantPar = 0
    NomsNotCompliantPar = 0
    NoParInTable = 0

    ' Check each paragraph
    For i = 1 To NoTotalPar
        Set para = ActiveDocument.Paragraphs(i)
        If Not para.Range.Information(wdWithInTable) Then   ' All outside the tables
            NoParInTable = NoParInTable + 1
            Set rng = para.Range
            rng.Find.ClearFormatting
            rng.Find.style = styleName
            rng.Find.Forward = True
            rng.Find.Wrap = wdFindStop
            rng.Find.Format = True
            rng.Find.MatchWildcards = False

            Do While rng.Find.Execute
                DoEvents    ' Force a screen refresh
                rng.HighlightColorIndex = wdYellow
                DoEvents    ' Force a screen refresh
                NomsNotCompliantPar = NomsNotCompliantPar + 1
                ' Insert a bookmark "ms_SearchedStyle_x" where x is incremented number
                If UserDecision = vbYes Then
                    bookmarkCounter = bookmarkCounter + 1
                    rng.Bookmarks.Add Name:=C_BM_SearchedStyle & bookmarkCounter
                End If
                rng.Collapse Direction:=wdCollapseEnd
            Loop
        End If

        ' Update progress label
        PerVal = (i / NoTotalPar) * 100 ' Calculate percentage value
        TemplateStyleValidation_Form.ProgressLabel = "Paragraph counter: " & i & " out of " & NoTotalPar & _
            " (" & Int(PerVal) & "%)" & vbNewLine & _
            "Compliant paragraph counter: " & NomsCompliantPar & vbNewLine & _
            "Non-compliant paragraph counter: " & NomsNotCompliantPar & vbNewLine & _
            "No. paragraphs in tables: " & NoParInTable
        
        ' Allow other events to be processed
        DoEvents
    Next i

    ' Display summary
    summaryMessage = "Total number of paragraphs: " & NoTotalPar & vbCrLf & _
                     "Number of paragraphs with the specified character style: " & NomsNotCompliantPar & vbCrLf & _
                     "Number of paragraphs without the specified character style: " & NomsCompliantPar & vbNewLine & _
                     "Number of not examined paragraphs in tables: " & NoParInTable
    MsgBox _
        Prompt:=summaryMessage, _
        Buttons:=vbInformation, _
        Title:=MsgBoxTitle
    
    ' Clear object variables
    Set para = Nothing
    Set rng = Nothing
End Sub


Private Function IsInCollection(col As Collection, item As Variant) As Boolean
    Dim var As Variant
    On Error Resume Next
    For Each var In col
        If var = item Then
            IsInCollection = True
            Exit Function
        End If
    Next var
    IsInCollection = False
End Function

' Manual removal of non-compliant bookmarks. It doesn't remove highlighting!
' This function works for both subs: ShowNonCompliantStylingInParagraphs() and ShowNonComplientStylingInTables()
' 2025-03-06 by ms and AI
Sub SearchNCstylingBookmarks()
'    Const C_BM_NCstylingP As String = "NCstylingP_"
'    Const C_BM_NCstylingT As String = "NCstylingT_"
    Dim Bm As bookmark
    Dim i As Integer
    Dim BookmarkName As String

    Dim FileName As String:     FileName = C_F_Macros
    Dim ModuleName As String:   ModuleName = C_M_Validation
    Dim MacroName As String:    MacroName = "SearchNCstylingBookmarks"
    Dim MsgBoxTitle As String:  MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName

    ' Loop through all bookmarks in the document
    For i = ActiveDocument.Bookmarks.count To 1 Step -1
        Set Bm = ActiveDocument.Bookmarks(i)
        BookmarkName = Bm.Name

        ' Check if the bookmark name starts with "NCstylingT_" or "NCstylingP_"
        If Left(BookmarkName, 11) = C_BM_NCstylingT Or Left(BookmarkName, 11) = C_BM_NCstylingP Then
            ' Check if the rest of the name is a number
            If IsNumeric(Mid(BookmarkName, 12)) Then
                ' Jump to the bookmark
                Bm.Range.Select
                
                ' Ask user if they want to remove the bookmark
                Dim UserDecision As VbMsgBoxResult
                Beep
                UserDecision = MsgBox( _
                    Prompt:="Do you want to remove the non-compliant bookmark '" & BookmarkName & "'?", _
                    Buttons:=vbYesNoCancel + vbQuestion, _
                    Title:=MsgBoxTitle)
                
                ' If user chooses Yes, remove the bookmark
                If UserDecision = vbYes Then
                    Bm.Delete
                End If
                
                If UserDecision = vbCancel Then
                    MsgBox _
                        Prompt:="Processing finished without removing of a bookmark.", _
                        Buttons:=vbInformation, _
                        Title:=MsgBoxTitle
                    Exit Sub
                End If
            End If
        End If
    Next i

    MsgBox _
        Prompt:="Finished processing bookmarks.", _
        Buttons:=vbInformation, _
        Title:=MsgBoxTitle
    
    ' Clear object variables
    Set Bm = Nothing
End Sub

' This is summary macro to widen / modify references to pictures and tables.
' The captions are replaced to written with capital or ordinary letter and point to the same destination.
' Multi step and complex procedure.
'   The summary macro is ModifyReferencesToPicTab (1 ÷ 5)
'   1. AddBookmarksToCaptions()     <->     DeleteBookmarksFromCaptions()
'   2. ModifyRefFields()            <->     RestoreModifiedRefFields()
'   3. AddModifiedRefCaption()
'   4. AddHyperlinkToModRefFields() <->     DeleteHyperlinksFromModRefFields()
'   5. RefToHyperlinks()            <->     ResetRefToHyperlinks()
' 2025-03-03 by ms
Sub ModifyReferencesToPicTab()
    Dim FileName As String:     FileName = C_F_Macros
    Dim ModuleName As String:   ModuleName = C_M_Validation
    Dim MacroName As String:    MacroName = "ModifyReferencesToPicTab"
    Dim MsgBoxTitle As String:  MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    ' Origin module: Validation
    Call AddBookmarksToCaptions
    ' Origin module: Validation
    Call ModifyRefFields
    ' Origin module: Validation
    Call AddModifiedRefCaption
    ' Origin module: Validation
    Call AddHyperlinkToModRefFields
    ' Origin module: Validation
    Call DeleteHiddenText
    
    MsgBox _
        Prompt:="Processing is finished", _
        Buttons:=vbInformation + vbOKOnly, _
        Title:=MsgBoxTitle
End Sub
