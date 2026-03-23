Attribute VB_Name = "Content"
' VBA Module name: Content.bas
' https://github.com/mslonik/Microsoft-Word-Configuration
'
'   License: MIT License.
'
'+----+-----------------------------------------+--------------+----------------+-----------------------------------------+
'| No | Sub name                                | Ribbon name  | Ribbon section | Ribbon button name                      |
'+----+-----------------------------------------+--------------+----------------+-----------------------------------------+
'| 3  | NewFileContent                          | Content_ms | Combos         | NewFileContent                          |
'+----+-----------------------------------------+--------------+----------------+-----------------------------------------+
'
' Subs related to keyboard shortcuts:
'| Ctrl + W  | UpdateAllFieldsAndCloseFile             |
'| Ctrl + F2 | CustomizedPrintPreviewAndPrint          |
'
' Outdated (no longer used):
'| 1  | ApplyDistanceBetweenNumberingAndHeading | Content_ms | Combos         | ApplyDistanceBetweenNumberingAndHeading |
'| 2  | ResetDistanceBetweenNumberingAndHeading | Content_ms | Combos         | ResetDistanceBetweenNumberingAndHeading |
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
Option Explicit
' Declare a module-level variable instead of a constant
Dim BetweenNumberAndText As String

Private Sub InitializeConstants()
    ' Initialize the variable in a subroutine
    BetweenNumberAndText = ChrW(8195) ' ChrW(8195) = em space
End Sub

' In order to work, headings styles 1 ÷ 8 must be correctly setup with zero length space between number and text.
' 2025-03-09 by ms
Sub ResetDistanceBetweenNumberingAndHeading()
    Dim FileName As String:     FileName = C_F_Macros
    Dim ModuleName As String:   ModuleName = C_M_Content
    Dim MacroName As String:    MacroName = "ResetDistanceBetweenNumberingAndHeading"
    Dim MsgBoxTitle As String:  MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    Dim UserDecision As VbMsgBoxResult
    Beep
    UserDecision = MsgBox( _
        Prompt:="To successfully run this function you must to manually change the Microsoft Word configuration." & vbNewLine & _
            "Enter: File -> Options -> Advanced ->  section: Cut, copy and paste -> button Settings" & vbNewLine & _
            "Disable: Adjust sentence and word spacing automatically" & vbNewLine & vbNewLine & _
            "Do you want to continue?", _
        Buttons:=vbQuestion + vbYesNo, _
        Title:=MsgBoxTitle)
    If UserDecision = vbNo Then
        Exit Sub
    End If
    
    Call Macros_ms.Content.InitializeConstants
    Call Macros_ms.StylesM.RemoveTextFromBeginningOfListParagraphs(textToRemove:=BetweenNumberAndText)
End Sub

' This macro is linked to Ctrl + S keyboard shortcut. Each time user runs it, it enters specific text character (em space) at the beginning of each paragraph style type list.
' Then the built-in command Save is run.
' 2025-03-09 by ms and AI
' 2026-01-17 by ms
Sub CustomizedFileSave()
    Dim FileName As String:      FileName = C_F_Macros
    Dim ModuleName As String:    ModuleName = C_M_Content
    Dim MacroName As String:     MacroName = "CustomizedFileSave"
    Dim MsgBoxTitle As String:   MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName

    ' Initialize em space constant as BetweenNumberAndText
'    Call Macros_ms.Content.InitializeConstants
'    Call Macros_ms.StylesM.InsertTextAtBeginningOfListParagraphs(textToInsert:=BetweenNumberAndText)
    
    ' Enable error handling in case that user presses 'Cancel' button.
'    On Error Resume Next
    Call Macros_ms.BuildingBlocks.BB_DeleteHeaderPar
    ' Execute the built-in Save command
    ActiveDocument.Save
    ' This statement turns off the error handling that was set by On Error Resume Next. It restores the default error handling behavior, which means that if an error occurs after this point, VBA will stop execution and display an error message.
'    On Error GoTo 0
    Application.StatusBar = MsgBoxTitle & " > " & "was running..."
    
End Sub

' Update all fields and then close the file.
' Associated to keyboard shortcut Ctrl + W.
' 2025-03-15
Sub UpdateAllFieldsAndCloseFile()
    Call Macros_ms.Validation.UpdateAllFields
    If Not CheckFieldsAgainstErrors Then  ' in module Validation
        Exit Sub                    ' exits if error was found
    End If
    Application.Run "DocClose"      ' call built-in Microsoft Word command
End Sub

' Combo fall forward: update all fields and then show print preview
' 2025-03-15
Sub CustomizedPrintPreviewAndPrint()
    Call Macros_ms.Validation.UpdateAllFields
    Application.CommandBars.ExecuteMso "PrintPreviewAndPrint"   ' call built-in Microsoft Word command
End Sub

' Insert full content: cover page, last page and example content: 3 sections in total.
' 2025-03-19 ms and AI
Sub NewFileContent()
    Dim FileName As String:     FileName = C_F_Macros
    Dim ModuleName As String:   ModuleName = C_M_Content
    Dim MacroName As String:    MacroName = "NewFileContent"
    Dim MsgBoxTitle As String:  MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
            
    ' Check if the add-in template is enabled
    Dim TemplateIndex As Integer
    TemplateIndex = Macros_ms.BuildingBlocks.GetTemplateIndex(C_F_BuildingBlocks)
      
    Dim UserDecision As VbMsgBoxResult
    Dim QuestionCounter As Byte
    QuestionCounter = 0
    Const QuestionTotal As Byte = 10
    
    ' 1. Setting up shortcuts
    QuestionCounter = QuestionCounter + 1
    Beep
    UserDecision = MsgBox( _
        Prompt:=QuestionCounter & "/" & QuestionTotal & " " & _
            "Would you like to add set of customized keyboard shortcuts?" & vbNewLine & vbNewLine & _
            "It is strongly recommended to do that.", _
        Buttons:=vbQuestion + vbYesNo + vbDefaultButton1, _
        Title:=MsgBoxTitle)
    If UserDecision = vbYes Then
        Call Macros_ms.Shortcuts.CreateActiveDocumentMacroShortcuts
    End If

    ' 2. Inserting customized styles
    QuestionCounter = QuestionCounter + 1
    Beep
    UserDecision = MsgBox( _
        Prompt:=QuestionCounter & "/" & QuestionTotal & " " & _
            "Would you like to insert set of customized styles?" & vbNewLine & vbNewLine & _
            "It is strongly recommended to do that.", _
        Buttons:=vbQuestion + vbYesNo + vbDefaultButton1, _
        Title:=MsgBoxTitle)
    If UserDecision = vbYes Then
        Call Macros_ms.StylesM.AddCompliantStyles
    End If

    ' 3. Setting up Theme file
    QuestionCounter = QuestionCounter + 1
    Beep
    UserDecision = MsgBox( _
        Prompt:=QuestionCounter & "/" & QuestionTotal & " " & _
            "Would you like to apply the specific theme " & vbNewLine & vbNewLine & _
            C_F_Theme & "?" & vbNewLine & vbNewLine & _
            "It is strongly recommended to do that.", _
        Buttons:=vbQuestion + vbYesNo + vbDefaultButton1, _
        Title:=MsgBoxTitle)
    If UserDecision = vbYes Then
        Call Macros_ms.Theme.AttachTheme
    End If

    ' 4. Setting up customized Microsoft Word options
    QuestionCounter = QuestionCounter + 1
    Beep
    UserDecision = MsgBox( _
        Prompt:=QuestionCounter & "/" & QuestionTotal & " " & _
            "Would you like to customize Microsoft Word options?", _
        Buttons:=vbYesNo + vbQuestion + vbDefaultButton1, _
        Title:=MsgBoxTitle)
    If UserDecision = vbYes Then
        Call Macros_ms.Tools.WordOptionsCustomize
    End If

    ' 5. Setting of active document margins
    Beep
    QuestionCounter = QuestionCounter + 1
    UserDecision = MsgBox( _
        Prompt:=QuestionCounter & "/" & QuestionTotal & " " & _
            "Would you like to set margins, headers and footers to specific sizes?", _
        Buttons:=vbYesNo + vbQuestion + vbDefaultButton1, _
        Title:=MsgBoxTitle)
    If UserDecision = vbYes Then
        Call Macros_ms.Tools.SetPageLayout_A4_V_1_2
    End If

    ' 6. Setting of active document custom properties
    Dim DocPropertiesFlag As Boolean    ' This flag will be set to true only if user decides to add custom DoC properties.
    DocPropertiesFlag = msoFalse
    Beep
    QuestionCounter = QuestionCounter + 1
    UserDecision = MsgBox( _
        Prompt:=QuestionCounter & "/" & QuestionTotal & " " & _
            "Would you like to set custom document properties?", _
        Buttons:=vbYesNo + vbQuestion + vbDefaultButton1, _
        Title:=MsgBoxTitle)
    If UserDecision = vbYes Then
        Call Macros_ms.Tools.DocPropertiesAddCustom
        DocPropertiesFlag = msoTrue
    End If
    If UserDecision = vbNo Then
        DocPropertiesFlag = msoFalse
    End If

    ' 7. Setting of Microsoft Word customized captions
    Beep
    QuestionCounter = QuestionCounter + 1
    If Not CaptionCheckCustomLabelsOnly() Then
        UserDecision = MsgBox( _
            Prompt:=QuestionCounter & "/" & QuestionTotal & " " & _
                "Would you like to add to Microsoft Word custom captions?" & vbNewLine & vbNewLine & _
                C_Caption_Pic & " and " & C_Caption_Tab, _
            Buttons:=vbYesNo + vbQuestion + vbDefaultButton1, _
            Title:=MsgBoxTitle)
        If UserDecision = vbYes Then
            Call Macros_ms.Tools.CaptionLabelDeleteCustomized
            Call Macros_ms.Tools.CaptionAddCustomized
        End If
    End If

    ' 8. Setting of document hyphenation
    Beep
    QuestionCounter = QuestionCounter + 1
    UserDecision = MsgBox( _
        Prompt:=QuestionCounter & "/" & QuestionTotal & " " & _
            "Would you like to set in this document text hyphenation?", _
        Buttons:=vbYesNo + vbQuestion + vbDefaultButton1, _
        Title:=MsgBoxTitle)
    If UserDecision = vbYes Then
        Call Macros_ms.Tools.SetHyphenation
    End If

    ' 9. Insertion of example content to the current document
    Beep
    QuestionCounter = QuestionCounter + 1
    UserDecision = MsgBox( _
        Prompt:=QuestionCounter & "/" & QuestionTotal & " " & _
            "Would you like to insert full content (Yes)," & _
            "or just basic content (No)," & _
            "or skip this step entirely (Cancel)?", _
        Buttons:=vbYesNoCancel + vbQuestion + vbDefaultButton1, _
        Title:=MsgBoxTitle)
    
    ' If basic content is selected, call the following set of macros
    If UserDecision = vbNo Then
        Call Macros_ms.Content.InsertBasicContent(TemplateIndex)
        Call Macros_ms.BuildingBlocks.BB_DeleteHeaderPar
        Call Macros_ms.Content.CleanHeadersFooters
    End If
    
    If UserDecision = vbCancel Then
        Exit Sub
    End If
        
    If UserDecision = vbYes Then
        If TemplateIndex = 0 Then
            MsgBox _
                Prompt:="Specified template name " & vbNewLine & vbNewLine & _
                    C_F_BuildingBlocks & vbNewLine & vbNewLine & " was not found." & vbNewLine & _
                    "Exiting.", _
                Buttons:=vbExclamation, _
                Title:=MsgBoxTitle
            Exit Sub
        End If
        Call Macros_ms.Content.InsertFullContent(TemplateIndex)
        Call Macros_ms.BuildingBlocks.BB_DeleteHeaderPar
        Call Macros_ms.Content.CleanHeadersFooters
    End If

    ' 10. Set document page background color to customized (grey).
    Beep
    QuestionCounter = QuestionCounter + 1
    UserDecision = MsgBox( _
        Prompt:=QuestionCounter & "/" & QuestionTotal & " " & _
            "Would you like to set document background color to grey?" & vbNewLine & vbNewLine & _
            "If you answer 'No' then page background color will be restored to default one.", _
        Buttons:=vbQuestion + vbDefaultButton1 + vbYesNo, _
        Title:=MsgBoxTitle)
    If UserDecision = vbYes Then
        Call Macros_ms.Tools.SetPageColorToCustom
    End If
    If UserDecision = vbNo Then
        Call Macros_ms.Tools.RestoreDefaultPageColor
    End If
    
    ' The following line is in my opinion bug in Microsoft Word. For unknown reason the C_S_ParNormal is set to AutomaticallyUpdate. So I prevent it silently.
    ActiveDocument.Styles(C_S_ParNormal).AutomaticallyUpdate = False
    
End Sub

' Clean headers and footers from unwanted empty paragraphs
' 2026-03-14 by ms
Sub CleanHeadersFooters()
    Dim sec As Section
    Dim hf As HeaderFooter
    Dim i As Integer
    Dim RemovedCount As Long
    
    Application.ScreenUpdating = False
    RemovedCount = 0
    
    For Each sec In ActiveDocument.Sections
        ' 1 = Primary, 2 = FirstPage, 3 = EvenPages
        For i = 1 To 3
            RemovedCount = RemovedCount + DeleteEmptyByStyle(sec.Headers(i), "Header")
            RemovedCount = RemovedCount + DeleteEmptyByStyle(sec.Footers(i), "Footer")
            RemovedCount = RemovedCount + DeleteEmptyByStyle(sec.Headers(i), C_S_ParNormal)
            RemovedCount = RemovedCount + DeleteEmptyByStyle(sec.Footers(i), C_S_ParNormal)
        Next i
    Next sec
    
    Application.ScreenUpdating = True
    MsgBox RemovedCount & " paragraph(s) removed without switching views.", vbInformation
End Sub

Private Function DeleteEmptyByStyle(hf As HeaderFooter, styleName As String) As Long
    Dim i As Long
    Dim p As Paragraph
    Dim count As Long: count = 0
    
    ' If the header doesn't exist (e.g. no different first page), skip it
    If Not hf.Exists Then Exit Function
    
    ' Loop backwards so the index (i) doesn't break when we delete a paragraph
    For i = hf.Range.Paragraphs.count To 1 Step -1
        Set p = hf.Range.Paragraphs(i)
        
        ' 1. Check if the style matches
        If InStr(1, p.style.NameLocal, styleName, vbTextCompare) > 0 Then
            
            ' 2. Check if it's empty (Length is 1 because of the paragraph mark)
            If Len(p.Range.Text) = 1 Then
                
                ' 3. Word requires at least ONE paragraph to stay in the header
                If hf.Range.Paragraphs.count > 1 Then
                    p.Range.Delete
                    count = count + 1
                End If
                
            End If
        End If
    Next i
    
    DeleteEmptyByStyle = count
End Function

' Insert full content into body of the ActiveDocument.
' 2025-07-19 by ms
' 2026-03-14 by ms
Private Sub InsertFullContent(TemplateIndex As Integer)
    Dim doc As Document: Set doc = ActiveDocument
    Dim InsertionPoint As Range
    Dim NewSection As Section
    Dim NewParagraph As Range
    Dim oTemplate As Template
    
    ' Set reference to the specific template
    Set oTemplate = Application.Templates(TemplateIndex)
    
    ' 1. Enable Odd & Even headers for the entire document
    doc.PageSetup.OddAndEvenPagesHeaderFooter = True
    
    ' --- SECTION 1: COVER PAGE ---
    With doc.Sections(1)
        ' Headers and Footers (No extra paragraphs needed here as they are in the H/F story)
        oTemplate.BuildingBlockEntries("HeaderCoverPage").Insert .Headers(wdHeaderFooterPrimary).Range
        oTemplate.BuildingBlockEntries("HeaderCoverPage").Insert .Headers(wdHeaderFooterEvenPages).Range
        oTemplate.BuildingBlockEntries("FooterCoverPage").Insert .Footers(wdHeaderFooterPrimary).Range
        oTemplate.BuildingBlockEntries("FooterCoverPage").Insert .Footers(wdHeaderFooterEvenPages).Range
        
        ' Insert Cover Table Building Block
        oTemplate.BuildingBlockEntries("CoverTable").Insert .Range
    End With
    
    ' Add empty paragraph after Cover Table
    Call AddStyledParagraph(doc, "ParNormal ms")
    
    ' --- SECTION 2: MAIN CONTENT ---
    Set InsertionPoint = doc.Range
    InsertionPoint.Collapse Direction:=wdCollapseEnd
    
    Set NewSection = doc.Sections.Add(Range:=InsertionPoint)
    With NewSection
        .Headers(wdHeaderFooterPrimary).LinkToPrevious = False
        .Headers(wdHeaderFooterEvenPages).LinkToPrevious = False
        .Footers(wdHeaderFooterPrimary).LinkToPrevious = False
        .Footers(wdHeaderFooterEvenPages).LinkToPrevious = False
        
        oTemplate.BuildingBlockEntries("HeaderOrdinary").Insert .Headers(wdHeaderFooterPrimary).Range
        oTemplate.BuildingBlockEntries("HeaderOrdinary").Insert .Headers(wdHeaderFooterEvenPages).Range
        oTemplate.BuildingBlockEntries("FooterOrdinary").Insert .Footers(wdHeaderFooterPrimary).Range
        oTemplate.BuildingBlockEntries("FooterOrdinary").Insert .Footers(wdHeaderFooterEvenPages).Range
    End With
    
    ' Insert DocumentInfoNew at start of Section 2
    Set InsertionPoint = NewSection.Range
    InsertionPoint.Collapse Direction:=wdCollapseStart
    oTemplate.BuildingBlockEntries("DocumentInfoNew").Insert InsertionPoint, True
    
    ' Add empty paragraph after DocumentInfoNew
    Call AddStyledParagraph(doc, "ParNormal ms")
    
    ' --- Lists ---
    ' List Of Content
    Set InsertionPoint = doc.Range
    InsertionPoint.Collapse Direction:=wdCollapseEnd
    oTemplate.BuildingBlockEntries("ListOfContent").Insert InsertionPoint
    Call AddStyledParagraph(doc, "ParNormal ms")
    
    ' List Of Pictures
    Set InsertionPoint = doc.Range
    InsertionPoint.Collapse Direction:=wdCollapseEnd
    oTemplate.BuildingBlockEntries("ListOfPictures").Insert InsertionPoint
    Call AddStyledParagraph(doc, "ParNormal ms")
    
    ' List Of Tables
    Set InsertionPoint = doc.Range
    InsertionPoint.Collapse Direction:=wdCollapseEnd
    oTemplate.BuildingBlockEntries("ListOfTables").Insert InsertionPoint, RichText:=True
    Call AddStyledParagraph(doc, "ParNormal ms")
    
    ' --- Heading 1 [Content] ---
    doc.Content.InsertAfter vbCr
    Set NewParagraph = doc.Paragraphs(doc.Paragraphs.count).Range
    With NewParagraph
        .style = C_S_Heading1
        .Text = "[Content]"
    End With
    
    ' Add empty paragraph after Heading 1
    Call AddStyledParagraph(doc, "ParNormal ms")
    
    ' Final Section 2 Paragraph (as per previous requirement)
    Call AddStyledParagraph(doc, C_S_ParNormal)
    
    ' --- SECTION 3: LAST PAGE ---
    Set InsertionPoint = doc.Range
    InsertionPoint.Collapse Direction:=wdCollapseEnd
    
    Set NewSection = doc.Sections.Add(Range:=InsertionPoint)
    With NewSection
        .Headers(wdHeaderFooterPrimary).LinkToPrevious = False
        .Headers(wdHeaderFooterEvenPages).LinkToPrevious = False
        .Footers(wdHeaderFooterPrimary).LinkToPrevious = False
        .Footers(wdHeaderFooterEvenPages).LinkToPrevious = False
        
        oTemplate.BuildingBlockEntries("HeaderLastPage").Insert .Headers(wdHeaderFooterPrimary).Range
        oTemplate.BuildingBlockEntries("HeaderLastPage").Insert .Headers(wdHeaderFooterEvenPages).Range
        oTemplate.BuildingBlockEntries("FooterLastPage").Insert .Footers(wdHeaderFooterPrimary).Range
        oTemplate.BuildingBlockEntries("FooterLastPage").Insert .Footers(wdHeaderFooterEvenPages).Range
    End With
    
    ' Cleanup
    Set doc = Nothing: Set InsertionPoint = Nothing: Set NewSection = Nothing: Set NewParagraph = Nothing
End Sub

' --- HELPER SUB TO REDUCE REPETITION ---
Private Sub AddStyledParagraph(ByRef doc As Document, ByVal styleName As Variant)
    doc.Content.InsertAfter vbCr
    doc.Paragraphs(doc.Paragraphs.count).Range.style = styleName
End Sub

' Insert basic content into ActiveDocument: no cover page and last page.
' 2025-04-27 by ms
' 2026-03-14 by ms
Private Sub InsertBasicContent(TemplateIndex As Integer)
    Dim doc As Document: Set doc = ActiveDocument
    Dim InsertionPoint As Range
    Dim NewParagraph As Range
    Dim oTemplate As Template
    
    ' Set reference to the specific template
    Set oTemplate = Application.Templates(TemplateIndex)
    
    ' Basic content usually implies Odd/Even headers as well
    doc.PageSetup.OddAndEvenPagesHeaderFooter = True
    
    ' --- SECTION 1: HEADERS & FOOTERS ---
    With doc.Sections(1)
        oTemplate.BuildingBlockEntries("HeaderOrdinary").Insert .Headers(wdHeaderFooterPrimary).Range
        oTemplate.BuildingBlockEntries("HeaderOrdinary").Insert .Headers(wdHeaderFooterEvenPages).Range
        
        oTemplate.BuildingBlockEntries("FooterOrdinary").Insert .Footers(wdHeaderFooterPrimary).Range
        oTemplate.BuildingBlockEntries("FooterOrdinary").Insert .Footers(wdHeaderFooterEvenPages).Range
    End With
    
    ' --- SECTION 1: BODY CONTENT ---
    ' Start at the beginning of the document
    Set InsertionPoint = doc.Range(0, 0)
    
    ' 1. DocumentInfo
    oTemplate.BuildingBlockEntries("DocumentInfo").Insert InsertionPoint, True
    Call AddStyledParagraph(doc, "ParNormal ms")
    
    ' 2. ListOfContent
    Set InsertionPoint = doc.Range
    InsertionPoint.Collapse Direction:=wdCollapseEnd
    oTemplate.BuildingBlockEntries("ListOfContent").Insert InsertionPoint
    Call AddStyledParagraph(doc, "ParNormal ms")
    
    ' 3. ListOfPictures
    Set InsertionPoint = doc.Range
    InsertionPoint.Collapse Direction:=wdCollapseEnd
    oTemplate.BuildingBlockEntries("ListOfPictures").Insert InsertionPoint
    Call AddStyledParagraph(doc, "ParNormal ms")
    
    ' 4. ListOfTables
    Set InsertionPoint = doc.Range
    InsertionPoint.Collapse Direction:=wdCollapseEnd
    oTemplate.BuildingBlockEntries("ListOfTables").Insert InsertionPoint, RichText:=True
    Call AddStyledParagraph(doc, "ParNormal ms")
    
    ' 5. Heading 1 [Content]
    doc.Content.InsertAfter vbCr
    Set NewParagraph = doc.Paragraphs(doc.Paragraphs.count).Range
    With NewParagraph
        .style = C_S_Heading1
        .Text = "[Content]"
    End With
    
    ' Add empty paragraph after Heading 1 as requested
    Call AddStyledParagraph(doc, "ParNormal ms")
    
    ' 6. Final trailing paragraph
    Call AddStyledParagraph(doc, C_S_ParNormal)
    
    ' Cleanup
    Set doc = Nothing
    Set InsertionPoint = Nothing
    Set NewParagraph = Nothing
End Sub
