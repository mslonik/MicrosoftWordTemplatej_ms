Attribute VB_Name = "Scenarios"
' VBA Module name: Scenarios.bas
' https://github.com/mslonik/Microsoft-Word-Configuration
'
'   License: MIT License.
'
'+----+-----------------------------------------+--------------+----------------+-----------------------------------------+
'| No | Sub name                                | Ribbon name  | Ribbon section | Ribbon button name                      |
'+----+-----------------------------------------+--------------+----------------+-----------------------------------------+
'| 1  | ApplyDistanceBetweenNumberingAndHeading | Scenarios_ms | Combos         | ApplyDistanceBetweenNumberingAndHeading |
'| 2  | ResetDistanceBetweenNumberingAndHeading | Scenarios_ms | Combos         | ResetDistanceBetweenNumberingAndHeading |
'| 3  | UpdateAllFieldsAndCloseFile             | Scenarios_ms | Combos         | UpdateAllFieldsAndCloseFile             |
'| 4  | DeleteAllVBAModulesExceptMacros         | Scenarios_ms | Combos         | DeleteAllVBAModulesExceptMacros         |
'| 5  | CustomizedPrintPreviewAndPrint          | Scenarios_ms | Combos         | CustomizedPrintPreviewAndPrint          |
'| 6  | NewFileConfAndContent                   | Scenarios_ms | Combos         | NewFileConfAndContent                   |
'+----+-----------------------------------------+--------------+----------------+-----------------------------------------+
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
    Dim ModuleName As String:   ModuleName = C_M_Scenarios
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
    
    Call InitializeConstants
    
    ' Origin module: Styles
    Call RemoveTextFromBeginningOfListParagraphs(textToRemove:=BetweenNumberAndText)
End Sub

' This macro is linked to Ctrl + S keyboard shortcut. Each time user runs it, it enters specific text character (em space) at the beginning of each paragraph style type list.
' Then the built-in command Save is run.
' 2025-03-09 by ms and AI
' 2026-01-17 by ms
Sub ApplyDistanceBetweenNumberingAndHeading()
    Dim FileName As String:    FileName = C_F_Macros
    Dim ModuleName As String:    ModuleName = C_M_Scenarios
    Dim MacroName As String:    MacroName = "ApplyDistanceBetweenNumberingAndHeading"
    Dim MsgBoxTitle As String:    MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName

    ' Initialize em space constant as BetweenNumberAndText
'    Call InitializeConstants
'    Call InsertTextAtBeginningOfListParagraphs(textToInsert:=BetweenNumberAndText)  ' in Styles
    
    ' Enable error handling in case that user presses 'Cancel' button.
    On Error Resume Next
    ' Execute the built-in Save command
    ActiveDocument.Save
    ' This statement turns off the error handling that was set by On Error Resume Next. It restores the default error handling behavior, which means that if an error occurs after this point, VBA will stop execution and display an error message.
    On Error GoTo 0
    Application.statusBar = MsgBoxTitle & " > " & "was running..."
    
End Sub

' Update all fields and then close the file.
' Associated to keyboard shortcut Ctrl + W.
' 2025-03-15
Sub UpdateAllFieldsAndCloseFile()
    Call UpdateAllFields            ' in module Validation
    If Not CheckFieldsAgainstErrors Then  ' in module Validation
        Exit Sub                    ' exits if error was found
    End If
    Application.Run "DocClose"      ' call built-in Microsoft Word command
End Sub

' Combo fall forward: update all fields and then show print preview
' 2025-03-15
Sub CustomizedPrintPreviewAndPrint()
    Call UpdateAllFields    ' in module Validation
    Application.CommandBars.ExecuteMso "PrintPreviewAndPrint"   ' call built-in Microsoft Word command
End Sub

' Insert full content: cover page, last page and example content: 3 sections in total.
' 2025-03-19 ms and AI
Sub NewFileConfAndContent()
    Dim FileName As String
    FileName = C_F_Macros
    
    Dim ModuleName As String
    ModuleName = C_M_Scenarios
    
    Dim MacroName As String
    MacroName = "NewFileConfAndContent"
    
    Dim MsgBoxTitle As String
    MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
            
    ' Check if the add-in template is enabled
    Dim TemplateIndex As Integer
    TemplateIndex = GetTemplateIndex(C_F_BuildingBlocks)                 ' module: BuildingBlocks
      
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
        Call CreateActiveDocumentMacroShortcuts                     ' module: Shortcuts
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
        Call CreateCustomStyles                                     ' module: Styles
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
        Call AttachTheme                    ' module: Theme
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
        Call WordOptionsCustomize            ' module: Tools
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
        Call SetMarginsDefault       ' module: Tools
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
        Call DocPropertiesUpdate            ' module: Tools
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
            Call CaptionLabelDeleteCustomized  ' module: Tools
            Call CapationAddCustomized       ' module: Tools
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
        Call SetHyphenation                 ' module: Tools
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
        Call InsertBasicContent(TemplateIndex) ' in module Scenarios
        Call BB_RemoveDefParagraphs             ' in module BuildingBlocks
        If DocPropertiesFlag = msoTrue Then
            Call DocPropertiesUserInput                 ' in module Tools, calls UpdateAllFields in module Validation
        End If
        Exit Sub
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
        Call InsertFullContent(TemplateIndex)  ' in module Scenarios
        Call BB_RemoveDefParagraphs             ' in module BuildingBlocks
        If DocPropertiesFlag = msoTrue Then
            Call DocPropertiesUserInput                 ' in module Tools, calls UpdateAllFields in module Validation
        End If
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
        Call SetPageColorToCustom                   'file: C_F_Macros, module: Tools
    End If
    If UserDecision = vbNo Then
        Call RestoreDefaultPageColor                'file: C_F_Macros, module: Tools
    End If
    
    ' The following line is in my opinion bug in Microsoft Word. For unknown reason the C_S_ParNormal is set to AutomaticallyUpdate. So I prevent it silently.
    ActiveDocument.Styles(C_S_ParNormal).AutomaticallyUpdate = False
    
End Sub

' 2025-07-19 by ms
' Insert full content into body of the ActiveDocument.
Private Sub InsertFullContent(TemplateIndex As Integer)
    Dim doc As Document
    Set doc = ActiveDocument
    'wdHeaderFooterPrimary: This constant is used to apply headers and footers to all pages in a section, except for the first page and even pages if they have their own headers and footers defined. This is the default header and footer type that is applied to pages in a section.
    ' Do not insert explicite Section 1 as it is already inserted on time a new file is created.
      
    ' Insert header and footer for Section 1
    Application.Templates(TemplateIndex).BuildingBlockEntries("HeaderCoverPage").Insert doc.Sections(1).Headers(wdHeaderFooterPrimary).Range
    Application.Templates(TemplateIndex).BuildingBlockEntries("FooterCoverPage").Insert doc.Sections(1).Footers(wdHeaderFooterPrimary).Range
    
    ' Insert BuildingBlocks for Section 1
    Application.Templates(TemplateIndex).BuildingBlockEntries("CoverTable").Insert doc.Sections(1).Range
    
    ' Insert Section 2
    doc.Sections.Add
    doc.Sections(2).Headers(wdHeaderFooterPrimary).LinkToPrevious = False
    doc.Sections(2).Footers(wdHeaderFooterPrimary).LinkToPrevious = False
    
    ' Insert header and footer for Section 2
    Application.Templates(TemplateIndex).BuildingBlockEntries("HeaderOrdinary").Insert doc.Sections(2).Headers(wdHeaderFooterPrimary).Range
    Application.Templates(TemplateIndex).BuildingBlockEntries("FooterOrdinary").Insert doc.Sections(2).Footers(wdHeaderFooterPrimary).Range
    
    ' Insert BuildingBlocks for Section 2
    Dim InsertionPoint As Range
    Set InsertionPoint = doc.Sections(2).Range
        
    InsertionPoint.Collapse Direction:=wdCollapseEnd
    Application.Templates(TemplateIndex).BuildingBlockEntries("DocumentInfoNew").Insert _
    InsertionPoint, True
    
    ' Define a new range for the just inserted paragraph
    Dim NewParagraph As Range
    
    ' Insert empty paragraph
    InsertionPoint.MoveEnd (wdSection)
    InsertionPoint.Collapse Direction:=wdCollapseEnd
    InsertionPoint.InsertParagraphAfter
    Set NewParagraph = doc.Paragraphs(doc.Paragraphs.count).Range
    NewParagraph.style = C_S_ParNormal
    
    InsertionPoint.MoveEnd (wdSection)
    InsertionPoint.Collapse Direction:=wdCollapseEnd
    Application.Templates(TemplateIndex).BuildingBlockEntries("ListOfContent").Insert _
        InsertionPoint
        
    InsertionPoint.MoveEnd (wdSection)
    InsertionPoint.Collapse Direction:=wdCollapseEnd
    Application.Templates(TemplateIndex).BuildingBlockEntries("ListOfPictures").Insert _
    InsertionPoint
    
    InsertionPoint.MoveEnd (wdSection)
    InsertionPoint.Collapse Direction:=wdCollapseEnd
    Application.Templates(TemplateIndex).BuildingBlockEntries("ListOfTables").Insert _
        InsertionPoint, RichText:=True
    
    ' Insert the specified field just before the 3rd section break
    ' This trick comes from the book WordTips_TheMacros_8E.pdf, section 7.07 "Automatic Blank Pages at the end of section"
    ' The "A4_Ver_BlankPage" must contain a page break character and empty paragraph afterwards.
    ' Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, Text:="IF { =INT({ PAGE } / 2) * 2 } = { PAGE } { AUTOTEXT ""A4_Ver_BlankPage"" } "" "" ", PreserveFormatting:=False
    ' https://stackoverflow.com/questions/15338309/setting-up-a-nested-field-in-word-using-vba
    ' BuildingBlocks: A4_Ver_BlankPage (Custom 1) + BlankPageFieldOddSection (Custom 1)
    
    ' Insert content paragraph
    With InsertionPoint
        .MoveEnd (wdSection)
        .Collapse Direction:=wdCollapseEnd
        .InsertParagraphAfter
    End With
    Set NewParagraph = doc.Paragraphs(doc.Paragraphs.count).Range
    With NewParagraph
        .style = C_S_Heading1
        .Text = "[Content]"
    End With
    
    ' Insert empty paragraph
    With InsertionPoint
        .MoveEnd (wdSection)
        .Collapse Direction:=wdCollapseEnd
        .InsertParagraphAfter
    End With
    Set NewParagraph = doc.Paragraphs(doc.Paragraphs.count).Range
    NewParagraph.style = C_S_ParNormal
    
    InsertionPoint.MoveEnd (wdSection)
    InsertionPoint.Collapse Direction:=wdCollapseEnd
    Application.Templates(TemplateIndex).BuildingBlockEntries("BlankPageFieldOddSection").Insert _
        InsertionPoint
    
    ' Insert Section 3
    doc.Sections.Add
    doc.Sections(3).Headers(wdHeaderFooterPrimary).LinkToPrevious = False
    doc.Sections(3).Footers(wdHeaderFooterPrimary).LinkToPrevious = False
    
    ' Insert header and footer for Section 3
    Application.Templates(TemplateIndex).BuildingBlockEntries("HeaderLastPage").Insert doc.Sections(3).Headers(wdHeaderFooterPrimary).Range
    Application.Templates(TemplateIndex).BuildingBlockEntries("FooterLastPage").Insert doc.Sections(3).Footers(wdHeaderFooterPrimary).Range
    
    ' Clear object variables
    Set doc = Nothing
    Set InsertionPoint = Nothing
    Set NewParagraph = Nothing
End Sub

' Insert basic content into ActiveDocument: no cover page and last page.
' 2025-04-27 by ms
Private Sub InsertBasicContent(TemplateIndex As Integer)
    Dim doc As Document
    Set doc = ActiveDocument

    ' Insert header and footer for Section 1
    Application.Templates(TemplateIndex).BuildingBlockEntries("HeaderOrdinary").Insert doc.Sections(1).Headers(wdHeaderFooterPrimary).Range
    Application.Templates(TemplateIndex).BuildingBlockEntries("FooterOrdinary").Insert doc.Sections(1).Footers(wdHeaderFooterPrimary).Range
    
    ' Insert BuildingBlocks for Section 1
    Dim InsertionPoint As Range
    Set InsertionPoint = doc.Sections(1).Range
        
    InsertionPoint.Collapse Direction:=wdCollapseEnd
    Application.Templates(TemplateIndex).BuildingBlockEntries("DocumentInfo").Insert _
    InsertionPoint, True
    
    ' Define a new range for the just inserted paragraph
    Dim NewParagraph As Range
    
    ' Insert empty paragraph
    InsertionPoint.MoveEnd (wdSection)
    InsertionPoint.Collapse Direction:=wdCollapseEnd
    InsertionPoint.InsertParagraphAfter
    Set NewParagraph = doc.Paragraphs(doc.Paragraphs.count).Range
    NewParagraph.style = C_S_ParNormal
    
    InsertionPoint.MoveEnd (wdSection)
    InsertionPoint.Collapse Direction:=wdCollapseEnd
    Application.Templates(TemplateIndex).BuildingBlockEntries("ListOfContent").Insert _
        InsertionPoint
        
    InsertionPoint.MoveEnd (wdSection)
    InsertionPoint.Collapse Direction:=wdCollapseEnd
    Application.Templates(TemplateIndex).BuildingBlockEntries("ListOfPictures").Insert _
    InsertionPoint
    
    InsertionPoint.MoveEnd (wdSection)
    InsertionPoint.Collapse Direction:=wdCollapseEnd
    Application.Templates(TemplateIndex).BuildingBlockEntries("ListOfTables").Insert _
        InsertionPoint, RichText:=True
    
    ' Insert content paragraph
    With InsertionPoint
        .MoveEnd (wdSection)
        .Collapse Direction:=wdCollapseEnd
        .InsertParagraphAfter
    End With
    Set NewParagraph = doc.Paragraphs(doc.Paragraphs.count).Range
    With NewParagraph
        .style = C_S_Heading1
        .Text = "[Content]"
    End With
    
    ' Insert empty paragraph
    With InsertionPoint
        .MoveEnd (wdSection)
        .Collapse Direction:=wdCollapseEnd
        .InsertParagraphAfter
    End With
    Set NewParagraph = doc.Paragraphs(doc.Paragraphs.count).Range
    NewParagraph.style = C_S_ParNormal
            
    With InsertionPoint
        .MoveEnd (wdSection)
        .Collapse Direction:=wdCollapseEnd
    End With
    
    Set doc = Nothing
    Set InsertionPoint = Nothing
    Set NewParagraph = Nothing
End Sub
