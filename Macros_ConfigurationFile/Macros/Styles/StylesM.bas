Attribute VB_Name = "StylesM"
' VBA Module name: Styles.bas
' https://github.com/mslonik/Microsoft-Word-Configuration
'
'   License: MIT License.
'
'
'+----+------------------------------------+-------------+------------------+------------------------------------+
'| No | Sub name                           | Ribbon name | Ribbon section   | Ribbon button name                 |
'+----+------------------------------------+-------------+------------------+------------------------------------+
'| 1  | DeleteStylesOtherThanInTemplate    | Styles_ms   | custom (no name) | DeleteStylesOtherThanInTemplate    |
'| 2  | DeleteUnusedStyles                 | Styles_ms   | custom (no name) | DeleteUnusedStyles                 |
'| 3  | CopyStylesFromTemplateToThisFile   | Styles_ms   | custom (no name) | CopyStylesFromTemplateToThisFile   |
'| 4  | SwitchOffAutoupdate                | Styles_ms   | custom (no name) | SwitchOffAutoupdate                |
'| 5  | DeleteAllNCstylingBookmarks        | Styles_ms   | custom (no name) | DeleteAllNCstylingBookmarks        |
'| 6  | DeleteNCHighlighting               | Styles_ms   | custom (no name) | DeleteNCHighlighting               |
'| 7  | DeleteCustomStyles_KeepOnlyDefined | Styles_ms   | custom (no name) | DeleteCustomStyles_KeepOnlyDefined |
'| 8  | ListBuiltInStyles                  | Styles_ms   | TextOutput       | ListBuiltInStyles                  |
'| 9  | ListNonBuiltInAndSuffixStyles      | Styles_ms   | TextOutput       | ListNonBuiltInAndSuffixStyles      |
'| 10 | ListStylesCurrentlyInUse           | Styles_ms   | TextOutput       | ListStylesCurrentlyInUse           |
'| 11 | ListCustomStylesToTxt              | Styles_ms   | TextOutput       | ListCustomStylesToTxt              |
'| 12 | ReapplyStylesFromTemplate          | Styles_ms   | Reapply          | ReapplyStylesFromTemplate          |
'| 13 | ReapplyStylesFromTemplateSimple    | Styles_ms   | Reapply          | ReapplyStylesFromTemplateSimple    |
'| 14 | CreateCustomStyles                 | Styles_ms   | custom (no name) | CreateCustomStyles                 |
'+----+------------------------------------+-------------+------------------+------------------------------------+
'| 15 | AttachTheme                        | Styles_ms   | Theme            | AttachTheme                        |
'+----+------------------------------------+-------------+------------------+------------------------------------+
'
' ShowListTemplates
' ResetAllListGalleries
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'   16. InsertTextAtBeginningOfListParagraphs() -> RemoveTextFromBeginningOfListParagraphs()
'   18. ToggleCharBoldStyle()
'   19. ToggleCharItalicStyle()
'   20. ToggleCharUnderlineStyle()
'   21. ToggleCharCrossoutStyle()
'   22. ToggleCharHiddenStyle()
'   23. ToggleCharSourceCode()
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
' Paragraph Styles:
'CreateStyle_Normal()
'CreateStyle_ParNormalMs
'CreateStyle_ParHeading1ms
'CreateStyle_ParHeading2ms
'CreateStyle_ParHeading3ms
'CreateStyle_ParHeading4ms
'CreateStyle_ParHeading5ms
'CreateStyle_ParHeading6ms
'CreateStyle_ParHeading7ms
'CreateStyle_ParHeading8ms
'CreateStyle_ParInTableMs
'CreateStyle_ParLegalNoteMs
'CreateStyle_ParLegendPictureMs
'CreateStyle_ParLegendTableMs
'CreateStyle_ParListHeadingMs
'CreateStyle_ParListIndent1Ms
'CreateStyle_ParListIndentB1Ms
'CreateStyle_ParListIndent2Ms
'CreateStyle_ParListIndentB2Ms
'CreateStyle_ParListIndent3Ms
'CreateStyle_ParListIndentB3Ms
'CreateStyle_ParListIndent4Ms
'CreateStyle_ParListIndentB4Ms
'CreateStyle_ParMinimalMs
'CreateStyle_ParNormalAboveMs
'CreateStyle_ParNormalBelowMs
'CreateStyle_ParNormalAboveBelowMs
'CreateStyle_ParNormalZeroMs
'CreateStyle_ParPictureCanvaMs
'CreateStyle_ParSourceCodeMs
'CreateStyle_TOC1
'CreateStyle_TOC2
'CreateStyle_TOC3
'CreateStyle_ParTextBoxesMs
'CreateStyle_ParNumRefMs
'CreateStyle_ParListInTableMs
'CreateStyle_ParIconMs
'
' Character templates:
'CreateStyle_CharBoldMs
'CreateStyle_CharCrossoutMs
'CreateStyle_CharDefaultMs
'CreateStyle_CharHiddenMs
'CreateStyle_CharItalicMs
'CreateStyle_CharLegalNoteMs
'CreateStyle_CharSourceCodeMs
'CreateStyle_CharUnderlineMs
'
' List Templates:
'Create_LT_Headings
'Create_LT_NumOrd
'Create_LT_Bullets
'Create_LT_SingleLevelListNumRefMs
'Create_LT_SingleLevelListInTableMs
'Create_LT_ToC
'
' Table styles:
'CreateStyle_TabTableMs()
'CreateStyle_TabTableNoGridMs()
'CreateStyle_TabTableNoPaddingMs
'
' External sources:
' https://bettersolutions.com/word/styles/index.htm
' https://addbalance.com/usersguide/styles.htm#Overview
' https://gregmaxey.com/word_tips.html
' https://www.clausebase.com/msword/numbering
' https://wordmvp.com/
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'' Used to enforce the explicit declaration of all variables in a module. When you include Option Explicit at the beginning of a module, it ensures that you must
' declare all variables using the Dim, Private, Public, ReDim, or Static statements before using them. This helps prevent errors caused by typos or
' undeclared variables.
Option Explicit
Dim ExampleResult As ColourDetails
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
' Run this macro as the first one, before the DeleteBuiltInStyles.
' It deletes all not built-in styles from the current file.
' Added logging and user oriented MsgBox communication.
' Warning, this macro deletes styles, without checking if they are actually applied. You've been warned.
' This macro can be helpful e.g. to determine if Building Blocks contain not wanted styles:
'   1. Run macro BuildingBlocks_ExportAll()
'   2. Run this macro.
'   3. Review content of the DOCX file created by BuildingBlocks_ExportAll(), if there are unwanted styles.
' Created by AI and ms on 2025-02-18.

Sub DeleteStylesOtherThanInTemplate()
    Dim TemplatePath As String
    Dim style As style
    Dim deletedStyles As String
    Dim TemplateStyles As Collection
    Dim templateStyleNames As Collection
    Dim styleName As Variant
    Dim CounterDeletedStyles As Integer
    
    Dim FileName As String:     FileName = C_F_Macros
    Dim ModuleName As String:   ModuleName = C_M_Styles
    Dim MacroName As String:    MacroName = "DeleteStylesOtherThanInTemplate"
    Dim MsgBoxTitle As String:  MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    ' Check if the currently opened document is a template file
    If ActiveDocument.AttachedTemplate.Path = "" Then
        MsgBox _
            Prompt:="Warning!" & vbNewLine & vbNewLine & _
                "The currently opened file is a template file. Please open a document file.", _
            Buttons:=vbExclamation, _
            Title:=MsgBoxTitle
        Exit Sub
    End If
        
    Dim UserDecision As VbMsgBoxResult
    Beep
    UserDecision = MsgBox( _
        Prompt:="Warning!" & vbNewLine & vbNewLine & _
            "This macro will delete specific styles that" & vbNewLine & _
            "are not defined within the attached template" & vbNewLine & vbNewLine & _
            ActiveDocument.AttachedTemplate.Name & vbNewLine & vbNewLine & _
            "from the current file." & vbNewLine & vbNewLine & _
            "Are you sure you want to proceed?", _
        Buttons:=vbYesNo + vbQuestion, _
        Title:=MsgBoxTitle _
        )
    
    If UserDecision = vbNo Then
        Exit Sub
    End If
          
    ' Open the template file as a document (not as an attached template). It will trigger Document_Open macro within that template.
    Dim TemplateDoc As Document
    Set TemplateDoc = Documents.Open(ActiveDocument.AttachedTemplate.FullName)
    
    ' Initialize collections for template styles
    Set TemplateStyles = New Collection
    Set templateStyleNames = New Collection
    
    ' Loop through all styles in the template and add to collections only not built-in styles
    For Each style In TemplateDoc.Styles
        If Not style.BuiltIn Then
            TemplateStyles.Add style
            templateStyleNames.Add style.NameLocal
        End If
    Next style
    
    ' Close the template file without saving changes
    TemplateDoc.Close SaveChanges:=wdDoNotSaveChanges
        
    ' Initialize the deleted styles string and counter
    deletedStyles = ""
    CounterDeletedStyles = 1
    
    ' Loop through all styles in the document and delete all the styles which are not present in the template
    ' Example: "Tabela bez krawedzi ms" is built-in.
    Dim DebugIndex As Integer
    DebugIndex = 0
    Dim FlagFound As Boolean
    FlagFound = False
    For Each style In ActiveDocument.Styles
        DebugIndex = DebugIndex + 1
        If InStr(1, style.NameLocal, "Tabela bez", vbTextCompare) > 0 Then
            FlagFound = True
        End If
        If Not style.BuiltIn And Not IsInCollection(coll:=templateStyleNames, item:=style.NameLocal) Then
            deletedStyles = deletedStyles & CounterDeletedStyles & ". " & style.NameLocal & vbCrLf
            style.Delete
            CounterDeletedStyles = CounterDeletedStyles + 1
        End If
    Next style
    
    Call SaveLog(MacroName:=MacroName, LoggedParameter:=deletedStyles)  ' in module Styles
    
    ' Clear object variable
    Set TemplateDoc = Nothing
    Set TemplateStyles = Nothing
    Set templateStyleNames = Nothing
    
    MsgBox _
        Prompt:="No. of deleted styles: " & CounterDeletedStyles, _
        Buttons:=vbInformation, _
        Title:=MsgBoxTitle
    
End Sub

' 2025-10-07 by ms and AI
Sub RestoreBuiltinStyleNames()
    Dim s As style
    Dim StyleOriginalName As String
    Dim StyleAlteredName As String
    Dim CommaPos As Long
    
    Dim FileName As String:     FileName = C_F_Macros
    Dim ModuleName As String:   ModuleName = C_M_Styles
    Dim MacroName As String:    MacroName = "RestoreBuiltinStyleNames"
    Dim MsgBoxTitle As String:  MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    For Each s In ActiveDocument.Styles
        If s.BuiltIn Then
            StyleAlteredName = s.NameLocal
            CommaPos = InStr(StyleAlteredName, ",")
            ' Check if the name contains a comma, indicating style name was altered
            If CommaPos > 0 Then
                ' Get the original built-in name
                StyleOriginalName = Left(StyleAlteredName, CommaPos - 1)
                On Error Resume Next
                s.NameLocal = StyleOriginalName
                On Error GoTo 0
            End If
        End If
    Next s
    
    MsgBox _
        Prompt:="Built-in style names with have been restored.", _
        Buttons:=vbInformation, _
        Title:=MsgBoxTitle
End Sub

' Internal function to simplify a code
Private Sub SaveLog(MacroName As String, LoggedParameter As String)
    Dim CurrentDate As String
    Dim CurrentTime As String
    Dim FilePath As String
    Dim filenum As Integer
    Dim DocName As String
    Dim TemplatePath As String

    ' Get current date and time
    CurrentDate = Format(Date, "yyyy-mm-dd")
    CurrentTime = Format(Time, "hh:mm")
    
    ' Path to the output text file
    FilePath = Options.DefaultFilePath(wdDocumentsPath) & "\" & MacroName & ".txt"
    
    ' Get document name
    DocName = ActiveDocument.Name
    
    ' Path to the template file
    TemplatePath = ActiveDocument.AttachedTemplate.FullName

    ' Check if the file exists and open it for appending if it does, or create a new file if it doesn't
    ' This line assigns the next available file number to the variable fileNum. The FreeFile function returns an integer representing the next file number that is not currently in use. This is necessary because when you open a file in VBA, you need to specify a file number to refer to that file.
    filenum = FreeFile
    If Dir(FilePath) <> "" Then
        ' This line opens the file specified by filePath for output (writing). The Open statement is used to open a file, and the For Output clause indicates that the file is being opened for writing. The As fileNum part associates the file with the file number stored in fileNum, which was obtained using the FreeFile function. This allows you to refer to the file using fileNum in subsequent operations, such as writing data to the file.
        Open FilePath For Append As filenum
    Else
        ' This line opens the file specified by filePath for output (writing). The Open statement is used to open a file, and the For Output clause indicates that the file is being opened for writing. The As fileNum part associates the file with the file number stored in fileNum, which was obtained using the FreeFile function. This allows you to refer to the file using fileNum in subsequent operations, such as writing data to the file.
        Open FilePath For Output As filenum
    End If
    
    ' Write the deleted styles to the text file
    Print #filenum, "Document Name: " & DocName
    Print #filenum, "Template Name: " & TemplatePath
    Print #filenum, MacroName
    Print #filenum, "Date: " & CurrentDate
    Print #filenum, "Time: " & CurrentTime
    Print #filenum, vbCrLf & LoggedParameter
    Close filenum
    
    ' Inform the user that the styles have been deleted and logged
    MsgBox _
        Prompt:="Information are logged to " & vbNewLine & FilePath, _
        Buttons:=vbInformation, _
        Title:=MacroName
End Sub

' Because the built-in styles cannot be deleted, this macro lists all built-in styles.
' Created by AI and ms on 2025-02-19.
Sub ListBuiltInStyles()
    Dim doc As Document
    Dim TemplatePath As String
    Dim style As style
    Dim BuiltInStyles As String
    Dim counter As Integer
    
    Dim FileName As String:     FileName = C_F_Macros
    Dim ModuleName As String:   ModuleName = C_M_Styles
    Dim MacroName As String:    MacroName = "ListBuiltInStyles"
    Dim MsgBoxTitle As String:  MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    ' Get the currently active document
    Set doc = ActiveDocument
    
    ' Get the path of the template attached to the document
    TemplatePath = doc.AttachedTemplate.FullName
    
    ' Initialize the deleted styles string and counters
    counter = 1
    
    ' Loop through all styles in the document
    For Each style In doc.Styles
        If style.BuiltIn Then
                BuiltInStyles = BuiltInStyles & counter & ". " & style.NameLocal & vbCrLf
                counter = counter + 1
        End If
    Next style
    
    Call SaveLog(MacroName:=MacroName, LoggedParameter:=BuiltInStyles)
    
    ' Clear object variables
    Set doc = Nothing
End Sub

Private Function IsInCollection(coll As Collection, item As Variant) As Boolean
    Dim i As Integer
    
    IsInCollection = False
    For i = 1 To coll.count
        If coll(i) = item Then
            IsInCollection = True
            Exit Function
        End If
    Next i
End Function

' Lists / logs only style names which contain the suffix C_StyleSuffix (" ms") in currently active document.
' Created by AI and ms on 2025-01-28.
' 2025-10-08 by ms, added additional conditions. Simplified sub.
Sub ListNonBuiltInAndSuffixStyles()
    Dim doc As Document
    Dim TemplatePath As String
    Dim style As style
    Dim styleInfo As String
    Dim FilePath As String
    Dim filenum As Integer
    Dim rowNum As Integer
    
    Dim ParagraphStyles As Collection   ' wdStyleTypeParagraph
    Dim CharacterStyles As Collection   ' wdStyleTypeCharacter
    Dim TableStyles As Collection       ' wdStyleTypeTable
    Dim ListStyles As Collection        ' wdStyleTypeList
    ' This style type practically does not exist in new Microsoft Word
    Dim LinkedStyles As Collection      ' wdStyleTypeLinked
    
    Dim styleName As Variant
    Dim StyleType As Variant
    
    Dim FileName As String:     FileName = C_F_Macros
    Dim ModuleName As String:   ModuleName = C_M_Styles
    Dim MacroName As String:    MacroName = "ListNonBuiltInAndSuffixStylesInTemplate"
    Dim MsgBoxTitle As String:  MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    ' Get the currently active document
    If Documents.count > 0 Then
        Set doc = ActiveDocument
    Else
        MsgBox _
            Prompt:="No document is currently open." & vbNewLine & vbNewLine & _
                "Exiting.", _
            Buttons:=vbCritical, _
            Title:=MsgBoxTitle
        Exit Sub
    End If
       
    ' Initialize the style information string
    styleInfo = "No. | Style Name | Type | Built-in" & vbCrLf
    rowNum = 1
    
    ' Initialize collections for each style type
    Set ParagraphStyles = New Collection
    Set CharacterStyles = New Collection
    Set TableStyles = New Collection
    Set ListStyles = New Collection
    Set LinkedStyles = New Collection
    
    ' Loop through all styles in the template
    For Each style In doc.Styles
        If Not style.BuiltIn Or InStr(style.NameLocal, C_StyleSuffix) > 0 Then
            styleInfo = styleInfo & rowNum & " | " & _
                        style.NameLocal & " | " & _
                        StyleTypeName(StyleType:=style.Type) & " | " & _
                        CStr(style.BuiltIn) & vbCrLf
            rowNum = rowNum + 1
            
            ' Add style names to respective collections
            Select Case style.Type
                Case wdStyleTypeParagraph
                    ParagraphStyles.Add Array(style.NameLocal, StyleTypeName(StyleType:=style.Type), CStr(style.BuiltIn))
                Case wdStyleTypeCharacter
                    CharacterStyles.Add Array(style.NameLocal, StyleTypeName(StyleType:=style.Type), CStr(style.BuiltIn))
                Case wdStyleTypeTable
                    TableStyles.Add Array(style.NameLocal, StyleTypeName(StyleType:=style.Type), CStr(style.BuiltIn))
                Case wdStyleTypeList
                    ListStyles.Add Array(style.NameLocal, StyleTypeName(StyleType:=style.Type), CStr(style.BuiltIn))
                Case wdStyleTypeLinked
                   ListStyles.Add Array(style.NameLocal, StyleTypeName(StyleType:=style.Type), CStr(style.BuiltIn))
            End Select
        End If
    Next style
       
    ' Paragraph styles: sort and add paragraph styles to styleInfo
    styleInfo = styleInfo & vbCrLf & "Paragraph Styles:" & vbCrLf
    Call SortCollection(coll:=ParagraphStyles)
    rowNum = 1
    For Each styleName In ParagraphStyles
        styleInfo = styleInfo & rowNum & " | " & styleName(0) & " | " & styleName(1) & " | " & styleName(2) & vbCrLf
        rowNum = rowNum + 1
    Next styleName
    
    ' Character styles: sort and add character styles to styleInfo
    styleInfo = styleInfo & vbCrLf & "Character Styles:" & vbCrLf
    Call SortCollection(coll:=CharacterStyles)
    rowNum = 1
    For Each styleName In CharacterStyles
        styleInfo = styleInfo & rowNum & " | " & styleName(0) & " | " & styleName(1) & " | " & styleName(2) & vbCrLf
        rowNum = rowNum + 1
    Next styleName
    
    ' Table styles: Sort and add table styles to styleInfo
    styleInfo = styleInfo & vbCrLf & "Table Styles:" & vbCrLf
    Call SortCollection(coll:=TableStyles)
    rowNum = 1
    For Each styleName In TableStyles
        styleInfo = styleInfo & rowNum & " | " & styleName(0) & " | " & styleName(1) & " | " & styleName(2) & vbCrLf
        rowNum = rowNum + 1
    Next styleName
    
    ' List styles: sort and add list styles to styleInfo
    styleInfo = styleInfo & vbCrLf & "List Styles:" & vbCrLf
    Call SortCollection(coll:=ListStyles)
    rowNum = 1
    For Each styleName In ListStyles
        styleInfo = styleInfo & rowNum & " | " & styleName(0) & " | " & styleName(1) & " | " & styleName(2) & vbCrLf
        rowNum = rowNum + 1
    Next styleName
    
    ' Linked styles: sort and add list styles to styleInfo
    styleInfo = styleInfo & vbCrLf & "Linked styles:" & vbCrLf
    Call SortCollection(coll:=LinkedStyles)
    rowNum = 1
    For Each styleName In LinkedStyles
        styleInfo = styleInfo & rowNum & " | " & styleName(0) & " | " & styleName(1) & " | " & styleName(2) & vbCrLf
        rowNum = rowNum + 1
    Next styleName
    
    Call SaveLog(MacroName:=MacroName, LoggedParameter:=styleInfo)
        
    ' Clear object variables
    Set doc = Nothing
    Set ParagraphStyles = Nothing
    Set CharacterStyles = Nothing
    Set TableStyles = Nothing
    Set ListStyles = Nothing
    Set LinkedStyles = Nothing
End Sub

' Deletes unused and not built-in styles in the current document.
' Expected state: corresponding log file will be empty.
' Reviewed by AI and ms on 2025-02-18
Sub DeleteUnusedStyles()
    Dim oStyle As style
    Dim doc As Document
    Dim counter As Integer
    Dim FilePath As String
    Dim filenum As Integer
    Dim deletedStyles As String
    
    ' Initialize counter and deletedStyles
    counter = 0
    deletedStyles = ""
    For Each oStyle In ActiveDocument.Styles
        'Only check out non-built-in styles
        If oStyle.BuiltIn = False Then
            With ActiveDocument.Content.Find
                .ClearFormatting
                .style = oStyle.NameLocal
                .Execute findText:="", Format:=True
                If .found = False Then
                    deletedStyles = counter & ". " & deletedStyles & oStyle.NameLocal & vbCrLf
                    oStyle.Delete
                    counter = counter + 1
                End If
            End With
        End If
    Next oStyle
        
    Dim FileName As String:     FileName = C_F_Macros
    Dim ModuleName As String:   ModuleName = C_M_Styles
    Dim MacroName As String:    MacroName = "DeleteUnusedStyles"
    Dim MsgBoxTitle As String:  MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
        
    Call SaveLog(MacroName:=MacroName, LoggedParameter:=deletedStyles)
End Sub

' Restores or reapplies settings of a style assigned to paragraph. Each paragraph has assigned certain paragraph type style. Such style can be altered on time of editing by a user. The purpose of this macro is to restore original settings of style which was ssigned to the paragraph.
' Warning: as side effect the additional styling added by a user will be lost. You've been warned.
' The recommended practice is to not change paragraph type styles. If you really need to do that, define a new style.
' 2025-02-13 by ms and AI
Sub ReapplyStylesFromTemplate()
    Dim para As Paragraph
    Dim sec As Section
    Dim hdr As HeaderFooter
    Dim ftr As HeaderFooter
    Dim doc As Document
    Dim templateStyle As style
    Dim CurrentStyle As style
    Dim diffCount As Integer
    Dim TemplateName As String
    Dim totalParagraphs As Integer
    Dim checkedParagraphs As Integer
    
    Dim FileName As String:     FileName = C_F_Macros
    Dim ModuleName As String:   ModuleName = C_M_Styles
    Dim MacroName As String:    MacroName = "ReapplyStylesFromTemplate"
    Dim MsgBoxTitle As String:  MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    ' Display warning message
    Dim UserDecision As VbMsgBoxResult
    Beep
    UserDecision = MsgBox( _
        Prompt:="Warning! This macro will remove any changes in formatting of document paragraphs added by a user." & vbNewLine & _
            "In other words, only default formatting aligned to styles of the attached template will be preserved." & vbNewLine & vbNewLine & _
            "You've been warned." & vbNewLine & vbNewLine & _
            "Do you want to proceed?", _
        Buttons:=vbYesNo + vbExclamation + vbQuestion, _
        Title:=MsgBoxTitle _
        )
    If UserDecision = vbNo Then
        Exit Sub
    End If
    
    ' Check if this macro could run correctly
    Call CheckWordVersion
    Call AddTemporaryBookmark
    
    ' Initialize variables
    Set doc = ActiveDocument
    diffCount = 0
    TemplateName = doc.AttachedTemplate.Name
    FileName = doc.Name
    totalParagraphs = doc.Paragraphs.count
    checkedParagraphs = 0
    
    ' When Application.ScreenUpdating is set to False, it turns off screen updating, which can significantly speed up the execution of a macro by preventing the screen from refreshing until the macro has finished running. This is particularly useful for macros that perform a lot of operations, as it reduces the time spent on rendering the screen.
    Application.ScreenUpdating = False
    Application.DisplayAlerts = wdAlertsNone
    
    ' ShowModal must be set to False in the corresponding Form
    ReapplyStylesFromTemplate_Form.Show
      
    ' Reapply styles for all paragraphs in the main document (document body)
    For Each para In doc.Paragraphs
        ' Get the default style from the template
        Set templateStyle = doc.Styles(para.style.NameLocal)
        ' Get the current style of the paragraph
        Set CurrentStyle = para.style
        ' Compare the styles
        If Not CompareParagraphFormatting(para:=para, style:=templateStyle) Then
            diffCount = diffCount + 1
            para.Range.style = templateStyle
        End If
        ' Increment checked paragraphs counter
        checkedParagraphs = checkedParagraphs + 1
        ReapplyStylesFromTemplate_Form.ProgressLabel = "Finished: " & checkedParagraphs & " out of " & totalParagraphs
        ' The DoEvents function in Visual Basic for Applications (VBA) for Microsoft Word is used to yield execution so that the operating system can process other events. This function allows the operating system to handle other tasks, such as updating the screen, responding to user inputs, or processing other events in the queue, while your macro is running
        DoEvents
    Next para
    
    ' Reapply styles for all sections
    For Each sec In ActiveDocument.Sections
        ' Reapply styles for headers
        For Each hdr In sec.Headers
            For Each para In hdr.Range.Paragraphs
                ' Get the default style from the template
                Set templateStyle = doc.Styles(para.style.NameLocal)
                ' Get the current style of the paragraph
                Set CurrentStyle = para.style
                ' Compare the styles
                If Not CompareParagraphFormatting(para:=para, style:=templateStyle) Then
                    diffCount = diffCount + 1
                    para.Range.style = templateStyle
                End If
                ' Increment checked paragraphs counter
                checkedParagraphs = checkedParagraphs + 1
                ReapplyStylesFromTemplate_Form.ProgressLabel = "Finished: " & checkedParagraphs & " out of " & totalParagraphs
                ' The DoEvents function in Visual Basic for Applications (VBA) for Microsoft Word is used to yield execution so that the operating system can process other events. This function allows the operating system to handle other tasks, such as updating the screen, responding to user inputs, or processing other events in the queue, while your macro is running
                DoEvents
            Next para
        Next hdr
        
        ' Reapply styles for footers
        For Each ftr In sec.Footers
            For Each para In ftr.Range.Paragraphs
                ' Get the default style from the template
                Set templateStyle = doc.Styles(para.style.NameLocal)
                ' Get the current style of the paragraph
                Set CurrentStyle = para.style
                ' Compare the styles
                If Not CompareParagraphFormatting(para:=para, style:=templateStyle) Then
                    diffCount = diffCount + 1
                    para.Range.style = templateStyle
                End If
                ' Increment checked paragraphs counter
                checkedParagraphs = checkedParagraphs + 1
                ReapplyStylesFromTemplate_Form.ProgressLabel = "Finished: " & checkedParagraphs & " out of " & totalParagraphs
                ' The DoEvents function in Visual Basic for Applications (VBA) for Microsoft Word is used to yield execution so that the operating system can process other events. This function allows the operating system to handle other tasks, such as updating the screen, responding to user inputs, or processing other events in the queue, while your macro is running
                DoEvents
            Next para
        Next ftr
    Next sec
    
    ActiveWindow.View.Type = wdPrintView
    Application.DisplayAlerts = wdAlertsAll
    Application.ScreenUpdating = True
    
    Unload ReapplyStylesFromTemplate_Form
    Call RemoveTemporaryBookmark
    
    MsgBox _
        Prompt:="Styles in this document have been successfuly reapplied from the current template." _
            & vbNewLine & vbNewLine & "Template Name: " & TemplateName _
            & vbNewLine & "File Name: " & FileName _
            & vbNewLine & vbNewLine & "Number of style differences: " & diffCount _
            & vbNewLine & "number of paragraphs in this document: " & totalParagraphs, _
        Buttons:=vbInformation, _
        Title:=MsgBoxTitle
    
    ' Clear object variables
    Set doc = Nothing
    Set templateStyle = Nothing
    Set CurrentStyle = Nothing
End Sub

' Function to compare paragraph formatting
Private Function CompareParagraphFormatting(para As Paragraph, style As style) As Boolean
    With para.Range.ParagraphFormat
        If .Alignment <> style.ParagraphFormat.Alignment Then CompareParagraphFormatting = False: Exit Function
        If .LeftIndent <> style.ParagraphFormat.LeftIndent Then CompareParagraphFormatting = False: Exit Function
        If .RightIndent <> style.ParagraphFormat.RightIndent Then CompareParagraphFormatting = False: Exit Function
        If .SpaceBefore <> style.ParagraphFormat.SpaceBefore Then CompareParagraphFormatting = False: Exit Function
        If .SpaceAfter <> style.ParagraphFormat.SpaceAfter Then CompareParagraphFormatting = False: Exit Function
        If .LineSpacingRule <> style.ParagraphFormat.LineSpacingRule Then CompareParagraphFormatting = False: Exit Function
        If .LineSpacing <> style.ParagraphFormat.LineSpacing Then CompareParagraphFormatting = False: Exit Function
    End With
    CompareParagraphFormatting = True
End Function

Private Sub CheckWordVersion()
    Dim FileName As String:     FileName = C_F_Macros
    Dim ModuleName As String:   ModuleName = C_M_Macros
    Dim MacroName As String:    MacroName = "CheckWordVersion"
    Dim MsgBoxTitle As String:  MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName

    If Application.Version <> "14.0" And Application.Version <> "16.0" Then
        MsgBox _
            Prompt:="This macro couldn't run with your version of Microsoft Office!", _
            Buttons:=vbCritical, _
            Title:=MsgBoxTitle
        Exit Sub
    End If
End Sub

Private Sub AddTemporaryBookmark()
    ' Adds a bookmark in place where cursor is present
    If ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument Then
        Selection.Bookmarks.Add (C_BM_LastCursorPosition)
    End If
End Sub

Private Sub RemoveTemporaryBookmark()
    ' Goes to a place where temporary bookmark was located and removes it afterwards
    If ActiveDocument.Bookmarks.Exists(C_BM_LastCursorPosition) Then
        Selection.GoTo What:=wdGoToBookmark, Name:=C_BM_LastCursorPosition
        ActiveDocument.Bookmarks(C_BM_LastCursorPosition).Delete
    Else
        Selection.HomeKey wdStory    ' it moves the selection (or cursor) to the very beginning of the document.
    End If
End Sub

' Restores or reapplies settings of a style assigned to paragraph. Each paragraph has assigned certain paragraph type style. Such style can be altered on time of editing by a user. The purpose of this macro is to restore original settings of style which was ssigned to the paragraph.
' Warning: as side effect the additional styling added by a user will be lost. You've been warned.
' This is simplified version of the sub ReapplyStylesFromTemplate(), which do not ask additional question and there is no progress bar.
' 2025-02-13 by ms
Sub ReapplyStylesFromTemplateSimple()
    Dim para As Paragraph
    Dim sec As Section
    Dim hdr As HeaderFooter
    Dim ftr As HeaderFooter
    
    Dim FileName As String:     FileName = C_F_Macros
    Dim ModuleName As String:   ModuleName = C_M_Styles
    Dim MacroName As String:    MacroName = "ReapplyStylesFromTemplateSimple"
    Dim MsgBoxTitle As String:  MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    ' Reapply styles for all paragraphs in the main document
    For Each para In ActiveDocument.Paragraphs
        para.Range.style = ActiveDocument.Styles(para.style)
    Next para
    
    ' Reapply styles for all sections
    For Each sec In ActiveDocument.Sections
        ' Reapply styles for headers
        For Each hdr In sec.Headers
            For Each para In hdr.Range.Paragraphs
                para.Range.style = ActiveDocument.Styles(para.style)
            Next para
        Next hdr
        
        ' Reapply styles for footers
        For Each ftr In sec.Footers
            For Each para In ftr.Range.Paragraphs
                para.Range.style = ActiveDocument.Styles(para.style)
            Next para
        Next ftr
    Next sec
    
    MsgBox _
        Prompt:="Styles reapplied from the current template.", _
        Buttons:=vbInformation, _
        Title:=MsgBoxTitle
End Sub

' Shows the same list of available styles as in "Organizer".
' The "Organizer" dialog in Word shows only the styles that are currently in use or have been modified in the document, rather than all available styles.
' Generated by AI and ms on 2025-01-28.
Sub ListStylesCurrentlyInUse()
    Dim doc As Document
    Dim TemplateName As String
    Dim style As style
    Dim styleInfo As String
    Dim FilePath As String
    Dim filenum As Integer
    Dim counter As Integer
    Dim ParagraphStyles As Collection
    Dim CharacterStyles As Collection
    Dim TableStyles As Collection
    Dim ListStyles As Collection
    Dim styleName As Variant
    
    Dim FileName As String:     FileName = C_F_Macros
    Dim ModuleName As String:   ModuleName = C_M_Styles
    Dim MacroName As String:    MacroName = "ListStylesCurrentlyInUse"
    Dim MsgBoxTitle As String:  MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    ' Get the currently active document
    Set doc = ActiveDocument
    
    ' Get the template attached to the document
    TemplateName = doc.AttachedTemplate.FullName
    
    ' Initialize the style information string
    styleInfo = "No. | Style Name | Type | Built-in" & vbCrLf
    counter = 1
    
    ' Initialize collections for each style type
    Set ParagraphStyles = New Collection
    Set CharacterStyles = New Collection
    Set TableStyles = New Collection
    Set ListStyles = New Collection
    
    ' Loop through all styles in the document
    For Each style In doc.Styles
        If style.InUse Then
            styleInfo = styleInfo & counter & " | " & _
                        style.NameLocal & " | " & _
                        StyleTypeName(StyleType:=style.Type) & " | " & _
                        CStr(style.BuiltIn) & vbCrLf
            counter = counter + 1
            
            ' Add style names to respective collections
            Select Case style.Type
                Case wdStyleTypeParagraph
                    ParagraphStyles.Add Array(style.NameLocal, StyleTypeName(StyleType:=style.Type), CStr(style.BuiltIn))
                Case wdStyleTypeCharacter
                    CharacterStyles.Add Array(style.NameLocal, StyleTypeName(StyleType:=style.Type), CStr(style.BuiltIn))
                Case wdStyleTypeTable
                    TableStyles.Add Array(style.NameLocal, StyleTypeName(StyleType:=style.Type), CStr(style.BuiltIn))
                Case wdStyleTypeList
                    ListStyles.Add Array(style.NameLocal, StyleTypeName(StyleType:=style.Type), CStr(style.BuiltIn))
            End Select
        End If
    Next style
    
    ' Sort and add paragraph styles to styleInfo
    styleInfo = styleInfo & vbCrLf & "Paragraph Styles:" & vbCrLf
    Call SortCollection(coll:=ParagraphStyles)
    counter = 1
    For Each styleName In ParagraphStyles
        styleInfo = styleInfo & counter & " | " & styleName(0) & " | " & styleName(1) & " | " & styleName(2) & vbCrLf
        counter = counter + 1
    Next styleName
    
    ' Sort and add character styles to styleInfo
    styleInfo = styleInfo & vbCrLf & "Character Styles:" & vbCrLf
    Call SortCollection(coll:=CharacterStyles)
    counter = 1
    For Each styleName In CharacterStyles
        styleInfo = styleInfo & counter & " | " & styleName(0) & " | " & styleName(1) & " | " & styleName(2) & vbCrLf
        counter = counter + 1
    Next styleName
    
    ' Sort and add table styles to styleInfo
    styleInfo = styleInfo & vbCrLf & "Table Styles:" & vbCrLf
    Call SortCollection(coll:=TableStyles)
    counter = 1
    For Each styleName In TableStyles
        styleInfo = styleInfo & counter & " | " & styleName(0) & " | " & styleName(1) & " | " & styleName(2) & vbCrLf
        counter = counter + 1
    Next styleName
    
    ' Sort and add list styles to styleInfo
    styleInfo = styleInfo & vbCrLf & "List Styles:" & vbCrLf
    Call SortCollection(coll:=ListStyles)
    counter = 1
    For Each styleName In ListStyles
        styleInfo = styleInfo & counter & " | " & styleName(0) & " | " & styleName(1) & " | " & styleName(2) & vbCrLf
        counter = counter + 1
    Next styleName
    
    Call SaveLog(MacroName:=MacroName, _
        LoggedParameter:=vbNewLine & _
        "Only the styles that are currently in use or have been modified in the document, rather than all available styles." & _
        vbNewLine & vbNewLine & styleInfo)
        
    ' Clear object variables
    Set doc = Nothing
    Set ParagraphStyles = Nothing
    Set CharacterStyles = Nothing
    Set TableStyles = Nothing
    Set ListStyles = Nothing
End Sub

Private Function StyleTypeName(StyleType As WdStyleType) As String
    Select Case StyleType
        Case wdStyleTypeParagraph
            StyleTypeName = "Paragraph"
        Case wdStyleTypeCharacter
            StyleTypeName = "Character"
        Case wdStyleTypeTable
            StyleTypeName = "Table"
        Case wdStyleTypeList
            StyleTypeName = "List"
        Case Else
            StyleTypeName = "Unknown"
    End Select
End Function

Private Sub SortCollection(ByRef coll As Collection)
    Dim i As Integer, j As Integer
    Dim temp As Variant
    
    ' Simple bubble sort
    For i = 1 To coll.count - 1
        For j = i + 1 To coll.count
            If coll(i)(0) > coll(j)(0) Then
                temp = coll(i)
                coll.Add coll(j), Before:=i
                coll.Add temp, Before:=j
                coll.Remove i + 1
                coll.Remove j + 1
            End If
        Next j
    Next i
End Sub

' Copy or replace styles from currentlty attached template file in currently active document.
' All styles which names are finished with " ms" C_StyleSuffix and additionaly "TOC 1", "TOC 2", "TOC 3", "TOC 4"
' 2025-02-20 by ms and AI
Sub CopyStylesFromTemplateToThisFile()
    Dim TemplateDoc As Document
    Dim activeDoc As Document
    Dim styleName As String
    Dim tocStyles As Variant
    Dim i As Integer
    Dim errorCounter As Integer
    Dim StyleCounter As Integer
    Dim MyStyle As style
    
    Dim FileName As String:     FileName = C_F_Macros
    Dim ModuleName As String:   ModuleName = C_M_Styles
    Dim MacroName As String:    MacroName = "CopyStylesFromTemplateToThisFile"
    Dim MsgBoxTitle As String:  MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    ' Define the TOC styles
    tocStyles = Array("TOC 1", "TOC 2", "TOC 3", "TOC 4")
    
    ' Get the currently active document
    Set activeDoc = ActiveDocument
    
    ' Get the currently attached template document
    Set TemplateDoc = Documents.Open(ActiveDocument.AttachedTemplate.FullName)
    
    ' Check if the active document is the same as the template document
    If activeDoc.FullName = TemplateDoc.FullName Then
        MsgBox _
            Prompt:="The active document is the same as the template document. Operation aborted.", _
            Buttons:=vbExclamation, _
            Title:=MsgBoxTitle
        Exit Sub
    End If
    
    ' Initialization of the counters
    errorCounter = 0
    StyleCounter = 0
    
    ' Loop through all styles in the template document
    For Each MyStyle In TemplateDoc.Styles
        styleName = MyStyle.NameLocal
        
        ' Check if the style name ends with the specified suffix or is one of the TOC styles
        If Right(styleName, Len(C_StyleSuffix)) = C_StyleSuffix Or IsInArray(Value:=styleName, ArrayPar:=tocStyles) Then
            ' Copy or overwrite the style in the active document
            On Error Resume Next
            Application.OrganizerCopy _
                Source:=TemplateDoc.FullName, _
                Destination:=activeDoc.FullName, _
                Name:=styleName, _
                Object:=wdOrganizerObjectStyles
            
            ' Check for errors and update counters
            If Err.Number <> 0 Then
                errorCounter = errorCounter + 1
                Err.Clear
            Else
                StyleCounter = StyleCounter + 1
            End If
            
            ' This statement turns off the error handling that was set by On Error Resume Next. It restores the default error handling behavior, which means that if an error occurs after this point, VBA will stop execution and display an error message.
            On Error GoTo 0
        End If
    Next MyStyle
    
    ' Close the template document without saving changes
    TemplateDoc.Close SaveChanges:=wdDoNotSaveChanges
    
    ' Inform the user that the styles have been copied/overwritten
    MsgBox _
        Prompt:="Styles have been copied/overwritten from the template to the active document." & vbCrLf & _
           "Successfully copied styles: " & StyleCounter & vbCrLf & _
           "Errors encountered: " & errorCounter, _
        Buttons:=vbInformation, _
        Title:=MsgBoxTitle
    
    ' Clear object variables
    Set activeDoc = Nothing
    Set TemplateDoc = Nothing
End Sub

' Function to check if a value is in an array
Private Function IsInArray(Value As String, ArrayPar As Variant) As Boolean
    Dim i As Integer
    IsInArray = False
    For i = LBound(ArrayPar) To UBound(ArrayPar)
        If ArrayPar(i) = Value Then
            IsInArray = True
            Exit Function
        End If
    Next i
End Function

' At the very beginning of a list parargaph inserts specified text string
' ChrW(&H2003) = em space
' 2025-03-08 by ms and AI
Sub InsertTextAtBeginningOfListParagraphs(textToInsert As String)
    Dim para As Paragraph

    Dim FileName As String:     FileName = C_F_Macros
    Dim ModuleName As String:   ModuleName = C_M_Styles
    Dim MacroName As String:    MacroName = "InsertTextAtBeginningOfListParagraphs"
    Dim MsgBoxTitle As String:  MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    Dim UserDecision As VbMsgBoxResult
    
    Application.ScreenUpdating = True
    
    ' Loop through all paragraphs in the document
    For Each para In ActiveDocument.Paragraphs
        ' Check if the paragraph is formatted as a list and has the specified style
        If para.Range.ListFormat.ListType <> wdListNoNumbering And _
            (para.style.NameLocal = C_S_Heading1 Or _
            para.style.NameLocal = C_S_Heading2 Or _
            para.style.NameLocal = C_S_Heading3 Or _
            para.style.NameLocal = C_S_Heading4 Or _
            para.style.NameLocal = C_S_Heading5 Or _
            para.style.NameLocal = C_S_Heading6 Or _
            para.style.NameLocal = C_S_Heading7 Or _
            para.style.NameLocal = C_S_Heading8 Or _
            para.style.NameLocal = C_S_ListLevel1 Or _
            para.style.NameLocal = C_S_ListLevel2 Or _
            para.style.NameLocal = C_S_ListLevel3 Or _
            para.style.NameLocal = C_S_ListLevel4 Or _
            para.style.NameLocal = C_S_ListLevelB1 Or _
            para.style.NameLocal = C_S_ListLevelB2 Or _
            para.style.NameLocal = C_S_ListLevelB3 Or _
            para.style.NameLocal = C_S_ListLevelB4 Or _
            para.style.NameLocal = C_S_ListNumRef Or _
            para.style.NameLocal = C_S_ListNumTable Or _
            para.style.NameLocal = C_S_ParNormal) Then
            ' Check if the paragraph starts with the specified text
            If Left(para.Range.Text, Len(textToInsert)) <> textToInsert Then
                ' Insert the specified text at the beginning of the paragraph
                DoEvents    ' Force a screen refresh
                para.Range.HighlightColorIndex = wdYellow
                On Error Resume Next
                para.Range.InsertBefore textToInsert
                
                ' Move the selection one character to the right
                Selection.MoveRight Unit:=wdCharacter, count:=1

                If Err.Number <> 0 Then
                    Beep
                    UserDecision = MsgBox( _
                        Prompt:="Error " & Err.Number & ": " & Err.Description & vbNewLine & vbNewLine & _
                            "Do you want to continue?" & vbNewLine & vbNewLine & _
                            "Perhaps reapply style 'ParInTable ms' to a table with merged cells.", _
                        Buttons:=vbYesNo + vbQuestion + vbDefaultButton1, _
                        Title:=MsgBoxTitle _
                        )
                    If UserDecision = vbYes Then
                        para.Range.HighlightColorIndex = wdNoHighlight
                    Else
                        para.Range.Select
                        Exit Sub
                    End If
                End If
                On Error GoTo 0
                para.Range.HighlightColorIndex = wdNoHighlight
            End If
        End If
    Next para
    Application.ScreenUpdating = False
End Sub


' Reverts action taken by InsertTextAtBeginningOfListParagraphs
' https://stackoverflow.com/questions/73719070/word-vba-range-characters-first-delete-delets-2-characters-if-the-second-charac
' Crucial setting: switch-off manually the "Adjust sentence and word spacing automatically."
' Unfortunately, the specific setting "Adjust sentence and word spacing automatically" cannot be directly controlled through VBA. This setting is part of the advanced options in Word and does not have a corresponding VBA property or method.
' 2025-03-09 by ms and AI
Sub RemoveTextFromBeginningOfListParagraphs(textToRemove As String)
    Dim para As Paragraph
    Dim textLen As Long
    
    ' Get the length of the text to remove
    textLen = Len(textToRemove)
    
    ' Loop through all paragraphs in the document
    For Each para In ActiveDocument.Paragraphs
        ' Check if the paragraph is formatted as a list
        If para.Range.ListFormat.ListType <> wdListNoNumbering And para.style = C_S_Heading1 Or _
            para.Range.ListFormat.ListType <> wdListNoNumbering And para.style = C_S_Heading2 Or _
            para.Range.ListFormat.ListType <> wdListNoNumbering And para.style = C_S_Heading3 Or _
            para.Range.ListFormat.ListType <> wdListNoNumbering And para.style = C_S_Heading4 Or _
            para.Range.ListFormat.ListType <> wdListNoNumbering And para.style = C_S_Heading5 Or _
            para.Range.ListFormat.ListType <> wdListNoNumbering And para.style = C_S_Heading6 Or _
            para.Range.ListFormat.ListType <> wdListNoNumbering And para.style = C_S_Heading7 Or _
            para.Range.ListFormat.ListType <> wdListNoNumbering And para.style = C_S_Heading8 Or _
            para.Range.ListFormat.ListType <> wdListNoNumbering And para.style = C_S_ListLevel1 Or _
            para.Range.ListFormat.ListType <> wdListNoNumbering And para.style = C_S_ListLevel1 Or _
            para.Range.ListFormat.ListType <> wdListNoNumbering And para.style = C_S_ListLevel3 Or _
            para.Range.ListFormat.ListType <> wdListNoNumbering And para.style = C_S_ListLevel4 Or _
            para.Range.ListFormat.ListType <> wdListNoNumbering And para.style = C_S_ParNormal Then
            ' Check if the paragraph starts with the specified text
            If Left(para.Range.Text, textLen) = textToRemove Then
                ' Remove the specified text from the beginning of the paragraph
                para.Range.Characters(1).Delete Unit:=wdCharacter, count:=textLen
            End If
        End If
    Next para
End Sub

' 2025-03-20 by ms and AI
' 2025-08-07 by ms
' 2026-01-15 by ms
Sub ToggleCharBoldStyle()
    Dim CurrentStyle As String
    
    ' Surprisingly this is enough to proceed further. If user selects few paragraphs with different styling, then Selectio.style.NameLocal is empty.
    On Error Resume Next
        CurrentStyle = Selection.style.NameLocal
    On Error GoTo 0
    
    Dim FileName As String:     FileName = C_F_Macros
    Dim ModuleName As String:   ModuleName = C_M_Styles
    Dim MacroName As String:    MacroName = "ToggleCharBoldStyle"
    Dim MsgBoxTitle As String:  MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    ' Check if the styles exist
    If Not StyleExists(C_S_Bold) Or Not StyleExists(C_S_CharDefault) Then
        MsgBox _
            Prompt:="One or both of the required styles do not exist in this document:" & vbNewLine & vbNewLine & _
                C_S_Bold & " or " & C_S_CharDefault, _
            Buttons:=vbExclamation, _
            Title:=MsgBoxTitle
        Exit Sub
    End If
    
    Dim rng As Word.Range
    Set rng = Selection.Range
    
    ' Toggle styles
    If rng.font.Bold = False Then
        ' TURN ON: Apply the specific Character Style
        rng.style = ActiveDocument.Styles(C_S_Bold)
        Application.statusBar = MsgBoxTitle & " > Applied: " & C_S_Bold
    Else
        ' TURN OFF: Clear back to Default Paragraph Font
        ' Note: Using wdStyleDefaultParagraphFont is safer than a custom string
        rng.style = ActiveDocument.Styles(wdStyleDefaultParagraphFont)
        ' IMPORTANT: Remove direct italic formatting in case it was applied manually
        rng.font.Bold = False
        Application.statusBar = MsgBoxTitle & " > Reset to default paragraph style font"
    End If
    
    Set rng = Nothing
    
End Sub

' 2025-08-05 by ms and ai
Function StyleExists(styleName As String) As Boolean
    On Error Resume Next
    StyleExists = Not ActiveDocument.Styles(styleName) Is Nothing
    On Error GoTo 0
End Function

' 2025-03-20 by ms and AI
' 2025-08-06 by ms
' 2026-01-14 by ms
Sub ToggleCharItalicStyle()
    Dim CurrentStyle As String
    
    ' Surprisingly this is enough to proceed further. If user selects few paragraphs with different styling, then Selectio.style.NameLocal is empty.
    On Error Resume Next
        CurrentStyle = Selection.style.NameLocal
    On Error GoTo 0
    
    Dim FileName As String:      FileName = C_F_Macros
    Dim ModuleName As String:    ModuleName = C_M_Styles
    Dim MacroName As String:     MacroName = "ToggleCharItalicStyle"
    Dim MsgBoxTitle As String:   MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    ' Check if the styles exist
    If Not StyleExists(C_S_Italic) Or Not StyleExists(C_S_CharDefault) Then
        MsgBox _
            Prompt:="One or both of the required styles do not exist in this document:" & vbNewLine & vbNewLine & _
                C_S_Italic & " or " & C_S_CharDefault, _
            Buttons:=vbExclamation, _
            Title:=MsgBoxTitle
        Exit Sub
    End If
    
    Dim rng As Word.Range
    Set rng = Selection.Range
    
    ' Toggle styles
    If rng.font.Italic = False Then
        ' TURN ON: Apply the specific Character Style
        rng.style = ActiveDocument.Styles(C_S_Italic)
        Application.statusBar = MsgBoxTitle & " > Applied: " & C_S_Italic
    Else
        ' TURN OFF: Clear back to Default Paragraph Font
        ' Note: Using wdStyleDefaultParagraphFont is safer than a custom string
        rng.style = ActiveDocument.Styles(wdStyleDefaultParagraphFont)
        ' IMPORTANT: Remove direct italic formatting in case it was applied manually
        rng.font.Italic = False
        Application.statusBar = MsgBoxTitle & " > Reset to default paragraph style font"
    End If
    
    Set rng = Nothing
End Sub

' 2025-03-20 by ms and AI
' 2025-08-06 by ms
' 2026-01-14 by ms and AI
Sub ToggleCharUnderlineStyle()
    Dim FileName As String:       FileName = C_F_Macros
    Dim ModuleName As String:     ModuleName = C_M_Styles
    Dim MacroName As String:      MacroName = "ToggleCharUnderlineStyle"
    Dim MsgBoxTitle As String:    MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    ' Check if the styles exist
    If Not StyleExists(C_S_Underline) Or Not StyleExists(C_S_CharDefault) Then
        MsgBox _
            Prompt:="One or both of the required styles do not exist in this document:" & vbNewLine & vbNewLine & _
                C_S_Underline & " or " & C_S_CharDefault, _
            Buttons:=vbExclamation, _
            Title:=MsgBoxTitle
        Exit Sub
    End If
    
    Dim rng As Word.Range
    Set rng = Selection.Range
    
    ' Toggle styles
    Select Case rng.font.Underline
        Case wdUnderlineNone, wdUndefined
            ' turn on: apply the character style that only sets underline
            rng.style = ActiveDocument.Styles(C_S_Underline)
            Application.statusBar = MsgBoxTitle & " > " & C_S_Underline
        Case Else
            ' turn off: clear the character-style overlay back to "default char"
            rng.style = ActiveDocument.Styles(C_S_CharDefault)
            ' ensure direct underline is off (if user applied underline directly)
            rng.font.Underline = wdUnderlineNone
            Application.statusBar = MsgBoxTitle & " > Reset to default paragraph style font"
    End Select
    Set rng = Nothing
End Sub

' 2025-03-21 by ms
' 2025-08-06 by ms
' 2026-01-15 by ms
Sub ToggleCharCrossoutStyle()
    Dim CurrentStyle As String
    
    ' Surprisingly this is enough to proceed further. If user selects few paragraphs with different styling, then Selectio.style.NameLocal is empty.
    On Error Resume Next
        CurrentStyle = Selection.style.NameLocal
    On Error GoTo 0
    
    Dim FileName As String:     FileName = C_F_Macros
    Dim ModuleName As String:   ModuleName = C_M_Styles
    Dim MacroName As String:    MacroName = "ToggleCharCrossoutStyle"
    Dim MsgBoxTitle As String:  MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    ' Check if the styles exist
    If Not StyleExists(C_S_CharCrossout) Or Not StyleExists(C_S_CharDefault) Then
        MsgBox _
            Prompt:="One or both of the required styles do not exist in this document:" & vbNewLine & vbNewLine & _
                C_S_CharCrossout & " or " & C_S_CharDefault, _
            Buttons:=vbExclamation, _
            Title:=MsgBoxTitle
        Exit Sub
    End If
    
    Dim rng As Word.Range
    Set rng = Selection.Range
    
    ' Toggle styles
    If rng.font.Strikethrough = False Then
        ' TURN ON: Apply the specific Character Style
        rng.style = ActiveDocument.Styles(C_S_CharCrossout)
        Application.statusBar = MsgBoxTitle & " > Applied: " & C_S_CharCrossout
    Else
        ' TURN OFF: Clear back to Default Paragraph Font
        ' Note: Using wdStyleDefaultParagraphFont is safer than a custom string
        rng.style = ActiveDocument.Styles(wdStyleDefaultParagraphFont)
        ' IMPORTANT: Remove direct italic formatting in case it was applied manually
        rng.font.Strikethrough = False
        Application.statusBar = MsgBoxTitle & " > Reset to default paragraph style font"
    End If
    
    Set rng = Nothing
End Sub

' Trick is, only extra Sub enableschange of the background shading. The VBA for character styles does not.
' 2025-03-21 by ms
' 2025-08-06 by ms
' 2026-01-15 by ms
Sub ToggleCharHiddenStyle()
    Dim CurrentStyle As String
    
    ' Surprisingly this is enough to proceed further. If user selects few paragraphs with different styling, then Selectio.style.NameLocal is empty.
    On Error Resume Next
        CurrentStyle = Selection.style.NameLocal
    On Error GoTo 0
    
    Dim FileName As String:     FileName = C_F_Macros
    Dim ModuleName As String:   ModuleName = C_M_Styles
    Dim MacroName As String:    MacroName = "ToggleCharHiddenStyle"
    Dim MsgBoxTitle As String:  MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    ' Check if the styles exist
    If Not StyleExists(C_S_CharHidden) Or Not StyleExists(C_S_CharDefault) Then
        MsgBox _
            Prompt:="One or both of the required styles do not exist in this document:" & vbNewLine & vbNewLine & _
                C_S_CharHidden & " or " & C_S_CharDefault, _
            Buttons:=vbExclamation, _
            Title:=MsgBoxTitle
        Exit Sub
    End If
    
    Dim rng As Word.Range
    Set rng = Selection.Range
    
    ' Toggle styles
    If rng.font.Hidden = False Then
        ' TURN ON: Apply the specific Character Style
        rng.style = ActiveDocument.Styles(C_S_CharHidden)
        Selection.Range.shading.BackgroundPatternColor = RGB(246, 192, 192)
        Application.statusBar = MsgBoxTitle & " > Applied: " & C_S_CharHidden
    Else
        ' TURN OFF: Clear back to Default Paragraph Font
        ' Note: Using wdStyleDefaultParagraphFont is safer than a custom string
        rng.style = ActiveDocument.Styles(wdStyleDefaultParagraphFont)
        ' IMPORTANT: Remove direct italic formatting in case it was applied manually
        rng.font.Hidden = False
        Application.statusBar = MsgBoxTitle & " > Reset to default paragraph style font"
    End If
    
    Set rng = Nothing
End Sub

' 2025-04-14 by ms
' 2025-08-06 by ms
' 2026-01-15 by ms
Sub ToggleCharSourceCode()
    Dim CurrentStyle As String
    
    ' Surprisingly this is enough to proceed further. If user selects few paragraphs with different styling, then Selectio.style.NameLocal is empty.
    On Error Resume Next
        CurrentStyle = Selection.style.NameLocal
    On Error GoTo 0
    
    Dim FileName As String:     FileName = C_F_Macros
    Dim ModuleName As String:   ModuleName = C_M_Styles
    Dim MacroName As String:    MacroName = "ToggleCharSourceCode"
    Dim MsgBoxTitle As String:  MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    ' Check if the styles exist
    If Not StyleExists(C_S_CharSourceCode) Or Not StyleExists(C_S_CharDefault) Then
        MsgBox _
            Prompt:="One or both of the required styles do not exist in this document:" & vbNewLine & vbNewLine & _
                C_S_CharSourceCode & " or " & C_S_CharDefault, _
            Buttons:=vbExclamation, _
            Title:=MsgBoxTitle
        Exit Sub
    End If
    
    ' Toggle styles
    Dim rng As Word.Range
    Set rng = Selection.Range
    
    ' Toggle styles
    If rng.style <> C_S_CharSourceCode Then
        ' TURN ON: Apply the specific Character Style
        rng.style = ActiveDocument.Styles(C_S_CharSourceCode)
        Application.statusBar = MsgBoxTitle & " > Applied: " & C_S_CharSourceCode
    Else
        ' TURN OFF: Clear back to Default Paragraph Font
        ' Note: Using wdStyleDefaultParagraphFont is safer than a custom string
        rng.style = ActiveDocument.Styles(wdStyleDefaultParagraphFont)
        ' IMPORTANT: Remove direct italic formatting in case it was applied manually
        Application.statusBar = MsgBoxTitle & " > Reset to default paragraph style font"
    End If
    
    Set rng = Nothing
    
End Sub

' 2025-08-07 by ms
' Switches off autoupdate feature to all styles in the template file.
Sub SwitchOffAutoupdate()
    Dim aStyle As style

    ' The following line is necessary as not all types of styles have property AutomaticallyUpdate.
    On Error Resume Next
    For Each aStyle In ActiveDocument.Styles
        Let aStyle.AutomaticallyUpdate = False
    Next aStyle
    Set aStyle = Nothing
    On Error GoTo -1
    
    Dim FileName As String:     FileName = C_F_Macros
    Dim ModuleName As String:   ModuleName = C_M_Styles
    Dim MacroName As String:    MacroName = "SwitchOffAutoupdate"
    Dim MsgBoxTitle As String:  MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    MsgBox _
        Prompt:="End of processing.", _
        Buttons:=vbInformation + vbOKOnly, _
        Title:=MsgBoxTitle
End Sub

' Order matters. At first the most basic styles must be created (C_S_ParNormal)
' 2025-11-16 by ms
' 2025-12-28 by ms
' 2026-01-15 by ms
Sub CreateCustomStyles()
    Dim IsSuccessful As Boolean
    IsSuccessful = False
    Dim CharCounter As Byte
    CharCounter = 0
    Dim ParagraphCounter As Byte
    ParagraphCounter = 0
    Dim ListTemplateCounter As Byte
    ListTemplateCounter = 0
    Dim TabCounter As Byte
    TabCounter = 0
    
    Dim FileName As String:     FileName = C_F_Macros
    Dim ModuleName As String:   ModuleName = C_M_Styles
    Dim MacroName As String:    MacroName = "CreateCustomStyles"
    Dim MsgBoxTitle As String:  MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    ' Paragraph Styles:
    IsSuccessful = CreateStyle_Normal()
    If IsSuccessful = False Then
        MsgBox _
            Prompt:="Error on time of creation the paragraph style " & vbNewLine & "Normal" & vbNewLine & ". Exiting.", _
            Buttons:=vbCritical, _
            Title:=MsgBoxTitle
        Exit Sub
    Else
        ParagraphCounter = ParagraphCounter + 1
    End If
    
    IsSuccessful = CreateStyle_ParNormalMs()
    If IsSuccessful = False Then
        MsgBox _
            Prompt:="Error on time of creation the paragraph style " & vbNewLine & C_S_ParNormal & vbNewLine & ". Exiting.", _
            Buttons:=vbCritical, _
            Title:=MsgBoxTitle
        Exit Sub
    Else
        ParagraphCounter = ParagraphCounter + 1
    End If
        
    IsSuccessful = CreateStyle_ParHeading1ms()
    If IsSuccessful = False Then
        MsgBox _
            Prompt:="Error on time of creation the paragraph style " & vbNewLine & C_S_Heading1 & vbNewLine & ". Exiting.", _
            Buttons:=vbCritical, _
            Title:=MsgBoxTitle
        Exit Sub
    Else
        ParagraphCounter = ParagraphCounter + 1
    End If
    
    IsSuccessful = CreateStyle_ParHeading2ms
    If IsSuccessful = False Then
        MsgBox _
            Prompt:="Error on time of creation the paragraph style " & vbNewLine & C_S_Heading2 & vbNewLine & ". Exiting.", _
            Buttons:=vbCritical, _
            Title:=MsgBoxTitle
        Exit Sub
    Else
        ParagraphCounter = ParagraphCounter + 1
    End If

    IsSuccessful = CreateStyle_ParHeading3ms
    If IsSuccessful = False Then
        MsgBox _
            Prompt:="Error on time of creation the paragraph style " & vbNewLine & C_S_Heading3 & vbNewLine & ". Exiting.", _
            Buttons:=vbCritical, _
            Title:=MsgBoxTitle
        Exit Sub
    Else
        ParagraphCounter = ParagraphCounter + 1
    End If
    
    IsSuccessful = CreateStyle_ParHeading4ms
    If IsSuccessful = False Then
        MsgBox _
            Prompt:="Error on time of creation the paragraph style " & vbNewLine & C_S_Heading4 & vbNewLine & ". Exiting.", _
            Buttons:=vbCritical, _
            Title:=MsgBoxTitle
        Exit Sub
    Else
        ParagraphCounter = ParagraphCounter + 1
    End If

    IsSuccessful = CreateStyle_ParHeading5ms
    If IsSuccessful = False Then
        MsgBox _
            Prompt:="Error on time of creation the paragraph style " & vbNewLine & C_S_Heading5 & vbNewLine & ". Exiting.", _
            Buttons:=vbCritical, _
            Title:=MsgBoxTitle
        Exit Sub
    Else
        ParagraphCounter = ParagraphCounter + 1
    End If

    IsSuccessful = CreateStyle_ParHeading6ms
    If IsSuccessful = False Then
        MsgBox _
            Prompt:="Error on time of creation the paragraph style " & vbNewLine & C_S_Heading6 & vbNewLine & ". Exiting.", _
            Buttons:=vbCritical, _
            Title:=MsgBoxTitle
        Exit Sub
    Else
        ParagraphCounter = ParagraphCounter + 1
    End If

    IsSuccessful = CreateStyle_ParHeading7ms
    If IsSuccessful = False Then
        MsgBox _
            Prompt:="Error on time of creation the paragraph style " & vbNewLine & C_S_Heading7 & vbNewLine & ". Exiting.", _
            Buttons:=vbCritical, _
            Title:=MsgBoxTitle
        Exit Sub
    Else
        ParagraphCounter = ParagraphCounter + 1
    End If

    IsSuccessful = CreateStyle_ParHeading8ms
    If IsSuccessful = False Then
        MsgBox _
            Prompt:="Error on time of creation the paragraph style " & vbNewLine & C_S_Heading8 & vbNewLine & ". Exiting.", _
            Buttons:=vbCritical, _
            Title:=MsgBoxTitle
        Exit Sub
    Else
        ParagraphCounter = ParagraphCounter + 1
    End If
    
    IsSuccessful = CreateStyle_ParInTableMs
    If IsSuccessful = False Then
        MsgBox _
            Prompt:="Error on time of creation the paragraph style " & vbNewLine & C_S_ParInTable & vbNewLine & ". Exiting.", _
            Buttons:=vbCritical, _
            Title:=MsgBoxTitle
        Exit Sub
    Else
        ParagraphCounter = ParagraphCounter + 1
    End If
    
    IsSuccessful = CreateStyle_ParLegalNoteMs
    If IsSuccessful = False Then
        MsgBox _
            Prompt:="Error on time of creation the paragraph style " & vbNewLine & C_S_ParLegalNote & vbNewLine & ". Exiting.", _
            Buttons:=vbCritical, _
            Title:=MsgBoxTitle
        Exit Sub
    Else
        ParagraphCounter = ParagraphCounter + 1
    End If

    IsSuccessful = CreateStyle_ParLegendPictureMs
    If IsSuccessful = False Then
        MsgBox _
            Prompt:="Error on time of creation the paragraph style " & vbNewLine & C_S_PictureLegend & vbNewLine & ". Exiting.", _
            Buttons:=vbCritical, _
            Title:=MsgBoxTitle
        Exit Sub
    Else
        ParagraphCounter = ParagraphCounter + 1
    End If
    
    IsSuccessful = CreateStyle_ParLegendTableMs
    If IsSuccessful = False Then
        MsgBox _
            Prompt:="Error on time of creation the paragraph style " & vbNewLine & C_S_TableLegend & vbNewLine & ". Exiting.", _
            Buttons:=vbCritical, _
            Title:=MsgBoxTitle
        Exit Sub
    Else
        ParagraphCounter = ParagraphCounter + 1
    End If

    IsSuccessful = CreateStyle_ParListHeadingMs
    If IsSuccessful = False Then
        MsgBox _
            Prompt:="Error on time of creation the paragraph style " & vbNewLine & C_S_ListHeading & vbNewLine & ". Exiting.", _
            Buttons:=vbCritical, _
            Title:=MsgBoxTitle
        Exit Sub
    Else
        ParagraphCounter = ParagraphCounter + 1
    End If

    IsSuccessful = CreateStyle_ParListIndent1Ms
    If IsSuccessful = False Then
        MsgBox _
            Prompt:="Error on time of creation the paragraph style " & vbNewLine & C_S_ListLevel1 & vbNewLine & ". Exiting.", _
            Buttons:=vbCritical, _
            Title:=MsgBoxTitle
        Exit Sub
    Else
        ParagraphCounter = ParagraphCounter + 1
    End If
    
    IsSuccessful = CreateStyle_ParListIndentB1Ms
    If IsSuccessful = False Then
        MsgBox _
            Prompt:="Error on time of creation the paragraph style " & vbNewLine & C_S_ListLevelB1 & vbNewLine & ". Exiting.", _
            Buttons:=vbCritical, _
            Title:=MsgBoxTitle
        Exit Sub
    Else
        ParagraphCounter = ParagraphCounter + 1
    End If
        
    IsSuccessful = CreateStyle_ParListIndent2Ms
    If IsSuccessful = False Then
        MsgBox _
            Prompt:="Error on time of creation the paragraph style " & vbNewLine & C_S_ListLevel2 & vbNewLine & ". Exiting.", _
            Buttons:=vbCritical, _
            Title:=MsgBoxTitle
        Exit Sub
    Else
        ParagraphCounter = ParagraphCounter + 1
    End If
    
    IsSuccessful = CreateStyle_ParListIndentB2Ms
    If IsSuccessful = False Then
        MsgBox _
            Prompt:="Error on time of creation the paragraph style " & vbNewLine & C_S_ListLevelB2 & vbNewLine & ". Exiting.", _
            Buttons:=vbCritical, _
            Title:=MsgBoxTitle
        Exit Sub
    Else
        ParagraphCounter = ParagraphCounter + 1
    End If
    
    IsSuccessful = CreateStyle_ParListIndent3Ms
    If IsSuccessful = False Then
        MsgBox _
            Prompt:="Error on time of creation the paragraph style " & vbNewLine & C_S_ListLevel3 & vbNewLine & ". Exiting.", _
            Buttons:=vbCritical, _
            Title:=MsgBoxTitle
        Exit Sub
    Else
        ParagraphCounter = ParagraphCounter + 1
    End If
    
    IsSuccessful = CreateStyle_ParListIndentB3Ms
    If IsSuccessful = False Then
        MsgBox _
            Prompt:="Error on time of creation the paragraph style " & vbNewLine & C_S_ListLevelB3 & vbNewLine & ". Exiting.", _
            Buttons:=vbCritical, _
            Title:=MsgBoxTitle
        Exit Sub
    Else
        ParagraphCounter = ParagraphCounter + 1
    End If
    
    IsSuccessful = CreateStyle_ParListIndent4Ms
    If IsSuccessful = False Then
        MsgBox _
            Prompt:="Error on time of creation the paragraph style " & vbNewLine & C_S_ListLevel4 & vbNewLine & ". Exiting.", _
            Buttons:=vbCritical, _
            Title:=MsgBoxTitle
        Exit Sub
    Else
        ParagraphCounter = ParagraphCounter + 1
    End If

    IsSuccessful = CreateStyle_ParListIndentB4Ms
    If IsSuccessful = False Then
        MsgBox _
            Prompt:="Error on time of creation the paragraph style " & vbNewLine & C_S_ListLevelB4 & vbNewLine & ". Exiting.", _
            Buttons:=vbCritical, _
            Title:=MsgBoxTitle
        Exit Sub
    Else
        ParagraphCounter = ParagraphCounter + 1
    End If

    IsSuccessful = CreateStyle_ParMinimalMs
    If IsSuccessful = False Then
        MsgBox _
            Prompt:="Error on time of creation the paragraph style " & vbNewLine & C_S_ParMinimal & vbNewLine & ". Exiting.", _
            Buttons:=vbCritical, _
            Title:=MsgBoxTitle
        Exit Sub
    Else
        ParagraphCounter = ParagraphCounter + 1
    End If

    IsSuccessful = CreateStyle_ParNormalAboveMs
    If IsSuccessful = False Then
        MsgBox _
            Prompt:="Error on time of creation the paragraph style " & vbNewLine & C_S_ParNormalAbove & vbNewLine & ". Exiting.", _
            Buttons:=vbCritical, _
            Title:=MsgBoxTitle
        Exit Sub
    Else
        ParagraphCounter = ParagraphCounter + 1
    End If
    
    IsSuccessful = CreateStyle_ParNormalBelowMs
    If IsSuccessful = False Then
        MsgBox _
            Prompt:="Error on time of creation the paragraph style " & vbNewLine & C_S_ParNormalBelow & vbNewLine & ". Exiting.", _
            Buttons:=vbCritical, _
            Title:=MsgBoxTitle
        Exit Sub
    Else
        ParagraphCounter = ParagraphCounter + 1
    End If
    
    IsSuccessful = CreateStyle_ParNormalAboveBelowMs
    If IsSuccessful = False Then
        MsgBox _
            Prompt:="Error on time of creation the paragraph style " & vbNewLine & C_S_ParNormalAB & vbNewLine & ". Exiting.", _
            Buttons:=vbCritical, _
            Title:=MsgBoxTitle
        Exit Sub
    Else
        ParagraphCounter = ParagraphCounter + 1
    End If
    
    IsSuccessful = CreateStyle_ParNormalZeroMs
    If IsSuccessful = False Then
        MsgBox _
            Prompt:="Error on time of creation the paragraph style " & vbNewLine & C_S_ParNormalZero & vbNewLine & ". Exiting.", _
            Buttons:=vbCritical, _
            Title:=MsgBoxTitle
        Exit Sub
    Else
        ParagraphCounter = ParagraphCounter + 1
    End If
    
    IsSuccessful = CreateStyle_ParPictureCanvaMs
    If IsSuccessful = False Then
        MsgBox _
            Prompt:="Error on time of creation the paragraph style " & vbNewLine & C_S_ParPictureCanva & vbNewLine & ". Exiting.", _
            Buttons:=vbCritical, _
            Title:=MsgBoxTitle
        Exit Sub
    Else
        ParagraphCounter = ParagraphCounter + 1
    End If
    
    IsSuccessful = CreateStyle_ParSourceCodeMs
    If IsSuccessful = False Then
        MsgBox _
            Prompt:="Error on time of creation the paragraph style " & vbNewLine & C_S_ParSourceCode & vbNewLine & ". Exiting.", _
            Buttons:=vbCritical, _
            Title:=MsgBoxTitle
        Exit Sub
    Else
        ParagraphCounter = ParagraphCounter + 1
    End If
    
    IsSuccessful = CreateStyle_TOC1
    If IsSuccessful = False Then
        MsgBox _
            Prompt:="Error on time of creation the paragraph style " & vbNewLine & wdStyleTOC1 & vbNewLine & ". Exiting.", _
            Buttons:=vbCritical, _
            Title:=MsgBoxTitle
        Exit Sub
    Else
        ParagraphCounter = ParagraphCounter + 1
    End If
    
    IsSuccessful = CreateStyle_TOC2
    If IsSuccessful = False Then
        MsgBox _
            Prompt:="Error on time of creation the paragraph style " & vbNewLine & wdStyleTOC2 & vbNewLine & ". Exiting.", _
            Buttons:=vbCritical, _
            Title:=MsgBoxTitle
        Exit Sub
    Else
        ParagraphCounter = ParagraphCounter + 1
    End If
    
    IsSuccessful = CreateStyle_TOC3
    If IsSuccessful = False Then
        MsgBox _
            Prompt:="Error on time of creation the paragraph style " & vbNewLine & wdStyleTOC3 & vbNewLine & ". Exiting.", _
            Buttons:=vbCritical, _
            Title:=MsgBoxTitle
        Exit Sub
    Else
        ParagraphCounter = ParagraphCounter + 1
    End If
    
    IsSuccessful = CreateStyle_ParTextBoxesMs
    If IsSuccessful = False Then
        MsgBox _
            Prompt:="Error on time of creation the paragraph style " & vbNewLine & C_S_TextBoxes & vbNewLine & ". Exiting.", _
            Buttons:=vbCritical, _
            Title:=MsgBoxTitle
        Exit Sub
    Else
        ParagraphCounter = ParagraphCounter + 1
    End If
   
   IsSuccessful = CreateStyle_ParNumRefMs
    If IsSuccessful = False Then
        MsgBox _
            Prompt:="Error on time of creation the paragraph style " & vbNewLine & C_S_ListNumRef & vbNewLine & ". Exiting.", _
            Buttons:=vbCritical, _
            Title:=MsgBoxTitle
        Exit Sub
    Else
        ParagraphCounter = ParagraphCounter + 1
    End If
   
    IsSuccessful = CreateStyle_ParListInTableMs()
    If IsSuccessful = False Then
        MsgBox _
            Prompt:="Error on time of creation the character style " & vbNewLine & C_S_ListNumTable & vbNewLine & ". Exiting.", _
            Buttons:=vbCritical, _
            Title:=MsgBoxTitle
        Exit Sub
    Else
        ParagraphCounter = ParagraphCounter + 1
    End If
    
    IsSuccessful = CreateStyle_ParIconMs()
    If IsSuccessful = False Then
        MsgBox _
            Prompt:="Error on time of creation the character style " & vbNewLine & C_S_ParIcon & vbNewLine & ". Exiting.", _
            Buttons:=vbCritical, _
            Title:=MsgBoxTitle
        Exit Sub
    Else
        ParagraphCounter = ParagraphCounter + 1
    End If
   
    ' Character Styles:
    IsSuccessful = CreateStyle_CharBoldMs()
    If IsSuccessful = False Then
        MsgBox _
            Prompt:="Error on time of creation the character style " & vbNewLine & C_S_Bold & vbNewLine & ". Exiting.", _
            Buttons:=vbCritical, _
            Title:=MsgBoxTitle
        Exit Sub
    Else
        CharCounter = CharCounter + 1
    End If
    
    IsSuccessful = CreateStyle_CharCrossoutMs()
    If IsSuccessful = False Then
        MsgBox _
            Prompt:="Error on time of creation the character style " & vbNewLine & C_S_CharCrossout & vbNewLine & ". Exiting.", _
            Buttons:=vbCritical, _
            Title:=MsgBoxTitle
        Exit Sub
    Else
        CharCounter = CharCounter + 1
    End If
    
    IsSuccessful = CreateStyle_CharDefaultMs()
    If IsSuccessful = False Then
        MsgBox _
            Prompt:="Error on time of creation the character style " & vbNewLine & C_S_CharDefault & vbNewLine & ". Exiting.", _
            Buttons:=vbCritical, _
            Title:=MsgBoxTitle
        Exit Sub
    Else
        CharCounter = CharCounter + 1
    End If
    
    IsSuccessful = CreateStyle_CharHiddenMs()
    If IsSuccessful = False Then
        MsgBox _
            Prompt:="Error on time of creation the character style " & vbNewLine & C_S_CharHidden & vbNewLine & ". Exiting.", _
            Buttons:=vbCritical, _
            Title:=MsgBoxTitle
        Exit Sub
    Else
        CharCounter = CharCounter + 1
    End If
    
    IsSuccessful = CreateStyle_CharItalicMs()
    If IsSuccessful = False Then
        MsgBox _
            Prompt:="Error on time of creation the character style " & vbNewLine & C_S_Italic & vbNewLine & ". Exiting.", _
            Buttons:=vbCritical, _
            Title:=MsgBoxTitle
        
        Exit Sub
    Else
        CharCounter = CharCounter + 1
    End If
        
    IsSuccessful = CreateStyle_CharLegalNoteMs()
    If IsSuccessful = False Then
        MsgBox _
            Prompt:="Error on time of creation the character style " & vbNewLine & C_S_CharLegalNote & vbNewLine & ". Exiting.", _
            Buttons:=vbCritical, _
            Title:=MsgBoxTitle
        Exit Sub
    Else
        CharCounter = CharCounter + 1
    End If
    
    IsSuccessful = CreateStyle_CharSourceCodeMs()
    If IsSuccessful = False Then
        MsgBox _
            Prompt:="Error on time of creation the character style " & vbNewLine & C_S_CharSourceCode & vbNewLine & ". Exiting.", _
            Buttons:=vbCritical, _
            Title:=MsgBoxTitle
        Exit Sub
    Else
        CharCounter = CharCounter + 1
    End If
    
    IsSuccessful = CreateStyle_CharUnderlineMs()
    If IsSuccessful = False Then
        MsgBox _
            Prompt:="Error on time of creation the character style " & vbNewLine & C_S_Underline & vbNewLine & ". Exiting.", _
            Buttons:=vbCritical, _
            Title:=MsgBoxTitle
        Exit Sub
    Else
        CharCounter = CharCounter + 1
    End If
        
    ' List Templates:
    IsSuccessful = Create_LT_Headings()
    If IsSuccessful = False Then
        MsgBox _
            Prompt:="Error on time of creation the 'list template' " & vbNewLine & C_LT_Headings & vbNewLine & ". Exiting.", _
            Buttons:=vbCritical, _
            Title:=MsgBoxTitle
        Exit Sub
    Else
        ListTemplateCounter = ListTemplateCounter + 1
    End If
    
    IsSuccessful = Create_LT_NumOrd()
    If IsSuccessful = False Then
        MsgBox _
            Prompt:="Error on time of creation the 'list template' " & vbNewLine & C_LT_NumOrd & vbNewLine & ". Exiting.", _
            Buttons:=vbCritical, _
            Title:=MsgBoxTitle
        Exit Sub
    Else
        ListTemplateCounter = ListTemplateCounter + 1
    End If
    
    IsSuccessful = Create_LT_Bullets()
    If IsSuccessful = False Then
        MsgBox _
            Prompt:="Error on time of creation the 'list template' " & vbNewLine & C_LT_Bullets & vbNewLine & ". Exiting.", _
            Buttons:=vbCritical, _
            Title:=MsgBoxTitle
        Exit Sub
    Else
        ListTemplateCounter = ListTemplateCounter + 1
    End If
    
    IsSuccessful = Create_LT_SingleLevelListNumRefMs()
    If IsSuccessful = False Then
        MsgBox _
            Prompt:="Error on time of creation the 'list template' " & vbNewLine & C_LT_ListNumRef & vbNewLine & ". Exiting.", _
            Buttons:=vbCritical, _
            Title:=MsgBoxTitle
        Exit Sub
    Else
        ListTemplateCounter = ListTemplateCounter + 1
    End If
    
    IsSuccessful = Create_LT_SingleLevelListInTableMs
    If IsSuccessful = False Then
        MsgBox _
            Prompt:="Error on time of creation the 'list template' " & vbNewLine & C_LT_ListNumTable & vbNewLine & ". Exiting.", _
            Buttons:=vbCritical, _
            Title:=MsgBoxTitle
        Exit Sub
    Else
        ListTemplateCounter = ListTemplateCounter + 1
    End If
    
    IsSuccessful = Create_LT_ToC
    If IsSuccessful = False Then
        MsgBox _
            Prompt:="Error on time of creation the 'list template' " & vbNewLine & C_LT_TOC & vbNewLine & ". Exiting.", _
            Buttons:=vbCritical, _
            Title:=MsgBoxTitle
        Exit Sub
    Else
        ListTemplateCounter = ListTemplateCounter + 1
    End If
    
    
    ' Table styles:
    IsSuccessful = CreateStyle_TabTableMs()
    If IsSuccessful = False Then
        MsgBox _
            Prompt:="Error on time of creation the 'table style' " & vbNewLine & C_S_TabTable & vbNewLine & ". Exiting.", _
            Buttons:=vbCritical, _
            Title:=MsgBoxTitle
        Exit Sub
    Else
        TabCounter = TabCounter + 1
    End If

    IsSuccessful = CreateStyle_TabTableNoGridMs()
    If IsSuccessful = False Then
        MsgBox _
            Prompt:="Error on time of creation the 'table style' " & vbNewLine & C_S_TabNoGrid & vbNewLine & ". Exiting.", _
            Buttons:=vbCritical, _
            Title:=MsgBoxTitle
        Exit Sub
    Else
        TabCounter = TabCounter + 1
    End If

    IsSuccessful = CreateStyle_TabTableNoPaddingMs
    If IsSuccessful = False Then
        MsgBox _
            Prompt:="Error on time of creation the 'table style' " & vbNewLine & C_S_TabNoPadding & vbNewLine & ". Exiting.", _
            Buttons:=vbCritical, _
            Title:=MsgBoxTitle
        Exit Sub
    Else
        TabCounter = TabCounter + 1
    End If

    ' summary:
    MsgBox _
        Prompt:="The following number of styles have been added to the current document:" & vbNewLine & _
            ActiveDocument.Name & vbNewLine & vbNewLine & _
            "Character styles: " & CharCounter & vbNewLine & _
            "Paragraph styles: " & ParagraphCounter & vbNewLine & _
            "List templates: " & ListTemplateCounter & vbNewLine & _
            "Table styles: " & TabCounter, _
        Buttons:=vbInformation, _
        Title:=MsgBoxTitle
End Sub

' 2026-01-18 by ms
Sub ResetAllListGalleries()
    Dim i As Integer
    Dim g As Integer
    
    ' Word has 3 main list galleries: 1-bullets, 2-numbered, 3-multileveled
    For g = 1 To 3
        For i = 1 To 7 ' each gallery has 7 slots
            ListGalleries(g).Reset (i)
        Next i
    Next g
    
    Dim FileName As String:     FileName = C_F_Macros
    Dim ModuleName As String:   ModuleName = C_M_Styles
    Dim MacroName As String:    MacroName = "ResetAllListGalleries"
    Dim MsgBoxTitle As String:  MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    MsgBox _
        Prompt:="All List Template galleries are clear now" & vbNewLine, _
        Buttons:=vbInformation + vbOKOnly, _
        Title:=MsgBoxTitle

End Sub

' Shows only defined / named List Templates.
' 2025-11-24 by ms
' 2026-01-19 by ms
Sub ListTemplatesShowNamed()
    Dim lt As ListTemplate
    Dim msg As String
    Dim AllCounter As Integer: AllCounter = 0
    Dim NamedCount As Integer: NamedCount = 0
    Dim i As Integer
    Dim lvl As ListLevel
    
    Dim FileName As String:     FileName = C_F_Macros
    Dim ModuleName As String:   ModuleName = C_M_Styles
    Dim MacroName As String:    MacroName = "ShowListTemplates"
    Dim MsgBoxTitle As String:  MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    msg = "List Templates in Active Document:" & vbNewLine & vbNewLine
    
    For Each lt In ActiveDocument.ListTemplates
        If lt.Name <> "" Then
            NamedCount = NamedCount + 1
            msg = msg & NamedCount & ". LT Name: " & lt.Name & vbNewLine
            
            For i = 1 To lt.ListLevels.count
                Set lvl = lt.ListLevels(i)
                If lvl.LinkedStyle <> "" Then
                    
                    msg = msg & "   - Lvl " & i & " LinkedStyle: " & lvl.LinkedStyle & vbNewLine
                End If
            Next i
            msg = msg & vbNewLine
        Else
            AllCounter = AllCounter + 1
        End If
    Next lt
    
    MsgBox _
        Prompt:=msg & vbNewLine & _
            "All ListTemplates in document: " & AllCounter, _
        Buttons:=vbInformation, _
        Title:=MsgBoxTitle
End Sub

' List, create file, only defined / named List Templates.
' 2025-11-24 by ms
' 2026-01-19 by ms
Sub ListTemplatesListNamed()
    Dim lt As ListTemplate
    Dim msg As String
    Dim AllCounter As Integer: AllCounter = 0
    Dim NamedCount As Integer: NamedCount = 0
    Dim i As Integer
    Dim lvl As ListLevel
    
    Dim FileName As String:     FileName = C_F_Macros
    Dim ModuleName As String:   ModuleName = C_M_Styles
    Dim MacroName As String:    MacroName = "ShowListTemplates"
    Dim MsgBoxTitle As String:  MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    msg = "List Templates in Active Document:" & vbNewLine & vbNewLine
    
    For Each lt In ActiveDocument.ListTemplates
        If lt.Name <> "" Then
            NamedCount = NamedCount + 1
            msg = msg & NamedCount & ". LT Name: " & lt.Name & vbNewLine
            
            For i = 1 To lt.ListLevels.count
                Set lvl = lt.ListLevels(i)
                If lvl.LinkedStyle <> "" Then
                    
                    msg = msg & "   - Lvl " & i & " LinkedStyle: " & lvl.LinkedStyle & vbNewLine
                End If
            Next i
            msg = msg & vbNewLine
        Else
            AllCounter = AllCounter + 1
        End If
    Next lt
    
    MsgBox _
        Prompt:=msg & vbNewLine & _
            "All ListTemplates in document: " & AllCounter, _
        Buttons:=vbInformation, _
        Title:=MsgBoxTitle
End Sub


' Unlink specific style from ListTemplate
' 2026-01-21 by ms
Sub ListTemplatesUnlinkStyle()
    ActiveDocument.Styles("List Paragraph").LinkToListTemplate Nothing
End Sub

' 2025-11-25 by ms
' 2026-01-18 by ms
Public Function Create_LT_SingleLevelListInTableMs() As Boolean
    Dim NewStyle As style
    Dim ListTemplate As ListTemplate
    
    Dim FileName As String:      FileName = C_F_Macros
    Dim ModuleName As String:    ModuleName = C_M_Styles
    Dim MacroName As String:     MacroName = "Create_LT_SingleLevelListInTableMs"
    Dim MsgBoxTitle As String:   MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    ' Check if template already exists in gallery
    Dim lt As ListTemplate
    Dim FlagFound As Boolean
    FlagFound = False
    
    For Each lt In ActiveDocument.ListTemplates
        If lt.Name = C_LT_ListNumTable Then
            Set ListTemplate = lt
            FlagFound = True
            Exit For
        End If
    Next lt
    
    ' If not found, create new template in Outline Numbered gallery
    If Not FlagFound Then
        Set ListTemplate = ActiveDocument.ListTemplates.Add( _
            Name:=C_LT_ListNumTable, _
            OutlineNumbered:=False) ' OutlineNumbered = False, creates a simple list template, not multilevel
    End If

   ' Configure Level 1
    With ListTemplate.ListLevels(1)
        .NumberFormat = "%1."                           ' custom format: 1., 2., 3., ...
        .NumberStyle = wdListNumberStyleArabic
        .Alignment = wdListLevelAlignLeft
        .TrailingCharacter = wdTrailingNone             ' wdTrailingNone | wdTrailingSpace | wdTrailingTab
        .NumberPosition = CentimetersToPoints(0)
        .TextPosition = CentimetersToPoints(0)
        '.TabPosition = .TextPosition
        .StartAt = 1
        .LinkedStyle = C_S_ListNumTable
        .font.Name = C_FT_Body
        .font.Size = C_BaseFontSize
        .font.color = wdColorAutomatic
    End With
    
    Create_LT_SingleLevelListInTableMs = True
End Function

' 2025-11-25 by ms
' 2026-01-18 by ms
Public Function Create_LT_SingleLevelListNumRefMs() As Boolean
    Dim NewStyle As style
    Dim ListTemplate As ListTemplate
    
    Dim FileName As String:      FileName = C_F_Macros
    Dim ModuleName As String:    ModuleName = C_M_Styles
    Dim MacroName As String:     MacroName = "Create_LT_SingleLevelListNumRefMs"
    Dim MsgBoxTitle As String:   MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    ' Check if template already exists in gallery
    Dim lt As ListTemplate
    Dim FlagFound As Boolean
    FlagFound = False
    
    For Each lt In ActiveDocument.ListTemplates
        If lt.Name = C_LT_ListNumRef Then
            Set ListTemplate = lt
            FlagFound = True
            Exit For
        End If
    Next lt
    
    ' If not found, create new template in Outline Numbered gallery
    If Not FlagFound Then
        Set ListTemplate = ActiveDocument.ListTemplates.Add( _
            Name:=C_LT_ListNumRef, _
            OutlineNumbered:=False) ' OutlineNumbered = False, creates a simple list template, not multilevel
    End If

   ' Configure Level 1
    With ListTemplate.ListLevels(1)
        .NumberFormat = "[%1]."                 ' custom format: [1]., [2]., [3]., ...
        .NumberStyle = wdListNumberStyleArabic
        .Alignment = wdListLevelAlignLeft
        .TrailingCharacter = wdTrailingTab     ' wdTrailingNone | wdTrailingSpace | wdTrailingTab
        .NumberPosition = CentimetersToPoints(0)
        .TextPosition = CentimetersToPoints(C_ListTabPosL1)
        .TabPosition = .TextPosition
        .StartAt = 1
        .LinkedStyle = C_S_ListNumRef
        .font.Name = C_FT_Body
        .font.Size = C_BaseFontSize
        .font.color = wdColorAutomatic
    End With
    
    Create_LT_SingleLevelListNumRefMs = True
End Function

' 2025-11-24 by ms
' 2026-01-18 by ms
Public Function Create_LT_Bullets() As Boolean
    Dim ListTemplate As ListTemplate
    Dim NumberFormat As String
    Dim StyleNames As Variant
            
    ' Define heading styles for each level
    StyleNames = Array( _
        C_S_ListLevelB1, _
        C_S_ListLevelB2, _
        C_S_ListLevelB3, _
        C_S_ListLevelB4)
    
    ' Check if template already exists in gallery
    Dim lt As ListTemplate
    Dim FlagFound As Boolean
    FlagFound = False
    
    For Each lt In ActiveDocument.ListTemplates
        If lt.Name = C_LT_Bullets Then
            Set ListTemplate = lt
            FlagFound = True
            Exit For
        End If
    Next lt
    
    ' If not found, create new template in Outline Numbered gallery
    If Not FlagFound Then
        Set ListTemplate = ActiveDocument.ListTemplates.Add( _
            Name:=C_LT_Bullets, _
            OutlineNumbered:=True)
    End If
    
    ' Configure each level
    With ListTemplate.ListLevels(1)
        .NumberFormat = ChrW(&HBB)                  ' Unicode for » (solid circle bullet)
        .NumberStyle = wdListNumberStyleBullet
        .Alignment = wdListLevelAlignLeft
        .TrailingCharacter = wdTrailingTab             ' wdTrailingNone | wdTrailingSpace | wdTrailingTab
        .NumberPosition = CentimetersToPoints(C_BaseIndent)
        .TextPosition = CentimetersToPoints(C_ListTabPosL1)
        .TabPosition = .TextPosition
        .StartAt = 1
        .LinkedStyle = StyleNames(0)
        .font.Name = C_FT_Body
        .font.Size = C_BaseFontSize
        .font.color = GetThemeColor(ThemeColorIndex:=wdThemeColorAccent1, TintAndShade:=0)  ' in module Theme
    End With
        
    With ListTemplate.ListLevels(2)
        .NumberFormat = ChrW(&HBB)                  ' Unicode for » (solid circle bullet)
        .NumberStyle = wdListNumberStyleBullet
        .Alignment = wdListLevelAlignLeft
        .TrailingCharacter = wdTrailingTab          ' wdTrailingNone | wdTrailingSpace | wdTrailingTab
        .NumberPosition = CentimetersToPoints(2 * C_BaseIndent)
        .TextPosition = CentimetersToPoints(C_ListTabPosL2)
        .TabPosition = .TextPosition
        .StartAt = 1
        .LinkedStyle = StyleNames(1)
        .font.Name = C_FT_Body
        .font.Size = C_BaseFontSize
        .font.color = GetThemeColor(ThemeColorIndex:=wdThemeColorAccent1, TintAndShade:=0)  ' in module Theme
    End With
    
    With ListTemplate.ListLevels(3)
        .NumberFormat = ChrW(&HBB)                  ' Unicode for » (solid circle bullet)
        .NumberStyle = wdListNumberStyleBullet
        .Alignment = wdListLevelAlignLeft
        .TrailingCharacter = wdTrailingTab          ' wdTrailingNone | wdTrailingSpace | wdTrailingTab
        .NumberPosition = CentimetersToPoints(3 * C_BaseIndent)
        .TextPosition = CentimetersToPoints(C_ListTabPosL3)
        .TabPosition = .TextPosition
        .StartAt = 1
        .LinkedStyle = StyleNames(2)
        .font.Name = C_FT_Body
        .font.Size = C_BaseFontSize
        .font.color = GetThemeColor(ThemeColorIndex:=wdThemeColorAccent1, TintAndShade:=0)  ' in module Theme
    End With
        
    With ListTemplate.ListLevels(4)
        .NumberFormat = ChrW(&HBB)                  ' Unicode for » (solid circle bullet)
        .NumberStyle = wdListNumberStyleBullet
        .Alignment = wdListLevelAlignLeft
        .TrailingCharacter = wdTrailingTab          ' wdTrailingNone | wdTrailingSpace | wdTrailingTab
        .NumberPosition = CentimetersToPoints(4 * C_BaseIndent)
        .TextPosition = CentimetersToPoints(C_ListTabPosL4)
        .TabPosition = .TextPosition
        .StartAt = 1
        .LinkedStyle = StyleNames(3)
        .font.Name = C_FT_Body
        .font.Size = C_BaseFontSize
        .font.color = GetThemeColor(ThemeColorIndex:=wdThemeColorAccent1, TintAndShade:=0)  ' in module Theme
    End With
                
                
    Create_LT_Bullets = True
End Function

' Create dedicated List Template just for ToC 1 to ToC 3.
' I define only indentations, no numbering. Numbering is inherited just from the field definition, I suppose:
'{ TOC \h \z \t \u "ParListHeading ms;1;ParHeading 1ms;1;ParHeading 2 ms;2;ParHeading 3 ms;3" }
' Remaining space after number is a mystery.
' Observe sub: Private Sub ResetTOCStylesNumbering()
' 2026-01-02 by ms
' 2026-01-18 by ms
Public Function Create_LT_ToC() As Boolean
    Dim ListTemplate As ListTemplate
    Dim NumberFormat As String
    Dim StyleNames As Variant
    
    ' Check if template already exists in gallery
    Dim lt As ListTemplate
    Dim FlagFound As Boolean
    FlagFound = False
    
    For Each lt In ActiveDocument.ListTemplates
        If lt.Name = C_LT_TOC Then
            Set ListTemplate = lt
            FlagFound = True
            Exit For
        End If
    Next lt
    
    ' If not found, create new template in Outline Numbered gallery
    If Not FlagFound Then
        Set ListTemplate = ActiveDocument.ListTemplates.Add( _
            Name:=C_LT_TOC, _
            OutlineNumbered:=True)
    End If
    
    ' Configure each level
    With ListTemplate.ListLevels(1)
'        .NumberPosition = CentimetersToPoints(0)
'        .TabPosition = CentimetersToPoints(0)
'        .TextPosition = CentimetersToPoints(0)
        .LinkedStyle = ActiveDocument.Styles(wdStyleTOC1).NameLocal
    End With
        
    With ListTemplate.ListLevels(2)
'        .NumberPosition = CentimetersToPoints(1 * C_BaseIndent)
'        .TabPosition = .TextPosition
'        .TextPosition = CentimetersToPoints(1#)
        .LinkedStyle = ActiveDocument.Styles(wdStyleTOC2).NameLocal
    End With
    
    With ListTemplate.ListLevels(3)
'        .NumberPosition = CentimetersToPoints(2 * C_BaseIndent)
'        .TextPosition = CentimetersToPoints(2#)
        .LinkedStyle = ActiveDocument.Styles(wdStyleTOC3).NameLocal
    End With
                
    Create_LT_ToC = True
End Function

' 2025-11-22 by ms
' 2026-01-18 by ms
Public Function Create_LT_NumOrd() As Boolean
    Dim ListTemplate As ListTemplate
    Dim NumberFormat As String
    Dim StyleNames As Variant
    
    ' Check if template already exists in gallery
    Dim lt As ListTemplate
    Dim FlagFound As Boolean
    FlagFound = False
    
    For Each lt In ActiveDocument.ListTemplates
        If lt.Name = C_LT_NumOrd Then
            Set ListTemplate = lt
            FlagFound = True
            Exit For
        End If
    Next lt
    
    ' If not found, create new template in Outline Numbered gallery
    If Not FlagFound Then
        Set ListTemplate = ActiveDocument.ListTemplates.Add( _
            Name:=C_LT_NumOrd, _
            OutlineNumbered:=True)
    End If
    
    ' Configure each level
    With ListTemplate.ListLevels(1)
        .NumberFormat = "%1."
        .NumberStyle = wdListNumberStyleArabic  '1., 2., 3., …
        .Alignment = wdListLevelAlignLeft
        .TrailingCharacter = wdTrailingTab             ' wdTrailingNone | wdTrailingSpace | wdTrailingTab
        .NumberPosition = CentimetersToPoints(C_BaseIndent)
        .TextPosition = CentimetersToPoints(C_ListTabPosL1)
        .TabPosition = .TextPosition
        .StartAt = 1
        .LinkedStyle = C_S_ListLevel1
        .font.Name = C_FT_Body
        .font.Size = C_BaseFontSize
        .font.color = GetThemeColor(ThemeColorIndex:=wdThemeColorAccent1, TintAndShade:=0)  ' in module Theme
    End With
        
    With ListTemplate.ListLevels(2)
        .NumberFormat = "%2."
        .NumberStyle = wdListNumberStyleLowercaseLetter ' a., b., c., …
        .Alignment = wdListLevelAlignLeft
        .TrailingCharacter = wdTrailingTab          ' wdTrailingNone | wdTrailingSpace | wdTrailingTab
        .NumberPosition = CentimetersToPoints(2 * C_BaseIndent)
        .TextPosition = CentimetersToPoints(C_ListTabPosL2)
        .TabPosition = .TextPosition
        .StartAt = 1
        .LinkedStyle = C_S_ListLevel2
        .font.Name = C_FT_Body
        .font.Size = C_BaseFontSize
        .font.color = GetThemeColor(ThemeColorIndex:=wdThemeColorAccent1, TintAndShade:=0)  ' in module Theme
    End With
    
    With ListTemplate.ListLevels(3)
        .NumberFormat = "%3."
        .NumberStyle = wdListNumberStyleLowercaseRoman ' i., ii., iii., …
        .Alignment = wdListLevelAlignLeft
        .TrailingCharacter = wdTrailingTab          ' wdTrailingNone | wdTrailingSpace | wdTrailingTab
        .NumberPosition = CentimetersToPoints(3 * C_BaseIndent)
        .TextPosition = CentimetersToPoints(C_ListTabPosL3)
        .TabPosition = .TextPosition
        .StartAt = 1
        .LinkedStyle = C_S_ListLevel3
        .font.Name = C_FT_Body
        .font.Size = C_BaseFontSize
        .font.color = GetThemeColor(ThemeColorIndex:=wdThemeColorAccent1, TintAndShade:=0)  ' in module Theme
    End With
        
    With ListTemplate.ListLevels(4)
        .NumberFormat = "00%4."
        .NumberStyle = wdListNumberStyleArabic ' 001., 002., 003., …
        .Alignment = wdListLevelAlignLeft
        .TrailingCharacter = wdTrailingTab          ' wdTrailingNone | wdTrailingSpace | wdTrailingTab
        .NumberPosition = CentimetersToPoints(4 * C_BaseIndent)
        .TextPosition = CentimetersToPoints(C_ListTabPosL4)
        .TabPosition = .TextPosition
        .StartAt = 1
        .LinkedStyle = C_S_ListLevel4
        .font.Name = C_FT_Body
        .font.Size = C_BaseFontSize
        .font.color = GetThemeColor(ThemeColorIndex:=wdThemeColorAccent1, TintAndShade:=0)  ' in module Theme
    End With
        
    Create_LT_NumOrd = True
End Function

' 2025-11-22 by ms
' 2026-01-18 by ms
Public Function Create_LT_Headings() As Boolean
    Dim ListTemplate As ListTemplate
    Dim i As Integer, j As Integer
    Dim NumberFormat As String
    Dim StyleNames As Variant
    Dim SafetyMargin As Double
    
    ' Define heading styles for each level
    StyleNames = Array( _
        C_S_Heading1, _
        C_S_Heading2, _
        C_S_Heading3, _
        C_S_Heading4, _
        C_S_Heading5, _
        C_S_Heading6, _
        C_S_Heading7, _
        C_S_Heading8)
    
    ' Check if template already exists in gallery
    Dim lt As ListTemplate
    Dim FlagFound As Boolean
    FlagFound = False
    
    For Each lt In ActiveDocument.ListTemplates
        If lt.Name = C_LT_Headings Then
            Set ListTemplate = lt
            FlagFound = True
            Exit For
        End If
    Next lt
    
    ' If not found, create new template in Outline Numbered gallery
    If Not FlagFound Then
        Set ListTemplate = ActiveDocument.ListTemplates.Add( _
            Name:=C_LT_Headings, _
            OutlineNumbered:=True)
    End If
    
    ' Configure each level
    For i = 1 To 8
        NumberFormat = ""
        For j = 1 To i
            NumberFormat = NumberFormat & "%" & j & "."
        Next j
        
        ' Dynamic approach, different separator for different levels, defined by trial and error approach
        Select Case i
            Case 1
                SafetyMargin = 1#   ' (eg. 1)
            Case 2
                SafetyMargin = 1.5  ' (eg. 1.2.)
            Case 3
                SafetyMargin = 2#  ' (eg. 1.2.3.)
            Case 4
                SafetyMargin = 2.5  ' (eg. 1.2.3.4.)
            Case 5
                SafetyMargin = 3#  ' (eg. 1.2.3.4.5.)
            Case 6
                SafetyMargin = 3.5  ' (eg. 1.2.3.4.5.6.)
            Case 7
                SafetyMargin = 4#   ' (eg. 1.2.3.4.5.6.7.)
            Case 8
                SafetyMargin = 4.5   ' (eg. 1.2.3.4.5.6.7.8.)
        End Select
        
        With ListTemplate.ListLevels(i)
            .NumberFormat = NumberFormat
            .NumberStyle = wdListNumberStyleArabic
            .Alignment = wdListLevelAlignLeft
            .TrailingCharacter = wdTrailingTab ' wdTrailingNone | wdTrailingSpace | wdTrailingTab
            .NumberPosition = CentimetersToPoints((i - 1) * C_BaseIndent)
            .TabPosition = CentimetersToPoints(SafetyMargin + (i - 1) * C_BaseIndent)
            .TextPosition = .TabPosition
            .ResetOnHigher = i - 1
            .StartAt = 1
            .LinkedStyle = StyleNames(i - 1)
            ' Apply font and color for number formatting
            .font.Name = C_FT_Headings
            .font.color = GetThemeColor(ThemeColorIndex:=wdThemeColorAccent1, TintAndShade:=0)   ' in module Template
        End With
    Next i
    
    Create_LT_Headings = True
End Function

' 2025-11-25 by ms
Private Function CreateStyle_ParListInTableMs() As Boolean
    Dim NewStyle As style
    
    Dim FileName As String:      FileName = C_F_Macros
    Dim ModuleName As String:    ModuleName = C_M_Styles
    Dim MacroName As String:     MacroName = "CreateStyle_ParListInTableMs"
    Dim MsgBoxTitle As String:   MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    ' Check if the style already exists in the document
    ' If an error occurs, skip the line that caused the error and continue with the next line of code.
    On Error Resume Next
    Set NewStyle = ActiveDocument.Styles(C_S_ListNumTable)
    '  Turns off any active error handler and restores default behavior
    On Error GoTo 0

    ' If the style doesn't exist, try to create it
    If NewStyle Is Nothing Then
        On Error GoTo StyleError
        Set NewStyle = ActiveDocument.Styles.Add(Name:=C_S_ListNumTable, Type:=wdStyleTypeParagraph)
        On Error GoTo 0
    End If

    If Not NewStyle Is Nothing Then
        With NewStyle
            .BaseStyle = C_S_ParNormal
            .NextParagraphStyle = C_S_ParInTable
            .AutomaticallyUpdate = False
            .QuickStyle = False
            .LanguageId = wdEnglishUS
            
            ' Font formatting
            With .font
                .Name = C_FT_Body
                .Size = C_BaseFontSize
                .Bold = False
                .Italic = False
                .color = wdColorAutomatic
            End With
            
            ' Paragraph formatting
            With .ParagraphFormat
                .Alignment = wdAlignParagraphLeft
                .LeftIndent = CentimetersToPoints(0)
                .RightIndent = CentimetersToPoints(0)
                .FirstLineIndent = 0
                .SpaceBefore = 3
                .SpaceAfter = 3
                .LineSpacing = NewStyle.font.Size   ' order matters: specify at first LineSpacing, next LineSpacingRule
                .LineSpacingRule = wdLineSpaceExactly
                .WidowControl = True
                .KeepWithNext = False
                .KeepTogether = False
                .PageBreakBefore = False
                .OutlineLevel = wdOutlineLevelBodyText
            End With
        End With

'       === Add keyboard shortcut for the style ===
'       On Error GoTo ShortcutError
'       CustomizationContext = ActiveDocument
'       KeyBindings.Add _
'            KeyCategory:=wdKeyCategoryStyle, _
'            Command:=C_S_ListNumTable, _
'            KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyN), _
'            KeyCode2:=wdKeyR
'       On Error GoTo 0

        CreateStyle_ParListInTableMs = True
    End If
    Exit Function
                
' === Error handler ===
StyleError:
    MsgBox _
        Prompt:="Error: Unable to create style '" & C_S_ListNumTable & "'." & vbNewLine & _
            "Error number: " & Err.Number & vbNewLine & _
            "Description: " & Err.Description, _
        Buttons:=vbCritical, _
        Title:=MsgBoxTitle
    CreateStyle_ParListInTableMs = False
    Exit Function
        
ShortcutError:
    MsgBox _
        Prompt:="Error adding shortcut for style '" & C_S_ListNumTable & "'." & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, _
        Buttons:=vbCritical, _
        Title:=MsgBoxTitle
    Resume Next

End Function

' 2025-11-25 by ms
Private Function CreateStyle_ParNumRefMs() As Boolean
    Dim NewStyle As style
    
    Dim FileName As String:      FileName = C_F_Macros
    Dim ModuleName As String:    ModuleName = C_M_Styles
    Dim MacroName As String:     MacroName = "CreateStyle_ParNumRefMs"
    Dim MsgBoxTitle As String:   MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    ' Check if the style already exists in the document
    ' If an error occurs, skip the line that caused the error and continue with the next line of code.
    On Error Resume Next
    Set NewStyle = ActiveDocument.Styles(C_S_ListNumRef)
    '  Turns off any active error handler and restores default behavior
    On Error GoTo 0

    ' If the style doesn't exist, try to create it
    If NewStyle Is Nothing Then
        On Error GoTo StyleError
        Set NewStyle = ActiveDocument.Styles.Add(Name:=C_S_ListNumRef, Type:=wdStyleTypeParagraph)
        On Error GoTo 0
    End If

    If Not NewStyle Is Nothing Then
        With NewStyle
            .BaseStyle = C_S_ParNormal
            .NextParagraphStyle = C_S_ParNormal
            .AutomaticallyUpdate = False
            .QuickStyle = False
            .LanguageId = wdEnglishUS
            
            ' Font formatting
            With .font
                .Name = C_FT_Body
                .Size = C_BaseFontSize
                .Bold = False
                .Italic = False
                .color = wdColorAutomatic
            End With
            
            ' Paragraph formatting
            With .ParagraphFormat
                .Alignment = wdAlignParagraphLeft
                .LeftIndent = CentimetersToPoints(0)
                .RightIndent = CentimetersToPoints(0)
                .FirstLineIndent = CentimetersToPoints(-0.9)
                .SpaceBefore = 0
                .SpaceAfter = 6
                .LineSpacing = NewStyle.font.Size   ' order matters: specify at first LineSpacing, next LineSpacingRule
                .LineSpacingRule = wdLineSpaceExactly
                .WidowControl = False
                .KeepWithNext = False
                .KeepTogether = False
                .PageBreakBefore = False
                .OutlineLevel = wdOutlineLevelBodyText
            End With
        End With

'       === Add keyboard shortcut for the style ===
       On Error GoTo ShortcutError
       CustomizationContext = ActiveDocument
       KeyBindings.Add _
            KeyCategory:=wdKeyCategoryStyle, _
            Command:=C_S_ListNumRef, _
            KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyN), _
            KeyCode2:=wdKeyR
       On Error GoTo 0

        CreateStyle_ParNumRefMs = True
    End If
    Exit Function
                
' === Error handler ===
StyleError:
    MsgBox _
        Prompt:="Error: Unable to create style '" & C_S_ListNumRef & "'." & vbNewLine & _
            "Error number: " & Err.Number & vbNewLine & _
            "Description: " & Err.Description, _
        Buttons:=vbCritical, _
        Title:=MsgBoxTitle
    CreateStyle_ParNumRefMs = False
    Exit Function
        
ShortcutError:
    MsgBox _
        Prompt:="Error adding shortcut for style '" & C_S_ListNumRef & "'." & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, _
        Buttons:=vbCritical, _
        Title:=MsgBoxTitle
    Resume Next

End Function

' 2025-11-22 by ms
Private Function CreateStyle_ParTextBoxesMs() As Boolean
    Dim NewStyle As style
    
    Dim FileName As String:      FileName = C_F_Macros
    Dim ModuleName As String:    ModuleName = C_M_Styles
    Dim MacroName As String:     MacroName = "CreateStyle_ParTextBoxesMs"
    Dim MsgBoxTitle As String:   MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    ' Check if the style already exists in the document
    ' If an error occurs, skip the line that caused the error and continue with the next line of code.
    On Error Resume Next
    Set NewStyle = ActiveDocument.Styles(C_S_TextBoxes)
    '  Turns off any active error handler and restores default behavior
    On Error GoTo 0

    ' If the style doesn't exist, try to create it
    If NewStyle Is Nothing Then
        On Error GoTo StyleError
        Set NewStyle = ActiveDocument.Styles.Add(Name:=C_S_TextBoxes, Type:=wdStyleTypeParagraph)
        On Error GoTo 0
    End If

    If Not NewStyle Is Nothing Then
        With NewStyle
            .BaseStyle = C_S_ParNormal
            .NextParagraphStyle = C_S_TextBoxes
            .AutomaticallyUpdate = False
            .QuickStyle = False
            .LanguageId = wdEnglishUS
            
            ' Font formatting
            With .font
                .Name = C_FT_Body
                .Size = C_BaseFontSize - 1
                .Bold = False
                .Italic = False
                .color = wdColorAutomatic
            End With
            
            ' Paragraph formatting
            With .ParagraphFormat
                .Alignment = wdAlignParagraphCenter
                .LeftIndent = CentimetersToPoints(0)
                .RightIndent = CentimetersToPoints(0)
                .FirstLineIndent = 0
                .SpaceBefore = 0
                .SpaceAfter = 0
                .LineSpacing = NewStyle.font.Size ' order matters: specify at first LineSpacing, next LineSpacingRule
                .LineSpacingRule = wdLineSpaceExactly
                .WidowControl = False
                .KeepWithNext = False
                .KeepTogether = False
                .PageBreakBefore = False
                .OutlineLevel = wdOutlineLevelBodyText
            End With
        End With

'       === Add keyboard shortcut for the style ===
'       On Error GoTo ShortcutError
'       CustomizationContext = ActiveDocument
'       KeyBindings.Add _
'            KeyCategory:=wdKeyCategoryStyle, _
'            Command:=C_S_TextBoxes, _
'            KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyN), _
'            KeyCode2:=wdKeyM
'       On Error GoTo 0

        CreateStyle_ParTextBoxesMs = True
    End If
    Exit Function
                
' === Error handler ===
StyleError:
    MsgBox _
        Prompt:="Error: Unable to create style '" & C_S_TextBoxes & "'." & vbNewLine & _
            "Error number: " & Err.Number & vbNewLine & _
            "Description: " & Err.Description, _
        Buttons:=vbCritical, _
        Title:=MsgBoxTitle
    CreateStyle_ParTextBoxesMs = False
    Exit Function
        
ShortcutError:
    MsgBox _
        Prompt:="Error adding shortcut for style '" & C_S_TextBoxes & "'." & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, _
        Buttons:=vbCritical, _
        Title:=MsgBoxTitle
    Resume Next

End Function

' Exception: built-in styles, which I modify. This is the only way which I know to keep Table of Content functionality of Microsoft Word.
' 2025-11-22 by ms
Private Function CreateStyle_TOC3() As Boolean
    Dim NewStyle As style
    
    Dim FileName As String:       FileName = C_F_Macros
    Dim ModuleName As String:     ModuleName = C_M_Styles
    Dim MacroName As String:      MacroName = "CreateStyle_TOC3"
    Dim MsgBoxTitle As String:    MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    ' Check if the style already exists in the document
    ' If an error occurs, skip the line that caused the error and continue with the next line of code.
    On Error Resume Next
    Set NewStyle = ActiveDocument.Styles(wdStyleTOC3)   ' exception, apply built-in style
    '  Turns off any active error handler and restores default behavior
    On Error GoTo 0
    
    ' If built-in style doesn't exist, warn user and exit.
    If NewStyle Is Nothing Then
        ' It should not happen, as this is built-in style
        MsgBox _
            Prompt:="Built-in style cannot be find: " & wdStyleTOC3 & vbNewLine & vbNewLine & _
                "Exiting.", _
            Buttons:=vbCritical, _
            Title:=MsgBoxTitle
        CreateStyle_TOC3 = False
        Exit Function
    End If

    With NewStyle
        .Visibility = False
        .UnhideWhenUsed = False
        .BaseStyle = C_S_ParNormal
        .NextParagraphStyle = C_S_ParNormal
        .AutomaticallyUpdate = False
        .QuickStyle = False
        .LanguageId = wdEnglishUS
        
        ' Font formatting
        With .font
            .Name = C_FT_Body
            .Size = C_BaseFontSize
            .Bold = False
            .Italic = False
            .color = wdColorAutomatic
        End With
        
        ' Paragraph formatting
        With .ParagraphFormat
'            .Alignment = wdAlignParagraphLeft
            .LeftIndent = CentimetersToPoints(2 * C_BaseIndent)
            .RightIndent = CentimetersToPoints(0)
'            .FirstLineIndent = CentimetersToPoints(-2 * C_BaseIndent)
            .SpaceBefore = 0
            .SpaceAfter = 0
            .LineSpacing = NewStyle.font.Size   ' order matters: specify at first LineSpacing, next LineSpacingRule
            .LineSpacingRule = wdLineSpaceExactly
            .WidowControl = True
            .KeepWithNext = False
            .KeepTogether = False
            .PageBreakBefore = False
            .OutlineLevel = wdOutlineLevelBodyText
        End With
    End With

'       === Add keyboard shortcut for the style ===
'       On Error GoTo ShortcutError
'       CustomizationContext = ActiveDocument
'       KeyBindings.Add _
'            KeyCategory:=wdKeyCategoryStyle, _
'            Command:=C_S_TOC1, _
'            KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyN), _
'            KeyCode2:=wdKeyM
'       On Error GoTo 0

    CreateStyle_TOC3 = True
    Exit Function
                
' === Error handler ===
        
ShortcutError:
    MsgBox _
        Prompt:="Error adding shortcut for style '" & wdStyleTOC3 & "'." & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, _
        Buttons:=vbCritical, _
        Title:=MsgBoxTitle
    Resume Next

End Function

' Exception: built-in styles, which I modify. This is the only way which I know to keep Table of Content functionality of Microsoft Word.
' 2025-11-22 by ms
Private Function CreateStyle_TOC2() As Boolean
    Dim NewStyle As style
    
    Dim FileName As String:     FileName = C_F_Macros
    Dim ModuleName As String:   ModuleName = C_M_Styles
    Dim MacroName As String:    MacroName = "CreateStyle_TOC2"
    Dim MsgBoxTitle As String:  MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    ' Check if the style already exists in the document
    ' If an error occurs, skip the line that caused the error and continue with the next line of code.
    On Error Resume Next
    Set NewStyle = ActiveDocument.Styles(wdStyleTOC2)   ' exception, apply built-in style
    '  Turns off any active error handler and restores default behavior
    On Error GoTo 0
    
    ' If built-in style doesn't exist, warn user and exit.
    If NewStyle Is Nothing Then
        ' It should not happen, as this is built-in style
        MsgBox _
            Prompt:="Built-in style cannot be find: " & wdStyleTOC2 & vbNewLine & vbNewLine & _
                "Exiting.", _
            Buttons:=vbCritical, _
            Title:=MsgBoxTitle
        CreateStyle_TOC2 = False
        Exit Function
    End If

    With NewStyle
        .Visibility = False
        .UnhideWhenUsed = False
        .BaseStyle = C_S_ParNormal
        .NextParagraphStyle = C_S_ParNormal
        .AutomaticallyUpdate = False
        .QuickStyle = False
        .LanguageId = wdEnglishUS
        
        ' Font formatting
        With .font
            .Name = C_FT_Body
            .Size = C_BaseFontSize
            .Bold = False
            .Italic = False
            .color = wdColorAutomatic
        End With
        
        ' Paragraph formatting
        With .ParagraphFormat
'            .Alignment = wdAlignParagraphLeft
            .LeftIndent = CentimetersToPoints(1 * C_BaseIndent)
            .RightIndent = CentimetersToPoints(0)
'            .FirstLineIndent = CentimetersToPoints(-1 * C_BaseIndent)
            .SpaceBefore = 0
            .SpaceAfter = 0
            .LineSpacing = NewStyle.font.Size   ' order matters: specify at first LineSpacing, next LineSpacingRule
            .LineSpacingRule = wdLineSpaceExactly
            .WidowControl = True
            .KeepWithNext = False
            .KeepTogether = False
            .PageBreakBefore = False
            .OutlineLevel = wdOutlineLevelBodyText
        End With
    End With

'       === Add keyboard shortcut for the style ===
'       On Error GoTo ShortcutError
'       CustomizationContext = ActiveDocument
'       KeyBindings.Add _
'            KeyCategory:=wdKeyCategoryStyle, _
'            Command:=C_S_TOC1, _
'            KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyN), _
'            KeyCode2:=wdKeyM
'       On Error GoTo 0

    CreateStyle_TOC2 = True

    Exit Function
                
' === Error handler ===
        
ShortcutError:
    MsgBox _
        Prompt:="Error adding shortcut for style '" & wdStyleTOC2 & "'." & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, _
        Buttons:=vbCritical, _
        Title:=MsgBoxTitle
    Resume Next

End Function

' Exception: built-in styles, which I modify. This is the only way which I know to keep Table of Content functionality of Microsoft Word.
' 2025-11-22 by ms
' 2026-01-18 by ms
Private Function CreateStyle_TOC1() As Boolean
    Dim NewStyle As style
    
    Dim FileName As String:     FileName = C_F_Macros
    Dim ModuleName As String:   ModuleName = C_M_Styles
    Dim MacroName As String:    MacroName = "CreateStyle_TOC1"
    Dim MsgBoxTitle As String:  MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    ' Check if the style already exists in the document
    ' If an error occurs, skip the line that caused the error and continue with the next line of code.
    On Error Resume Next
    Set NewStyle = ActiveDocument.Styles(wdStyleTOC1)   ' exception, apply built-in style
    '  Turns off any active error handler and restores default behavior
    On Error GoTo 0

    ' If built-in style doesn't exist, warn user and exit.
    If NewStyle Is Nothing Then
        ' It should not happen, as this is built-in style
        MsgBox _
            Prompt:="Built-in style cannot be find: " & wdStyleTOC1 & vbNewLine & vbNewLine & _
                "Exiting.", _
            Buttons:=vbCritical, _
            Title:=MsgBoxTitle
        CreateStyle_TOC1 = False
        Exit Function
    End If

    With NewStyle
        .Visibility = False
        .UnhideWhenUsed = False
        .BaseStyle = C_S_ParNormal
        .NextParagraphStyle = C_S_ParNormal
        .AutomaticallyUpdate = False
        .QuickStyle = False
        .LanguageId = wdEnglishUS
        
        ' Font formatting
        With .font
            .Name = C_FT_Body
            .Size = C_BaseFontSize + 3
            .Bold = True
            .Italic = False
            .color = wdColorAutomatic
        End With
        
        ' Paragraph formatting
        With .ParagraphFormat
'            .Alignment = wdAlignParagraphLeft
            .LeftIndent = CentimetersToPoints(0)
            .RightIndent = CentimetersToPoints(0)
'            .FirstLineIndent = CentimetersToPoints(-C_BaseIndent)
            .SpaceBefore = 3
            .SpaceAfter = 3
            .LineSpacing = NewStyle.font.Size ' order matters: specify at first LineSpacing, next LineSpacingRule
            .LineSpacingRule = wdLineSpaceExactly
            .WidowControl = True
            .KeepWithNext = False
            .KeepTogether = False
            .PageBreakBefore = False
            .OutlineLevel = wdOutlineLevelBodyText
        End With
    End With

'       === Add keyboard shortcut for the style ===
'       On Error GoTo ShortcutError
'       CustomizationContext = ActiveDocument
'       KeyBindings.Add _
'            KeyCategory:=wdKeyCategoryStyle, _
'            Command:=C_S_TOC1, _
'            KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyN), _
'            KeyCode2:=wdKeyM
'       On Error GoTo 0

    CreateStyle_TOC1 = True
    Exit Function
                
' === Error handler ===
        
ShortcutError:
    MsgBox _
        Prompt:="Error adding shortcut for style '" & wdStyleTOC1 & "'." & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, _
        Buttons:=vbCritical, _
        Title:=MsgBoxTitle
    Resume Next

End Function

' 2025-11-18 by ms
Private Function CreateStyle_ParMinimalMs() As Boolean
    Dim NewStyle As style
    
    Dim FileName As String:      FileName = C_F_Macros
    Dim ModuleName As String:    ModuleName = C_M_Styles
    Dim MacroName As String:     MacroName = "CreateStyle_ParMinimalMs"
    Dim MsgBoxTitle As String:   MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    ' Check if the style already exists in the document
    ' If an error occurs, skip the line that caused the error and continue with the next line of code.
    On Error Resume Next
    Set NewStyle = ActiveDocument.Styles(C_S_ParMinimal)
    '  Turns off any active error handler and restores default behavior
    On Error GoTo 0

    ' If the style doesn't exist, try to create it
    If NewStyle Is Nothing Then
        On Error GoTo StyleError
        Set NewStyle = ActiveDocument.Styles.Add(Name:=C_S_ParMinimal, Type:=wdStyleTypeParagraph)
        On Error GoTo 0
    End If

    If Not NewStyle Is Nothing Then
        With NewStyle
            .BaseStyle = C_S_ParNormal
            .NextParagraphStyle = C_S_ParNormal
            .AutomaticallyUpdate = False
            .QuickStyle = False
            .LanguageId = wdEnglishUS
            
            ' Font formatting
            With .font
                .Name = C_FT_Body
                .Size = 1
                .Bold = False
                .Italic = False
                .color = wdColorAutomatic
            End With
            
            ' Paragraph formatting
            With .ParagraphFormat
                .Alignment = wdAlignParagraphLeft
                .LeftIndent = CentimetersToPoints(0)
                .RightIndent = CentimetersToPoints(0)
                .FirstLineIndent = 0
                .SpaceBefore = 0
                .SpaceAfter = 0
                .LineSpacing = NewStyle.font.Size   ' order matters: specify at first LineSpacing, next LineSpacingRule
                .LineSpacingRule = wdLineSpaceExactly
                .WidowControl = False
                .KeepWithNext = False
                .KeepTogether = False
                .PageBreakBefore = False
                .OutlineLevel = wdOutlineLevelBodyText
            End With
        End With

'       === Add keyboard shortcut for the style ===
       On Error GoTo ShortcutError
       CustomizationContext = ActiveDocument
       KeyBindings.Add _
            KeyCategory:=wdKeyCategoryStyle, _
            Command:=C_S_ParMinimal, _
            KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyN), _
            KeyCode2:=wdKeyM
       On Error GoTo 0

    CreateStyle_ParMinimalMs = True
    End If
    Exit Function
                
' === Error handler ===
StyleError:
    MsgBox _
        Prompt:="Error: Unable to create style '" & C_S_ParMinimal & "'." & vbNewLine & _
            "Error number: " & Err.Number & vbNewLine & _
            "Description: " & Err.Description, _
        Buttons:=vbCritical, _
        Title:=MsgBoxTitle
    CreateStyle_ParMinimalMs = False
    Exit Function
        
ShortcutError:
    MsgBox _
        Prompt:="Error adding shortcut for style '" & C_S_ParMinimal & "'." & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, _
        Buttons:=vbCritical, _
        Title:=MsgBoxTitle
    Resume Next

End Function

' 2025-11-25 by ms
' Additional style for bulleted list only
Private Function CreateStyle_ParListIndentB4Ms() As Boolean
    Dim NewStyle As style
    
    Dim FileName As String:     FileName = C_F_Macros
    Dim ModuleName As String:   ModuleName = C_M_Styles
    Dim MacroName As String:    MacroName = "CreateStyle_ParListIndentB4Ms"
    Dim MsgBoxTitle As String:  MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    ' Check if the style already exists in the document
    ' If an error occurs, skip the line that caused the error and continue with the next line of code.
    On Error Resume Next
    Set NewStyle = ActiveDocument.Styles(C_S_ListLevelB4)
    '  Turns off any active error handler and restores default behavior
    On Error GoTo 0

    ' If the style doesn't exist, try to create it
    If NewStyle Is Nothing Then
        On Error GoTo StyleError
        Set NewStyle = ActiveDocument.Styles.Add(Name:=C_S_ListLevelB4, Type:=wdStyleTypeParagraph)
        On Error GoTo 0
    End If

    If Not NewStyle Is Nothing Then
        With NewStyle
            .BaseStyle = C_S_ParNormal
            .NextParagraphStyle = C_S_ListLevelB4
            .AutomaticallyUpdate = False
            .QuickStyle = False
            .LanguageId = wdEnglishUS
            
            ' Font formatting
            With .font
                .Name = C_FT_Body
                .Size = C_BaseFontSize
                .Bold = False
                .Italic = False
                .color = wdColorAutomatic
            End With
            
            ' Paragraph formatting
            With .ParagraphFormat
                .Alignment = wdAlignParagraphLeft
                .LeftIndent = CentimetersToPoints(1.2)
                .RightIndent = CentimetersToPoints(-1.2)
                .FirstLineIndent = 0
                .SpaceBefore = C_DistParBAList
                .SpaceAfter = C_DistParBAList
                .LineSpacing = NewStyle.font.Size   ' order matters: specify at first LineSpacing, next LineSpacingRule
                .LineSpacingRule = wdLineSpaceExactly
                .WidowControl = True
                .KeepWithNext = False
                .KeepTogether = False
                .PageBreakBefore = False
                .OutlineLevel = wdOutlineLevelBodyText
            End With
        End With

'       === Add keyboard shortcut for the style ===
       On Error GoTo ShortcutError
       CustomizationContext = ActiveDocument
       KeyBindings.Add _
            KeyCategory:=wdKeyCategoryStyle, _
            Command:=C_S_ListLevelB4, _
            KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyB), _
            KeyCode2:=wdKey4
       On Error GoTo 0

    CreateStyle_ParListIndentB4Ms = True
    End If
    Exit Function
                
' === Error handler ===
StyleError:
    MsgBox _
        Prompt:="Error: Unable to create style '" & C_S_ListLevelB4 & "'." & vbNewLine & _
            "Error number: " & Err.Number & vbNewLine & _
            "Description: " & Err.Description, _
        Buttons:=vbCritical, _
        Title:=MsgBoxTitle
    CreateStyle_ParListIndentB4Ms = False
    Exit Function
        
ShortcutError:
    MsgBox _
        Prompt:="Error adding shortcut for style '" & C_S_ListLevelB4 & "'." & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, _
        Buttons:=vbCritical, _
        Title:=MsgBoxTitle
    Resume Next

End Function

' 2025-11-18 by ms
Private Function CreateStyle_ParListIndent4Ms() As Boolean
    Dim NewStyle As style
    
    Dim FileName As String:     FileName = C_F_Macros
    Dim ModuleName As String:   ModuleName = C_M_Styles
    Dim MacroName As String:    MacroName = "CreateStyle_ParListIndent4Ms"
    Dim MsgBoxTitle As String:  MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    ' Check if the style already exists in the document
    ' If an error occurs, skip the line that caused the error and continue with the next line of code.
    On Error Resume Next
    Set NewStyle = ActiveDocument.Styles(C_S_ListLevel4)
    '  Turns off any active error handler and restores default behavior
    On Error GoTo 0

    ' If the style doesn't exist, try to create it
    If NewStyle Is Nothing Then
        On Error GoTo StyleError
        Set NewStyle = ActiveDocument.Styles.Add(Name:=C_S_ListLevel4, Type:=wdStyleTypeParagraph)
        On Error GoTo 0
    End If

    If Not NewStyle Is Nothing Then
        With NewStyle
            .BaseStyle = C_S_ParNormal
            .NextParagraphStyle = C_S_ListLevel4
            .AutomaticallyUpdate = False
            .QuickStyle = False
            .LanguageId = wdEnglishUS
            
            ' Font formatting
            With .font
                .Name = C_FT_Body
                .Size = C_BaseFontSize
                .Bold = False
                .Italic = False
                .color = wdColorAutomatic
            End With
            
            ' Paragraph formatting
            With .ParagraphFormat
                .Alignment = wdAlignParagraphLeft
                .LeftIndent = CentimetersToPoints(1.2)
                .RightIndent = CentimetersToPoints(0)
                .FirstLineIndent = CentimetersToPoints(-1.2)
                .SpaceBefore = C_DistParBAList
                .SpaceAfter = C_DistParBAList
                .LineSpacing = NewStyle.font.Size   ' order matters: specify at first LineSpacing, next LineSpacingRule
                .LineSpacingRule = wdLineSpaceExactly
                .WidowControl = True
                .KeepWithNext = False
                .KeepTogether = False
                .PageBreakBefore = False
                .OutlineLevel = wdOutlineLevelBodyText
            End With
        End With

'       === Add keyboard shortcut for the style ===
       On Error GoTo ShortcutError
       CustomizationContext = ActiveDocument
       KeyBindings.Add _
            KeyCategory:=wdKeyCategoryStyle, _
            Command:=C_S_ListLevel4, _
            KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyL), _
            KeyCode2:=wdKey4
       On Error GoTo 0

    CreateStyle_ParListIndent4Ms = True
    End If
    Exit Function
                
' === Error handler ===
StyleError:
    MsgBox _
        Prompt:="Error: Unable to create style '" & C_S_ListLevel4 & "'." & vbNewLine & _
            "Error number: " & Err.Number & vbNewLine & _
            "Description: " & Err.Description, _
        Buttons:=vbCritical, _
        Title:=MsgBoxTitle
    CreateStyle_ParListIndent4Ms = False
    Exit Function
        
ShortcutError:
    MsgBox _
        Prompt:="Error adding shortcut for style '" & C_S_ListLevel4 & "'." & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, _
        Buttons:=vbCritical, _
        Title:=MsgBoxTitle
    Resume Next

End Function

' 2025-11-25 by ms
' Additional style for bulleted list only
Private Function CreateStyle_ParListIndentB3Ms() As Boolean
    Dim NewStyle As style
    
    Dim FileName As String:     FileName = C_F_Macros
    Dim ModuleName As String:   ModuleName = C_M_Styles
    Dim MacroName As String:    MacroName = "CreateStyle_ParListIndentB3Ms"
    Dim MsgBoxTitle As String:  MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    ' Check if the style already exists in the document
    ' If an error occurs, skip the line that caused the error and continue with the next line of code.
    On Error Resume Next
    Set NewStyle = ActiveDocument.Styles(C_S_ListLevelB3)
    '  Turns off any active error handler and restores default behavior
    On Error GoTo 0

    ' If the style doesn't exist, try to create it
    If NewStyle Is Nothing Then
        On Error GoTo StyleError
        Set NewStyle = ActiveDocument.Styles.Add(Name:=C_S_ListLevelB3, Type:=wdStyleTypeParagraph)
        On Error GoTo 0
    End If

    If Not NewStyle Is Nothing Then
        With NewStyle
            .BaseStyle = C_S_ParNormal
            .NextParagraphStyle = C_S_ListLevelB3
            .AutomaticallyUpdate = False
            .QuickStyle = False
            .LanguageId = wdEnglishUS
            
            ' Font formatting
            With .font
                .Name = C_FT_Body
                .Size = C_BaseFontSize
                .Bold = False
                .Italic = False
                .color = wdColorAutomatic
            End With
            
            ' Paragraph formatting
            With .ParagraphFormat
                .Alignment = wdAlignParagraphLeft
                .LeftIndent = CentimetersToPoints(0.9)
                .RightIndent = CentimetersToPoints(0)
                .FirstLineIndent = CentimetersToPoints(-0.9)
                .SpaceBefore = C_DistParBAList
                .SpaceAfter = C_DistParBAList
                .LineSpacing = NewStyle.font.Size   ' order matters: specify at first LineSpacing, next LineSpacingRule
                .LineSpacingRule = wdLineSpaceExactly
                .WidowControl = True
                .KeepWithNext = False
                .KeepTogether = False
                .PageBreakBefore = False
                .OutlineLevel = wdOutlineLevelBodyText
            End With
        End With

'       === Add keyboard shortcut for the style ===
       On Error GoTo ShortcutError
       CustomizationContext = ActiveDocument
       KeyBindings.Add _
            KeyCategory:=wdKeyCategoryStyle, _
            Command:=C_S_ListLevelB3, _
            KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyB), _
            KeyCode2:=wdKey3
       On Error GoTo 0

    CreateStyle_ParListIndentB3Ms = True
    End If
    Exit Function
                
' === Error handler ===
StyleError:
    MsgBox _
        Prompt:="Error: Unable to create style '" & C_S_ListLevelB2 & "'." & vbNewLine & _
            "Error number: " & Err.Number & vbNewLine & _
            "Description: " & Err.Description, _
        Buttons:=vbCritical, _
        Title:=MsgBoxTitle
    CreateStyle_ParListIndentB3Ms = False
    Exit Function
        
ShortcutError:
    MsgBox _
        Prompt:="Error adding shortcut for style '" & C_S_ListLevelB2 & "'." & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, _
        Buttons:=vbCritical, _
        Title:=MsgBoxTitle
    Resume Next

End Function

' 2025-11-18 by ms
Private Function CreateStyle_ParListIndent3Ms() As Boolean
    Dim NewStyle As style
    
    Dim FileName As String:     FileName = C_F_Macros
    Dim ModuleName As String:   ModuleName = C_M_Styles
    Dim MacroName As String:    MacroName = "CreateStyle_ParListIndent3Ms"
    Dim MsgBoxTitle As String:  MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    ' Check if the style already exists in the document
    ' If an error occurs, skip the line that caused the error and continue with the next line of code.
    On Error Resume Next
    Set NewStyle = ActiveDocument.Styles(C_S_ListLevel3)
    '  Turns off any active error handler and restores default behavior
    On Error GoTo 0

    ' If the style doesn't exist, try to create it
    If NewStyle Is Nothing Then
        On Error GoTo StyleError
        Set NewStyle = ActiveDocument.Styles.Add(Name:=C_S_ListLevel3, Type:=wdStyleTypeParagraph)
        On Error GoTo 0
    End If

    If Not NewStyle Is Nothing Then
        With NewStyle
            .BaseStyle = C_S_ParNormal
            .NextParagraphStyle = C_S_ListLevel3
            .AutomaticallyUpdate = False
            .QuickStyle = False
            .LanguageId = wdEnglishUS
            
            ' Font formatting
            With .font
                .Name = C_FT_Body
                .Size = C_BaseFontSize
                .Bold = False
                .Italic = False
                .color = wdColorAutomatic
            End With
            
            ' Paragraph formatting
            With .ParagraphFormat
                .Alignment = wdAlignParagraphLeft
                .LeftIndent = CentimetersToPoints(0.9)
                .RightIndent = CentimetersToPoints(0)
                .FirstLineIndent = CentimetersToPoints(-0.9)
                .SpaceBefore = C_DistParBAList
                .SpaceAfter = C_DistParBAList
                .LineSpacing = NewStyle.font.Size   ' order matters: specify at first LineSpacing, next LineSpacingRule
                .LineSpacingRule = wdLineSpaceExactly
                .WidowControl = True
                .KeepWithNext = False
                .KeepTogether = False
                .PageBreakBefore = False
                .OutlineLevel = wdOutlineLevelBodyText
            End With
        End With

'       === Add keyboard shortcut for the style ===
       On Error GoTo ShortcutError
       CustomizationContext = ActiveDocument
       KeyBindings.Add _
            KeyCategory:=wdKeyCategoryStyle, _
            Command:=C_S_ListLevel3, _
            KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyL), _
            KeyCode2:=wdKey3
       On Error GoTo 0

    CreateStyle_ParListIndent3Ms = True
    End If
    Exit Function
                
' === Error handler ===
StyleError:
    MsgBox _
        Prompt:="Error: Unable to create style '" & C_S_ListLevel3 & "'." & vbNewLine & _
            "Error number: " & Err.Number & vbNewLine & _
            "Description: " & Err.Description, _
        Buttons:=vbCritical, _
        Title:=MsgBoxTitle
    CreateStyle_ParListIndent3Ms = False
    Exit Function
        
ShortcutError:
    MsgBox _
        Prompt:="Error adding shortcut for style '" & C_S_ListLevel3 & "'." & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, _
        Buttons:=vbCritical, _
        Title:=MsgBoxTitle
    Resume Next

End Function

' 2025-11-25 by ms
' Additional style for bulleted list only
Private Function CreateStyle_ParListIndentB2Ms() As Boolean
    Dim NewStyle As style
    
    Dim FileName As String:     FileName = C_F_Macros
    Dim ModuleName As String:   ModuleName = C_M_Styles
    Dim MacroName As String:    MacroName = "CreateStyle_ParListIndentB2Ms"
    Dim MsgBoxTitle As String:  MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    ' Check if the style already exists in the document
    ' If an error occurs, skip the line that caused the error and continue with the next line of code.
    On Error Resume Next
    Set NewStyle = ActiveDocument.Styles(C_S_ListLevelB2)
    '  Turns off any active error handler and restores default behavior
    On Error GoTo 0

    ' If the style doesn't exist, try to create it
    If NewStyle Is Nothing Then
        On Error GoTo StyleError
        Set NewStyle = ActiveDocument.Styles.Add(Name:=C_S_ListLevelB2, Type:=wdStyleTypeParagraph)
        On Error GoTo 0
    End If

    If Not NewStyle Is Nothing Then
        With NewStyle
            .BaseStyle = C_S_ParNormal
            .NextParagraphStyle = C_S_ListLevelB2
            .AutomaticallyUpdate = False
            .QuickStyle = False
            .LanguageId = wdEnglishUS
            
            ' Font formatting
            With .font
                .Name = C_FT_Body
                .Size = C_BaseFontSize
                .Bold = False
                .Italic = False
                .color = wdColorAutomatic
            End With
            
            ' Paragraph formatting
            With .ParagraphFormat
                .Alignment = wdAlignParagraphLeft
                .LeftIndent = CentimetersToPoints(0.6)
                .RightIndent = CentimetersToPoints(0)
                .FirstLineIndent = CentimetersToPoints(-0.6)
                .SpaceBefore = C_DistParBAList
                .SpaceAfter = C_DistParBAList
                .LineSpacing = NewStyle.font.Size   ' order matters: specify at first LineSpacing, next LineSpacingRule
                .LineSpacingRule = wdLineSpaceExactly
                .WidowControl = True
                .KeepWithNext = False
                .KeepTogether = False
                .PageBreakBefore = False
                .OutlineLevel = wdOutlineLevelBodyText
            End With
        End With

'       === Add keyboard shortcut for the style ===
       On Error GoTo ShortcutError
       CustomizationContext = ActiveDocument
       KeyBindings.Add _
            KeyCategory:=wdKeyCategoryStyle, _
            Command:=C_S_ListLevelB2, _
            KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyB), _
            KeyCode2:=wdKey2
       On Error GoTo 0

    CreateStyle_ParListIndentB2Ms = True
    End If
    Exit Function
                
' === Error handler ===
StyleError:
    MsgBox _
        Prompt:="Error: Unable to create style '" & C_S_ListLevelB2 & "'." & vbNewLine & _
            "Error number: " & Err.Number & vbNewLine & _
            "Description: " & Err.Description, _
        Buttons:=vbCritical, _
        Title:=MsgBoxTitle
    CreateStyle_ParListIndentB2Ms = False
    Exit Function
        
ShortcutError:
    MsgBox _
        Prompt:="Error adding shortcut for style '" & C_S_ListLevelB2 & "'." & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, _
        Buttons:=vbCritical, _
        Title:=MsgBoxTitle
    Resume Next

End Function

' 2025-11-18 by ms
Private Function CreateStyle_ParListIndent2Ms() As Boolean
    Dim NewStyle As style
    
    Dim FileName As String:     FileName = C_F_Macros
    Dim ModuleName As String:   ModuleName = C_M_Styles
    Dim MacroName As String:    MacroName = "CreateStyle_ParListIndent2Ms"
    Dim MsgBoxTitle As String:  MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    ' Check if the style already exists in the document
    ' If an error occurs, skip the line that caused the error and continue with the next line of code.
    On Error Resume Next
    Set NewStyle = ActiveDocument.Styles(C_S_ListLevel2)
    '  Turns off any active error handler and restores default behavior
    On Error GoTo 0

    ' If the style doesn't exist, try to create it
    If NewStyle Is Nothing Then
        On Error GoTo StyleError
        Set NewStyle = ActiveDocument.Styles.Add(Name:=C_S_ListLevel2, Type:=wdStyleTypeParagraph)
        On Error GoTo 0
    End If

    If Not NewStyle Is Nothing Then
        With NewStyle
            .BaseStyle = C_S_ParNormal
            .NextParagraphStyle = C_S_ListLevel2
            .AutomaticallyUpdate = False
            .QuickStyle = False
            .LanguageId = wdEnglishUS
            
            ' Font formatting
            With .font
                .Name = C_FT_Body
                .Size = C_BaseFontSize
                .Bold = False
                .Italic = False
                .color = wdColorAutomatic
            End With
            
            ' Paragraph formatting
            With .ParagraphFormat
                .Alignment = wdAlignParagraphLeft
                .LeftIndent = CentimetersToPoints(0.6)
                .RightIndent = CentimetersToPoints(0)
                .FirstLineIndent = CentimetersToPoints(-0.6)
                .SpaceBefore = C_DistParBAList
                .SpaceAfter = C_DistParBAList
                .LineSpacing = NewStyle.font.Size   ' order matters: specify at first LineSpacing, next LineSpacingRule
                .LineSpacingRule = wdLineSpaceExactly
                .WidowControl = True
                .KeepWithNext = False
                .KeepTogether = False
                .PageBreakBefore = False
                .OutlineLevel = wdOutlineLevelBodyText
            End With
        End With

'       === Add keyboard shortcut for the style ===
       On Error GoTo ShortcutError
       CustomizationContext = ActiveDocument
       KeyBindings.Add _
            KeyCategory:=wdKeyCategoryStyle, _
            Command:=C_S_ListLevel2, _
            KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyL), _
            KeyCode2:=wdKey2
       On Error GoTo 0

    CreateStyle_ParListIndent2Ms = True
    End If
    Exit Function
                
' === Error handler ===
StyleError:
    MsgBox _
        Prompt:="Error: Unable to create style '" & C_S_ListLevel2 & "'." & vbNewLine & _
            "Error number: " & Err.Number & vbNewLine & _
            "Description: " & Err.Description, _
        Buttons:=vbCritical, _
        Title:=MsgBoxTitle
    CreateStyle_ParListIndent2Ms = False
    Exit Function
        
ShortcutError:
    MsgBox _
        Prompt:="Error adding shortcut for style '" & C_S_ListLevel2 & "'." & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, _
        Buttons:=vbCritical, _
        Title:=MsgBoxTitle
    Resume Next

End Function

' 2025-11-25 by ms
' Additional style for bulleted list only
Private Function CreateStyle_ParListIndentB1Ms() As Boolean
    Dim NewStyle As style
    
    Dim FileName As String:     FileName = C_F_Macros
    Dim ModuleName As String:   ModuleName = C_M_Styles
    Dim MacroName As String:    MacroName = "CreateStyle_ParListIndentB1Ms"
    Dim MsgBoxTitle As String:  MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    ' Check if the style already exists in the document
    ' If an error occurs, skip the line that caused the error and continue with the next line of code.
    On Error Resume Next
    Set NewStyle = ActiveDocument.Styles(C_S_ListLevelB1)
    '  Turns off any active error handler and restores default behavior
    On Error GoTo 0

    ' If the style doesn't exist, try to create it
    If NewStyle Is Nothing Then
        On Error GoTo StyleError
        Set NewStyle = ActiveDocument.Styles.Add(Name:=C_S_ListLevelB1, Type:=wdStyleTypeParagraph)
        On Error GoTo 0
    End If

    If Not NewStyle Is Nothing Then
        With NewStyle
            .BaseStyle = C_S_ParNormal
            .NextParagraphStyle = C_S_ListLevelB1
            .AutomaticallyUpdate = False
            .QuickStyle = False
            .LanguageId = wdEnglishUS
            
            ' Font formatting
            With .font
                .Name = C_FT_Body
                .Size = C_BaseFontSize
                .Bold = False
                .Italic = False
                .color = wdColorAutomatic
            End With
            
            ' Paragraph formatting
            With .ParagraphFormat
                .Alignment = wdAlignParagraphLeft
                .LeftIndent = CentimetersToPoints(0.3)
                .RightIndent = CentimetersToPoints(0)
                .FirstLineIndent = CentimetersToPoints(-0.3)
                .SpaceBefore = C_DistParBAList
                .SpaceAfter = C_DistParBAList
                .LineSpacing = NewStyle.font.Size   ' order matters: specify at first LineSpacing, next LineSpacingRule
                .LineSpacingRule = wdLineSpaceExactly
                .WidowControl = True
                .KeepWithNext = False
                .KeepTogether = False
                .PageBreakBefore = False
                .OutlineLevel = wdOutlineLevelBodyText
            End With
        End With

'       === Add keyboard shortcut for the style ===
       On Error GoTo ShortcutError
       CustomizationContext = ActiveDocument
       KeyBindings.Add _
            KeyCategory:=wdKeyCategoryStyle, _
            Command:=C_S_ListLevelB1, _
            KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyB), _
            KeyCode2:=wdKey1
       On Error GoTo 0

    CreateStyle_ParListIndentB1Ms = True
    End If
    Exit Function
                
' === Error handler ===
StyleError:
    MsgBox _
        Prompt:="Error: Unable to create style '" & C_S_ListLevelB1 & "'." & vbNewLine & _
            "Error number: " & Err.Number & vbNewLine & _
            "Description: " & Err.Description, _
        Buttons:=vbCritical, _
        Title:=MsgBoxTitle
    CreateStyle_ParListIndentB1Ms = False
    Exit Function
        
ShortcutError:
    MsgBox _
        Prompt:="Error adding shortcut for style '" & C_S_ListLevelB1 & "'." & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, _
        Buttons:=vbCritical, _
        Title:=MsgBoxTitle
    Resume Next

End Function

' 2025-11-18 by ms
Private Function CreateStyle_ParListIndent1Ms() As Boolean
    Dim NewStyle As style
    
    Dim FileName As String:     FileName = C_F_Macros
    Dim ModuleName As String:   ModuleName = C_M_Styles
    Dim MacroName As String:    MacroName = "CreateStyle_ParListIndent1Ms"
    Dim MsgBoxTitle As String:  MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    ' Check if the style already exists in the document
    ' If an error occurs, skip the line that caused the error and continue with the next line of code.
    On Error Resume Next
    Set NewStyle = ActiveDocument.Styles(C_S_ListLevel1)
    '  Turns off any active error handler and restores default behavior
    On Error GoTo 0

    ' If the style doesn't exist, try to create it
    If NewStyle Is Nothing Then
        On Error GoTo StyleError
        Set NewStyle = ActiveDocument.Styles.Add(Name:=C_S_ListLevel1, Type:=wdStyleTypeParagraph)
        On Error GoTo 0
    End If

    If Not NewStyle Is Nothing Then
        With NewStyle
            .BaseStyle = C_S_ParNormal
            .NextParagraphStyle = C_S_ListLevel1
            .AutomaticallyUpdate = False
            .QuickStyle = False
            .LanguageId = wdEnglishUS
            
            ' Font formatting
            With .font
                .Name = C_FT_Body
                .Size = C_BaseFontSize
                .Bold = False
                .Italic = False
                .color = wdColorAutomatic
            End With
            
            ' Paragraph formatting
            With .ParagraphFormat
                .Alignment = wdAlignParagraphLeft
                .LeftIndent = CentimetersToPoints(0.3)
                .RightIndent = CentimetersToPoints(0)
                .FirstLineIndent = CentimetersToPoints(-0.3)
                .SpaceBefore = C_DistParBAList
                .SpaceAfter = C_DistParBAList
                .LineSpacing = NewStyle.font.Size   ' order matters: specify at first LineSpacing, next LineSpacingRule
                .LineSpacingRule = wdLineSpaceExactly
                .WidowControl = True
                .KeepWithNext = False
                .KeepTogether = False
                .PageBreakBefore = False
                .OutlineLevel = wdOutlineLevelBodyText
            End With
        End With

'       === Add keyboard shortcut for the style ===
       On Error GoTo ShortcutError
       CustomizationContext = ActiveDocument
       KeyBindings.Add _
            KeyCategory:=wdKeyCategoryStyle, _
            Command:=C_S_ListLevel1, _
            KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyL), _
            KeyCode2:=wdKey1
       On Error GoTo 0

    CreateStyle_ParListIndent1Ms = True
    End If
    Exit Function
                
' === Error handler ===
StyleError:
    MsgBox _
        Prompt:="Error: Unable to create style '" & C_S_ListLevel1 & "'." & vbNewLine & _
            "Error number: " & Err.Number & vbNewLine & _
            "Description: " & Err.Description, _
        Buttons:=vbCritical, _
        Title:=MsgBoxTitle
    CreateStyle_ParListIndent1Ms = False
    Exit Function
        
ShortcutError:
    MsgBox _
        Prompt:="Error adding shortcut for style '" & C_S_ListLevel1 & "'." & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, _
        Buttons:=vbCritical, _
        Title:=MsgBoxTitle
    Resume Next

End Function

' 2025-11-18 by ms
Private Function CreateStyle_ParListHeadingMs() As Boolean
    Dim NewStyle As style
    
    Dim FileName As String:     FileName = C_F_Macros
    Dim ModuleName As String:   ModuleName = C_M_Styles
    Dim MacroName As String:    MacroName = "CreateStyle_ParListHeadingMs"
    Dim MsgBoxTitle As String:  MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    ' Check if the style already exists in the document
    ' If an error occurs, skip the line that caused the error and continue with the next line of code.
    On Error Resume Next
    Set NewStyle = ActiveDocument.Styles(C_S_ListHeading)
    '  Turns off any active error handler and restores default behavior
    On Error GoTo 0

    ' If the style doesn't exist, try to create it
    If NewStyle Is Nothing Then
        On Error GoTo StyleError
        Set NewStyle = ActiveDocument.Styles.Add(Name:=C_S_ListHeading, Type:=wdStyleTypeParagraph)
        On Error GoTo 0
    End If

    If Not NewStyle Is Nothing Then
        With NewStyle
            .BaseStyle = C_S_ParNormal
            .NextParagraphStyle = C_S_ParNormal
            .AutomaticallyUpdate = False
            .QuickStyle = False
            .LanguageId = wdEnglishUS
            
            ' Font formatting
            With .font
                .Name = C_FT_Headings
                .Size = C_BaseFontSize + 3
                .Bold = False
                .Italic = False
                .color = GetThemeColor(ThemeColorIndex:=wdThemeColorAccent1, TintAndShade:=0)   ' in module Template
            End With
            
            ' Paragraph formatting
            With .ParagraphFormat
                .Alignment = wdAlignParagraphLeft
                .LeftIndent = CentimetersToPoints(0)
                .RightIndent = CentimetersToPoints(0)
                .FirstLineIndent = 0
                .SpaceBefore = 24
                .SpaceAfter = 0
                .LineSpacing = NewStyle.font.Size ' order matters: specify at first LineSpacing, next LineSpacingRule
                .LineSpacingRule = wdLineSpaceExactly
                .WidowControl = True
                .KeepWithNext = True
                .KeepTogether = False
                .PageBreakBefore = False
                .OutlineLevel = wdOutlineLevel1
            End With
        End With

'       === Add keyboard shortcut for the style ===
'       On Error GoTo ShortcutError
'       CustomizationContext = ActiveDocument
'       KeyBindings.Add _
'            KeyCategory:=wdKeyCategoryStyle, _
'            Command:=C_S_TableLegend, _
'            KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyL), _
'            KeyCode2:=wdKeyT
'       On Error GoTo 0

    ' Success message
    CreateStyle_ParListHeadingMs = True
    End If
    Exit Function
                
' === Error handler ===
StyleError:
    MsgBox _
        Prompt:="Error: Unable to create style '" & C_S_ListHeading & "'." & vbNewLine & _
            "Error number: " & Err.Number & vbNewLine & _
            "Description: " & Err.Description, _
        Buttons:=vbCritical, _
        Title:=MsgBoxTitle
    CreateStyle_ParListHeadingMs = False
    Exit Function
        
ShortcutError:
    MsgBox _
        Prompt:="Error adding shortcut for style '" & C_S_ListHeading & "'." & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, _
        Buttons:=vbCritical, _
        Title:=MsgBoxTitle
    Resume Next

End Function

' 2025-11-18 by ms
Private Function CreateStyle_ParLegendTableMs() As Boolean
    Dim NewStyle As style
    
    Dim FileName As String:      FileName = C_F_Macros
    Dim ModuleName As String:    ModuleName = C_M_Styles
    Dim MacroName As String:     MacroName = "CreateStyle_ParLegendTableMs"
    Dim MsgBoxTitle As String:   MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    ' Check if the style already exists in the document
    ' If an error occurs, skip the line that caused the error and continue with the next line of code.
    On Error Resume Next
    Set NewStyle = ActiveDocument.Styles(C_S_TableLegend)
    '  Turns off any active error handler and restores default behavior
    On Error GoTo 0

    ' If the style doesn't exist, try to create it
    If NewStyle Is Nothing Then
        On Error GoTo StyleError
        Set NewStyle = ActiveDocument.Styles.Add(Name:=C_S_TableLegend, Type:=wdStyleTypeParagraph)
        On Error GoTo 0
    End If

    If Not NewStyle Is Nothing Then
        With NewStyle
            .BaseStyle = C_S_ParNormal
            .NextParagraphStyle = C_S_ParNormal
            .AutomaticallyUpdate = False
            .QuickStyle = False
            .LanguageId = wdEnglishUS
            
            ' Font formatting
            With .font
                .Name = C_FT_Body
                .Size = C_BaseFontSize
                .Bold = True
                .Italic = False
                .color = wdColorAutomatic
            End With
            
            ' Paragraph formatting
            With .ParagraphFormat
                .Alignment = wdAlignParagraphLeft
                .LeftIndent = CentimetersToPoints(0)
                .RightIndent = CentimetersToPoints(0)
                .FirstLineIndent = 0
                .SpaceBefore = 12
                .SpaceAfter = 3
                .LineSpacing = NewStyle.font.Size   ' order matters: specify at first LineSpacing, next LineSpacingRule
                .LineSpacingRule = wdLineSpaceExactly
                .WidowControl = True
                .KeepWithNext = True
                .KeepTogether = False
                .PageBreakBefore = False
                .OutlineLevel = wdOutlineLevelBodyText
            End With
        End With

'       === Add keyboard shortcut for the style ===
       On Error GoTo ShortcutError
       CustomizationContext = ActiveDocument
       KeyBindings.Add _
            KeyCategory:=wdKeyCategoryStyle, _
            Command:=C_S_TableLegend, _
            KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyL), _
            KeyCode2:=wdKeyT
       On Error GoTo 0

    ' Success message
    CreateStyle_ParLegendTableMs = True
    End If
    Exit Function
                
' === Error handler ===
StyleError:
    MsgBox _
        Prompt:="Error: Unable to create style '" & C_S_TableLegend & "'." & vbNewLine & _
            "Error number: " & Err.Number & vbNewLine & _
            "Description: " & Err.Description, _
        Buttons:=vbCritical, _
        Title:=MsgBoxTitle
    CreateStyle_ParLegendTableMs = False
    Exit Function
        
ShortcutError:
    MsgBox _
        Prompt:="Error adding shortcut for style '" & C_S_TableLegend & "'." & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, _
        Buttons:=vbCritical, _
        Title:=MsgBoxTitle
    Resume Next

End Function

' 2025-11-18 by ms
Private Function CreateStyle_ParLegendPictureMs() As Boolean
    Dim NewStyle As style
    
    Dim FileName As String:     FileName = C_F_Macros
    Dim ModuleName As String:   ModuleName = C_M_Styles
    Dim MacroName As String:    MacroName = "CreateStyle_ParLegendPictureMs"
    Dim MsgBoxTitle As String:  MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    ' Check if the style already exists in the document
    ' If an error occurs, skip the line that caused the error and continue with the next line of code.
    On Error Resume Next
    Set NewStyle = ActiveDocument.Styles(C_S_PictureLegend)
    '  Turns off any active error handler and restores default behavior
    On Error GoTo 0

    ' If the style doesn't exist, try to create it
    If NewStyle Is Nothing Then
        On Error GoTo StyleError
        Set NewStyle = ActiveDocument.Styles.Add(Name:=C_S_PictureLegend, Type:=wdStyleTypeParagraph)
        On Error GoTo 0
    End If

    If Not NewStyle Is Nothing Then
        With NewStyle
            .BaseStyle = C_S_ParNormal
            .NextParagraphStyle = C_S_ParNormal
            .AutomaticallyUpdate = False
            .QuickStyle = False
            .LanguageId = wdEnglishUS
            
            ' Font formatting
            With .font
                .Name = C_FT_Body
                .Size = C_BaseFontSize
                .Bold = True
                .Italic = False
                .color = wdColorAutomatic
            End With
            
            ' Paragraph formatting
            With .ParagraphFormat
                .Alignment = wdAlignParagraphCenter
                .LeftIndent = CentimetersToPoints(0)
                .RightIndent = CentimetersToPoints(0)
                .FirstLineIndent = 0
                .SpaceBefore = 6
                .SpaceAfter = 12
                .LineSpacing = NewStyle.font.Size   ' order matters: specify at first LineSpacing, next LineSpacingRule
                .LineSpacingRule = wdLineSpaceExactly
                .WidowControl = True
                .KeepWithNext = False
                .KeepTogether = False
                .PageBreakBefore = False
                .OutlineLevel = wdOutlineLevelBodyText
            End With
        End With

'       === Add keyboard shortcut for the style ===
       On Error GoTo ShortcutError
       CustomizationContext = ActiveDocument
       KeyBindings.Add _
            KeyCategory:=wdKeyCategoryStyle, _
            Command:=C_S_PictureLegend, _
            KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyL), _
            KeyCode2:=wdKeyP
       On Error GoTo 0

    CreateStyle_ParLegendPictureMs = True
    End If
    Exit Function
                
' === Error handler ===
StyleError:
    MsgBox _
        Prompt:="Error: Unable to create style '" & C_S_PictureLegend & "'." & vbNewLine & _
            "Error number: " & Err.Number & vbNewLine & _
            "Description: " & Err.Description, _
        Buttons:=vbCritical, _
        Title:=MsgBoxTitle
    CreateStyle_ParLegendPictureMs = False
    Exit Function
        
ShortcutError:
    MsgBox _
        Prompt:="Error adding shortcut for style '" & C_S_PictureLegend & "'." & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, _
        Buttons:=vbCritical, _
        Title:=MsgBoxTitle
    Resume Next

End Function

' 2025-11-18 by ms
Private Function CreateStyle_ParLegalNoteMs() As Boolean
    Dim NewStyle As style
    
    Dim FileName As String:      FileName = C_F_Macros
    Dim ModuleName As String:    ModuleName = C_M_Styles
    Dim MacroName As String:     MacroName = "CreateStyle_ParLegalNoteMs"
    Dim MsgBoxTitle As String:   MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    ' Check if the style already exists in the document
    ' If an error occurs, skip the line that caused the error and continue with the next line of code.
    On Error Resume Next
    Set NewStyle = ActiveDocument.Styles(C_S_ParLegalNote)
    '  Turns off any active error handler and restores default behavior
    On Error GoTo 0

    ' If the style doesn't exist, try to create it
    If NewStyle Is Nothing Then
        On Error GoTo StyleError
        Set NewStyle = ActiveDocument.Styles.Add(Name:=C_S_ParLegalNote, Type:=wdStyleTypeParagraph)
        On Error GoTo 0
    End If

    If Not NewStyle Is Nothing Then
        With NewStyle
            .BaseStyle = C_S_ParNormal
            .NextParagraphStyle = C_S_ParLegalNote
            .AutomaticallyUpdate = False
            .QuickStyle = False
            .LanguageId = wdEnglishUS
            
            ' Font formatting
            With .font
                .Name = C_FT_Body
                .Size = C_BaseFontSize
                .Bold = False
                .Italic = False
                .color = GetThemeColor(ThemeColorIndex:=wdThemeColorAccent6, TintAndShade:=0)   ' in module Template
            End With
            
            ' Paragraph formatting
            With .ParagraphFormat
                .Alignment = wdAlignParagraphLeft
                .LeftIndent = CentimetersToPoints(0)
                .RightIndent = CentimetersToPoints(0)
                .FirstLineIndent = 0
                .SpaceBefore = 0
                .SpaceAfter = 6
                .LineSpacing = NewStyle.font.Size   ' order matters: specify at first LineSpacing, next LineSpacingRule
                .LineSpacingRule = wdLineSpaceExactly
                .WidowControl = True
                .KeepWithNext = False
                .KeepTogether = False
                .PageBreakBefore = False
                .OutlineLevel = wdOutlineLevelBodyText
            End With
        End With

'       === Add keyboard shortcut for the style ===
'       On Error GoTo ShortcutError
'       CustomizationContext = ActiveDocument
'       KeyBindings.Add _
'            KeyCategory:=wdKeyCategoryStyle, _
'            Command:=C_S_ParLegalNote, _
'            KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyN), _
'            KeyCode2:=wdKeyT
'       On Error GoTo 0

    CreateStyle_ParLegalNoteMs = True
    End If
    Exit Function
                
' === Error handler ===
StyleError:
    MsgBox _
        Prompt:="Error: Unable to create style '" & C_S_ParLegalNote & "'." & vbNewLine & _
            "Error number: " & Err.Number & vbNewLine & _
            "Description: " & Err.Description, _
        Buttons:=vbCritical, _
        Title:=MsgBoxTitle
    CreateStyle_ParLegalNoteMs = False
    Exit Function
        
ShortcutError:
    MsgBox _
        Prompt:="Error adding shortcut for style '" & C_S_ParLegalNote & "'." & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, _
        Buttons:=vbCritical, _
        Title:=MsgBoxTitle
    Resume Next

End Function

' 2025-11-18 by ms
Private Function CreateStyle_ParInTableMs() As Boolean
    Dim NewStyle As style
    
    Dim FileName As String:     FileName = C_F_Macros
    Dim ModuleName As String:   ModuleName = C_M_Styles
    Dim MacroName As String:    MacroName = "CreateStyle_ParInTableMs"
    Dim MsgBoxTitle As String:  MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    ' Check if the style already exists in the document
    ' If an error occurs, skip the line that caused the error and continue with the next line of code.
    On Error Resume Next
    Set NewStyle = ActiveDocument.Styles(C_S_ParInTable)
    '  Turns off any active error handler and restores default behavior
    On Error GoTo 0

    ' If the style doesn't exist, try to create it
    If NewStyle Is Nothing Then
        On Error GoTo StyleError
        Set NewStyle = ActiveDocument.Styles.Add(Name:=C_S_ParInTable, Type:=wdStyleTypeParagraph)
        On Error GoTo 0
    End If

    If Not NewStyle Is Nothing Then
        With NewStyle
            .BaseStyle = C_S_ParNormal
            .NextParagraphStyle = C_S_ParInTable
            .AutomaticallyUpdate = False
            .QuickStyle = False
            .LanguageId = wdEnglishUS
            
            ' Font formatting
            With .font
                .Name = C_FT_Body
                .Size = C_BaseFontSize
                .Bold = False
                .Italic = False
                .color = wdColorAutomatic
            End With
            
            ' Paragraph formatting
            With .ParagraphFormat
                .Alignment = wdAlignParagraphCenter
                .LeftIndent = CentimetersToPoints(0)
                .RightIndent = CentimetersToPoints(0)
                .FirstLineIndent = 0
                .SpaceBefore = 3
                .SpaceAfter = 3
                .LineSpacing = NewStyle.font.Size   ' order matters: specify at first LineSpacing, next LineSpacingRule
                .LineSpacingRule = wdLineSpaceExactly
                .WidowControl = True
                .KeepWithNext = False
                .KeepTogether = False
                .PageBreakBefore = False
                .OutlineLevel = wdOutlineLevelBodyText
            End With
        End With

'       === Add keyboard shortcut for the style ===
       On Error GoTo ShortcutError
       CustomizationContext = ActiveDocument
       KeyBindings.Add _
            KeyCategory:=wdKeyCategoryStyle, _
            Command:=C_S_ParInTable, _
            KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyN), _
            KeyCode2:=wdKeyT
       On Error GoTo 0

    CreateStyle_ParInTableMs = True
    End If
    Exit Function
                
' === Error handler ===
StyleError:
    MsgBox _
        Prompt:="Error: Unable to create style '" & C_S_ParInTable & "'." & vbNewLine & _
            "Error number: " & Err.Number & vbNewLine & _
            "Description: " & Err.Description, _
        Buttons:=vbCritical, _
        Title:=MsgBoxTitle
    CreateStyle_ParInTableMs = False
    Exit Function
        
ShortcutError:
    MsgBox _
        Prompt:="Error adding shortcut for style '" & C_S_ParInTable & "'." & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, _
        Buttons:=vbCritical, _
        Title:=MsgBoxTitle
    Resume Next

End Function

' 2025-11-18 by ms
Private Function CreateStyle_ParHeading8ms() As Boolean
    Dim NewStyle As style
    
    Dim FileName As String:     FileName = C_F_Macros
    Dim ModuleName As String:   ModuleName = C_M_Styles
    Dim MacroName As String:    MacroName = "CreateStyle_ParHeading8ms"
    Dim MsgBoxTitle As String:  MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    ' Check if the style already exists in the document
    ' If an error occurs, skip the line that caused the error and continue with the next line of code.
    On Error Resume Next
    Set NewStyle = ActiveDocument.Styles(C_S_Heading8)
    '  Turns off any active error handler and restores default behavior
    On Error GoTo 0

    ' If the style doesn't exist, try to create it
    If NewStyle Is Nothing Then
        On Error GoTo StyleError
        Set NewStyle = ActiveDocument.Styles.Add(Name:=C_S_Heading8, Type:=wdStyleTypeParagraph)
        On Error GoTo 0
    End If

    If Not NewStyle Is Nothing Then
        With NewStyle
            .BaseStyle = C_S_ParNormal
            .NextParagraphStyle = C_S_ParNormal
            .AutomaticallyUpdate = False
            .QuickStyle = False
            .LanguageId = wdEnglishUS
            
            ' Font formatting
            With .font
                .Name = C_FT_Headings
                .Size = C_BaseFontSize
                .Bold = False
                .Italic = False
                .color = GetThemeColor(ThemeColorIndex:=wdThemeColorAccent1, TintAndShade:=0)   ' in module Template
            End With
            
            ' Paragraph formatting
            With .ParagraphFormat
                .Alignment = wdAlignParagraphLeft
                .LeftIndent = CentimetersToPoints(0)
                .RightIndent = CentimetersToPoints(0)
                .FirstLineIndent = 0
                .SpaceBefore = 24
                .SpaceAfter = 6
                .LineSpacing = NewStyle.font.Size   ' order matters: specify at first LineSpacing, next LineSpacingRule
                .LineSpacingRule = wdLineSpaceExactly
                .WidowControl = True
                .KeepWithNext = True
                .KeepTogether = False
                .PageBreakBefore = False
                .OutlineLevel = wdOutlineLevel8
            End With
        End With

'       === Add keyboard shortcut for the style ===
       On Error GoTo ShortcutError
       CustomizationContext = ActiveDocument
       KeyBindings.Add _
            KeyCategory:=wdKeyCategoryStyle, _
            Command:=C_S_Heading8, _
            KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyControl, wdKey8)
       On Error GoTo 0

    CreateStyle_ParHeading8ms = True
    End If
    Exit Function
                
' === Error handler ===
StyleError:
    MsgBox _
        Prompt:="Error: Unable to create style '" & C_S_Heading8 & "'." & vbNewLine & _
            "Error number: " & Err.Number & vbNewLine & _
            "Description: " & Err.Description, _
        Buttons:=vbCritical, _
        Title:=MsgBoxTitle
    CreateStyle_ParHeading8ms = False
    Exit Function
        
ShortcutError:
    MsgBox _
        Prompt:="Error adding shortcut for style '" & C_S_Heading8 & "'." & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, _
        Buttons:=vbCritical, _
        Title:=MsgBoxTitle
    Resume Next

End Function

' 2025-11-18 by ms
Private Function CreateStyle_ParHeading7ms() As Boolean
    Dim NewStyle As style
    
    Dim FileName As String:     FileName = C_F_Macros
    Dim ModuleName As String:   ModuleName = C_M_Styles
    Dim MacroName As String:    MacroName = "CreateStyle_ParHeading7ms"
    Dim MsgBoxTitle As String:  MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    ' Check if the style already exists in the document
    ' If an error occurs, skip the line that caused the error and continue with the next line of code.
    On Error Resume Next
    Set NewStyle = ActiveDocument.Styles(C_S_Heading7)
    '  Turns off any active error handler and restores default behavior
    On Error GoTo 0

    ' If the style doesn't exist, try to create it
    If NewStyle Is Nothing Then
        On Error GoTo StyleError
        Set NewStyle = ActiveDocument.Styles.Add(Name:=C_S_Heading7, Type:=wdStyleTypeParagraph)
        On Error GoTo 0
    End If

    If Not NewStyle Is Nothing Then
        With NewStyle
            .BaseStyle = C_S_ParNormal
            .NextParagraphStyle = C_S_ParNormal
            .AutomaticallyUpdate = False
            .QuickStyle = False
            .LanguageId = wdEnglishUS
            
            ' Font formatting
            With .font
                .Name = C_FT_Headings
                .Size = C_BaseFontSize
                .Bold = False
                .Italic = False
                .color = GetThemeColor(ThemeColorIndex:=wdThemeColorAccent1, TintAndShade:=0)   ' in module Template
            End With
            
            ' Paragraph formatting
            With .ParagraphFormat
                .Alignment = wdAlignParagraphLeft
                .LeftIndent = CentimetersToPoints(0)
                .RightIndent = CentimetersToPoints(0)
                .FirstLineIndent = 0
                .SpaceBefore = 24
                .SpaceAfter = 6
                .LineSpacing = NewStyle.font.Size   ' order matters: specify at first LineSpacing, next LineSpacingRule
                .LineSpacingRule = wdLineSpaceExactly
                .WidowControl = True
                .KeepWithNext = True
                .KeepTogether = False
                .PageBreakBefore = False
                .OutlineLevel = wdOutlineLevel7
            End With
        End With

'       === Add keyboard shortcut for the style ===
       On Error GoTo ShortcutError
       CustomizationContext = ActiveDocument
       KeyBindings.Add _
            KeyCategory:=wdKeyCategoryStyle, _
            Command:=C_S_Heading7, _
            KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyControl, wdKey7)
       On Error GoTo 0

    CreateStyle_ParHeading7ms = True
    End If
    Exit Function
                
' === Error handler ===
StyleError:
    MsgBox _
        Prompt:="Error: Unable to create style '" & C_S_Heading7 & "'." & vbNewLine & _
            "Error number: " & Err.Number & vbNewLine & _
            "Description: " & Err.Description, _
        Buttons:=vbCritical, _
        Title:=MsgBoxTitle
    CreateStyle_ParHeading7ms = False
    Exit Function
        
ShortcutError:
    MsgBox _
        Prompt:="Error adding shortcut for style '" & C_S_Heading7 & "'." & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, _
        Buttons:=vbCritical, _
        Title:=MsgBoxTitle
    Resume Next

End Function

' 2025-11-18 by ms
Private Function CreateStyle_ParHeading6ms() As Boolean
    Dim NewStyle As style
    
    Dim FileName As String:     FileName = C_F_Macros
    Dim ModuleName As String:   ModuleName = C_M_Styles
    Dim MacroName As String:    MacroName = "CreateStyle_ParHeading6ms"
    Dim MsgBoxTitle As String:  MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    ' Check if the style already exists in the document
    ' If an error occurs, skip the line that caused the error and continue with the next line of code.
    On Error Resume Next
    Set NewStyle = ActiveDocument.Styles(C_S_Heading6)
    '  Turns off any active error handler and restores default behavior
    On Error GoTo 0

    ' If the style doesn't exist, try to create it
    If NewStyle Is Nothing Then
        On Error GoTo StyleError
        Set NewStyle = ActiveDocument.Styles.Add(Name:=C_S_Heading6, Type:=wdStyleTypeParagraph)
        On Error GoTo 0
    End If

    If Not NewStyle Is Nothing Then
        With NewStyle
            .BaseStyle = C_S_ParNormal
            .NextParagraphStyle = C_S_ParNormal
            .AutomaticallyUpdate = False
            .QuickStyle = False
            .LanguageId = wdEnglishUS
            
            ' Font formatting
            With .font
                .Name = C_FT_Headings
                .Size = C_BaseFontSize
                .Bold = False
                .Italic = False
                .color = GetThemeColor(ThemeColorIndex:=wdThemeColorAccent1, TintAndShade:=0)   ' in module Template
            End With
            
            ' Paragraph formatting
            With .ParagraphFormat
                .Alignment = wdAlignParagraphLeft
                .LeftIndent = CentimetersToPoints(0)
                .RightIndent = CentimetersToPoints(0)
                .FirstLineIndent = 0
                .SpaceBefore = 24
                .SpaceAfter = 6
                .LineSpacing = NewStyle.font.Size   ' order matters: specify at first LineSpacing, next LineSpacingRule
                .LineSpacingRule = wdLineSpaceExactly
                .WidowControl = True
                .KeepWithNext = True
                .KeepTogether = False
                .PageBreakBefore = False
                .OutlineLevel = wdOutlineLevel6
            End With
        End With

'       === Add keyboard shortcut for the style ===
       On Error GoTo ShortcutError
       CustomizationContext = ActiveDocument
       KeyBindings.Add _
            KeyCategory:=wdKeyCategoryStyle, _
            Command:=C_S_Heading6, _
            KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyControl, wdKey6)
       On Error GoTo 0

    CreateStyle_ParHeading6ms = True
    End If
    Exit Function
                
' === Error handler ===
StyleError:
    MsgBox _
        Prompt:="Error: Unable to create style '" & C_S_Heading6 & "'." & vbNewLine & _
            "Error number: " & Err.Number & vbNewLine & _
            "Description: " & Err.Description, _
        Buttons:=vbCritical, _
        Title:=MsgBoxTitle
    CreateStyle_ParHeading6ms = False
    Exit Function
        
ShortcutError:
    MsgBox _
        Prompt:="Error adding shortcut for style '" & C_S_Heading6 & "'." & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, _
        Buttons:=vbCritical, _
        Title:=MsgBoxTitle
    Resume Next

End Function

' 2025-11-18 by ms
Private Function CreateStyle_ParHeading5ms() As Boolean
    Dim NewStyle As style
    
    Dim FileName As String:     FileName = C_F_Macros
    Dim ModuleName As String:   ModuleName = C_M_Styles
    Dim MacroName As String:    MacroName = "CreateStyle_ParHeading5ms"
    Dim MsgBoxTitle As String:  MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    ' Check if the style already exists in the document
    ' If an error occurs, skip the line that caused the error and continue with the next line of code.
    On Error Resume Next
    Set NewStyle = ActiveDocument.Styles(C_S_Heading5)
    '  Turns off any active error handler and restores default behavior
    On Error GoTo 0

    ' If the style doesn't exist, try to create it
    If NewStyle Is Nothing Then
        On Error GoTo StyleError
        Set NewStyle = ActiveDocument.Styles.Add(Name:=C_S_Heading5, Type:=wdStyleTypeParagraph)
        On Error GoTo 0
    End If

    If Not NewStyle Is Nothing Then
        With NewStyle
            .BaseStyle = C_S_ParNormal
            .NextParagraphStyle = C_S_ParNormal
            .AutomaticallyUpdate = False
            .QuickStyle = False
            .LanguageId = wdEnglishUS
            
            ' Font formatting
            With .font
                .Name = C_FT_Headings
                .Size = C_BaseFontSize
                .Bold = False
                .Italic = False
                .color = GetThemeColor(ThemeColorIndex:=wdThemeColorAccent1, TintAndShade:=0)   ' in module Template
            End With
            
            ' Paragraph formatting
            With .ParagraphFormat
                .Alignment = wdAlignParagraphLeft
                .LeftIndent = CentimetersToPoints(0)
                .RightIndent = CentimetersToPoints(0)
                .FirstLineIndent = 0
                .SpaceBefore = 24
                .SpaceAfter = 6
                .LineSpacing = NewStyle.font.Size   ' order matters: specify at first LineSpacing, next LineSpacingRule
                .LineSpacingRule = wdLineSpaceExactly
                .WidowControl = True
                .KeepWithNext = True
                .KeepTogether = False
                .PageBreakBefore = False
                .OutlineLevel = wdOutlineLevel5
            End With
        End With

'       === Add keyboard shortcut for the style ===
       On Error GoTo ShortcutError
       CustomizationContext = ActiveDocument
       KeyBindings.Add _
            KeyCategory:=wdKeyCategoryStyle, _
            Command:=C_S_Heading5, _
            KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyControl, wdKey5)
       On Error GoTo 0

    CreateStyle_ParHeading5ms = True
    End If
    Exit Function
                
' === Error handler ===
StyleError:
    MsgBox _
        Prompt:="Error: Unable to create style '" & C_S_Heading5 & "'." & vbNewLine & _
            "Error number: " & Err.Number & vbNewLine & _
            "Description: " & Err.Description, _
        Buttons:=vbCritical, _
        Title:=MsgBoxTitle
    CreateStyle_ParHeading5ms = False
    Exit Function
        
ShortcutError:
    MsgBox _
        Prompt:="Error adding shortcut for style '" & C_S_Heading5 & "'." & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, _
        Buttons:=vbCritical, _
        Title:=MsgBoxTitle
    Resume Next

End Function

' 2025-11-18 by ms
Private Function CreateStyle_ParHeading4ms() As Boolean
    Dim NewStyle As style
    
    Dim FileName As String:     FileName = C_F_Macros
    Dim ModuleName As String:   ModuleName = C_M_Styles
    Dim MacroName As String:    MacroName = "CreateStyle_ParHeading4ms"
    Dim MsgBoxTitle As String:  MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    ' Check if the style already exists in the document
    ' If an error occurs, skip the line that caused the error and continue with the next line of code.
    On Error Resume Next
    Set NewStyle = ActiveDocument.Styles(C_S_Heading4)
    '  Turns off any active error handler and restores default behavior
    On Error GoTo 0

    ' If the style doesn't exist, try to create it
    If NewStyle Is Nothing Then
        On Error GoTo StyleError
        Set NewStyle = ActiveDocument.Styles.Add(Name:=C_S_Heading4, Type:=wdStyleTypeParagraph)
        On Error GoTo 0
    End If

    If Not NewStyle Is Nothing Then
        With NewStyle
            .BaseStyle = C_S_ParNormal
            .NextParagraphStyle = C_S_ParNormal
            .AutomaticallyUpdate = False
            .QuickStyle = False
            .LanguageId = wdEnglishUS
            
            ' Font formatting
            With .font
                .Name = C_FT_Headings
                .Size = C_BaseFontSize
                .Bold = False
                .Italic = False
                .color = GetThemeColor(ThemeColorIndex:=wdThemeColorAccent1, TintAndShade:=0)   ' in module Template
            End With
            
            ' Paragraph formatting
            With .ParagraphFormat
                .Alignment = wdAlignParagraphLeft
                .LeftIndent = CentimetersToPoints(0)
                .RightIndent = CentimetersToPoints(0)
                .FirstLineIndent = 0
                .SpaceBefore = 24
                .SpaceAfter = 6
                .LineSpacing = NewStyle.font.Size   ' order matters: specify at first LineSpacing, next LineSpacingRule
                .LineSpacingRule = wdLineSpaceExactly
                .WidowControl = True
                .KeepWithNext = True
                .KeepTogether = False
                .PageBreakBefore = False
                .OutlineLevel = wdOutlineLevel4
            End With
        End With

'       === Add keyboard shortcut for the style ===
       On Error GoTo ShortcutError
       CustomizationContext = ActiveDocument
       KeyBindings.Add _
            KeyCategory:=wdKeyCategoryStyle, _
            Command:=C_S_Heading4, _
            KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyControl, wdKey4)
       On Error GoTo 0

    CreateStyle_ParHeading4ms = True
    End If
    Exit Function
                
' === Error handler ===
StyleError:
    MsgBox _
        Prompt:="Error: Unable to create style '" & C_S_Heading4 & "'." & vbNewLine & _
            "Error number: " & Err.Number & vbNewLine & _
            "Description: " & Err.Description, _
        Buttons:=vbCritical, _
        Title:=MsgBoxTitle
    CreateStyle_ParHeading4ms = False
    Exit Function
        
ShortcutError:
    MsgBox _
        Prompt:="Error adding shortcut for style '" & C_S_Heading4 & "'." & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, _
        Buttons:=vbCritical, _
        Title:=MsgBoxTitle
    Resume Next

End Function

' 2025-11-18 by ms
Private Function CreateStyle_ParHeading3ms() As Boolean
    Dim NewStyle As style
    
    Dim FileName As String:     FileName = C_F_Macros
    Dim ModuleName As String:   ModuleName = C_M_Styles
    Dim MacroName As String:    MacroName = "CreateStyle_ParHeading3ms"
    Dim MsgBoxTitle As String:  MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    ' Check if the style already exists in the document
    ' If an error occurs, skip the line that caused the error and continue with the next line of code.
    On Error Resume Next
    Set NewStyle = ActiveDocument.Styles(C_S_Heading3)
    '  Turns off any active error handler and restores default behavior
    On Error GoTo 0

    ' If the style doesn't exist, try to create it
    If NewStyle Is Nothing Then
        On Error GoTo StyleError
        Set NewStyle = ActiveDocument.Styles.Add(Name:=C_S_Heading3, Type:=wdStyleTypeParagraph)
        On Error GoTo 0
    End If

    If Not NewStyle Is Nothing Then
        With NewStyle
            .BaseStyle = C_S_ParNormal
            .NextParagraphStyle = C_S_ParNormal
            .AutomaticallyUpdate = False
            .QuickStyle = False
            .LanguageId = wdEnglishUS
            
            ' Font formatting
            With .font
                .Name = C_FT_Headings
                .Size = C_BaseFontSize + 2
                .Bold = False
                .Italic = False
                .color = GetThemeColor(ThemeColorIndex:=wdThemeColorAccent1, TintAndShade:=0)   ' in module Template
            End With
            
            ' Paragraph formatting
            With .ParagraphFormat
                .Alignment = wdAlignParagraphLeft
                .LeftIndent = CentimetersToPoints(0)
                .RightIndent = CentimetersToPoints(0)
                .FirstLineIndent = 0
                .SpaceBefore = 26
                .SpaceAfter = 6
                .LineSpacing = NewStyle.font.Size   ' order matters: specify at first LineSpacing, next LineSpacingRule
                .LineSpacingRule = wdLineSpaceExactly
                .WidowControl = True
                .KeepWithNext = True
                .KeepTogether = False
                .PageBreakBefore = False
                .OutlineLevel = wdOutlineLevel3
            End With
        End With

'       === Add keyboard shortcut for the style ===
       On Error GoTo ShortcutError
       CustomizationContext = ActiveDocument
       KeyBindings.Add _
            KeyCategory:=wdKeyCategoryStyle, _
            Command:=C_S_Heading3, _
            KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyControl, wdKey3)
       On Error GoTo 0

    CreateStyle_ParHeading3ms = True
    End If
    Exit Function
                
' === Error handler ===
StyleError:
    MsgBox _
        Prompt:="Error: Unable to create style '" & C_S_Heading3 & "'." & vbNewLine & _
            "Error number: " & Err.Number & vbNewLine & _
            "Description: " & Err.Description, _
        Buttons:=vbCritical, _
        Title:=MsgBoxTitle
    CreateStyle_ParHeading3ms = False
    Exit Function
        
ShortcutError:
    MsgBox _
        Prompt:="Error adding shortcut for style '" & C_S_Heading3 & "'." & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, _
        Buttons:=vbCritical, _
        Title:=MsgBoxTitle
    Resume Next

End Function

' 2025-11-17 by ms
Private Function CreateStyle_ParHeading2ms() As Boolean
    Dim NewStyle As style
    
    Dim FileName As String:     FileName = C_F_Macros
    Dim ModuleName As String:   ModuleName = C_M_Styles
    Dim MacroName As String:    MacroName = "CreateStyle_ParHeading2ms"
    Dim MsgBoxTitle As String:  MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    ' Check if the style already exists in the document
    ' If an error occurs, skip the line that caused the error and continue with the next line of code.
    On Error Resume Next
    Set NewStyle = ActiveDocument.Styles(C_S_Heading2)
    '  Turns off any active error handler and restores default behavior
    On Error GoTo 0

    ' If the style doesn't exist, try to create it
    If NewStyle Is Nothing Then
        On Error GoTo StyleError
        Set NewStyle = ActiveDocument.Styles.Add(Name:=C_S_Heading2, Type:=wdStyleTypeParagraph)
        On Error GoTo 0
    End If

    If Not NewStyle Is Nothing Then
        With NewStyle
            .BaseStyle = C_S_ParNormal
            .NextParagraphStyle = C_S_ParNormal
            .AutomaticallyUpdate = False
            .QuickStyle = False
            .LanguageId = wdEnglishUS
            
            ' Font formatting
            With .font
                .Name = C_FT_Headings
                .Size = C_BaseFontSize + 2
                .Bold = False
                .Italic = False
                .color = GetThemeColor(ThemeColorIndex:=wdThemeColorAccent1, TintAndShade:=0)   ' in module Template
            End With
            
            ' Paragraph formatting
            With .ParagraphFormat
                .Alignment = wdAlignParagraphLeft
                .LeftIndent = CentimetersToPoints(0)
                .RightIndent = CentimetersToPoints(0)
                .FirstLineIndent = 0
                .SpaceBefore = 24
                .SpaceAfter = 6
                .LineSpacing = NewStyle.font.Size   ' order matters: specify at first LineSpacing, next LineSpacingRule
                .LineSpacingRule = wdLineSpaceAtLeast
                .WidowControl = True
                .KeepWithNext = True
                .KeepTogether = False
                .PageBreakBefore = False
                .OutlineLevel = wdOutlineLevel2
            End With
        End With

'       === Add keyboard shortcut for the style ===
       On Error GoTo ShortcutError
       CustomizationContext = ActiveDocument
       KeyBindings.Add _
            KeyCategory:=wdKeyCategoryStyle, _
            Command:=C_S_Heading2, _
            KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyControl, wdKey2)
       On Error GoTo 0

    CreateStyle_ParHeading2ms = True
    End If
    Exit Function
                
' === Error handler ===
StyleError:
    MsgBox _
        Prompt:="Error: Unable to create style '" & C_S_Heading2 & "'." & vbNewLine & _
            "Error number: " & Err.Number & vbNewLine & _
            "Description: " & Err.Description, _
        Buttons:=vbCritical, _
        Title:=MsgBoxTitle
    CreateStyle_ParHeading2ms = False
    Exit Function
        
ShortcutError:
    MsgBox _
        Prompt:="Error adding shortcut for style '" & C_S_Heading2 & "'." & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, _
        Buttons:=vbCritical, _
        Title:=MsgBoxTitle
    Resume Next

End Function

' 2025-11-17 by ms
Private Function CreateStyle_ParHeading1ms() As Boolean
    Dim NewStyle As style
    
    Dim FileName As String:      FileName = C_F_Macros
    Dim ModuleName As String:    ModuleName = C_M_Styles
    Dim MacroName As String:     MacroName = "CreateStyle_ParHeading1ms"
    Dim MsgBoxTitle As String:   MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    ' Check if the style already exists in the document
    ' If an error occurs, skip the line that caused the error and continue with the next line of code.
    On Error Resume Next
    Set NewStyle = ActiveDocument.Styles(C_S_Heading1)
    '  Turns off any active error handler and restores default behavior
    On Error GoTo 0

    ' If the style doesn't exist, try to create it
    If NewStyle Is Nothing Then
        On Error GoTo StyleError
        Set NewStyle = ActiveDocument.Styles.Add(Name:=C_S_Heading1, Type:=wdStyleTypeParagraph)
        On Error GoTo 0
    End If

    If Not NewStyle Is Nothing Then
        With NewStyle
            .BaseStyle = C_S_ParNormal
            .NextParagraphStyle = C_S_ParNormal
            .AutomaticallyUpdate = False
            .QuickStyle = False
            .LanguageId = wdEnglishUS
            
            ' Font formatting
            With .font
                .Name = C_FT_Headings
                .Size = C_BaseFontSize + 3
                .Bold = False
                .Italic = False
                .color = GetThemeColor(ThemeColorIndex:=wdThemeColorAccent1, TintAndShade:=0)   ' in module Template
            End With
            
            ' Paragraph formatting
            With .ParagraphFormat
                .Alignment = wdAlignParagraphLeft
                .LeftIndent = CentimetersToPoints(0)
                .RightIndent = CentimetersToPoints(0)
                .FirstLineIndent = 0
                .SpaceBefore = 24
                .SpaceAfter = 6
                .LineSpacing = NewStyle.font.Size  ' order matters: specify at first LineSpacing, next LineSpacingRule
                .LineSpacingRule = wdLineSpaceAtLeast
                .WidowControl = True
                .KeepWithNext = True
                .KeepTogether = True
                .PageBreakBefore = True
                .OutlineLevel = wdOutlineLevel1
            End With
        End With

'       === Add keyboard shortcut for the style ===
       On Error GoTo ShortcutError
       CustomizationContext = ActiveDocument
       KeyBindings.Add _
            KeyCategory:=wdKeyCategoryStyle, _
            Command:=C_S_Heading1, _
            KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyControl, wdKey1)
       On Error GoTo 0

    CreateStyle_ParHeading1ms = True
    End If
    Exit Function
                
' === Error handler ===
StyleError:
    MsgBox _
        Prompt:="Error: Unable to create style '" & C_S_Heading1 & "'." & vbNewLine & _
            "Error number: " & Err.Number & vbNewLine & _
            "Description: " & Err.Description, _
        Buttons:=vbCritical, _
        Title:=MsgBoxTitle
    CreateStyle_ParHeading1ms = False
    Exit Function
        
ShortcutError:
    MsgBox _
        Prompt:="Error adding shortcut for style '" & C_S_Heading1 & "'." & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, _
        Buttons:=vbCritical, _
        Title:=MsgBoxTitle
    Resume Next

End Function

' One additional style, to get rid of default styling for "Normal"
' 2026-01-15 by ms
Private Function CreateStyle_Normal() As Boolean
    Dim NewStyle As style
    
    Dim FileName As String:     FileName = C_F_Macros
    Dim ModuleName As String:   ModuleName = C_M_Styles
    Dim MacroName As String:    MacroName = "CreateStyle_Normal"
    Dim MsgBoxTitle As String:  MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    ' Check if the style already exists in the document
    ' If an error occurs, skip the line that caused the error and continue with the next line of code.
    On Error Resume Next
    Set NewStyle = ActiveDocument.Styles("Normal")
    '  Turns off any active error handler and restores default behavior
    On Error GoTo 0

    ' If the style doesn't exist, try to create it
    If NewStyle Is Nothing Then
        On Error GoTo StyleError
        Set NewStyle = ActiveDocument.Styles.Add(Name:="Normal", Type:=wdStyleTypeParagraph)
        On Error GoTo 0
    End If

    If Not NewStyle Is Nothing Then
        With NewStyle
            ' .BaseStyle = "Normal" run-time error 5648: the built-in styles Normal and Default Paragraph Font cannot be based on any style
            .NextParagraphStyle = "Normal"
            .AutomaticallyUpdate = False
            .QuickStyle = False
            .LanguageId = wdEnglishUS
            
            ' Font formatting
            With .font
                .Name = C_FT_Body
                .Size = C_BaseFontSize
                .Bold = False
                .Italic = False
                .color = wdColorAutomatic
            End With
            
            ' Paragraph formatting
            With .ParagraphFormat
                .Alignment = wdAlignParagraphLeft
                .LeftIndent = CentimetersToPoints(0)
                .RightIndent = CentimetersToPoints(0)
                .FirstLineIndent = 0
                .SpaceBefore = 0
                .SpaceAfter = 0
                .LineSpacing = NewStyle.font.Size   ' order matters: specify at first LineSpacing, next LineSpacingRule
                .LineSpacingRule = wdLineSpaceExactly
                .WidowControl = True
                .KeepWithNext = False
                .KeepTogether = False
                .PageBreakBefore = False
                .OutlineLevel = wdOutlineLevelBodyText
            End With
        End With

'       === Add keyboard shortcut for the style ===
'       On Error GoTo ShortcutError
'       CustomizationContext = ActiveDocument
'       KeyBindings.Add _
'            KeyCategory:=wdKeyCategoryStyle, _
'            Command:=C_S_ParNormal, _
'            KeyCode:=BuildKeyCode(wdKeyControl, wdKeyShift, wdKeyN)
'       On Error GoTo 0

    CreateStyle_Normal = True
    End If
    Exit Function
                
' === Error handler ===
StyleError:
    MsgBox _
        Prompt:="Error: Unable to create style '" & C_S_ParNormal & "'." & vbNewLine & _
            "Error number: " & Err.Number & vbNewLine & _
            "Description: " & Err.Description, _
        Buttons:=vbCritical, _
        Title:=MsgBoxTitle
    CreateStyle_Normal = False
    Exit Function
        
ShortcutError:
    MsgBox _
        Prompt:="Error adding shortcut for style '" & "Normal" & "'." & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, _
        Buttons:=vbCritical, _
        Title:=MsgBoxTitle
    Resume Next

End Function

' 2025-11-16 by ms
Private Function CreateStyle_ParNormalMs() As Boolean
    Dim NewStyle As style
    
    Dim FileName As String:     FileName = C_F_Macros
    Dim ModuleName As String:   ModuleName = C_M_Styles
    Dim MacroName As String:    MacroName = "CreateStyle_ParNormalMs"
    Dim MsgBoxTitle As String:  MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    ' Check if the style already exists in the document
    ' If an error occurs, skip the line that caused the error and continue with the next line of code.
    On Error Resume Next
    Set NewStyle = ActiveDocument.Styles(C_S_ParNormal)
    '  Turns off any active error handler and restores default behavior
    On Error GoTo 0

    ' If the style doesn't exist, try to create it
    If NewStyle Is Nothing Then
        On Error GoTo StyleError
        Set NewStyle = ActiveDocument.Styles.Add(Name:=C_S_ParNormal, Type:=wdStyleTypeParagraph)
        On Error GoTo 0
    End If

    If Not NewStyle Is Nothing Then
        With NewStyle
            .BaseStyle = "Normal"
            .NextParagraphStyle = C_S_ParNormal
            .AutomaticallyUpdate = False
            .QuickStyle = False
            .LanguageId = wdEnglishUS
            
            ' Font formatting
            With .font
                .Name = C_FT_Body
                .Size = C_BaseFontSize
                .Bold = False
                .Italic = False
                .color = wdColorAutomatic
            End With
            
            ' Paragraph formatting
            With .ParagraphFormat
                .Alignment = wdAlignParagraphLeft
                .LeftIndent = CentimetersToPoints(0)
                .RightIndent = CentimetersToPoints(0)
                .FirstLineIndent = 0
                .SpaceBefore = 0
                .SpaceAfter = 6
                .LineSpacing = NewStyle.font.Size   ' order matters: specify at first LineSpacing, next LineSpacingRule
                .LineSpacingRule = wdLineSpaceExactly
                .WidowControl = True
                .KeepWithNext = False
                .KeepTogether = False
                .PageBreakBefore = False
                .OutlineLevel = wdOutlineLevelBodyText
            End With
        End With

'       === Add keyboard shortcut for the style ===
       On Error GoTo ShortcutError
       CustomizationContext = ActiveDocument
       KeyBindings.Add _
            KeyCategory:=wdKeyCategoryStyle, _
            Command:=C_S_ParNormal, _
            KeyCode:=BuildKeyCode(wdKeyControl, wdKeyShift, wdKeyN)
       On Error GoTo 0

    CreateStyle_ParNormalMs = True
    End If
    Exit Function
                
' === Error handler ===
StyleError:
    MsgBox _
        Prompt:="Error: Unable to create style '" & C_S_ParNormal & "'." & vbNewLine & _
            "Error number: " & Err.Number & vbNewLine & _
            "Description: " & Err.Description, _
        Buttons:=vbCritical, _
        Title:=MsgBoxTitle
    CreateStyle_ParNormalMs = False
    Exit Function
        
ShortcutError:
    MsgBox _
        Prompt:="Error adding shortcut for style '" & C_S_ParNormal & "'." & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, _
        Buttons:=vbCritical, _
        Title:=MsgBoxTitle
    Resume Next

End Function

' 2025-10-07 by ms
' 2025-11-15 by ms
Private Function CreateStyle_ParSourceCodeMs() As Boolean
    Dim NewStyle As style
    
    Dim FileName As String:      FileName = C_F_Macros
    Dim ModuleName As String:    ModuleName = C_M_Styles
    Dim MacroName As String:     MacroName = "CreateStyle_ParSourceCodeMs"
    Dim MsgBoxTitle As String:   MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    ' Check if the style already exists in the document
    ' If an error occurs, skip the line that caused the error and continue with the next line of code.
    On Error Resume Next
    Set NewStyle = ActiveDocument.Styles(C_S_ParSourceCode)
    '  Turns off any active error handler and restores default behavior
    On Error GoTo 0

    ' If the style doesn't exist, try to create it
    If NewStyle Is Nothing Then
        On Error GoTo StyleError
        Set NewStyle = ActiveDocument.Styles.Add(Name:=C_S_ParSourceCode, Type:=wdStyleTypeParagraph)
        On Error GoTo 0
    End If

    If Not NewStyle Is Nothing Then
        With NewStyle
            .BaseStyle = C_S_ParNormal
            .NextParagraphStyle = C_S_ParNormal
            .AutomaticallyUpdate = False
            .QuickStyle = False
            .LanguageId = wdEnglishUS
            
            ' Font formatting
            With .font
                .Name = C_FT_AntiHomoglyph
                .Size = C_BaseFontSize
                .Bold = False
                .Italic = False
                .color = wdColorAutomatic
            End With
            
            ' Paragraph formatting
            With .ParagraphFormat
                .Alignment = wdAlignParagraphLeft
                .LeftIndent = CentimetersToPoints(0.3)
                .RightIndent = CentimetersToPoints(0.3)
                .FirstLineIndent = 0
                .SpaceBefore = 12
                .SpaceAfter = 12
                .LineSpacing = NewStyle.font.Size   ' order matters: specify at first LineSpacing, next LineSpacingRule
                .LineSpacingRule = wdLineSpaceExactly
                .WidowControl = True
                .KeepWithNext = False
                .KeepTogether = False
                .PageBreakBefore = False
                .OutlineLevel = wdOutlineLevelBodyText
            End With
        End With
        
        With NewStyle.ParagraphFormat.borders
            .Enable = True
            .OutsideLineStyle = wdLineStyleSingle
            .OutsideLineWidth = wdLineWidth100pt
            .OutsideColor = GetThemeColor(ThemeColorIndex:=wdThemeColorText2, TintAndShade:=0)   ' in module Template
        End With

'       === Add keyboard shortcut for the style ===
'       On Error GoTo ShortcutError
'       CustomizationContext = ActiveDocument
'       KeyBindings.Add _
'            KeyCategory:=wdKeyCategoryStyle, _
'            Command:=C_S_ParSourceCode, _
'            KeyCode:=BuildKeyCode(wdKeyControl, wdKeyAlt, wdKeyS)
'       On Error GoTo 0

    CreateStyle_ParSourceCodeMs = True
    End If
    Exit Function
                
' === Error handler ===
StyleError:
    MsgBox _
        Prompt:="Error: Unable to create style '" & C_S_ParSourceCode & "'." & vbNewLine & _
            "Error number: " & Err.Number & vbNewLine & _
            "Description: " & Err.Description, _
        Buttons:=vbCritical, _
        Title:=MsgBoxTitle
    CreateStyle_ParSourceCodeMs = False
    Exit Function
        
ShortcutError:
    MsgBox _
        Prompt:="Error adding shortcut for style '" & C_S_ParSourceCode & "'." & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, _
        Buttons:=vbCritical, _
        Title:=MsgBoxTitle
    Resume Next
End Function

' 2025-11-15 by ms
Private Function CreateStyle_ParPictureCanvaMs() As Boolean
    Dim NewStyle As style
    
    Dim FileName As String:     FileName = C_F_Macros
    Dim ModuleName As String:   ModuleName = C_M_Styles
    Dim MacroName As String:    MacroName = "CreateStyle_ParPictureCanvaMs"
    Dim MsgBoxTitle As String:  MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    ' Check if the style already exists in the document
    ' If an error occurs, skip the line that caused the error and continue with the next line of code.
    On Error Resume Next
    Set NewStyle = ActiveDocument.Styles(C_S_ParPictureCanva)
    '  Turns off any active error handler and restores default behavior
    On Error GoTo 0

    ' If the style doesn't exist, try to create it
    If NewStyle Is Nothing Then
        On Error GoTo StyleError
        Set NewStyle = ActiveDocument.Styles.Add(Name:=C_S_ParPictureCanva, Type:=wdStyleTypeParagraph)
        On Error GoTo 0
    End If

    If Not NewStyle Is Nothing Then
        With NewStyle
            .BaseStyle = C_S_ParNormal
            .NextParagraphStyle = C_S_ParNormal
            .AutomaticallyUpdate = False
            .QuickStyle = False
            .LanguageId = wdEnglishUS
            
            ' Font formatting
            With .font
                .Name = C_FT_Body
                .Size = C_BaseFontSize
                .Bold = False
                .Italic = False
                .color = wdColorAutomatic
            End With
            
            ' Paragraph formatting
            With .ParagraphFormat
                .Alignment = wdAlignParagraphCenter
                .LeftIndent = CentimetersToPoints(0)
                .RightIndent = CentimetersToPoints(0)
                .FirstLineIndent = 0
                .SpaceBefore = 12
                .SpaceAfter = 6
                .LineSpacingRule = wdLineSpaceSingle        ' exception for this style, without it pictures / canvas are strangly inserted in document content
                .WidowControl = True
                .KeepWithNext = True
                .KeepTogether = False
                .PageBreakBefore = False
                .OutlineLevel = wdOutlineLevelBodyText
            End With
        End With

'       === Add keyboard shortcut for the style ===
'       On Error GoTo ShortcutError
'       CustomizationContext = ActiveDocument
'       KeyBindings.Add _
'            KeyCategory:=wdKeyCategoryStyle, _
'            Command:=C_S_ParSourceCode, _
'            KeyCode:=BuildKeyCode(wdKeyControl, wdKeyAlt, wdKeyS)
'       On Error GoTo 0

    CreateStyle_ParPictureCanvaMs = True
    End If
    Exit Function
                
' === Error handler ===
StyleError:
    MsgBox _
        Prompt:="Error: Unable to create style '" & C_S_ParPictureCanva & "'." & vbNewLine & _
            "Error number: " & Err.Number & vbNewLine & _
            "Description: " & Err.Description, _
        Buttons:=vbCritical, _
        Title:=MsgBoxTitle
    CreateStyle_ParPictureCanvaMs = False
    Exit Function
        
ShortcutError:
    MsgBox _
        Prompt:="Error adding shortcut for style '" & C_S_ParPictureCanva & "'." & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, _
        Buttons:=vbCritical, _
        Title:=MsgBoxTitle
    Resume Next
End Function

' 2025-12-24 by ms
Private Function CreateStyle_ParIconMs() As Boolean
    Dim NewStyle As style
    
    Dim FileName As String:     FileName = C_F_Macros
    Dim ModuleName As String:   ModuleName = C_M_Styles
    Dim MacroName As String:    MacroName = "CreateStyle_ParIconMs"
    Dim MsgBoxTitle As String:  MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    ' Check if the style already exists in the document
    ' If an error occurs, skip the line that caused the error and continue with the next line of code.
    On Error Resume Next
    Set NewStyle = ActiveDocument.Styles(C_S_ParIcon)
    '  Turns off any active error handler and restores default behavior
    On Error GoTo 0

    ' If the style doesn't exist, try to create it
    If NewStyle Is Nothing Then
        On Error GoTo StyleError
        Set NewStyle = ActiveDocument.Styles.Add(Name:=C_S_ParIcon, Type:=wdStyleTypeParagraph)
        On Error GoTo 0
    End If

    If Not NewStyle Is Nothing Then
        With NewStyle
            .BaseStyle = C_S_ParNormal
            .NextParagraphStyle = C_S_ParNormal
            .AutomaticallyUpdate = False
            .QuickStyle = False
            .LanguageId = wdEnglishUS
            
            ' Font formatting
            With .font
                .Name = C_FT_Body
                .Size = C_BaseFontSize
                .Bold = False
                .Italic = False
                .color = wdColorAutomatic
            End With
            
            ' Paragraph formatting
            With .ParagraphFormat
                .Alignment = wdAlignParagraphCenter
                .LeftIndent = CentimetersToPoints(0)
                .RightIndent = CentimetersToPoints(0)
                .FirstLineIndent = 0
                .SpaceBefore = 0
                .SpaceAfter = 0
                .LineSpacingRule = wdLineSpaceSingle        ' exception for this style, without it pictures are strangly inserted in document content
                .WidowControl = True
                .KeepWithNext = False
                .KeepTogether = False
                .PageBreakBefore = False
                .OutlineLevel = wdOutlineLevelBodyText
            End With
        End With

'       === Add keyboard shortcut for the style ===
'       On Error GoTo ShortcutError
'       CustomizationContext = ActiveDocument
'       KeyBindings.Add _
'            KeyCategory:=wdKeyCategoryStyle, _
'            Command:=C_S_ParSourceCode, _
'            KeyCode:=BuildKeyCode(wdKeyControl, wdKeyAlt, wdKeyS)
'       On Error GoTo 0

    CreateStyle_ParIconMs = True
    End If
    Exit Function
                
' === Error handler ===
StyleError:
    MsgBox _
        Prompt:="Error: Unable to create style '" & C_S_ParIcon & "'." & vbNewLine & _
            "Error number: " & Err.Number & vbNewLine & _
            "Description: " & Err.Description, _
        Buttons:=vbCritical, _
        Title:=MsgBoxTitle
    CreateStyle_ParIconMs = False
    Exit Function
        
ShortcutError:
    MsgBox _
        Prompt:="Error adding shortcut for style '" & C_S_ParIcon & "'." & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, _
        Buttons:=vbCritical, _
        Title:=MsgBoxTitle
    Resume Next
End Function

' 2025-10-07 by ms
Private Function CreateStyle_ParNormalZeroMs() As Boolean
    Dim NewStyle As style
    
    Dim FileName As String:     FileName = C_F_Macros
    Dim ModuleName As String:   ModuleName = C_M_Styles
    Dim MacroName As String:    MacroName = "CreateStyle_ParNormalZeroMs"
    Dim MsgBoxTitle As String:  MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    ' Check if the style already exists in the document
    ' If an error occurs, skip the line that caused the error and continue with the next line of code.
    On Error Resume Next
    Set NewStyle = ActiveDocument.Styles(C_S_ParNormalZero)
    '  Turns off any active error handler and restores default behavior
    On Error GoTo 0

    ' If the style doesn't exist, try to create it
    If NewStyle Is Nothing Then
        On Error GoTo StyleError
        Set NewStyle = ActiveDocument.Styles.Add(Name:=C_S_ParNormalZero, Type:=wdStyleTypeParagraph)
        On Error GoTo 0
    End If

    If Not NewStyle Is Nothing Then
        With NewStyle
            .BaseStyle = C_S_ParNormal
            .NextParagraphStyle = C_S_ParNormalZero
            .AutomaticallyUpdate = False
            .QuickStyle = False
            .LanguageId = wdEnglishUS
            
            ' Font formatting
            With .font
                .Name = C_FT_Body
                .Size = C_BaseFontSize
                .Bold = False
                .Italic = False
                .color = wdColorAutomatic
            End With
            
            ' Paragraph formatting
            With .ParagraphFormat
                .Alignment = wdAlignParagraphLeft
                .LeftIndent = CentimetersToPoints(0)
                .RightIndent = CentimetersToPoints(0)
                .FirstLineIndent = 0
                .SpaceBefore = 0
                .SpaceAfter = 0
                .LineSpacing = NewStyle.font.Size   ' order matters: specify at first LineSpacing, next LineSpacingRule
                .LineSpacingRule = wdLineSpaceExactly
                .WidowControl = True
                .KeepWithNext = False
                .KeepTogether = False
                .PageBreakBefore = False
                .OutlineLevel = wdOutlineLevelBodyText
            End With
        End With

'       === Add keyboard shortcut for the style ===
'       On Error GoTo ShortcutError
'       CustomizationContext = ActiveDocument
'       KeyBindings.Add _
'            KeyCategory:=wdKeyCategoryStyle, _
'            Command:=C_S_ParSourceCode, _
'            KeyCode:=BuildKeyCode(wdKeyControl, wdKeyAlt, wdKeyS)
'       On Error GoTo 0

    CreateStyle_ParNormalZeroMs = True
    End If
    Exit Function
                
' === Error handler ===
StyleError:
    MsgBox _
        Prompt:="Error: Unable to create style '" & C_S_ParNormalZero & "'." & vbNewLine & _
            "Error number: " & Err.Number & vbNewLine & _
            "Description: " & Err.Description, _
        Buttons:=vbCritical, _
        Title:=MsgBoxTitle
    CreateStyle_ParNormalZeroMs = False
    Exit Function
        
ShortcutError:
    MsgBox _
        Prompt:="Error adding shortcut for style '" & C_S_ParNormalZero & "'." & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, _
        Buttons:=vbCritical, _
        Title:=MsgBoxTitle
    Resume Next

End Function

' 2025-10-07 by ms
Private Function CreateStyle_ParNormalBelowMs() As Boolean
    Dim NewStyle As style
    
    Dim FileName As String:     FileName = C_F_Macros
    Dim ModuleName As String:   ModuleName = C_M_Styles
    Dim MacroName As String:    MacroName = "CreateStyle_ParNormalBelowMs"
    Dim MsgBoxTitle As String:  MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    ' Check if the style already exists in the document
    ' If an error occurs, skip the line that caused the error and continue with the next line of code.
    On Error Resume Next
    Set NewStyle = ActiveDocument.Styles(C_S_ParNormalBelow)
    '  Turns off any active error handler and restores default behavior
    On Error GoTo 0

    ' If the style doesn't exist, try to create it
    If NewStyle Is Nothing Then
        On Error GoTo StyleError
        Set NewStyle = ActiveDocument.Styles.Add(Name:=C_S_ParNormalBelow, Type:=wdStyleTypeParagraph)
        On Error GoTo 0
    End If

    If Not NewStyle Is Nothing Then
        With NewStyle
            .BaseStyle = C_S_ParNormal
            .NextParagraphStyle = C_S_ParNormalBelow
            .AutomaticallyUpdate = False
            .QuickStyle = False
            .LanguageId = wdEnglishUS
            
            ' Font formatting
            With .font
                .Name = C_FT_Body
                .Size = C_BaseFontSize
                .Bold = False
                .Italic = False
                .color = wdColorAutomatic
            End With
            
            ' Paragraph formatting
            With .ParagraphFormat
                .Alignment = wdAlignParagraphLeft
                .LeftIndent = CentimetersToPoints(0)
                .RightIndent = CentimetersToPoints(0)
                .FirstLineIndent = 0
                .SpaceBefore = 0
                .SpaceAfter = 15
                .LineSpacing = NewStyle.font.Size   ' order matters: specify at first LineSpacing, next LineSpacingRule
                .LineSpacingRule = wdLineSpaceExactly
                .WidowControl = True
                .KeepWithNext = False
                .KeepTogether = False
                .PageBreakBefore = False
                .OutlineLevel = wdOutlineLevelBodyText
            End With
        End With

'       === Add keyboard shortcut for the style ===
       On Error GoTo ShortcutError
       CustomizationContext = ActiveDocument
       KeyBindings.Add _
            KeyCategory:=wdKeyCategoryStyle, _
            Command:=C_S_ParNormalBelow, _
            KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyN), _
            KeyCode2:=wdKeyB
       On Error GoTo 0

    CreateStyle_ParNormalBelowMs = True
    End If
    Exit Function
                
' === Error handler ===
StyleError:
    MsgBox _
        Prompt:="Error: Unable to create style '" & C_S_ParNormalBelow & "'." & vbNewLine & _
            "Error number: " & Err.Number & vbNewLine & _
            "Description: " & Err.Description, _
        Buttons:=vbCritical, _
        Title:=MsgBoxTitle
    CreateStyle_ParNormalBelowMs = False
    Exit Function
        
ShortcutError:
    MsgBox _
        Prompt:="Error adding shortcut for style '" & C_S_ParNormalBelow & "'." & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, _
        Buttons:=vbCritical, _
        Title:=MsgBoxTitle
    Resume Next

End Function

' 2025-10-07 by ms
Private Function CreateStyle_ParNormalAboveMs() As Boolean
    Dim NewStyle As style
    
    Dim FileName As String:     FileName = C_F_Macros
    Dim ModuleName As String:   ModuleName = C_M_Styles
    Dim MacroName As String:    MacroName = "CreateStyle_ParNormalAboveMs"
    Dim MsgBoxTitle As String:  MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    ' Check if the style already exists in the document
    ' If an error occurs, skip the line that caused the error and continue with the next line of code.
    On Error Resume Next
    Set NewStyle = ActiveDocument.Styles(C_S_ParNormalAbove)
    '  Turns off any active error handler and restores default behavior
    On Error GoTo 0

    ' If the style doesn't exist, try to create it
    If NewStyle Is Nothing Then
        On Error GoTo StyleError
        Set NewStyle = ActiveDocument.Styles.Add(Name:=C_S_ParNormalAbove, Type:=wdStyleTypeParagraph)
        On Error GoTo 0
    End If

    If Not NewStyle Is Nothing Then
        With NewStyle
            .BaseStyle = C_S_ParNormal
            .NextParagraphStyle = C_S_ParNormal
            .AutomaticallyUpdate = False
            .QuickStyle = False
            .LanguageId = wdEnglishUS
            
            ' Font formatting
            With .font
                .Name = C_FT_Body
                .Size = C_BaseFontSize
                .Bold = False
                .Italic = False
                .color = wdColorAutomatic
            End With
            
            ' Paragraph formatting
            With .ParagraphFormat
                .Alignment = wdAlignParagraphLeft
                .LeftIndent = CentimetersToPoints(0)
                .RightIndent = CentimetersToPoints(0)
                .FirstLineIndent = 0
                .SpaceBefore = 8
                .SpaceAfter = 6
                .LineSpacing = NewStyle.font.Size   ' order matters: specify at first LineSpacing, next LineSpacingRule
                .LineSpacingRule = wdLineSpaceExactly
                .WidowControl = True
                .KeepWithNext = False
                .KeepTogether = False
                .PageBreakBefore = False
                .OutlineLevel = wdOutlineLevelBodyText
            End With
        End With

'       === Add keyboard shortcut for the style ===
       On Error GoTo ShortcutError
       CustomizationContext = ActiveDocument
       KeyBindings.Add _
            KeyCategory:=wdKeyCategoryStyle, _
            Command:=C_S_ParNormalAbove, _
            KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyN), _
            KeyCode2:=wdKeyA
       On Error GoTo 0

    CreateStyle_ParNormalAboveMs = True
    End If
    Exit Function
                
' === Error handler ===
StyleError:
    MsgBox _
        Prompt:="Error: Unable to create style '" & C_S_ParNormalAbove & "'." & vbNewLine & _
            "Error number: " & Err.Number & vbNewLine & _
            "Description: " & Err.Description, _
        Buttons:=vbCritical, _
        Title:=MsgBoxTitle
    CreateStyle_ParNormalAboveMs = False
    Exit Function
        
ShortcutError:
    MsgBox _
        Prompt:="Error adding shortcut for style '" & C_S_ParNormalAbove & "'." & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, _
        Buttons:=vbCritical, _
        Title:=MsgBoxTitle
    Resume Next

End Function

' 2025-12-08 by ms
Private Function CreateStyle_ParNormalAboveBelowMs() As Boolean
    Dim NewStyle As style
    
    Dim FileName As String:     FileName = C_F_Macros
    Dim ModuleName As String:   ModuleName = C_M_Styles
    Dim MacroName As String:    MacroName = "CreateStyle_ParNormalAboveBelowMs"
    Dim MsgBoxTitle As String:  MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    ' Check if the style already exists in the document
    ' If an error occurs, skip the line that caused the error and continue with the next line of code.
    On Error Resume Next
    Set NewStyle = ActiveDocument.Styles(C_S_ParNormalAB)
    '  Turns off any active error handler and restores default behavior
    On Error GoTo 0

    ' If the style doesn't exist, try to create it
    If NewStyle Is Nothing Then
        On Error GoTo StyleError
        Set NewStyle = ActiveDocument.Styles.Add(Name:=C_S_ParNormalAB, Type:=wdStyleTypeParagraph)
        On Error GoTo 0
    End If

    If Not NewStyle Is Nothing Then
        With NewStyle
            .BaseStyle = C_S_ParNormal
            .NextParagraphStyle = C_S_ParNormal
            .AutomaticallyUpdate = False
            .QuickStyle = False
            .LanguageId = wdEnglishUS
            
            ' Font formatting
            With .font
                .Name = C_FT_Body
                .Size = C_BaseFontSize
                .Bold = False
                .Italic = False
                .color = wdColorAutomatic
            End With
            
            ' Paragraph formatting
            With .ParagraphFormat
                .Alignment = wdAlignParagraphLeft
                .LeftIndent = CentimetersToPoints(0)
                .RightIndent = CentimetersToPoints(0)
                .FirstLineIndent = 0
                .SpaceBefore = 8
                .SpaceAfter = 8
                .LineSpacing = NewStyle.font.Size   ' order matters: specify at first LineSpacing, next LineSpacingRule
                .LineSpacingRule = wdLineSpaceExactly
                .WidowControl = True
                .KeepWithNext = False
                .KeepTogether = False
                .PageBreakBefore = False
                .OutlineLevel = wdOutlineLevelBodyText
            End With
        End With

'       === Add keyboard shortcut for the style ===
       On Error GoTo ShortcutError
       CustomizationContext = ActiveDocument
       KeyBindings.Add _
            KeyCategory:=wdKeyCategoryStyle, _
            Command:=C_S_ParNormalAB, _
            KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyN), _
            KeyCode2:=wdKeyC
       On Error GoTo 0

    CreateStyle_ParNormalAboveBelowMs = True
    End If
    Exit Function
                
' === Error handler ===
StyleError:
    MsgBox _
        Prompt:="Error: Unable to create style '" & C_S_ParNormalAB & "'." & vbNewLine & _
            "Error number: " & Err.Number & vbNewLine & _
            "Description: " & Err.Description, _
        Buttons:=vbCritical, _
        Title:=MsgBoxTitle
    CreateStyle_ParNormalAboveBelowMs = False
    Exit Function
        
ShortcutError:
    MsgBox _
        Prompt:="Error adding shortcut for style '" & C_S_ParNormalAB & "'." & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, _
        Buttons:=vbCritical, _
        Title:=MsgBoxTitle
    Resume Next

End Function

' 2025-11-19 by ms
Private Function CreateStyle_CharBoldMs() As Boolean
    Dim NewStyle As style
    
    Dim FileName As String:       FileName = C_F_Macros
    Dim ModuleName As String:     ModuleName = C_M_Styles
    Dim MacroName As String:      MacroName = "CreateStyle_CharBoldMs"
    Dim MsgBoxTitle As String:    MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName

    ' Check if the style already exists
    On Error Resume Next
    Set NewStyle = ActiveDocument.Styles(C_S_Bold)
    On Error GoTo 0

    ' If the style doesn't exist, create it
    If NewStyle Is Nothing Then
        On Error GoTo StyleError
        Set NewStyle = ActiveDocument.Styles.Add(Name:=C_S_Bold, Type:=wdStyleTypeCharacter)
        On Error GoTo 0
    End If

    ' Apply character formatting
    If Not NewStyle Is Nothing Then
        NewStyle.BaseStyle = ActiveDocument.Styles(wdStyleDefaultParagraphFont)
        With NewStyle.font
            .Bold = True              ' Bold text
        End With
        CreateStyle_CharBoldMs = True
    End If
    Exit Function

StyleError:
    MsgBox _
        Prompt:="Error creating character style '" & C_S_Bold & "'." & vbNewLine & _
           "Error Number: " & Err.Number & vbNewLine & _
           "Description: " & Err.Description, _
        Buttons:=vbCritical, _
        Title:=MsgBoxTitle
    CreateStyle_CharBoldMs = False
End Function

' 2025-11-19 by ms
Private Function CreateStyle_CharCrossoutMs() As Boolean
    Dim NewStyle As style
    
    Dim FileName As String:       FileName = C_F_Macros
    Dim ModuleName As String:     ModuleName = C_M_Styles
    Dim MacroName As String:      MacroName = "CreateStyle_CharCrossoutMs"
    Dim MsgBoxTitle As String:    MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName

    ' Check if the style already exists
    On Error Resume Next
    Set NewStyle = ActiveDocument.Styles(C_S_CharCrossout)
    On Error GoTo 0

    ' If the style doesn't exist, create it
    If NewStyle Is Nothing Then
        On Error GoTo StyleError
        Set NewStyle = ActiveDocument.Styles.Add(Name:=C_S_CharCrossout, Type:=wdStyleTypeCharacter)
        On Error GoTo 0
    End If

    ' Apply character formatting
    If Not NewStyle Is Nothing Then
        NewStyle.BaseStyle = ActiveDocument.Styles(wdStyleDefaultParagraphFont)
        With NewStyle.font
            .Strikethrough = True
        End With
                
        CreateStyle_CharCrossoutMs = True
    End If
    Exit Function

StyleError:
    MsgBox _
        Prompt:="Error creating character style '" & C_S_CharCrossout & "'." & vbNewLine & _
           "Error Number: " & Err.Number & vbNewLine & _
           "Description: " & Err.Description, _
        Buttons:=vbCritical, _
        Title:=MsgBoxTitle
    CreateStyle_CharCrossoutMs = False
End Function

' 2025-11-19 by ms
' 2026-01-14 by ms and AI
Private Function CreateStyle_CharDefaultMs() As Boolean
    Dim NewStyle As style
    
    Dim FileName As String:       FileName = C_F_Macros
    Dim ModuleName As String:     ModuleName = C_M_Styles
    Dim MacroName As String:      MacroName = "CreateStyle_CharDefaultMs"
    Dim MsgBoxTitle As String:    MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName

    ' Check if the style already exists
    On Error Resume Next
    Set NewStyle = ActiveDocument.Styles(C_S_CharDefault)
    On Error GoTo 0

    ' If the style doesn't exist, create it
    If NewStyle Is Nothing Then
        On Error GoTo StyleError
        Set NewStyle = ActiveDocument.Styles.Add(Name:=C_S_CharDefault, Type:=wdStyleTypeCharacter)
        On Error GoTo 0
    End If

    ' Apply character formatting
    If Not NewStyle Is Nothing Then
        NewStyle.BaseStyle = ActiveDocument.Styles(wdStyleDefaultParagraphFont)
        CreateStyle_CharDefaultMs = True
    End If
    Exit Function

StyleError:
    MsgBox _
        Prompt:="Error creating character style '" & C_S_CharDefault & "'." & vbNewLine & _
           "Error Number: " & Err.Number & vbNewLine & _
           "Description: " & Err.Description, _
        Buttons:=vbCritical, _
        Title:=MsgBoxTitle
    CreateStyle_CharDefaultMs = False
End Function

' 2025-11-19 by ms
'Unfortunately it is not possible in VBA style definition to apply shading color (#F6C0C0) light red. Therefore another macro is involved which is run by shortcut Shift + Ctrl + H to toggle shading color of the selected content.
Private Function CreateStyle_CharHiddenMs() As Boolean
    Dim NewStyle As style
        
    Dim FileName As String:       FileName = C_F_Macros
    Dim ModuleName As String:     ModuleName = C_M_Styles
    Dim MacroName As String:      MacroName = "CreateStyle_CharHiddenMs"
    Dim MsgBoxTitle As String:    MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName

    ' Check if the style already exists
    On Error Resume Next
    Set NewStyle = ActiveDocument.Styles(C_S_CharHidden)
    On Error GoTo 0

    ' If the style doesn't exist, create it
    If NewStyle Is Nothing Then
        On Error GoTo StyleError
        Set NewStyle = ActiveDocument.Styles.Add(Name:=C_S_CharHidden, Type:=wdStyleTypeCharacter)
        On Error GoTo 0
    End If

    ' Apply character formatting
    If Not NewStyle Is Nothing Then
        With NewStyle.font
            .Hidden = True
        End With
        
        ' Remove all borders
        Dim i As Integer
        For i = wdBorderTop To wdBorderRight
            NewStyle.borders(i).LineStyle = wdLineStyleNone
        Next i
                
' Unfortunately it is not possible in VBA style definition to apply shading color (#F6C0C0) light red. Therefore another macro is involved which is run by shortcut Shift + Ctrl + H to toggle shading color of the selected content.
'        With NewStyle.shading
'            .BackgroundPatternColor = RGB(246, 192, 192) ' #F6C0C0 light red tu jestem
'            .Texture = wdTextureNone
'        End With
                
        CreateStyle_CharHiddenMs = True
    End If
    Exit Function

StyleError:
    MsgBox _
        Prompt:="Error creating character style '" & C_S_CharHidden & "'." & vbNewLine & _
           "Error Number: " & Err.Number & vbNewLine & _
           "Description: " & Err.Description, _
        Buttons:=vbCritical, _
        Title:=MsgBoxTitle
    CreateStyle_CharHiddenMs = False
End Function

' 2025-11-19 by ms
Private Function CreateStyle_CharItalicMs() As Boolean
    Dim NewStyle As style
    
    Dim FileName As String:       FileName = C_F_Macros
    Dim ModuleName As String:     ModuleName = C_M_Styles
    Dim MacroName As String:      MacroName = "CreateStyle_CharItalicMs"
    Dim MsgBoxTitle As String:    MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName

    ' Check if the style already exists
    On Error Resume Next
    Set NewStyle = ActiveDocument.Styles(C_S_Italic)
    On Error GoTo 0

    ' If the style doesn't exist, create it
    If NewStyle Is Nothing Then
        On Error GoTo StyleError
        Set NewStyle = ActiveDocument.Styles.Add(Name:=C_S_Italic, Type:=wdStyleTypeCharacter)
        On Error GoTo 0
    End If

    ' Apply character formatting
    If Not NewStyle Is Nothing Then
        ' Make sure the style inherits and only overrides underline
        NewStyle.BaseStyle = ActiveDocument.Styles(wdStyleDefaultParagraphFont)
    
        With NewStyle.font
            .Italic = True
        End With
        CreateStyle_CharItalicMs = True
    End If
    Exit Function

StyleError:
    MsgBox _
        Prompt:="Error creating character style '" & C_S_Italic & "'." & vbNewLine & _
           "Error Number: " & Err.Number & vbNewLine & _
           "Description: " & Err.Description, _
        Buttons:=vbCritical, _
        Title:=MsgBoxTitle
    CreateStyle_CharItalicMs = False
End Function

' 2025-11-19 by ms
Private Function CreateStyle_CharLegalNoteMs() As Boolean
    Dim NewStyle As style
    
    Dim FileName As String:       FileName = C_F_Macros
    Dim ModuleName As String:     ModuleName = C_M_Styles
    Dim MacroName As String:      MacroName = "CreateStyle_CharLegalNoteMs"
    Dim MsgBoxTitle As String:    MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName

    ' Check if the style already exists
    On Error Resume Next
    Set NewStyle = ActiveDocument.Styles(C_S_CharLegalNote)
    On Error GoTo 0

    ' If the style doesn't exist, create it
    If NewStyle Is Nothing Then
        On Error GoTo StyleError
        Set NewStyle = ActiveDocument.Styles.Add(Name:=C_S_CharLegalNote, Type:=wdStyleTypeCharacter)
        On Error GoTo 0
    End If

    ' Apply character formatting
    If Not NewStyle Is Nothing Then
        With NewStyle.font
            .Name = C_FT_Body        ' Font name, optional for Character style type definition
            .Size = C_BaseFontSize               ' Font size, optional for Character style type definition
            '.Bold = False             ' Bold text
            '.Strikethrough = True
            '.Italic = True
            '.Hidden = True
            .color = GetThemeColor(ThemeColorIndex:=wdThemeColorText2, TintAndShade:=0)   ' in module Template
        End With
        CreateStyle_CharLegalNoteMs = True
    End If
    Exit Function

StyleError:
    MsgBox _
        Prompt:="Error creating character style '" & C_S_CharLegalNote & "'." & vbNewLine & _
           "Error Number: " & Err.Number & vbNewLine & _
           "Description: " & Err.Description, _
        Buttons:=vbCritical, _
        Title:=MsgBoxTitle
    CreateStyle_CharLegalNoteMs = False
End Function

' 2025-11-19 by ms
Private Function CreateStyle_CharSourceCodeMs() As Boolean
    Dim NewStyle As style
    
    Dim FileName As String:      FileName = C_F_Macros
    Dim ModuleName As String:    ModuleName = C_M_Styles
    Dim MacroName As String:     MacroName = "CreateStyle_CharSourceCodeMs"
    Dim MsgBoxTitle As String:   MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName

    ' Check if the style already exists
    On Error Resume Next
    Set NewStyle = ActiveDocument.Styles(C_S_CharSourceCode)
    On Error GoTo 0

    ' If the style doesn't exist, create it
    If NewStyle Is Nothing Then
        On Error GoTo StyleError
        Set NewStyle = ActiveDocument.Styles.Add(Name:=C_S_CharSourceCode, Type:=wdStyleTypeCharacter)
        On Error GoTo 0
    End If

    ' Apply character formatting
    If Not NewStyle Is Nothing Then
        With NewStyle.font
            .Name = C_FT_AntiHomoglyph  ' Font name, optional for Character style type definition
            .Size = C_BaseFontSize                  ' Font size, optional for Character style type definition
            '.Bold = False              ' Bold text
            '.Strikethrough = True
            '.Italic = True
            '.Hidden = True
            .color = wdColorAutomatic
        End With
                       
        CreateStyle_CharSourceCodeMs = True
    End If
    Exit Function

StyleError:
    MsgBox _
        Prompt:="Error creating character style '" & C_S_CharSourceCode & "'." & vbNewLine & _
           "Error Number: " & Err.Number & vbNewLine & _
           "Description: " & Err.Description, _
        Buttons:=vbCritical, _
        Title:=MsgBoxTitle
    CreateStyle_CharSourceCodeMs = False
End Function

' 2025-11-19 by ms
Private Function CreateStyle_CharUnderlineMs() As Boolean
    Dim NewStyle As style
    
    Dim FileName As String:     FileName = C_F_Macros
    Dim ModuleName As String:   ModuleName = C_M_Styles
    Dim MacroName As String:    MacroName = "CreateStyle_CharUnderlineMs"
    Dim MsgBoxTitle As String:  MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName

    ' Check if the style already exists
    On Error Resume Next
    Set NewStyle = ActiveDocument.Styles(C_S_Underline)
    On Error GoTo 0

    ' If the style doesn't exist, create it
    If NewStyle Is Nothing Then
        On Error GoTo StyleError
        Set NewStyle = ActiveDocument.Styles.Add(Name:=C_S_Underline, Type:=wdStyleTypeCharacter)
        On Error GoTo 0
    End If

    ' Apply character formatting
    If Not NewStyle Is Nothing Then
        ' Make sure the style inherits and only overrides underline
        NewStyle.BaseStyle = ActiveDocument.Styles(wdStyleDefaultParagraphFont)
        
        With NewStyle.font
            .Underline = wdUnderlineSingle
        End With
                                
        CreateStyle_CharUnderlineMs = True
    End If
    Exit Function

StyleError:
    MsgBox _
        Prompt:="Error creating character style '" & C_S_Underline & "'." & vbNewLine & _
           "Error Number: " & Err.Number & vbNewLine & _
           "Description: " & Err.Description, _
        Buttons:=vbCritical, _
        Title:=MsgBoxTitle
    CreateStyle_CharUnderlineMs = False
End Function

' 2025-11-26 by ms
Private Function CreateStyle_TabTableMs() As Boolean
    Dim NewStyle As style
    
    Dim FileName As String:     FileName = C_F_Macros
    Dim ModuleName As String:   ModuleName = C_M_Styles
    Dim MacroName As String:    MacroName = "CreateStyle_TabTableMs"
    Dim MsgBoxTitle As String:  MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    ' Create or get the style
    On Error Resume Next
    Set NewStyle = ActiveDocument.Styles(C_S_TabTable)
    On Error GoTo 0
    
    If NewStyle Is Nothing Then
        Set NewStyle = ActiveDocument.Styles.Add(Name:=C_S_TabTable, Type:=wdStyleTypeTable)
    End If

    ' --- WHOLE TABLE SETTINGS ---
    With NewStyle.Table
        ' Table-level options (padding, spacing)
        .LeftPadding = 3 '  pt = points
        .RightPadding = 3
        .TopPadding = 0
        .BottomPadding = 0

        ' Borders: box with single line 1 pt
        .borders.Enable = False
'        .borders.OutsideLineStyle = wdLineStyleSingle
'        .borders.OutsideLineWidth = wdLineWidth100pt
'        .borders.OutsideColor = GetThemeColor(ThemeColorIndex:=wdThemeColorAccent1, TintAndShade:=0)  ' in module Theme
'        .borders.InsideLineStyle = wdLineStyleSingle
'        .borders.InsideLineWidth = wdLineWidth100pt
'        .borders.InsideColor = GetThemeColor(ThemeColorIndex:=wdThemeColorAccent1, TintAndShade:=0)  ' in module Theme

        With .Condition(wdFirstRow)
            .shading.BackgroundPatternColor = GetThemeColor(ThemeColorIndex:=wdThemeColorAccent2, TintAndShade:=0)  ' in module Theme
            .borders(wdBorderVertical).LineStyle = wdLineStyleSingle
            .borders(wdBorderVertical).LineWidth = wdLineWidth100pt
            .borders(wdBorderVertical).color = GetThemeColor(ThemeColorIndex:=wdThemeColorAccent1, TintAndShade:=0)  ' in module Theme
            .borders(wdBorderLeft).LineStyle = wdLineStyleSingle
            .borders(wdBorderLeft).LineWidth = wdLineWidth100pt
            .borders(wdBorderLeft).color = GetThemeColor(ThemeColorIndex:=wdThemeColorAccent1, TintAndShade:=0)  ' in module Theme
            .borders(wdBorderRight).LineStyle = wdLineStyleSingle
            .borders(wdBorderRight).LineWidth = wdLineWidth100pt
            .borders(wdBorderRight).color = GetThemeColor(ThemeColorIndex:=wdThemeColorAccent1, TintAndShade:=0)  ' in module Theme
            .borders(wdBorderTop).LineStyle = wdLineStyleSingle
            .borders(wdBorderTop).LineWidth = wdLineWidth100pt
            .borders(wdBorderTop).color = wdColorAutomatic
            .borders(wdBorderBottom).LineStyle = wdLineStyleSingle
            .borders(wdBorderBottom).LineWidth = wdLineWidth100pt
            .borders(wdBorderBottom).color = wdColorAutomatic
        End With

        With .Condition(wdEvenRowBanding)
            .shading.BackgroundPatternColor = GetThemeColor(ThemeColorIndex:=wdThemeColorAccent3, TintAndShade:=0)  ' in module Theme
            .borders(wdBorderVertical).LineStyle = wdLineStyleSingle
            .borders(wdBorderVertical).LineWidth = wdLineWidth100pt
            .borders(wdBorderVertical).color = GetThemeColor(ThemeColorIndex:=wdThemeColorAccent1, TintAndShade:=0)  ' in module Theme
            .borders(wdBorderLeft).LineStyle = wdLineStyleSingle
            .borders(wdBorderLeft).LineWidth = wdLineWidth100pt
            .borders(wdBorderLeft).color = GetThemeColor(ThemeColorIndex:=wdThemeColorAccent1, TintAndShade:=0)  ' in module Theme
            .borders(wdBorderRight).LineStyle = wdLineStyleSingle
            .borders(wdBorderRight).LineWidth = wdLineWidth100pt
            .borders(wdBorderRight).color = GetThemeColor(ThemeColorIndex:=wdThemeColorAccent1, TintAndShade:=0)  ' in module Theme
            .borders(wdBorderBottom).LineStyle = wdLineStyleSingle
            .borders(wdBorderBottom).LineWidth = wdLineWidth100pt
            .borders(wdBorderBottom).color = wdColorAutomatic
        End With
        
        With .Condition(wdOddRowBanding)
            .shading.BackgroundPatternColor = GetThemeColor(ThemeColorIndex:=wdThemeColorMainLight1, TintAndShade:=0)  ' in module Theme
            .borders(wdBorderVertical).LineStyle = wdLineStyleSingle
            .borders(wdBorderVertical).LineWidth = wdLineWidth100pt
            .borders(wdBorderVertical).color = GetThemeColor(ThemeColorIndex:=wdThemeColorAccent1, TintAndShade:=0)  ' in module Theme
            .borders(wdBorderLeft).LineStyle = wdLineStyleSingle
            .borders(wdBorderLeft).LineWidth = wdLineWidth100pt
            .borders(wdBorderLeft).color = GetThemeColor(ThemeColorIndex:=wdThemeColorAccent1, TintAndShade:=0)  ' in module Theme
            .borders(wdBorderRight).LineStyle = wdLineStyleSingle
            .borders(wdBorderRight).LineWidth = wdLineWidth100pt
            .borders(wdBorderRight).color = GetThemeColor(ThemeColorIndex:=wdThemeColorAccent1, TintAndShade:=0)  ' in module Theme
            .borders(wdBorderBottom).LineStyle = wdLineStyleSingle
            .borders(wdBorderBottom).LineWidth = wdLineWidth100pt
            .borders(wdBorderBottom).color = wdColorAutomatic
        End With
                
        .AllowPageBreaks = True    ' Applies to table styles (or sometimes entire tables when using styles). True: The table can split across pages.
        .AllowBreakAcrossPage = True   ' Applies to individual rows. True: Word allows the content of a row to break across two pages. False: Row stays intact on one page.
        .RowStripe = 1 ' band every 1 row (used by Odd/Even Row style elements)
        .ColumnStripe = 0 ' no column banding unless set in elements
    End With

    CreateStyle_TabTableMs = True
End Function

' 2025-11-30 by ms
Private Function CreateStyle_TabTableNoGridMs() As Boolean
    Dim NewStyle As style
    
    Dim FileName As String:     FileName = C_F_Macros
    Dim ModuleName As String:   ModuleName = C_M_Styles
    Dim MacroName As String:    MacroName = "CreateStyle_TabTableNoGridMs"
    Dim MsgBoxTitle As String:  MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    ' Create or get the style
    On Error Resume Next
    Set NewStyle = ActiveDocument.Styles(C_S_TabNoGrid)
    On Error GoTo 0
    
    If NewStyle Is Nothing Then
        Set NewStyle = ActiveDocument.Styles.Add(Name:=C_S_TabNoGrid, Type:=wdStyleTypeTable)
    End If

    ' --- WHOLE TABLE SETTINGS ---
    With NewStyle.Table
        ' Table-level options (padding, spacing)
        .LeftPadding = 3 '  pt = points
        .RightPadding = 3
        .TopPadding = 2
        .BottomPadding = 2

        ' Borders: box with single line 1 pt
        .borders.Enable = False

        .AllowPageBreaks = True    ' Applies to table styles (or sometimes entire tables when using styles). True: The table can split across pages.
        .AllowBreakAcrossPage = False   ' Applies to individual rows. False: Row stays intact on one page.
        .RowStripe = 0 ' band every 1 row (used by Odd/Even Row style elements)
        .ColumnStripe = 0 ' no column banding unless set in elements
    End With

    CreateStyle_TabTableNoGridMs = True
End Function

' 2025-11-30 by ms
Private Function CreateStyle_TabTableNoPaddingMs() As Boolean
    Dim NewStyle As style
    
    Dim FileName As String:     FileName = C_F_Macros
    Dim ModuleName As String:   ModuleName = C_M_Styles
    Dim MacroName As String:    MacroName = "CreateStyle_TabTableNoPaddingMs"
    Dim MsgBoxTitle As String:  MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    ' Create or get the style
    On Error Resume Next
    Set NewStyle = ActiveDocument.Styles(C_S_TabNoPadding)
    On Error GoTo 0
    
    If NewStyle Is Nothing Then
        Set NewStyle = ActiveDocument.Styles.Add(Name:=C_S_TabNoPadding, Type:=wdStyleTypeTable)
    End If

    ' --- WHOLE TABLE SETTINGS ---
    With NewStyle.Table
        ' Table-level options (padding, spacing)
        .LeftPadding = 0 '  pt = points
        .RightPadding = 0
        .TopPadding = 0
        .BottomPadding = 0

        ' Borders: box with single line 1 pt
        .borders.Enable = False

        .AllowPageBreaks = True    ' Applies to table styles (or sometimes entire tables when using styles). True: The table can split across pages.
        .AllowBreakAcrossPage = False   ' Applies to individual rows. False: Row stays intact on one page.
        .RowStripe = 0 ' band every 1 row (used by Odd/Even Row style elements)
        .ColumnStripe = 0 ' no column banding unless set in elements
    End With

    CreateStyle_TabTableNoPaddingMs = True
End Function

' The purpose of this Sub is to have a kind of text written report showing at least the most
' important features of the customized styles (specific suffix).
' Actually it can't be done properly. In other words it isn't possible to export reliably all
' the settings of styles from Microsoft Word to text files.
' There are numerous resons for that statement to be true at 2025.
' Work in progress.
' 2025-02-16 by ms
Sub ListCustomStylesToTxt()
    Dim style As style
    Dim FilePath As String
    Dim filenum As Integer
    Dim TemplateName As String
    Dim i As Byte

    Dim FileName As String:     FileName = C_F_Macros
    Dim ModuleName As String:   ModuleName = C_M_Template
    Dim MacroName As String:    MacroName = "ListCustomStylesToTxt"
    Dim MsgBoxTitle As String:  MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    ' Set the CustomizationContext to the currently active document
    CustomizationContext = ActiveDocument
    
    ' Set the file path to the default file location
    FilePath = Options.DefaultFilePath(wdDocumentsPath) & "\" & MacroName & ".txt"
    
    ' Open the file for writing
    filenum = FreeFile
    Open FilePath For Output As #filenum
    
    ' Add the header information
    TemplateName = ActiveDocument.AttachedTemplate.Name
    Print #filenum, "Template name: " & TemplateName
    Print #filenum, "This file contains styles definitions."
    Print #filenum, "It was made by the macro: " & MsgBoxTitle
    Print #filenum, "----------------------------------------"
    
    ' Loop through all styles in the active document
    For Each style In ActiveDocument.Styles
        ' Check if the style name ends with " ms"
        If Right(style.NameLocal, 3) = C_StyleSuffix Then
            ' Write the style name and type to the file
            Call PrintStyleBasicProperties(filenum, style)
            
            ' Char type styles selected features:
            If style.Type = wdStyleTypeCharacter Then
                Print #filenum, "Language ID: " & style.LanguageId
                
                Call PrintFontProperties(filenum, style)
                ' Shortcut key: in KeyBindings
                ' Text effects: future, if necessary
            End If
            
            ' Paragraph type styles selected features:
            If style.Type = wdStyleTypeParagraph Then
                Print #filenum, "Automatically update: " & style.AutomaticallyUpdate
                Print #filenum, "Language ID: " & style.LanguageId

                ' Font:
                Call PrintFontProperties(filenum, style)
                
                ' Paragraph:
                Call PrintParagraphProperties(filenum, style)
            End If
            
            ' Linked type styles selected features:
            If style.Type = wdStyleTypeLinked Then
                ' Surprisingly no style here in my template.
            End If
            
            If style.Type = wdStyleTypeParagraphOnly Then
                ' Surprisingly no style here in my template.
            End If
            
            If style.Type = wdStyleTypeTable Then
                Print #filenum, "Style base name local: " & style.BaseStyle.NameLocal
                Print #filenum, "Language ID: " & style.LanguageId
                Print #filenum, "Style description: " & style.Description
                Print #filenum, "Is this a quick style: " & style.QuickStyle
                Print #filenum, "Number of borders: " & style.Table.borders.count
                Print #filenum, "AllowBreakAcrossPage: " & GetBooleanName(style.Table.AllowBreakAcrossPage)
                Print #filenum, "AllowPageBreaks: " & style.Table.AllowPageBreaks
                
                Call InitializeColourDetails(MyVar:=ExampleResult)
                ExampleResult.ColourValue = style.Table.Condition(wdFirstRow).shading.BackgroundPatternColor
                ExampleResult = QueryColour(ExampleResult.ColourValue)
                Print #filenum, "Heading row fill (shading) [Long]: " & ExampleResult.ColourValue
                Print #filenum, "Heading row fill (shading) ColourType: " & GetColourTypeName(ExampleResult.ColourType)
                If ExampleResult.ColourType <> ColourTypeAutomatic Then
                    Print #filenum, "Heading row fill (shading) ThemeColorIndex: " & ExampleResult.ThemeColorIndex
                    Print #filenum, "Heading row fill (shading) ThemeColorText: " & ExampleResult.ThemeColorText
                    Print #filenum, "Heading row fill (shading) TintAndShade: " & ExampleResult.TintAndShade
                    Print #filenum, "Heading row fill (shading) TintAndShadeText: " & ExampleResult.TintAndShadeText
                    Print #filenum, "Heading row fill (shading) RGB: " & ExampleResult.RGB
                    Print #filenum, "Heading row fill (shading) RGB_Hex: " & ExampleResult.RGB_Hex
                    Print #filenum, "Heading row fill (shading) Red: " & ExampleResult.Red
                    Print #filenum, "Heading row fill (shading) Green: " & ExampleResult.Green
                    Print #filenum, "Heading row fill (shading) Blue: " & ExampleResult.Blue
                End If
                
                Call InitializeColourDetails(MyVar:=ExampleResult)
                ExampleResult.ColourValue = style.Table.Condition(wdOddRowBanding).shading.BackgroundPatternColor
                ExampleResult = QueryColour(ExampleResult.ColourValue)
                Print #filenum, "Odd row banding fill (shading) [Long]: " & ExampleResult.ColourValue
                Print #filenum, "Odd row banding fill (shading) ColourType: " & GetColourTypeName(ExampleResult.ColourType)
                If ExampleResult.ColourType <> ColourTypeAutomatic Then
                    Print #filenum, "Odd row banding fill (shading) ThemeColorIndex: " & ExampleResult.ThemeColorIndex
                    Print #filenum, "Odd row banding fill (shading) ThemeColorText: " & ExampleResult.ThemeColorText
                    Print #filenum, "Odd row banding fill (shading) TintAndShade: " & ExampleResult.TintAndShade
                    Print #filenum, "Odd row banding fill (shading) TintAndShadeText: " & ExampleResult.TintAndShadeText
                    Print #filenum, "Odd row banding fill (shading) RGB: " & ExampleResult.RGB
                    Print #filenum, "Odd row banding fill (shading) RGB_Hex: " & ExampleResult.RGB_Hex
                    Print #filenum, "Odd row banding fill (shading) Red: " & ExampleResult.Red
                    Print #filenum, "Odd row banding fill (shading) Green: " & ExampleResult.Green
                    Print #filenum, "Odd row banding fill (shading) Blue: " & ExampleResult.Blue
                End If
                
                Call InitializeColourDetails(MyVar:=ExampleResult)
                ExampleResult.ColourValue = style.Table.Condition(wdEvenRowBanding).shading.BackgroundPatternColor
                ExampleResult = QueryColour(ExampleResult.ColourValue)
                Print #filenum, "Even row banding fill (shading) [Long]: " & ExampleResult.ColourValue
                Print #filenum, "Even row banding fill (shading) ColourType: " & GetColourTypeName(ExampleResult.ColourType)
                If ExampleResult.ColourType <> ColourTypeAutomatic Then
                    Print #filenum, "Even row banding fill (shading) ThemeColorIndex: " & ExampleResult.ThemeColorIndex
                    Print #filenum, "Even row banding fill (shading) ThemeColorText: " & ExampleResult.ThemeColorText
                    Print #filenum, "Even row banding fill (shading) TintAndShade: " & ExampleResult.TintAndShade
                    Print #filenum, "Even row banding fill (shading) TintAndShadeText: " & ExampleResult.TintAndShadeText
                    Print #filenum, "Even row banding fill (shading) RGB: " & ExampleResult.RGB
                    Print #filenum, "Even row banding fill (shading) RGB_Hex: " & ExampleResult.RGB_Hex
                    Print #filenum, "Even row banding fill (shading) Red: " & ExampleResult.Red
                    Print #filenum, "Even row banding fill (shading) Green: " & ExampleResult.Green
                    Print #filenum, "Even row banding fill (shading) Blue: " & ExampleResult.Blue
                End If
                                
                Call InitializeColourDetails(MyVar:=ExampleResult)
                ExampleResult.ColourValue = style.Table.borders(wdBorderBottom).color
                ExampleResult = QueryColour(ExampleResult.ColourValue)
                Print #filenum, "Border bottom color [Long]: " & ExampleResult.ColourValue
                Print #filenum, "Border bottom color ColourType: " & GetColourTypeName(ExampleResult.ColourType)
                If ExampleResult.ColourType <> ColourTypeAutomatic Then
                    Print #filenum, "Border bottom color ThemeColorIndex: " & ExampleResult.ThemeColorIndex
                    Print #filenum, "Border bottom color ThemeColorText: " & ExampleResult.ThemeColorText
                    Print #filenum, "Border bottom color TintAndShade: " & ExampleResult.TintAndShade
                    Print #filenum, "Border bottom color TintAndShadeText: " & ExampleResult.TintAndShadeText
                    Print #filenum, "Border bottom color RGB: " & ExampleResult.RGB
                    Print #filenum, "Border bottom color RGB_Hex: " & ExampleResult.RGB_Hex
                    Print #filenum, "Border bottom color Red: " & ExampleResult.Red
                    Print #filenum, "Border bottom color Green: " & ExampleResult.Green
                    Print #filenum, "Border bottom color Blue: " & ExampleResult.Blue
                End If
                Print #filenum, "Border bottom line style: "; style.Table.borders(wdBorderBottom).LineStyle
                Print #filenum, "Border bottom line width [pt]: "; style.Table.borders(wdBorderBottom).LineWidth
                                
                Call InitializeColourDetails(MyVar:=ExampleResult)
                ExampleResult.ColourValue = style.Table.borders(wdBorderVertical).color
                ExampleResult = QueryColour(ExampleResult.ColourValue)
                Print #filenum, "Border vertical color [Long]: " & ExampleResult.ColourValue
                Print #filenum, "Border vertical color ColourType: " & GetColourTypeName(ExampleResult.ColourType)
                If ExampleResult.ColourType <> ColourTypeAutomatic Then
                    Print #filenum, "Border vertical color ThemeColorIndex: " & ExampleResult.ThemeColorIndex
                    Print #filenum, "Border vertical color ThemeColorText: " & ExampleResult.ThemeColorText
                    Print #filenum, "Border vertical color TintAndShade: " & ExampleResult.TintAndShade
                    Print #filenum, "Border vertical color TintAndShadeText: " & ExampleResult.TintAndShadeText
                    Print #filenum, "Border vertical color RGB: " & ExampleResult.RGB
                    Print #filenum, "Border vertical color RGB_Hex: " & ExampleResult.RGB_Hex
                    Print #filenum, "Border vertical color Red: " & ExampleResult.Red
                    Print #filenum, "Border vertical color Green: " & ExampleResult.Green
                    Print #filenum, "Border vertical color Blue: " & ExampleResult.Blue
                End If
                Print #filenum, "Border vertical line style: " & style.Table.borders(wdBorderBottom).LineStyle
                Print #filenum, "Border vertical line width [pt]: " & style.Table.borders(wdBorderBottom).LineWidth
                                
            End If
            
            If style.Type = wdStyleTypeList Then
                For i = 1 To 9
                    Print #filenum, "List level " & i & ", alignment: " & GetAlignmentName(style.ListTemplate.ListLevels(i).Alignment)
                    Print #filenum, "List level " & i & ", linked style: " & style.ListTemplate.ListLevels(i).LinkedStyle
                    Print #filenum, "List level " & i & ", number format: " & style.ListTemplate.ListLevels(i).NumberFormat
                    Print #filenum, "List level " & i & ", number position [pt]: " & style.ListTemplate.ListLevels(i).NumberPosition & " or " & Round(style.ListTemplate.ListLevels(i).NumberPosition * C_PointsToCm, 2) & " [cm]"
                    Print #filenum, "List level " & i & ", number style: " & GetNumberingStyle(style.ListTemplate.ListLevels(i).NumberStyle)
                    Print #filenum, "List level " & i & ", start numbering at: " & style.ListTemplate.ListLevels(i).StartAt
                    Print #filenum, "List level " & i & ", text position [pt]: " & style.ListTemplate.ListLevels(i).TextPosition & " or " & Round(style.ListTemplate.ListLevels(i).TextPosition * C_PointsToCm, 2) & " [cm]"
                    Print #filenum, "List level " & i & ", trailing character: " & GetTrailingCharStyle(style.ListTemplate.ListLevels(i).TrailingCharacter)
                Next i
            End If
                        
            Print #filenum, "----------------------------------------"
        End If
    
    ' TOC 1 ÷ TOC 3 were just adapted, TOC 8 is part of ListOfPictures, TOC 9 is part of ListOfTables
    If style.NameLocal = "TOC 1" Or _
        style.NameLocal = "TOC 2" Or _
        style.NameLocal = "TOC 3" Or _
        style.NameLocal = "TOC 8" Or _
        style.NameLocal = "TOC 9" Then
        If style.Type = wdStyleTypeParagraph Then
            Print #filenum, "----------------------------------------"
            Call PrintStyleBasicProperties(filenum, style)
            Print #filenum, "Automatically update: " & style.AutomaticallyUpdate
            Print #filenum, "Language ID: " & style.LanguageId

            ' Font:
            Call PrintFontProperties(filenum, style)
            
            ' Paragraph:
            Call PrintParagraphProperties(filenum, style)
        End If
    End If
    
    Next style
    
    ' Close the file
    Close #filenum
    
    MsgBox _
        Prompt:="Styles saved to:" & vbNewLine _
            & FilePath, _
        Buttons:=vbInformation, _
        Title:=MsgBoxTitle
        
End Sub

Private Sub PrintStyleBasicProperties(filenum As Integer, style As style)
    Print #filenum, "Style Name: " & style.NameLocal
    Print #filenum, "Style Type: " & GetStyleType(style.Type)
    Print #filenum, "Base Style: " & style.BaseStyle.NameLocal
    Print #filenum, "Built-In: " & GetBooleanName(style.BuiltIn)
    Print #filenum, "Decription: " & style.Description
End Sub

Private Sub PrintParagraphProperties(filenum As Integer, style As style)
    Print #filenum, "Alignment: " & GetAlignmentName(style.ParagraphFormat.Alignment)
    Print #filenum, "Outline level: " & style.ParagraphFormat.OutlineLevel
    Print #filenum, "Indentation left [pt]: " & style.ParagraphFormat.LeftIndent & " or " & Round(style.ParagraphFormat.LeftIndent * C_PointsToCm, 2) & " [cm]"
    Print #filenum, "Indentation right [pt]: " & style.ParagraphFormat.RightIndent & " or " & Round(style.ParagraphFormat.RightIndent * C_PointsToCm, 2) & " [cm]"
    Print #filenum, "Indentation special, first line [pt]: " & style.ParagraphFormat.FirstLineIndent
    Print #filenum, "Indentation special, hanging: " & GetBooleanName(style.ParagraphFormat.HangingPunctuation)
    
    Print #filenum, "Spacing before, auto: " & GetBooleanName(style.ParagraphFormat.SpaceBeforeAuto)
    Print #filenum, "Spacing before [pt]: " & style.ParagraphFormat.SpaceBefore
    Print #filenum, "Spacing after, auto: " & GetBooleanName(style.ParagraphFormat.SpaceAfterAuto)
    Print #filenum, "Spacing after [pt]: " & style.ParagraphFormat.SpaceAfter
    Print #filenum, "Line spacing [pt]: " & style.ParagraphFormat.LineSpacing
    ' https://stackoverflow.com/questions/23418243/programmatically-change-dont-add-space-between-paragraphs-of-the-same-style
    Print #filenum, "Don't add space between paragraphs of the same style: " & GetBooleanName(style.NoSpaceBetweenParagraphsOfSameStyle)
    
    Print #filenum, "Window/Orphan control: " & GetBooleanName(style.ParagraphFormat.WidowControl)
    Print #filenum, "Keep with next: " & GetBooleanName(style.ParagraphFormat.KeepWithNext)
    Print #filenum, "Keep lines together: " & GetBooleanName(style.ParagraphFormat.KeepTogether)
    Print #filenum, "Page break before: " & GetBooleanName(style.ParagraphFormat.PageBreakBefore)
    Print #filenum, "Hyphenation: " & GetBooleanName(style.ParagraphFormat.Hyphenation)
    
    Print #filenum, "Linked: " & style.Linked
    Print #filenum, "Link Style: " & style.LinkStyle
    Print #filenum, "List Level Number: " & style.ListLevelNumber
    Print #filenum, "Next Paragraph Style: " & style.NextParagraphStyle
    Print #filenum, "Quick Style: " & style.QuickStyle
End Sub

Private Sub PrintFontProperties(filenum As Integer, style As style)
    Print #filenum, "Font Name: " & style.font.Name
    Print #filenum, "Font Name other: " & style.font.NameOther
    Print #filenum, "Font Size: " & style.font.Size
    Print #filenum, "Font Bold: " & GetBooleanName(style.font.Bold)
    Print #filenum, "Font Italic: " & GetBooleanName(style.font.Italic)
    Print #filenum, "Font Underline: " & GetBooleanName(style.font.Underline)
    Print #filenum, "Font All Caps: " & GetBooleanName(style.font.AllCaps)
    Print #filenum, "Font DoubleStrikeThrough: " & GetBooleanName(style.font.DoubleStrikeThrough)
    Print #filenum, "Font Hidden: " & GetBooleanName(style.font.Hidden)
    Print #filenum, "Font Strikethrough: " & GetBooleanName(style.font.Strikethrough)
    ' Initialize variable
    Call InitializeColourDetails(MyVar:=ExampleResult)
    ExampleResult.ColourValue = style.font.TextColor.RGB
    ExampleResult = QueryColour(ExampleResult.ColourValue)
    Print #filenum, "Font Color [Long]: " & ExampleResult.ColourValue
    Print #filenum, "Font Color ColourType: " & GetColourTypeName(ExampleResult.ColourType)
    If ExampleResult.ColourType <> ColourTypeAutomatic Then
        Print #filenum, "Font Color ThemeColorIndex: " & ExampleResult.ThemeColorIndex
        Print #filenum, "Font Color ThemeColorText: " & ExampleResult.ThemeColorText
        Print #filenum, "Font Color TintAndShade: " & ExampleResult.TintAndShade
        Print #filenum, "Font Color TintAndShadeText: " & ExampleResult.TintAndShadeText
        Print #filenum, "Font Color RGB: " & ExampleResult.RGB
        Print #filenum, "Font Color RGB_Hex: " & ExampleResult.RGB_Hex
        Print #filenum, "Font Color Red: " & ExampleResult.Red
        Print #filenum, "Font Color Green: " & ExampleResult.Green
        Print #filenum, "Font Color Blue: " & ExampleResult.Blue
    End If
    Call InitializeColourDetails(MyVar:=ExampleResult)
    ExampleResult.ColourValue = style.font.shading.BackgroundPatternColor
    ExampleResult = QueryColour(ExampleResult.ColourValue)
    Print #filenum, "Shading Font Color [Long]: " & ExampleResult.ColourValue
    Print #filenum, "Shading Font Color ColourType: " & GetColourTypeName(ExampleResult.ColourType)
    If ExampleResult.ColourType <> ColourTypeAutomatic Then
        Print #filenum, "Shading Font Color ThemeColorIndex: " & ExampleResult.ThemeColorIndex
        Print #filenum, "Shading Font Color ThemeColorText: " & ExampleResult.ThemeColorText
        Print #filenum, "Shading Font Color TintAndShade: " & ExampleResult.TintAndShade
        Print #filenum, "Shading Font Color TintAndShadeText: " & ExampleResult.TintAndShadeText
        Print #filenum, "Shading Font Color RGB: " & ExampleResult.RGB
        Print #filenum, "Shading Font Color RGB_Hex: " & ExampleResult.RGB_Hex
        Print #filenum, "Shading Font Color Red: " & ExampleResult.Red
        Print #filenum, "Shading Font Color Green: " & ExampleResult.Green
        Print #filenum, "Shading Font Color Blue: " & ExampleResult.Blue
    End If
End Sub

Private Function GetStyleType(StyleType As WdStyleType) As String
    Select Case StyleType
        Case wdStyleTypeParagraph
            GetStyleType = "Paragraph"
        Case wdStyleTypeCharacter
            GetStyleType = "Character"
        Case wdStyleTypeTable
            GetStyleType = "Table"
        Case wdStyleTypeList
            GetStyleType = "List"
        Case wdStyleTypeLinked
            GetStyleType = "Linked"
        Case wdStyleTypeParagraphOnly
            GetStyleType = "Paragraph Only"
        Case Else
            GetStyleType = "Unknown"
    End Select
End Function

Private Function GetBooleanName(CurrentVar As Variant) As String
    Select Case CurrentVar
        Case True:      GetBooleanName = "True"
        Case False:     GetBooleanName = "False"
    End Select
End Function

Private Function GetAlignmentName(CurrentVar As Variant) As String
    Select Case CurrentVar
        Case wdAlignParagraphCenter:    GetAlignmentName = "centered"
        Case wdAlignParagraphJustify:   GetAlignmentName = "justified"
        Case wdAlignParagraphLeft:       GetAlignmentName = "left"
        Case wdAlignParagraphRight:     GetAlignmentName = "right"
    End Select
End Function

Private Function GetNumberingStyle(CurrentVar As Variant) As String
    Select Case CurrentVar
        Case wdListNumberStyleArabic:           GetNumberingStyle = "Arabic numbers (1, 2, 3, ...)"
        Case wdListNumberStyleUppercaseRoman:   GetNumberingStyle = "Uppercase Roman numerals (I, II, III, ...)"
        Case wdListNumberStyleLowercaseRoman:   GetNumberingStyle = "Lowercase Roman numerals (i, ii, iii, ...)"
        Case wdListNumberStyleUppercaseLetter:  GetNumberingStyle = "Uppercase letters (A, B, C, ...)"
        Case wdListNumberStyleLowercaseLetter:  GetNumberingStyle = "Lowercase letters (a, b, c, ...)"
    End Select
End Function

Private Function GetTrailingCharStyle(CurrentVar As Variant) As String
    Select Case CurrentVar
        Case wdTrailingTab:     GetTrailingCharStyle = "Tab character"
        Case wdTrailingSpace:   GetTrailingCharStyle = "Space character"
        Case wdTrailingNone:    GetTrailingCharStyle = "No character"
    End Select
End Function

' Combo for validation of styling.
' 2025-03-07 by ms
Sub ShowNonComplientStyling()
    Call ShowNonCompliantStylingInParagraphs ' in module: Styles
    Call ShowNonComplientStylingInTables ' in module: Styles
End Sub

' Shows in every paragraph except tables as yellow highlighting all paragraphs which are not compliant to styles " ms" and TOCs.
' 2025-03-06 by ms
' 2025-12-28 by ms
Private Sub ShowNonCompliantStylingInParagraphs()
    'Const C_BM_NCstylingP As String = "NCstylingP_"
    Dim msStyleCollection As Collection
    Dim para As Paragraph
    Dim NoTotalPar As Integer
    Dim NomsNotCompliantPar As Integer
    Dim NomsCompliantPar As Integer
    Dim summaryMessage As String
    Dim i As Integer
    Dim PerVal As Double
    Dim NoParInTable As Integer
    Dim UserBMDecision As VbMsgBoxResult    ' User Bookmark Decision
    
    Dim FileName As String:     FileName = C_F_Macros
    Dim ModuleName As String:   ModuleName = C_M_Styles
    Dim MacroName As String:    MacroName = "ShowNonCompliantStylingInParagraphs"
    Dim MsgBoxTitle As String:  MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    ' Initialize collection
    Set msStyleCollection = New Collection
    ' Collect styles with " ms" suffix
    Call CollectmsStylesAndTOC(msStyleCollection)
    
    ' Initialize counters
    NoTotalPar = ActiveDocument.Paragraphs.count
    NomsCompliantPar = 0
    NomsNotCompliantPar = 0
    NoParInTable = 0
    
    Beep
    UserBMDecision = MsgBox( _
        Prompt:="Do you want to add bookmarks " & vbNewLine & vbNewLine & _
            C_BM_NCstylingP & vbNewLine & vbNewLine & _
            " in paragraphs where non-compliance is detected?", _
            Buttons:=vbYesNo + vbQuestion, _
        Title:=MsgBoxTitle)

    Call AddLastCursorPositionBookmark

    ' Initialization of the dedicated Form
    TemplateStyleValidation_Form.Show vbModeless ' sets ShowModal to False in the corresponding Form

    ' Check each paragraph
    For i = 1 To NoTotalPar
        Set para = ActiveDocument.Paragraphs(i)
        If Not para.Range.Information(wdWithInTable) Then   ' All outside the tables
            NoParInTable = NoParInTable + 1
            If Not IsInCollection(msStyleCollection, para.style.NameLocal) Then
                DoEvents    ' Force a screen refresh
                para.Range.HighlightColorIndex = wdYellow
                DoEvents    ' Force a screen refresh
                NomsNotCompliantPar = NomsNotCompliantPar + 1
                ' Insert a bookmark "NCstylingP_x" where x is equal to NomsNotCompliantPar
                If UserBMDecision = vbYes Then
                    para.Range.Bookmarks.Add Name:=C_BM_NCstylingP & NomsNotCompliantPar
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
        
        ' The DoEvents function in Visual Basic for Applications (VBA) for Microsoft Word is used to yield execution so that the operating system can process other events. This function allows the operating system to handle other tasks, such as updating the screen, responding to user inputs, or processing other events in the queue, while your macro is running
        ' Allow other events to be processed
        DoEvents
    Next i
    
    Unload TemplateStyleValidation_Form
    Call RemoveLastCursorPositionBookmark
    
    ' Display summary
    summaryMessage = "Total number of paragraphs: " & NoTotalPar & vbCrLf & _
                     "Number of ms compliant paragraphs: " & NomsCompliantPar & vbCrLf & _
                     "Number of ms non-compliant paragraphs: " & NomsNotCompliantPar & vbNewLine & _
                     "Number of not examined paragraphs in tables: " & NoParInTable
    MsgBox _
        Prompt:=summaryMessage, _
        Buttons:=vbInformation, _
        Title:=MsgBoxTitle
    
    ' Clear object variables
    Set msStyleCollection = Nothing
    Set para = Nothing
End Sub

' Shows only in tables as yellow highlighting all paragraphs, which are not compliant to " ms" styling and TOCs.
' Sets a bookmark at every paragraph which is not compliant.
' 2025-03-06 by ms and AI
' 2025-12-28 by ms
Private Sub ShowNonComplientStylingInTables()
'Const C_BM_NCstylingT As String = "NCstylingT_"
    Dim tbl As Table
    Dim SingleCell As Cell
    Dim SinglePar As Paragraph
    Dim msStyleCollection As Collection
    Dim TblIndex As Integer
    Dim CellIndex As Integer
    Dim ParaIndex As Integer
    
    Dim NoParInTables As Integer
    Dim NoTotalPar As Integer
    Dim NomsNotCompliantPar As Integer
    Dim NomsCompliantPar As Integer
    
    Dim UserBMDecision As VbMsgBoxResult    ' User Bookmark Decision
    
    Dim FileName As String:     FileName = C_F_Macros
    Dim ModuleName As String:   ModuleName = C_M_Styles
    Dim MacroName As String:    MacroName = "ShowNonComplientStylingInTables"
    Dim MsgBoxTitle As String:  MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    Beep
    UserBMDecision = MsgBox( _
        Prompt:="Do you want to add bookmarks" & vbNewLine & vbNewLine & _
            C_BM_NCstylingT & vbNewLine & vbNewLine & _
            "in paragraphs where non-compliance is detected?", _
        Buttons:=vbYesNo + vbQuestion, _
        Title:=MsgBoxTitle)
                
    ' Initialization
    NoParInTables = 0
    NoTotalPar = ActiveDocument.Paragraphs.count
    NomsNotCompliantPar = 0
    NomsCompliantPar = 0
    
    ' Initialize collection
    Set msStyleCollection = New Collection
    ' Collect styles with " ms" suffix
    Call CollectmsStylesAndTOC(msStyleCollection)
    
    ' Loop through all tables in the document
    For TblIndex = 1 To ActiveDocument.Tables.count
        Set tbl = ActiveDocument.Tables(TblIndex)
        ' Loop through all cells in the table
        For CellIndex = 1 To tbl.Range.Cells.count
            Set SingleCell = tbl.Range.Cells(CellIndex)
            ' Loop through all paragraphs in the cell
            For ParaIndex = 1 To SingleCell.Range.Paragraphs.count
                Set SinglePar = SingleCell.Range.Paragraphs(ParaIndex)
                NoParInTables = NoParInTables + 1
                If Not IsInCollection(msStyleCollection, SinglePar.style.NameLocal) Then
                    DoEvents    ' Force a screen refresh
                    SinglePar.Range.HighlightColorIndex = wdYellow
                    DoEvents    ' Force a screen refresh
                    NomsNotCompliantPar = NomsNotCompliantPar + 1
                    ' Insert a bookmark "NCstylingT_x" where x is equal to NomsNotCompliantPar
                    If UserBMDecision = vbYes Then
                        SinglePar.Range.Bookmarks.Add Name:=C_BM_NCstylingT & NomsNotCompliantPar
                    End If
                Else
                    SinglePar.Range.HighlightColorIndex = wdNoHighlight
                End If
                    NomsCompliantPar = NomsCompliantPar + 1
            Next ParaIndex
        Next CellIndex
    Next TblIndex
    
    MsgBox _
        Prompt:="Checked all paragraphs within tables." & vbNewLine & _
            "No. of paragraphs in document: " & NoTotalPar & vbNewLine & _
            "No. of table paragraphs: " & NoParInTables & vbNewLine & _
            "No. of compliant table paragraphs: " & NomsCompliantPar & vbNewLine & _
            "No. of non-compliant table paragraphs: " & NomsNotCompliantPar, _
        Buttons:=vbInformation, _
        Title:=MsgBoxTitle
    
    ' Clear object variables
    Set msStyleCollection = Nothing
    Set tbl = Nothing
    Set SingleCell = Nothing
    Set SinglePar = Nothing
End Sub

' Remove all bookmarks added by the sub ShowNonCompliantStylingInParagraphs() and ShowNonComplientStylingInTables()
' 2025-03-06 by ms and AI
' 2025-12-28 by ms
Sub DeleteAllNCstylingBookmarks()
'Const C_BM_NCstylingP As String = "NCstylingP_"
'Const C_BM_NCstylingT As String = "NCstylingT_"
    Dim Bm As bookmark
    Dim i As Integer
    Dim BookmarkName As String
    
    Dim FileName As String:      FileName = C_F_Macros
    Dim ModuleName As String:    ModuleName = C_M_Styles
    Dim MacroName As String:     MacroName = "DeleteAllNCstylingBookmarks"
    Dim MsgBoxTitle As String:   MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName

    ' Loop through all bookmarks in the document
    For i = ActiveDocument.Bookmarks.count To 1 Step -1
        Set Bm = ActiveDocument.Bookmarks(i)
        BookmarkName = Bm.Name

        ' Check if the bookmark name starts with "NCstyling_"
        If Left(BookmarkName, 11) = C_BM_NCstylingP Or Left(BookmarkName, 11) = C_BM_NCstylingT Then
            ' Check if the rest of the name is a number
            If IsNumeric(Mid(BookmarkName, 12)) Then
                ' Remove the bookmark
                Bm.Delete
            End If
        End If
    Next i

    MsgBox _
        Prompt:="All NCstyling bookmarks" & vbNewLine & vbNewLine & _
            C_BM_NCstylingP & " and " & C_BM_NCstylingT & vbNewLine & vbNewLine & _
            "have been removed.", _
        Buttons:=vbInformation, _
        Title:=MsgBoxTitle
    
    ' Clear object variables
    Set Bm = Nothing
End Sub

' Reverse the macro ShowNonCompliantStylingInParagraphs() and ShowNonComplientStylingInTables()
' 2025-03-01 reworked by ms
' 2025-03-06 reworked by ms
' 2025-12-28 by ms
Sub DeleteNCHighlighting()
    Dim para As Paragraph
    Dim NoTotalPar As Integer
    Dim i As Integer
    Dim PerVal As Double
    
    Dim FileName As String:      FileName = C_F_Macros
    Dim ModuleName As String:    ModuleName = C_M_Styles
    Dim MacroName As String:     MacroName = "DeleteNCHighlighting"
    Dim MsgBoxTitle As String:   MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    Call MacroBeginning
    Call AddLastCursorPositionBookmark
    Call CheckMicrosoftWordVersion(MacroName)
    
    NoTotalPar = ActiveDocument.Paragraphs.count
        
    ' Future: change Form caption
    TemplateStyleValidation_Form.Show vbModeless ' sets ShowModal to False in the corresponding Form
    For i = 1 To NoTotalPar
        Set para = ActiveDocument.Paragraphs(i)
        
        If para.Range.HighlightColorIndex = wdYellow Then
            para.Range.HighlightColorIndex = wdNoHighlight
        End If
        
        ' Update progress label
        PerVal = (i / NoTotalPar) * 100 ' Calculate percentage value
        TemplateStyleValidation_Form.ProgressLabel = "Paragraph counter: " & i & " out of " & NoTotalPar & _
            " (" & Int(PerVal) & "%)"
        ' The DoEvents function in Visual Basic for Applications (VBA) for Microsoft Word is used to yield execution so that the operating system can process other events. This function allows the operating system to handle other tasks, such as updating the screen, responding to user inputs, or processing other events in the queue, while your macro is running
        DoEvents
    Next i
    
    Unload TemplateStyleValidation_Form
    Call MacroFinish
    Call RemoveLastCursorPositionBookmark

    MsgBox _
        Prompt:="Finished processing.", _
        Buttons:=vbInformation, _
        Title:=MsgBoxTitle
    
    ' Clear object variables
    Set para = Nothing
End Sub

' Collect styles with " ms" suffix and additionally TOC 1 ÷ TOC 4
' 2025-03-06 by ms
Private Sub CollectmsStylesAndTOC(ByRef msStyleCollection As Collection)
'   Const C_StyleSuffix As String = " ms"
    Dim style As style
    
    For Each style In ActiveDocument.Styles
        If style.Type = wdStyleTypeParagraph And InStr(style.NameLocal, C_StyleSuffix) > 0 Then
            msStyleCollection.Add style.NameLocal
        End If
    Next style
    msStyleCollection.Add "TOC 1"
    msStyleCollection.Add "TOC 2"
    msStyleCollection.Add "TOC 3"
    msStyleCollection.Add "TOC 4"
End Sub

' 2025-12-08 by ms and AI
' 2025-12-28 by ms
Public Sub DeleteCustomStyles_KeepOnlyDefined()
    Dim deletedStyles As Long, skippedStyles As Long, errStyles As Long
    Dim deletedLists  As Long, skippedLists  As Long, errLists  As Long
    
    Dim FileName As String:     FileName = C_F_Macros
    Dim ModuleName As String:   ModuleName = C_M_Styles
    Dim MacroName As String:    MacroName = "DeleteCustomStyles_KeepOnlyDefined"
    Dim MsgBoxTitle As String:  MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    Dim sty As style
    Dim ls  As String
    Dim keep As Boolean
    Dim report As String
    Dim WhichDeleted As String
    WhichDeleted = ""
    
    ' 1) Initialize the four tables
    Call InitParagraphStyles
    Call InitCharacterStyles
    Call InitTableStyles
    Call InitListTemplates
    
    ' Optional confirmation (destructive operation!)
    Beep
    If MsgBox( _
        Prompt:="Delete all custom styles and list styles not in the defined tables?" & vbCrLf & _
              "This cannot be undone.", _
        Buttons:=vbQuestion + vbYesNo, _
        Title:=MsgBoxTitle) <> vbYes Then
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    On Error GoTo CleanExit   ' safety guard
    
    ' 2) Styles: delete all *custom* styles not in allowlists
    For Each sty In ActiveDocument.Styles
        ' Skip built-ins (not deletable)
        If sty.BuiltIn Then
            skippedStyles = skippedStyles + 1
        Else
            ' Keep if listed in any of your tables
            keep = _
                IsAllowedStyleName(sty.NameLocal, ParagraphStyles) Or _
                IsAllowedStyleName(sty.NameLocal, CharacterStyles) Or _
                IsAllowedStyleName(sty.NameLocal, TableStyles)
            
            If keep Then
                skippedStyles = skippedStyles + 1
            Else
                On Error Resume Next
                WhichDeleted = WhichDeleted & sty.NameLocal & vbNewLine
                sty.Delete
                If Err.Number <> 0 Then
                    errStyles = errStyles + 1
                    Err.Clear
                Else
                    deletedStyles = deletedStyles + 1
                End If
                On Error GoTo 0
            End If
        End If
    Next sty
    
CleanExit:
    Application.ScreenUpdating = True
    
    ' 3) Summary
    report = "Custom Styles:" & vbCrLf & _
             "  Deleted: " & deletedStyles & vbCrLf & _
             "  Which deleted: " & vbNewLine & _
                WhichDeleted & vbNewLine & _
             "  Skipped (allowed or built-in): " & skippedStyles & vbCrLf & _
             "  Errors (couldn't delete): " & errStyles & vbCrLf & vbCrLf & _
             "List Styles/Templates:" & vbCrLf & _
             "  Deleted: " & deletedLists & vbCrLf & _
             "  Skipped (allowed or built-in): " & skippedLists & vbCrLf & _
             "  Errors (couldn't delete): " & errLists
    MsgBox _
        Prompt:=report, _
        Buttons:=vbInformation, _
        Title:=MsgBoxTitle
End Sub

' Returns True if styleName matches any second element (Style Name) in the given table (2D Variant-of-Variants)
' 2025-12-08 by ms and AI
Private Function IsAllowedStyleName(ByVal styleName As String, ByVal tableArr As Variant) As Boolean
    Dim i As Long
    Dim s As String, query As String
    query = LCase$(Trim$(styleName))
    On Error GoTo Bail
    For i = LBound(tableArr) To UBound(tableArr)
        s = LCase$(Trim$(tableArr(i)(1))) ' second element holds the Style Name
        If s = query Then
            IsAllowedStyleName = True
            Exit Function
        End If
    Next i
Bail:
End Function

' Tool to clear all styles numbering.
' 2026-01-02 by ms and AI
Private Sub ResetTOCStylesNumbering()
    Dim i As Integer
    Dim styleName As Variant
    Dim stylesToClean As Variant
    
    ' Lista stylÃ³w do wyczyszczenia
    stylesToClean = Array(wdStyleTOC1, wdStyleTOC2, wdStyleTOC3)
    
    For Each styleName In stylesToClean
        With ActiveDocument.Styles(styleName)
            ' To polecenie caÅ‚kowicie usuwa powiÄ…zanie stylu z jakÄ…kolwiek listÄ…
            .LinkToListTemplate ListTemplate:=Nothing
        End With
    Next styleName
    
    ' OdÅ›wieÅ¼ spis treÅ›ci
    If ActiveDocument.TablesOfContents.count > 0 Then
        ActiveDocument.TablesOfContents(1).Update
    End If
End Sub
