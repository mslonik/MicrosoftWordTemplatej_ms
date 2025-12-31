Attribute VB_Name = "Tools"
' VBA Module name: Tools.bas
' https://github.com/mslonik/Microsoft-Word-Configuration
'
'   License: MIT License.
'
'
'+----+----------------------------+-------------+-----------------+----------------------------+
'| No | Sub name                   | Ribbon name | Ribbon section  | Ribbon button name         |
'+----+----------------------------+-------------+-----------------+----------------------------+
'| 1  | DocPropertiesUpdate        | Tools_ms    | DocProperties   | AttachTheme                |
'| 2  | DocPropertiesUserInput     | Tools_ms    | DocProperties   | DocPropertiesUserInput     |
'| 3  | SetMarginsDefault          | Tools_ms    | Document        | SetMarginsDefault          |
'| 4  | SetMarginsMinimal          | Tools_ms    | Document        | SetMarginsMinimal          |
'| 5  | SetHyphenation             | Tools_ms    | Document        | SetHyphenation             |
'| 6  | SetLanguageToEnglishUS     | Tools_ms    | Document        | SetLanguageToEnglishUS     |
'| 7  | SetPageColorToCustom       | Tools_ms    | Document        | SetPageColorToCustom       |
'| 8  | ShowAllTemplates           | Tools_ms    | Document        | ShowAllTemplates           |
'| 9  | CommentAddNumber           | Tools_ms    | Comments        | CommentAddNumber           |
'| 10 | CommentDeleteNumber        | Tools_ms    | Comments        | CommentDeleteNumber        |
'| 11 | CommentCountByUser         | Tools_ms    | Comments        | CommentCountByUser         |
'| 12 | DeleteAllUserBookmarks     | Tools_ms    | Before printing | DeleteAllUserBookmarks     |
'| 13 | ParDistAtNewSectionCheck   | Tools_ms    | Before printing | ParDistAtNewSectionCheck   |
'| 14 | ParDistAtNewSectionReduce  | Tools_ms    | Before printing | ParDistAtNewSectionReduce  |
'| 15 | ParDistAtNewSectionRestore | Tools_ms    | Before printing | ParDistAtNewSectionRestore |
'| 16 | CanvaFormatTextBoxes       | Tools_ms    | Canva           | CanvaFormatTextBoxes       |
'| 17 | CanvaToggleBorder          | Tools_ms    | Canva           | CanvaToggleBorder          |
'| 18 | CanvaInsertPNGfiles        | Tools_ms    | Canva           | CanvaInsertPNGfiles        |
'+----+----------------------------+-------------+-----------------+----------------------------+
'
'
'   Captions:
'   19. CaptionShow()
'   20. CapationAddCustomized()
'   21. CaptionLabelDeleteCustomized()
'
'
'   Word options:
'   22. WordOptionsCustomize()
'   23. WordOptionsRestore()
'   24. WordOptionsDisableAutoFormat()
'   25. WordOptionsRestoreAutoFormat
'   26. WordOptionsDisableAutoCorrect()
'   27. WordOptionsToggleAutoCorrect()
'
'   Section related to shortcuts:
'   31. ToggleSpecificFormatting()
'   32. SaveDocumentAsPDFWithSettings()
'   33. ReapplyTemplateStyle()
'   34. RestartListNumbering()
'   35. ToggleHeadingCollapseExpand()
'   36. CustomizedOvertype()
'   37. InsertCrossRef()                <- keyboard shortcut F7
'   38. InsertCrossReferences            -> DeleteCrossReferences()
'   40. CustomizedPrinting()
'   41. CustomizedSaveAs()
'   42. CustomizedToggleFieldCodes()
'   43. Strikethrough()
'   44. Italic()
'   45. Underline()
'   46. Bold()
'   47. JumpToNextList()
'   48. JumpToNextTable()
'   49. JumpToNextCanvas()
'
'   Others / legacy:
'   50. InsertSVNCommitNumber()
'   51. AttachBuildingBlocks()
'   52. AutoExec()
'
'   ResetHyphenation()
'   RestoreDefaultPageColor()
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
' Used to enforce the explicit declaration of all variables in a module. When you include Option Explicit at the beginning of a module, it ensures that you must
' declare all variables using the Dim, Private, Public, ReDim, or Static statements before using them. This helps prevent errors caused by typos or
' undeclared variables.
Option Explicit

' The following enum is used in InsertCrossRef()
Private Enum RefType
    RefTypeNotDefined = 0
    RefTypeC_Caption_Pic = 1
    RefTypeC_Caption_Tab = 2
    RefTypeReference = 3
    RefTypeHeading = 4
End Enum

Dim WordAppEvents As ClsAppEvents
' For sleep / delay function
Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)

' Inserts PNG files from the specified folder.
' ms and AI on 2025-02-07
' ms on 2025-03-15 added BuildingBlocks "Legend Picture" and source name.
Sub CanvaInsertPNGfiles()
    Dim folderPath As String
    Dim pngFiles As Collection
    Dim file As Variant
    Dim doc As Document
    Dim canvasShape As Shape
    Dim pictureShape As Shape
    Dim rng As Range
    Dim totalFiles As Integer
    Dim estimatedTime As Double
    Dim processingTime As Double
    Dim totalTime As Double
    Dim TemplateName As String
    
    Dim FileName As String
    FileName = C_F_Macros
    
    Dim ModuleName As String
    ModuleName = C_M_Tools
    
    Dim MacroName As String
    MacroName = "CanvaInsertPNGfiles"
    
    Dim MsgBoxTitle As String
    MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    ' Initialize
    Set pngFiles = New Collection
    Set rng = Selection.Range
    Set doc = ActiveDocument
    TemplateName = doc.AttachedTemplate.Name
    
    ' Open folder selection dialog
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select Folder Containing PNG Files"
        If .Show = -1 Then
            folderPath = .SelectedItems(1)
        Else
            MsgBox _
                Prompt:="No folder selected. Macro terminated.", _
                Buttons:=vbExclamation + vbOKOnly, _
                Title:=MsgBoxTitle
            Exit Sub
        End If
    End With
    
    ' Get all PNG files in the folder
    Set pngFiles = GetPNGFilesInFolder(folderPath)
    
    ' Check if any PNG files were found
    totalFiles = pngFiles.count
    If totalFiles = 0 Then
        MsgBox _
            Prompt:="No PNG files found in the selected folder.", _
            Buttons:=vbExclamation + vbOKOnly, _
            Title:=MsgBoxTitle
        Exit Sub
    End If
    
    ' Calculate estimated time
    estimatedTime = totalFiles * 2
    processingTime = totalFiles * 0.5
    totalTime = estimatedTime + processingTime
    
    ' Show message box with the number of files and estimated time
    MsgBox _
        Prompt:="Number of PNG files found: " & totalFiles & vbCrLf & _
            "Estimated time of insertion: " & estimatedTime & " seconds" & vbCrLf & _
            "Processing time: " & processingTime & " seconds" & vbCrLf & _
            "Total time: " & totalTime & " seconds", _
        Buttons:=vbInformation + vbOKOnly, _
        Title:=MsgBoxTitle
    
    ' Insert each PNG file into a separate canvas with an empty paragraph in between
    For Each file In pngFiles
        ' Insert empty paragraph and format with style "Normal ms"
        rng.Collapse Direction:=wdCollapseEnd
        rng.InsertParagraphAfter
        rng.style = C_S_ParNormal
        rng.Collapse Direction:=wdCollapseEnd
        
        ' Insert next empty paragraph and format with style "PictureCanva ms"
        rng.InsertParagraphAfter
        rng.style = C_S_ParPictureCanva
        rng.Collapse Direction:=wdCollapseEnd
        
        ' Move back (up) to the empty paragraph formatted with style "PictureCanva ms"
        rng.MoveStart Unit:=wdParagraph, count:=-1
        rng.Select
        
        ' Add a new canvas
        Set canvasShape = doc.Shapes.AddCanvas(0, 0, 500, 500)
        
        ' Set Format Drawing Canvas Fill to "No fill"
        canvasShape.Fill.Transparency = 1#
        
        ' Add the picture to the canvas
        Set pictureShape = canvasShape.CanvasItems.AddPicture(folderPath & "\" & file)
        
        ' Set Layout Option to 'With Text Wrapping' and 'In Line With Text'
        canvasShape.WrapFormat.Type = wdWrapInline
        
        ' Insert empty paragraph and BuildingBlock "LegendPicture"
        rng.Collapse Direction:=wdCollapseEnd
        rng.InsertParagraphAfter
        
        Dim bb As BuildingBlock
        Dim bbe As BuildingBlockEntries
        Set bbe = ReturnBuildingBlockEntries() ' in module Shortcuts
        Set bb = bbe(C_BB_LegendPicture)
        bb.Insert rng, True
        
        ' Here do text processing
        Dim prevPara As Range
        Set prevPara = rng.Paragraphs(rng.Paragraphs.count).Range
        With prevPara.Find
            .Text = "[source: ]"
            .Replacement.Text = "[source: " & file & "]"
            .Forward = False
            .Wrap = wdFindStop
            .Format = True
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
            .Execute Replace:=wdReplaceAll
        End With
        ' Delay to ensure proper insertion
        Dim startTime As Single
        startTime = Timer
        Do While Timer < startTime + 2
            DoEvents
        Loop
    Next file
       
    ' Clear object variables
    Set pngFiles = Nothing
    Set doc = Nothing
    Set rng = Nothing
    Set pictureShape = Nothing
    Set canvasShape = Nothing
    Set prevPara = Nothing
    Set bbe = Nothing
    Set bb = Nothing
       
    ' Show message box indicating successful completion
    MsgBox _
        Prompt:="Macro '" & MacroName & "' completed successfully." & vbCrLf & _
            "Template: " & TemplateName, _
        Buttons:=vbInformation + vbOKOnly, _
        Title:=MsgBoxTitle
End Sub

Private Function GetPNGFilesInFolder(folderPath As String) As Collection
    Dim pngFiles As Collection
    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    Dim FileName As String
    FileName = C_F_Macros
    
    Dim ModuleName As String
    ModuleName = C_M_Tools
    
    Dim MacroName As String
    MacroName = "GetPNGFilesInFolder"
    
    Dim MsgBoxTitle As String
    MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    ' Initialize
    Set pngFiles = New Collection
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Check if the folder exists
    If Not fso.FolderExists(folderPath) Then
        MsgBox _
            Prompt:="The specified folder does not exist.", _
            Buttons:=vbExclamation + vbOKOnly, _
            Title:=MsgBoxTitle
        Exit Function
    End If
    
    ' Get all PNG files in the folder
    Set folder = fso.GetFolder(folderPath)
    For Each file In folder.Files
        If LCase(fso.GetExtensionName(file.Name)) = "png" Then
            pngFiles.Add file.Name
        End If
    Next file
    
    ' Return the collection of PNG files
    Set GetPNGFilesInFolder = pngFiles
    
    ' Clear object variables
    Set pngFiles = Nothing
    Set fso = Nothing
    Set folder = Nothing
End Function

' Minimal headers and footers, mainly to print pictures.
' 2025-08-21 by ms
Sub SetMarginsMinimal()
    Dim MarginInside As Double
    Dim MarginOutside As Double
    Dim MirrorMarginsDecision As Boolean
    Dim HFDistance As Double
    Dim GutterSize As Double

    Dim FileName As String
    FileName = C_F_Macros
    
    Dim ModuleName As String
    ModuleName = C_M_Tools
    
    Dim MacroName As String
    MacroName = "SetMarginsMinimal"
    
    Dim MsgBoxTitle As String
    MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
        
    MarginInside = 0.5  ' cm
    MarginOutside = 0.5 ' cm
    MirrorMarginsDecision = False
    HFDistance = 0#     ' cm
    GutterSize = 0#     ' cm
     
     With ActiveDocument.PageSetup
        .TopMargin = CentimetersToPoints(MarginInside)
        .BottomMargin = CentimetersToPoints(MarginInside)
        .LeftMargin = CentimetersToPoints(MarginInside) ' This sets the inside margin
        .RightMargin = CentimetersToPoints(MarginOutside) ' This sets the outside margin
        .Orientation = wdOrientPortrait
        .MirrorMargins = MirrorMarginsDecision
        .PaperSize = wdPaperA4
        .HeaderDistance = CentimetersToPoints(HFDistance)
        .FooterDistance = CentimetersToPoints(HFDistance)
        .Gutter = CentimetersToPoints(GutterSize)
    End With
    
    MsgBox _
        Prompt:="Margins, headers and footers were set to the required values:" & vbNewLine & vbNewLine & _
            "top margin = " & MarginInside & " cm" & vbNewLine & _
            "bottom margin = " & MarginInside & " cm" & vbNewLine & _
            "left margin = " & MarginOutside & " cm" & vbNewLine & _
            "right margin = " & MarginInside & " cm" & vbNewLine & vbNewLine & _
            "mirror margins = " & MirrorMarginsDecision & vbNewLine & vbNewLine & _
            "header distance = " & HFDistance & " cm" & vbNewLine & _
            "footer distance = " & HFDistance & " cm" & vbNewLine & vbNewLine & _
            "gutter size = " & GutterSize & " cm", _
        Buttons:=vbInformation + vbOKOnly, _
        Title:=MsgBoxTitle
End Sub


' Sets nominal values of margins, header and footer.
' Unfortunately this is one of two ways to store information about such parametersi in the template file.
' (Second way is to apply Document Variables, which is nearly identical).
' 2025-02-02 by ms and AI
' 2025-03-02 by ms and AI
' 2025-08-05 by ms added gutter size
Sub SetMarginsDefault()
    Dim MarginInside As Double
    Dim MarginOutside As Double
    Dim MirrorMarginsDecision As Boolean
    Dim HFDistance As Double
    Dim GutterSize As Double

    Dim FileName As String
    FileName = C_F_Macros
    
    Dim ModuleName As String
    ModuleName = C_M_Tools
    
    Dim MacroName As String
    MacroName = "SetMarginsDefault"
    
    Dim MsgBoxTitle As String
    MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
        
    MarginInside = 1.2  ' cm
    MarginOutside = 2.2 ' cm
    MirrorMarginsDecision = True
    HFDistance = 0.5    ' cm
    GutterSize = 0#     ' cm
     
     With ActiveDocument.PageSetup
        .TopMargin = CentimetersToPoints(MarginInside)
        .BottomMargin = CentimetersToPoints(MarginInside)
        .LeftMargin = CentimetersToPoints(MarginInside) ' This sets the inside margin
        .RightMargin = CentimetersToPoints(MarginOutside) ' This sets the outside margin
        .Orientation = wdOrientPortrait
        .MirrorMargins = MirrorMarginsDecision
        .PaperSize = wdPaperA4
        .HeaderDistance = CentimetersToPoints(HFDistance)
        .FooterDistance = CentimetersToPoints(HFDistance)
        .Gutter = CentimetersToPoints(GutterSize)
    End With
    
    MsgBox _
        Prompt:="Margins, headers and footers were set to the required values:" & vbNewLine & vbNewLine & _
            "top margin = " & MarginInside & " cm" & vbNewLine & _
            "bottom margin = " & MarginInside & " cm" & vbNewLine & _
            "left margin = " & MarginOutside & " cm" & vbNewLine & _
            "right margin = " & MarginInside & " cm" & vbNewLine & vbNewLine & _
            "mirror margins = " & MirrorMarginsDecision & vbNewLine & vbNewLine & _
            "header distance = " & HFDistance & " cm" & vbNewLine & _
            "footer distance = " & HFDistance & " cm" & vbNewLine & vbNewLine & _
            "gutter size = " & GutterSize & " cm", _
        Buttons:=vbInformation + vbOKOnly, _
        Title:=MsgBoxTitle
End Sub


' Switches active document view properties in a loop, 4x views are available.
' Keyboard shortcut: F4.
' Settings are stored within ActiveDocument to restore them next time document is opened.
' Reworked by ms on 2025-02-11
' Reworked by ms on 2025-07-29
' 2025-12-31 by ms
Sub ToggleSpecificFormatting()
    Dim FileName As String:       FileName = C_F_Macros
    Dim ModuleName As String:     ModuleName = C_M_Tools
    Dim MacroName As String:      MacroName = "ToggleSpecificFormatting"
    Dim MsgBoxTitle As String:    MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    Call CheckMicrosoftWordVersion(MacroName)   ' in module 'Tools'
    
    Dim oView As View
    Set oView = ActiveDocument.ActiveWindow.View
    
    Static FormattingToggle As Boolean      ' static initial value: false
    ' View Mode = local counter: <1, 4>.
    ' 1 = toggle formatting, visibility of gridline.
    ' 2 = toggle page color: grey / white.
    ' 3 = toggle formatting, visibility of gridline.
    ' 4 = toggle page color: grey / white.
    Static ViewMode As Byte                 ' static initial value: 0
    
    Dim DocVarName As String: DocVarName = "DocVarToggleSpecificFormatting"
    Dim DocVarTemp As Variable
    Dim FlagDocVarExists As Boolean:    FlagDocVarExists = False
    Dim FlagUpdateDocVar As Boolean:    FlagUpdateDocVar = False
    
    ' Check if document variable named DocVarToggleSpecificFormatting exists in ActiveDocument.
    For Each DocVarTemp In ActiveDocument.Variables
        If DocVarTemp.Name = DocVarName Then
            FlagDocVarExists = True
            Exit For
        End If
    Next DocVarTemp
    
    Dim UserDecision As VbMsgBoxResult
    If FlagDocVarExists Then
        ViewMode = CByte(ActiveDocument.Variables(DocVarName).Value)
        FlagUpdateDocVar = True
    Else
        ' If such document variable doesn't exist, ask user if it should be created and saved
        UserDecision = MsgBox( _
                            Prompt:="Document variable: " & DocVarName & " doesn't exist in the ActiveDocument." & vbNewLine & vbNewLine & _
                                "Do you want to create it to persist the view state?" & vbNewLine & vbNewLine & _
                                "It is strongly recommended to do so.", _
                            Buttons:=vbQuestion + vbYesNo, _
                            Title:=MsgBoxTitle)
        If UserDecision = vbYes Then
            ActiveDocument.Variables.Add _
                Name:=DocVarName, _
                Value:=ViewMode
        End If
    End If
    
    ViewMode = ViewMode + 1
    If FlagUpdateDocVar Then ActiveDocument.Variables(DocVarName).Value = ViewMode
    
    If ViewMode = 1 Then
        FormattingToggle = Not FormattingToggle
        If FormattingToggle = False Then
            oView.ShowTextBoundaries = True
            oView.FieldShading = wdFieldShadingAlways
            oView.ShowHiddenText = True
            oView.ShowAll = True
            ActiveWindow.View.TableGridlines = True
            ActiveWindow.View.ShowCropMarks = True
        Else
            oView.ShowTextBoundaries = False
            oView.FieldShading = wdFieldShadingWhenSelected
            oView.ShowHiddenText = False
            oView.ShowAll = False
            ActiveWindow.View.TableGridlines = False
            ActiveWindow.View.ShowCropMarks = False
        End If
        MsgBox _
            Prompt:="Specific formatting was just toggled:" & ViewMode & vbNewLine & vbNewLine & _
                "ShowTextBoundaries" & vbNewLine & _
                "FieldShading" & vbNewLine & _
                "ShowHiddenText" & vbNewLine & _
                "ShowAll" & vbNewLine & _
                "TableGridlines" & vbNewLine & _
                "ShowCropMarks", _
            Buttons:=vbInformation + vbOKOnly, _
            Title:=MsgBoxTitle
    End If
    
    If ViewMode = 2 Then
        If FormattingToggle = False Then
            Call SetPageColorToCustom               ' in module Tools
        Else
            Call RestoreDefaultPageColor            ' in module Tools
        End If
        MsgBox _
            Prompt:="Specific formatting was just toggled:" & ViewMode & vbNewLine & vbNewLine & _
                "Page background color was just toggled.", _
            Buttons:=vbInformation + vbOKOnly, _
            Title:=MsgBoxTitle
    End If
    
    If ViewMode = 3 Then
        FormattingToggle = Not FormattingToggle
        If FormattingToggle = False Then
            oView.ShowTextBoundaries = True
            oView.FieldShading = wdFieldShadingAlways
            oView.ShowHiddenText = True
            oView.ShowAll = True
            ActiveWindow.View.TableGridlines = True
            ActiveWindow.View.ShowCropMarks = True
        Else
            oView.ShowTextBoundaries = False
            oView.FieldShading = wdFieldShadingWhenSelected
            oView.ShowHiddenText = False
            oView.ShowAll = False
            ActiveWindow.View.TableGridlines = False
            ActiveWindow.View.ShowCropMarks = False
        End If
        MsgBox _
            Prompt:="Specific formatting was just toggled:" & ViewMode & vbNewLine & vbNewLine & _
                "ShowTextBoundaries" & vbNewLine & _
                "FieldShading" & vbNewLine & _
                "ShowHiddenText" & vbNewLine & _
                "ShowAll" & vbNewLine & _
                "TableGridlines" & vbNewLine & _
                "ShowCropMarks", _
            Buttons:=vbInformation + vbOKOnly, _
            Title:=MsgBoxTitle
    End If
    
    If ViewMode = 4 Then
        If FormattingToggle = False Then
            Call SetPageColorToCustom               ' in module Tools
        Else
            Call RestoreDefaultPageColor            ' in module Tools
        End If
        MsgBox _
            Prompt:="Specific formatting was just toggled:" & ViewMode & vbNewLine & vbNewLine & _
                "Page background color was just toggled.", _
            Buttons:=vbInformation + vbOKOnly, _
            Title:=MsgBoxTitle
        ViewMode = 0
        If FlagUpdateDocVar Then ActiveDocument.Variables(DocVarName).Value = ViewMode
    End If
    
    ' Clear object variables
    Set oView = Nothing
    
End Sub

Private Sub CheckMicrosoftWordVersion(MacroName As String)
    Dim FileName As String
    FileName = C_F_Macros
    
    Dim ModuleName As String
    ModuleName = C_M_Tools
    
    Dim MsgBoxTitle As String
    MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName

    If Application.Version <> "14.0" And Application.Version <> "16.0" Then
        MsgBox _
            Prompt:="This macro is not compatible to this version of Office!", _
            Buttons:=vbCritical + vbOKOnly, _
            Title:=MsgBoxTitle
        Exit Sub ' Exit the subroutine
    End If

End Sub

' Sets the margins of all text boxes to 0
' The style of text boxes is hardcoded to C_S_TextBoxes
' 2025-02-11 by ms
Sub CanvaFormatTextBoxes()
    Dim MyShape As Shape
    Dim canvasitem As Object
    Dim groupitem As Object
    Dim Pole As Field
    Dim i As Integer
    Dim j As Integer
    
    Dim FileName As String
    FileName = C_F_Macros
    
    Dim ModuleName As String
    ModuleName = C_M_Tools
    
    Dim MacroName As String
    MacroName = "CanvaFormatTextBoxes"
    
    Dim MsgBoxTitle As String
    MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    Call CheckMicrosoftWordVersion(MacroName)
    
    ' When Application.ScreenUpdating is set to False, it turns off screen updating, which can significantly speed up the execution of a macro by preventing the screen from refreshing until the macro has finished running. This is particularly useful for macros that perform a lot of operations, as it reduces the time spent on rendering the screen.
    Application.ScreenUpdating = False
    Application.DisplayAlerts = wdAlertsNone
    
    ' Saves last cursor position as a temporary bookmark
    Call AddLastCursorPositionBookmark
    
    i = ActiveDocument.Shapes.count
    j = 0
    CanvaFormatTextBoxes_Form.ProgressLabel = "Finished: " & j & " out of " & i
    ' ShowModal must be set to False in the corresponding Form
    CanvaFormatTextBoxes_Form.Show vbModeless

    For Each MyShape In ActiveDocument.Shapes
        If MyShape.Type = msoAutoShape Or MyShape.Type = msoTextBox Then
               With MyShape.TextFrame
                    .MarginBottom = 0
                    .MarginLeft = 0
                    .MarginRight = 0
                    .MarginTop = 0
                    .TextRange.Select
               End With
               Selection.style = C_S_TextBoxes
        End If
        
        If MyShape.Type = msoCanvas Then
            For Each canvasitem In MyShape.CanvasItems
                If canvasitem.Type = msoGroup Then
                    For Each groupitem In canvasitem.GroupItems
                        If groupitem.Type = msoTextBox Or groupitem.Type = msoAutoShape Then
                            With groupitem.TextFrame
                                .MarginBottom = 0
                                .MarginLeft = 0
                                .MarginRight = 0
                                .MarginTop = 0
                                .TextRange.Select
                            End With
                            Selection.style = C_S_TextBoxes
                        End If
                    Next
                End If
            
                If canvasitem.Type = msoTextBox Then
                    With canvasitem.TextFrame
                        .MarginBottom = 0
                        .MarginLeft = 0
                        .MarginRight = 0
                        .MarginTop = 0
                        .TextRange.Select
                    End With
                    Selection.style = C_S_TextBoxes
                End If

                If canvasitem.Type = msoAutoShape Then
                    With canvasitem.TextFrame
                        .MarginBottom = 0
                        .MarginLeft = 0
                        .MarginRight = 0
                        .MarginTop = 0
                        .TextRange.Select
                    End With
                    Selection.style = C_S_TextBoxes
                End If
            On Error Resume Next
            Next canvasitem
        End If
        CanvaFormatTextBoxes_Form.ProgressLabel = "Finished: " & j & " out of " & i
        j = j + 1
        DoEvents
    Next MyShape
    
    Unload CanvaFormatTextBoxes_Form
    ActiveWindow.View.Type = wdPrintView
    Application.ScreenRefresh
    
    ' Goes to a place where temporary bookmark was located and removes it afterwards
    Call RemoveLastCursorPositionBookmark
    
    MsgBox _
        Prompt:="Processing is finished.", _
        Buttons:=vbInformation + vbOKOnly, _
        Title:=MsgBoxTitle
End Sub

' Numbering of comments.
' 2025-03-01 by ms and AI
Sub CommentAddNumber()
    Dim i As Long
    Dim rngComment As Range
    Dim FileName As String
    FileName = C_F_Macros
    
    Dim ModuleName As String
    ModuleName = C_M_Tools
    
    Dim MacroName As String
    MacroName = "CommentAddNumber"
    
    Dim MsgBoxTitle As String
    MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    With ActiveDocument
        For i = 1 To .Comments.count
            Set rngComment = .Comments(i).Range
            If Left(rngComment.Text, 7) = "Comment" Then
                rngComment.Text = "Comment " & i & Mid(rngComment.Text, InStr(rngComment.Text, ":"))
            Else
                rngComment.Text = "Comment " & i & ": " & rngComment.Text
            End If
        Next i
    End With
    
    ' Clear object variables
    Set rngComment = Nothing
    
    MsgBox _
        Prompt:="Numbers have been added to all the comments.", _
        Buttons:=vbInformation + vbOKOnly, _
        Title:=MsgBoxTitle
End Sub

' Deleting the comment numbers.
' 2025-03-02 by ms and AI.
Sub CommentDeleteNumber()
    Dim i As Long
    Dim rngComment As Range
    Dim commentText As String
    Dim colonPos As Long
    
    Dim FileName As String
    FileName = C_F_Macros
    
    Dim ModuleName As String
    ModuleName = C_M_Tools
    
    Dim MacroName As String
    MacroName = "CommentDeleteNumber"
    
    Dim MsgBoxTitle As String
    MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    With ActiveDocument
        For i = 1 To .Comments.count
            Set rngComment = .Comments(i).Range
            commentText = rngComment.Text
            
            ' Check if the comment starts with "Comment" followed by a number
            If Left(commentText, 7) = "Comment" Then
                colonPos = InStr(commentText, ":")
                If colonPos > 0 Then
                    ' Remove the "Comment X: " part
                    rngComment.Text = Mid(commentText, colonPos + 2)
                End If
            End If
        Next i
    End With
    
    ' Clear object variables
    Set rngComment = Nothing
    
    MsgBox _
        Prompt:="Numbers of comments have been just deleted.", _
        Buttons:=vbInformation + vbOKOnly, _
        Title:=MsgBoxTitle
End Sub

' Count number of comments added to this document per user.
' 2025-03-01 by ms and AI
Sub CommentCountByUser()
    Dim doc As Document
    Dim comment As comment
    Dim userNames As Collection
    Dim UserName As Variant
    Dim commentCount As Long
    Dim result As String
    Dim i As Long
    
    Dim FileName As String
    FileName = C_F_Macros
    
    Dim ModuleName As String
    ModuleName = C_M_Tools
    
    Dim MacroName As String
    MacroName = "CommentCountByUser"
    
    Dim MsgBoxTitle As String
    MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    Set doc = ActiveDocument
    Set userNames = New Collection
    
    ' Collect unique user names
    On Error Resume Next
    For Each comment In doc.Comments
        userNames.Add comment.Author, comment.Author
    Next comment
    On Error GoTo 0
    
    ' Count comments for each user
    result = ""
    For i = 1 To userNames.count
        UserName = userNames(i)
        commentCount = 0
        For Each comment In doc.Comments
            If comment.Author = UserName Then
                commentCount = commentCount + 1
            End If
        Next comment
        result = result & UserName & ": " & commentCount & " comments" & vbCrLf
    Next i
    
    ' Clear object variables
    Set doc = Nothing
    Set userNames = Nothing
    
    MsgBox _
        Prompt:=result, _
        Buttons:=vbInformation + vbOKOnly, _
        Title:=MsgBoxTitle
End Sub


Public Sub AddLastCursorPositionBookmark()
    ' Adds a bookmark in place where cursor is present
    Dim rng As Range
    
    If ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument Then
        Set rng = Selection.Range
        rng.Bookmarks.Add C_BM_LastCursorPosition
    End If
    
    ' Clear object variable
    Set rng = Nothing
End Sub

Public Sub RemoveLastCursorPositionBookmark()
    If ActiveDocument.Bookmarks.Exists(C_BM_LastCursorPosition) Then
        Selection.GoTo What:=wdGoToBookmark, Name:=C_BM_LastCursorPosition
        ActiveDocument.Bookmarks(C_BM_LastCursorPosition).Delete
    Else
        ActiveDocument.GoTo wdStory ' it moves the selection (or cursor) to the very beginning of the document.
    End If
End Sub

' Fixes a bug in Microsoft Word where function "ViewFieldCodes" jumps over the document body each time it is called.
' Keyboard shortcut to this macro is set or reset in Module "Shortcut".
' 2025-03-01 by ms and AI
' 2025-03-06 by ms added line with customization context.
' 2025-03-08 by ms, separated macro code from shortcut
' 2025-03-27 by ms, Range instead of Select
Sub CustomizedToggleFieldCodes()
    Call AddLastCursorPositionBookmark
    Application.Run "ViewFieldCodes" ' call built-in Microsoft Word command
    Call RemoveLastCursorPositionBookmark
End Sub

' 2025-08-19 by ms and AI
Function CaptionCheckCustomLabelsOnly() As Boolean
    Dim IfPicOrTabExists As Boolean
    Dim i As Integer
    Dim LabelName As String
    Dim BuiltInCounter As Byte
    
    IfPicOrTabExists = False
    BuiltInCounter = 0
    
    For i = 1 To CaptionLabels.count
        With CaptionLabels(i)
            If Not .BuiltIn Then
                BuiltInCounter = BuiltInCounter + 1
                LabelName = .Name
                If LabelName = C_Caption_Tab Or LabelName = C_Caption_Pic Then
                    IfPicOrTabExists = True
                End If
            End If
        End With
    Next i
    
    If BuiltInCounter = 2 And IfPicOrTabExists Then
        CaptionCheckCustomLabelsOnly = True
    Else
        CaptionCheckCustomLabelsOnly = False
    End If
End Function


' The captions aren't stored in the template, so they must be defined within a macro.
' When new document is created from the template body ("enter"), then the specific captions will be available in such document.
' When you attach this template to existing document and want captions to be moved to that document file, you need to run the macro CaptionLabelCopyFromTemplate.
' 2025-03-04 by ms and AI
Sub CapationAddCustomized()
    Dim FileName As String
    FileName = C_F_Macros
    
    Dim ModuleName As String
    ModuleName = C_M_Macros
    
    Dim MacroName As String
    MacroName = "CapationAddCustomized"
    
    Dim MsgBoxTitle As String
    MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    ' Add the new caption labels
    CaptionLabels.Add Name:=C_Caption_Tab
    CaptionLabels.Add Name:=C_Caption_Pic
    
    MsgBox _
        Prompt:="New caption labels " & C_Caption_Tab & " and " & C_Caption_Pic & " have been added to the application.", _
        Buttons:=vbInformation + vbOKOnly, _
        Title:=MsgBoxTitle
End Sub

' Delete not built-in caption labels
' 2025-04-27 by ms and AI
Sub CaptionLabelDeleteCustomized()
    Dim Caption As CaptionLabel
    Dim i As Byte
    Dim NotBuiltinCaptions As String
    Dim BuiltinCaptions As String
            
    Let i = 1
    Let NotBuiltinCaptions = ""
    For Each Caption In Application.CaptionLabels
        If Not Caption.BuiltIn Then
            NotBuiltinCaptions = NotBuiltinCaptions & i & ". " & Caption.Name & vbNewLine
            Caption.Delete
            i = i + 1
        End If
    Next Caption
    
    Let i = 1
    Let BuiltinCaptions = ""
    For Each Caption In Application.CaptionLabels
        BuiltinCaptions = BuiltinCaptions & i & ". " & Caption.Name & vbNewLine
        i = i + 1
    Next Caption

    Dim FileName As String
    FileName = C_F_Macros
    
    Dim ModuleName As String
    ModuleName = C_M_Macros
    
    Dim MacroName As String
    MacroName = "CaptionLabelDeleteCustomized"
    
    Dim MsgBoxTitle As String
    MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    MsgBox _
        Prompt:="Deleted caption labels: " & vbNewLine & NotBuiltinCaptions & vbNewLine & vbNewLine & _
                "Built in caption labels: " & vbNewLine & BuiltinCaptions, _
        Buttons:=vbInformation + vbOKOnly, _
        Title:=MsgBoxTitle
End Sub


' The specific captions are stored within the file, not within a template. But they can be copied from the template to that file.
' This macro copies all the captions from the template file to specific file.
' 2025-03-04 by ms and AI
Sub CaptionShow()
    Dim Label As CaptionLabel
    Dim Info As String
    Dim i As Byte
    
    Dim FileName As String
    FileName = C_F_Macros
    
    Dim ModuleName As String
    ModuleName = C_M_Tools
    
    Dim MacroName As String
    MacroName = "CaptionShow"
    
    Dim MsgBoxTitle As String
    MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    i = 0
    For Each Label In Application.CaptionLabels
        i = i + 1
        Info = Info & i & ". " & Label.Name & " : " & IIf(Label.BuiltIn, "Built-in", "Custom") & vbNewLine
    Next Label
    
    MsgBox _
        Prompt:="Caption labels in this Microsoft Word: " & vbNewLine & vbNewLine & Info, _
        Buttons:=vbInformation, _
        Title:=MsgBoxTitle
End Sub

' Delete all user bookmarks. Don't delete hidden bookmarks.
' 2025-03-08 by ms and AI
Sub DeleteAllUserBookmarks()
    Dim Bm As bookmark
    Dim bmCount As Integer
    
    Dim FileName As String
    FileName = C_F_Macros
    
    Dim ModuleName As String
    ModuleName = C_M_Tools
    
    Dim MacroName As String
    MacroName = "DeleteAllUserBookmarks"
    
    Dim MsgBoxTitle As String
    MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    ' Ask user for permission
    Dim UserDecision As VbMsgBoxResult
    Beep
    UserDecision = MsgBox( _
        Prompt:="This action will delete all user-added bookmarks in the currently opened document. This action cannot be undone." & vbNewLine & vbNewLine & _
            "Do you want to proceed?", _
        Buttons:=vbYesNo + vbQuestion + vbDefaultButton2, _
        Title:=MsgBoxTitle)
    
    ' If user answers No, exit sub
    If UserDecision = vbNo Then
        Exit Sub
    End If
    
    ' Count the number of user bookmarks
    bmCount = 0
    For Each Bm In ActiveDocument.Bookmarks
        ' Check if the bookmark is not hidden
        If Not Bm.Range.BookmarkID Like "\*" Then
            bmCount = bmCount + 1
        End If
    Next Bm
    
    ' Delete all user bookmarks
    For Each Bm In ActiveDocument.Bookmarks
        ' Check if the bookmark is not hidden
        If Not Bm.Range.BookmarkID Like "\*" Then
            Bm.Delete
        End If
    Next Bm
    
    ' Display the summary in a message box
    MsgBox _
        Prompt:="Processing complete." & vbNewLine & vbNewLine & "Deleted " & bmCount & " user bookmarks.", _
        Buttons:=vbInformation, _
        Title:=MsgBoxTitle
End Sub

' Jumps to the next paragraph type list.
' Shortcuts to this sub are set in the module "Shortcuts".
' 2025-03-08 by ms and AI
Sub JumpToNextList()
    Dim para As Paragraph
    Dim found As Boolean
    Dim startPos As Long
        
    Dim FileName As String
    FileName = C_F_Macros
    
    Dim ModuleName As String
    ModuleName = C_M_Tools
    
    Dim MacroName As String
    MacroName = "JumpToNextList"
    
    Dim MsgBoxTitle As String
    MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
        
    ' Get the starting position of the current selection
    startPos = Selection.Start
    found = False
    
    ' Loop through all paragraphs in the document starting from the current selection
    For Each para In ActiveDocument.Paragraphs
        If para.Range.Start > startPos Then
            ' Check if the paragraph is formatted as a list
            If para.Range.ListFormat.ListType <> wdListNoNumbering Then
                ' Move the selection to the start of the paragraph
                para.Range.Select
                found = True
                Exit For
            End If
        End If
    Next para
    
    ' Inform the user if no list was found
    If Not found Then
        MsgBox _
            Prompt:="No next list found in the document.", _
            Buttons:=vbInformation, _
            Title:=MsgBoxTitle
    End If
End Sub

' Jumps to the next paragraph type table.
' Shortcuts to this sub are set in the module "Shortcuts".
' 2025-03-08 by ms and AI
Sub JumpToNextTable()
    Dim tbl As Table
    Dim found As Boolean
    Dim startPos As Long
    
    Dim FileName As String
    FileName = C_F_Macros
    
    Dim ModuleName As String
    ModuleName = C_M_Tools
    
    Dim MacroName As String
    MacroName = "JumpToNextTable"
    
    Dim MsgBoxTitle As String
    MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    ' Get the starting position of the current selection
    startPos = Selection.Start
    found = False
    
    ' Loop through all tables in the document
    For Each tbl In ActiveDocument.Tables
        ' Check if the table is after the current selection
        If tbl.Range.Start > startPos Then
            ' Move the selection to the start of the table
            tbl.Range.Select
            found = True
            Exit For
        End If
    Next tbl
    
    ' Inform the user if no table was found
    If Not found Then
        MsgBox _
            Prompt:="No next table found in the document.", _
            Buttons:=vbInformation, _
            Title:=MsgBoxTitle
    End If
End Sub

' 2025-04-11 by ms and AI
Sub JumpToNextCanvas()
    Dim shp As Shape
    Dim found As Boolean
    Dim startPos As Long
    
    Dim FileName As String
    FileName = C_F_Macros
    
    Dim ModuleName As String
    ModuleName = C_M_Tools
    
    Dim MacroName As String
    MacroName = "JumpToNextCanvas"
    
    Dim MsgBoxTitle As String
    MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    ' Get the starting position of the current selection
    startPos = Selection.Start
    found = False
    
    ' Loop through all shapes in the document
    For Each shp In ActiveDocument.Shapes
        ' Check if the shape is a canvas and is after the current selection
        If shp.Type = msoCanvas And shp.Anchor.Start > startPos Then
            ' Move the selection to the start of the canvas
            shp.Select
            found = True
            Exit For
        End If
    Next shp
    
    ' Inform the user if no canvas was found
    If Not found Then
        MsgBox _
            Prompt:="No next canvas found in the document.", _
            Buttons:=vbInformation, _
            Title:=MsgBoxTitle
    End If
End Sub

' If new section starts with a heading style and this style has a gap, that gap is respected by Microsoft Word. I consider it a bug.
' This macro removes this bug by changing locally that paragraph distance to previous paragraph to 0.
' 2025-03-13 reworked by ms and AI
Sub ParDistAtNewSectionReduce()
    Dim Message As String
    Dim sec As Section
    Dim para As Paragraph
    Dim CounterFound As Integer
    Dim CounterChanged As Integer
    Dim CounterUnchanged As Integer
    Dim BookmarkName As String
    
    Dim FileName As String
    FileName = C_F_Macros
    
    Dim ModuleName As String
    ModuleName = C_M_Tools
    
    Dim MacroName As String
    MacroName = "ParDistAtNewSectionReduce"
    
    Dim MsgBoxTitle As String
    MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    ActiveWindow.View.Type = wdPrintView
    Application.ScreenUpdating = True
    
    ' Saves last cursor position as a temporary bookmark
    Call AddLastCursorPositionBookmark

    CounterFound = 0
    CounterChanged = 0
    CounterUnchanged = 0
    
    For Each sec In ActiveDocument.Sections
        Set para = sec.Range.Paragraphs.First
        
        If para.style Like "Heading*" Then
            CounterFound = CounterFound + 1
            DoEvents    ' Force a screen refresh
            para.Range.HighlightColorIndex = wdYellow
            DoEvents    ' Force a screen refresh
            Dim UserDecision As VbMsgBoxResult
            Beep
            UserDecision = MsgBox( _
                Prompt:="Do you want to reduce distance at the beginning of a new section to 0?", _
                Buttons:=vbYesNo + vbQuestion + vbDefaultButton1, _
                Title:=MsgBoxTitle)
        
            If UserDecision = vbYes Then
                para.SpaceBefore = 0
                CounterChanged = CounterChanged + 1
                BookmarkName = C_BM_ReducedDistance & CounterChanged
                ActiveDocument.Bookmarks.Add Name:=BookmarkName, Range:=para.Range.Characters.First
            Else
                CounterUnchanged = CounterUnchanged + 1
            End If
            para.Range.HighlightColorIndex = wdNoHighlight
        End If
    Next sec

    Message = "Number of paragraphs found: " & CounterFound & vbCrLf & _
              "Number of paragraphs changed: " & CounterChanged & vbCrLf & _
              "Number of paragraphs not changed: " & CounterUnchanged & "."
    MsgBox _
        Prompt:=Message, _
        Buttons:=vbInformation, _
        Title:=MsgBoxTitle

    Application.ScreenUpdating = False

    ' Goes to a place where temporary bookmark was located and removes it afterwards
    Call RemoveLastCursorPositionBookmark
    ActiveWindow.View.Type = wdPrintView
    
    ' Clear object variable
    Set para = Nothing

End Sub

' Find previously set bookmarks and ask user if restore original formatting.
' 2025-03-13 by ms and AI
Sub ParDistAtNewSectionRestore()
    Dim Bm As bookmark
    Dim CounterRestored As Integer
    Dim CounterUnchanged As Integer
    
    Dim FileName As String
    FileName = C_F_Macros
    
    Dim ModuleName As String
    ModuleName = C_M_Tools
    
    Dim MacroName As String
    MacroName = "ParDistAtNewSectionRestore"
    
    Dim MsgBoxTitle As String
    MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    CounterRestored = 0
    CounterUnchanged = 0
    
    ActiveWindow.View.Type = wdPrintView
    Application.ScreenUpdating = True
    
    ' Saves last cursor position as a temporary bookmark
    Call AddLastCursorPositionBookmark
    
    For Each Bm In ActiveDocument.Bookmarks
        If Bm.Name Like C_BM_ReducedDistance & "*" Then
            DoEvents    ' Force a screen refresh
            Bm.Range.Paragraphs(1).Range.HighlightColorIndex = wdYellow
            DoEvents    ' Force a screen refresh
            Dim UserDecision As VbMsgBoxResult
            Beep
            UserDecision = MsgBox( _
                Prompt:="Do you want to restore formatting for the paragraph with bookmark " & Bm.Name & "?", _
                Buttons:=vbYesNo + vbQuestion + vbDefaultButton1, _
                Title:=MsgBoxTitle)
            
            If UserDecision = vbYes Then
                Bm.Range.ParagraphFormat.SpaceBefore = Bm.Range.Paragraphs(1).style.ParagraphFormat.SpaceBefore
                Bm.Range.Paragraphs(1).Range.HighlightColorIndex = wdNoHighlight
                Bm.Delete
                CounterRestored = CounterRestored + 1
            Else
                CounterUnchanged = CounterUnchanged + 1
                Bm.Range.Paragraphs(1).Range.HighlightColorIndex = wdNoHighlight
            End If
            
        End If
    Next Bm
    
    ' Goes to a place where temporary bookmark was located and removes it afterwards
    Call RemoveLastCursorPositionBookmark
    Application.ScreenUpdating = False
    MsgBox _
        Prompt:="Number of paragraphs restored to default formatting: " & CounterRestored & vbCrLf & _
            "Number of paragraphs left unchanged: " & CounterUnchanged, _
        Buttons:=vbInformation, _
        Title:=MsgBoxTitle
End Sub

' Only check how many paragraphs at the beginning of each section are styled as headings.
' 2025-03-13 by ms and AI
Sub ParDistAtNewSectionCheck()
    Dim sec As Section
    Dim para As Paragraph
    Dim CounterFound As Integer
    
    Dim FileName As String
    FileName = C_F_Macros
    
    Dim ModuleName As String
    ModuleName = C_M_Tools
    
    Dim MacroName As String
    MacroName = "ParDistAtNewSectionCheck"
    
    Dim MsgBoxTitle As String
    MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    ActiveWindow.View.Type = wdPrintView
    Application.ScreenUpdating = False
    
    CounterFound = 0
    
    For Each sec In ActiveDocument.Sections
        Set para = sec.Range.Paragraphs.First
        
        If para.style Like "Heading*" Then
            CounterFound = CounterFound + 1
        End If
    Next sec

    ' Clear object variable
    Set para = Nothing

    MsgBox _
        Prompt:="Number of paragraphs meeting the criterion: " & CounterFound, _
        Buttons:=vbInformation, _
        Title:=MsgBoxTitle

    Application.ScreenUpdating = True

End Sub

' 2025-08-03 by ms
' 2025-11-16 by ms
' Saves file in the default directory with specific PDF settings
Sub SaveDocumentAsPDFWithSettings()
    Dim FilePath As String
    Dim DefaultPath As String
    
    Dim FileName As String
    FileName = C_F_Macros
    
    Dim ModuleName As String
    ModuleName = C_M_Tools
    
    Dim MacroName As String
    MacroName = "SaveDocumentAsPDFWithSettings"
    
    Dim MsgBoxTitle As String
    MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    ' Define the file path and name for the PDF file
    ' This example uses the document's name and saves the PDF in the default directory.
    DefaultPath = Options.DefaultFilePath(wdDocumentsPath)
    Dim BaseName As String
    ' Remove extension from current document name
    ' InStrRev in VBA is a string function that searches for a substring within another string, starting from the end (right side) of the string and moving backward. It returns the position of the first occurrence found when searching from the right.
    ' Left in VBA is a string function that returns a specified number of characters from the beginning (left side) of a string.
    If InStrRev(ActiveDocument.Name, ".") > 0 Then
        BaseName = Left(ActiveDocument.Name, InStrRev(ActiveDocument.Name, ".") - 1)
    Else
        BaseName = ActiveDocument.Name
    End If
    ' Construct full path with .pdf extension
    FilePath = DefaultPath & "\" & BaseName & ".pdf"
    
    ' Check if the document is saved (it needs a file path)
    If FilePath = "" Then
        MsgBox _
            Prompt:="Please check the Options.DefaultFilePath before exporting it in PDF format." & vbNewLine & vbNewLine & _
                "Exiting.", _
            Buttons:=vbExclamation + vbOKOnly, _
            Title:=MsgBoxTitle
        Exit Sub
    End If
    
    ' Export the document as PDF with specified settings
    ActiveDocument.ExportAsFixedFormat _
        OutputFileName:=FilePath, _
        ExportFormat:=wdExportFormatPDF, _
        OpenAfterExport:=False, _
        OptimizeFor:=wdExportOptimizeForPrint, _
        Range:=wdExportAllDocument, _
        From:=1, _
        To:=1, _
        item:=wdExportDocumentContent, _
        IncludeDocProps:=False, _
        KeepIRM:=True, _
        CreateBookmarks:=wdExportCreateHeadingBookmarks, _
        DocStructureTags:=False, _
        BitmapMissingFonts:=True, _
        UseISO19005_1:=False
    
    MsgBox _
        Prompt:="Document exported as PDF to:" & vbNewLine & vbNewLine & _
            FilePath, _
        Buttons:=vbInformation + vbOKOnly, _
        Title:=MsgBoxTitle
End Sub

' Management of customer properties:
' 1. Check if required properties exist, add them if they don't
' 2. Set specific values for certain properties
' 3. Remove all other custom properties
' 2025-03-14 by ms
' 2025-07-18 by ms
Sub DocPropertiesUpdate()
    Dim doc As Document
    Dim Prop As Variant
    Dim PropToDelete As Boolean
    Dim ReqProp As Variant
    
    Dim FileName As String
    FileName = C_F_Macros
    
    Dim ModuleName As String
    ModuleName = C_M_Tools
    
    Dim MacroName As String
    MacroName = "DocPropertiesUpdate"
    
    Dim MsgBoxTitle As String
    MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    ' Set the document to the currently opened document
    Set doc = ActiveDocument
    
    ' List of required properties
    Dim requiredProperties As Variant
    requiredProperties = Array( _
        C_CPN_1, _
        C_CPN_2, _
        C_CPN_3, _
        C_CPN_4, _
        C_CPN_5, _
        C_CPN_6, _
        C_CPN_7, _
        C_CPN_8, _
        C_CPN_9, _
        C_CPN_10 _
        )
    
    ' Check if required properties exist, add them if they don't
    For Each Prop In requiredProperties
        On Error Resume Next
        If doc.CustomDocumentProperties(Prop).Name = "" Then
            doc.CustomDocumentProperties.Add Name:=Prop, LinkToContent:=False, Value:="", Type:=msoPropertyTypeString
        End If
        On Error GoTo 0
    Next Prop
    
    ' Set specific values for certain properties
    doc.CustomDocumentProperties(C_CPN_1).Value = ""
    doc.CustomDocumentProperties(C_CPN_2).Value = C_CPV_2
    doc.CustomDocumentProperties(C_CPN_3).Value = C_CPV_3
    doc.CustomDocumentProperties(C_CPN_4).Value = ""
    doc.CustomDocumentProperties(C_CPN_5).Value = ""
    doc.CustomDocumentProperties(C_CPN_6).Value = ""
    doc.CustomDocumentProperties(C_CPN_7).Value = C_CPV_7
    doc.CustomDocumentProperties(C_CPN_8).Value = ""
    doc.CustomDocumentProperties(C_CPN_9).Value = ""
    doc.CustomDocumentProperties(C_CPN_10).Value = ""
    
    ' Remove all other custom properties
    For Each Prop In doc.CustomDocumentProperties
        PropToDelete = True
        For Each ReqProp In requiredProperties
            If Prop.Name = ReqProp Then
                PropToDelete = False
                Exit For
            End If
        Next ReqProp
        
        If PropToDelete Then
            Prop.Delete
        End If
    Next Prop
    
    Dim InfoForUser As String
    InfoForUser = ""
    InfoForUser = InfoForUser & C_CPN_1 & ": " & C_CPV_1 & vbNewLine
    InfoForUser = InfoForUser & C_CPN_2 & ": " & C_CPV_2 & vbNewLine
    InfoForUser = InfoForUser & C_CPN_3 & ": " & C_CPV_3 & vbNewLine
    InfoForUser = InfoForUser & C_CPN_4 & ": " & C_CPV_4 & vbNewLine
    InfoForUser = InfoForUser & C_CPN_5 & ": " & C_CPV_5 & vbNewLine
    InfoForUser = InfoForUser & C_CPN_6 & ": " & C_CPV_6 & vbNewLine
    InfoForUser = InfoForUser & C_CPN_7 & ": " & C_CPV_7 & vbNewLine
    InfoForUser = InfoForUser & C_CPN_8 & ": " & C_CPV_8 & vbNewLine
    InfoForUser = InfoForUser & C_CPN_9 & ": " & C_CPV_9 & vbNewLine
    InfoForUser = InfoForUser & C_CPN_10 & ": " & C_CPV_10 & vbNewLine
        
    MsgBox _
        Prompt:="Document custom properties updated successfully in:" & vbNewLine & _
            ActiveDocument.Name & vbNewLine & vbNewLine & _
            InfoForUser, _
        Buttons:=vbInformation, _
        Title:=MsgBoxTitle

    ' Clear the object variables
    Set doc = Nothing
End Sub

' Set Hyphenation in currently open document
' 2025-03-23 by ms
Sub SetHyphenation()
    Dim doc As Document
    Dim para As Paragraph
    
    Dim FileName As String
    FileName = C_F_Macros
    
    Dim ModuleName As String
    ModuleName = C_M_Tools
    
    Dim MacroName As String
    MacroName = "SetHyphenation"
    
    Dim MsgBoxTitle As String
    MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    ' Initialize object variables
    Set doc = ActiveDocument
    
    ' Enable hyphenation for the entire document
    doc.Range.LanguageId = wdEnglishUS ' Set the language to English (US) or another language as needed
    doc.Range.NoProofing = False ' Ensure proofing is enabled
    ' Set hyphenation options
    doc.HyphenationZone = 17 ' Set the hyphenation zone to 0.6 cm (approximately 17 points)
    doc.ConsecutiveHyphensLimit = 2 ' Set the limit for consecutive hyphens
    
    ' Loop through each paragraph in the document and disable hyphenation
    For Each para In doc.Paragraphs
        para.Range.ParagraphFormat.Hyphenation = True
    Next para
    
    ' Enable automatic hyphenation for the active document
    doc.AutoHyphenation = True
    doc.HyphenateCaps = True
    
    ' Clear object variables
    Set doc = Nothing
    
    MsgBox _
        Prompt:="Hyphenation was enabled in the active document.", _
        Buttons:=vbInformation + vbOKOnly, _
        Title:=MsgBoxTitle
End Sub

' 2025-03-23 by ms and AI
Sub ResetHyphenation()
    Dim para As Paragraph
    Dim doc As Document
    
    Dim FileName As String
    FileName = C_F_Macros
    
    Dim ModuleName As String
    ModuleName = C_M_Tools
    
    Dim MacroName As String
    MacroName = "ResetHyphenation"
    
    Dim MsgBoxTitle As String
    MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    Set doc = ActiveDocument
    
    ' Loop through each paragraph in the document and disable hyphenation
    For Each para In doc.Paragraphs
        para.Range.ParagraphFormat.Hyphenation = False
    Next para
    
    ' Clear object data
    Set doc = Nothing
    
    MsgBox _
        Prompt:="Hyphenation was disabled.", _
        Buttons:=vbInformation + vbOKOnly, _
        Title:=MsgBoxTitle
End Sub

' 2025-03-27 by ms and AI
' Applies the style "CharBold ms" to the selected content if not already applied
Sub Bold()
    Dim CharBoldExists As Boolean
    Dim var As Variable
    
    CharBoldExists = False
    On Error Resume Next
    ' Check if the document variable "CharBold_ms" exists
    For Each var In ActiveDocument.Variables
        If var.Name = C_S_Bold Then
            CharBoldExists = True
            Exit For
        End If
    Next var
    
    ' If the document variable doesn't exist or is set to False, run the "Bold" command
    If Not CharBoldExists Or ActiveDocument.Variables(C_S_Bold).Value = False Then
        Selection.font.Bold = wdToggle
        Exit Sub
    End If
    
    ' If the document variable exists and is set to True, proceed with applying styles
    On Error GoTo 0
    If ActiveDocument.Variables(C_S_Bold).Value = True Then
        If Selection.Type <> wdNoSelection Then
            If Selection.style = C_S_Bold Then
                Selection.style = C_S_CharDefault
            Else
                Selection.style = C_S_Bold
            End If
        End If
    End If
End Sub

' 2025-03-27 by ms and AI
Sub Underline()
    Dim CharUnderlineExists As Boolean
    Dim var As Variable
    
    On Error Resume Next
    ' Check if the document variable "CharUnderline_ms" exists
    CharUnderlineExists = False
    For Each var In ActiveDocument.Variables
        If var.Name = C_S_Underline Then
            CharUnderlineExists = True
            Exit For
        End If
    Next var
    
    ' If the document variable doesn't exist or is set to False, run the "Underline" command
    If Not CharUnderlineExists Or ActiveDocument.Variables(C_S_Underline).Value = False Then
        Selection.font.Underline = wdToggle
        Exit Sub
    End If
    
    ' If the document variable exists and is set to True, proceed with applying styles
    On Error GoTo 0
    If ActiveDocument.Variables(C_S_Underline).Value = True Then
        If Selection.Type <> wdNoSelection Then
            If Selection.style = C_S_Underline Then
                Selection.style = C_S_CharDefault
            Else
                Selection.style = C_S_Underline
            End If
        End If
    End If
End Sub

' 2025-03-27 by ms and AI
Sub Italic()
    Dim CharItalicExists As Boolean
    Dim var As Variable
    
    On Error Resume Next
    ' Check if the document variable "CharUnderline_ms" exists
    CharItalicExists = False
    For Each var In ActiveDocument.Variables
        If var.Name = C_S_Italic Then
            CharItalicExists = True
            Exit For
        End If
    Next var
    
    ' If the document variable doesn't exist or is set to False, run the "Underline" command
    If Not CharItalicExists Or ActiveDocument.Variables(C_S_Italic).Value = False Then
        Selection.font.Italic = wdToggle
        Exit Sub
    End If
    
    ' If the document variable exists and is set to True, proceed with applying styles
    On Error GoTo 0
    If ActiveDocument.Variables(C_S_Italic).Value = True Then
        If Selection.Type <> wdNoSelection Then
            If Selection.style = C_S_Italic Then
                Selection.style = C_S_CharDefault
            Else
                Selection.style = C_S_Italic
            End If
        End If
    End If
End Sub

' Check if the document variable "CharStrikethrough_ms" exists
' If the document variable doesn't exist or is set to False, run the "Strikethrough" command
' If the document variable exists and is set to True, proceed with applying styles
' 2025-03-27 by ms and AI
Sub Strikethrough()
    Dim CharStrikethrough As Boolean
    Dim var As Variable
    
    On Error Resume Next
    ' Check if the document variable "CharStrikethrough_ms" exists
    CharStrikethrough = False
    For Each var In ActiveDocument.Variables
        If var.Name = C_S_CharCrossout Then
            CharStrikethrough = True
            Exit For
        End If
    Next var
    
    ' If the document variable doesn't exist or is set to False, run the "Strikethrough" command
    If Not CharStrikethrough Or ActiveDocument.Variables(C_S_CharCrossout).Value = False Then
        Selection.font.Strikethrough = wdToggle
        Exit Sub
    End If
    
    ' If the document variable exists and is set to True, proceed with applying styles
    On Error GoTo 0
    If ActiveDocument.Variables(C_S_CharCrossout).Value = True Then
        If Selection.Type <> wdNoSelection Then
            If Selection.style = C_S_CharCrossout Then
                Selection.style = C_S_CharDefault
            Else
                Selection.style = C_S_CharCrossout
            End If
        End If
    End If
End Sub

' Dirty tricks, as described here: https://answers.microsoft.com/en-us/msoffice/forum/all/how-to-reset-a-border-to-microsoft-word-canvas/b5ebbc0c-a304-419b-bf75-2cbb893e5a99?rtAction=1743443617273
' 2025-03-27 by ms and AI
Sub CanvaToggleBorder()
    Dim MyCanvas As Shape
    Dim RngAnchor As Range
    
    Dim FileName As String
    FileName = C_F_Macros
    
    Dim ModuleName As String
    ModuleName = C_M_Tools
    
    Dim MacroName As String
    MacroName = "CanvaToggleBorder"
    
    Dim MsgBoxTitle As String
    MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    ' Loop through all shapes in the active document
    For Each MyCanvas In ActiveDocument.Shapes
        ' Check if the shape is a canvas
        If MyCanvas.Type = msoCanvas Then
            ' Highlight the paragraph containing the canvas in yellow color
            DoEvents    ' Force a screen refresh
            Set RngAnchor = MyCanvas.Anchor.Paragraphs(1).Range
            RngAnchor.HighlightColorIndex = wdYellow
            
            ' Ask user if they wish to switch line color of the canvas
            Dim UserDecision As VbMsgBoxResult
            Beep
            UserDecision = MsgBox( _
                Prompt:="Do you wish to switch the border line color of the canvas?", _
                Buttons:=vbYesNo + vbQuestion, _
                Title:=MsgBoxTitle)
            
            If UserDecision = vbYes Then
                ' If user answers yes and the current canvas line has set the line color, switch it off
                Call ToggleCanvaBrightness(MyCanvas)
            End If
            RngAnchor.HighlightColorIndex = wdNoHighlight
        End If
    Next
    
    ' Clear object data
    Set RngAnchor = Nothing
End Sub

Private Sub ToggleCanvaBrightness(ByVal Canvas As Shape)
    With Canvas.Line
        With .ForeColor
            If .Parent.Weight = 0 Then
                .Brightness = 0 ' neutral
                .SchemeColor = msoThemeColorAccent1
                .Parent.Weight = 1.5
            Else
                .Brightness = 1 ' lightest
                .Parent.Weight = 0
            End If
        End With
    End With
End Sub

' Insert SVN commit number to Document Properties.
' 2025-04-01 by ms
Sub InsertSVNCommitNumber()
    Dim CommitFilePath As String
    Dim CommitNumber As String
    Dim CommitFile As Integer
    Dim doc As Document
    Dim rng As Range
    Const CommitFilename As String = "next_commit.txt"
    Const DocumentVariableName As String = "DV_CommitFilePath"
    Const DocCP_SVN_Revision As String = "ms_SVN_Revision" ' Document Custom Property
    
    Dim MyCanvas As Shape
    Dim RngAnchor As Range
    
    Dim FileName As String
    FileName = C_F_Macros
    
    Dim ModuleName As String
    ModuleName = C_M_Tools
    
    Dim MacroName As String
    MacroName = "InsertSVNCommitNumber"
    
    Dim MsgBoxTitle As String
    MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    ' Initialize commitFilePath with Document Variable DV_CommitFilePath
    On Error Resume Next
    CommitFilePath = ActiveDocument.Variables(DocumentVariableName).Value
    On Error GoTo 0
    
    ' Check if DV_CommitFilePath is empty
    If CommitFilePath = "" Then
        ' Display dialog for user to select folder containing next_commit.txt
        With Application.FileDialog(msoFileDialogFolderPicker)
            .Title = "Select folder containing " & CommitFilename
            If .Show = -1 Then
                CommitFilePath = .SelectedItems(1) & "\" & CommitFilename
                ActiveDocument.Variables(DocumentVariableName).Value = CommitFilePath
            Else
                MsgBox _
                    Prompt:="No folder selected. Operation cancelled.", _
                    Buttons:=vbInformation, _
                    Title:=MsgBoxTitle
                Exit Sub
            End If
        End With
    Else
        ' Display MsgBox with specified DV_CommitFilePath and ask if user wishes to change it
        Dim UserDecision As VbMsgBoxResult
        Beep
        UserDecision = MsgBox( _
            Prompt:="Current commit file path: " & CommitFilePath & vbCrLf & _
                "Do you wish to change it?", _
            Buttons:=vbYesNo + vbQuestion, _
            Title:=MsgBoxTitle)
        If UserDecision = vbYes Then
            ' Display dialog for user to select folder containing next_commit.txt
            With Application.FileDialog(msoFileDialogFolderPicker)
                .Title = "Select folder containing " & CommitFilename
                If .Show = -1 Then
                    CommitFilePath = .SelectedItems(1) & "\" & CommitFilename
                    ActiveDocument.Variables(DocumentVariableName).Value = CommitFilePath
                Else
                    MsgBox _
                        Prompt:="No folder selected. Operation cancelled.", _
                        Buttons:=vbInformation, _
                        Title:=MsgBoxTitle
                    Exit Sub
                End If
            End With
        End If
    End If
    
    ' Open the text file and read the commit number
    CommitFile = FreeFile
    Open CommitFilePath For Input As CommitFile
    Input #CommitFile, CommitNumber
    Close CommitFile
    

    ' Store the commit number in a document custom property named ms_SVN_Revision
    Set doc = ActiveDocument
    doc.CustomDocumentProperties(DocCP_SVN_Revision).Value = CommitNumber
    Call UpdateAllFields    ' module: Validation
    
    ' Clear object data
    Set doc = Nothing
    
    MsgBox _
        Prompt:="Document custom property '" & DocCP_SVN_Revision & "' was updated with SVN revision no. " & CommitNumber & "." & vbNewLine & _
            "All fields in document have been updated.", _
        Buttons:=vbInformation, _
        Title:=MsgBoxTitle
End Sub

' Restore customized Microsoft Word options.
' 2025-08-17 by ms
Sub WordOptionsRestore()
    ActiveWindow.View.FieldShading = wdFieldShadingWhenSelected
    ActiveWindow.View.ShowDrawings = msoTrue
    ActiveWindow.View.ShowBookmarks = msoFalse
    ActiveWindow.View.ShowTextBoundaries = msoFalse
    ActiveWindow.View.ShowCropMarks = msoFalse
    ActiveWindow.StyleAreaWidth = 0
    ActiveDocument.Compatibility(wdSuppressBottomSpacing) = msoFalse
    ActiveDocument.Compatibility(wdSuppressTopSpacing) = msoFalse
    ActiveWindow.View.PageMovementType = wdVertical
    
    Dim FileName As String
    FileName = C_F_Macros
    
    Dim ModuleName As String
    ModuleName = C_M_Tools
    
    Dim MacroName As String
    MacroName = "WordOptionsRestore"
    
    Dim MsgBoxTitle As String
    MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    MsgBox _
        Prompt:="The following Microsoft Word properties have been set to default:" & vbNewLine & vbNewLine & _
            "ActiveWindow.View.FieldShading: " & ActiveWindow.View.FieldShading & vbNewLine & _
            "ActiveWindow.View.ShowDrawings: " & ActiveWindow.View.ShowDrawings & vbNewLine & _
            "ActiveWindow.View.ShowBookmarks: " & ActiveWindow.View.ShowBookmarks & vbNewLine & _
            "ActiveWindow.View.ShowTextBoundaries: " & ActiveWindow.View.ShowTextBoundaries & vbNewLine & _
            "ActiveWindow.View.ShowCropMarks: " & ActiveWindow.View.ShowCropMarks & vbNewLine & _
            "ActiveWindow.StyleAreaWidth: " & ActiveWindow.StyleAreaWidth & vbNewLine & _
            "ActiveDocument.Compatibility(wdSuppressBottomSpacing): " & ActiveDocument.Compatibility(wdSuppressBottomSpacing) & vbNewLine & _
            "ActiveDocument.Compatibility(wdSuppressTopSpacing): " & ActiveDocument.Compatibility(wdSuppressTopSpacing) & vbNewLine & _
            "ActiveWindow.View.PageMovementType: " & ActiveWindow.View.PageMovementType, _
        Buttons:=vbInformation, _
        Title:=MsgBoxTitle
    
End Sub

' If returns true, then all options were already set as expected.
' 2025-08-14 by ms
' 2025-10-02 by ms CheckCustomizeWordOptions -> WordOptionsSetAsExpected
Private Function WordOptionsSetAsExpected() As Boolean
    WordOptionsSetAsExpected = True
    
    Dim FileName As String
    FileName = C_F_Macros
    
    Dim ModuleName As String
    ModuleName = C_M_Tools
    
    Dim MacroName As String
    MacroName = "WordOptionsSetAsExpected"
    
    Dim MsgBoxTitle As String
    MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
       
    If ActiveWindow.View.FieldShading <> wdFieldShadingAlways Then
        WordOptionsSetAsExpected = False
        MsgBox _
            Prompt:="ActiveWindow.View.FieldShading <> wdFieldShadingAlways", _
            Buttons:=vbExclamation + vbOKOnly, _
            Title:=MsgBoxTitle
    End If
    
    If Not ActiveWindow.View.ShowDrawings = True Then
        WordOptionsSetAsExpected = False
        MsgBox _
            Prompt:="Not ActiveWindow.View.ShowDrawings = True", _
            Buttons:=vbExclamation + vbOKOnly, _
            Title:=MsgBoxTitle
    End If
    
    If Not ActiveWindow.View.ShowBookmarks = True Then
        WordOptionsSetAsExpected = False
        MsgBox _
            Prompt:="Not ActiveWindow.View.ShowBookmarks = True", _
            Buttons:=vbExclamation + vbOKOnly, _
            Title:=MsgBoxTitle
    End If
    
    If Not ActiveWindow.View.ShowTextBoundaries = True Then
        WordOptionsSetAsExpected = False
        MsgBox _
            Prompt:="Not ActiveWindow.View.ShowTextBoundaries = True", _
            Buttons:=vbExclamation + vbOKOnly, _
            Title:=MsgBoxTitle
    End If
    
    If Not ActiveWindow.View.ShowCropMarks = True Then
        WordOptionsSetAsExpected = False
        MsgBox _
            Prompt:="Not ActiveWindow.View.ShowCropMarks = True", _
            Buttons:=vbExclamation + vbOKOnly, _
            Title:=MsgBoxTitle
    End If
    
    ' Int truncates the decimal
    If ActiveWindow.StyleAreaWidth <> Int(CentimetersToPoints(5.3)) Then
        WordOptionsSetAsExpected = False
        MsgBox _
            Prompt:="ActiveWindow.StyleAreaWidth <> Int(CentimetersToPoints(5.3))", _
            Buttons:=vbExclamation + vbOKOnly, _
            Title:=MsgBoxTitle
    End If
    
    If ActiveDocument.Compatibility(wdSuppressBottomSpacing) Then
        WordOptionsSetAsExpected = False
        MsgBox _
            Prompt:="Not ActiveDocument.Compatibility(wdSuppressBottomSpacing)", _
            Buttons:=vbExclamation + vbOKOnly, _
            Title:=MsgBoxTitle
    End If
    
    If ActiveDocument.Compatibility(wdSuppressTopSpacing) Then
        WordOptionsSetAsExpected = False
        MsgBox _
            Prompt:="Not ActiveDocument.Compatibility(wdSuppressTopSpacing)", _
            Buttons:=vbExclamation + vbOKOnly, _
            Title:=MsgBoxTitle
    End If
    
    With ActiveWindow.View
        If .Type <> wdPrintView Then
            .Type = wdPrintView
        End If
    
        If .PageMovementType <> wdSideToSide Then
            WordOptionsSetAsExpected = False
            MsgBox _
                Prompt:="ActiveWindow.View.PageMovementType <> wdSideToSide", _
                Buttons:=vbExclamation + vbOKOnly, _
                Title:=MsgBoxTitle
        End If
    End With
    
End Function

' Microsoft Word customized settings
' 2025-04-02 by ms and AI
Sub WordOptionsCustomize()
    Dim FileName As String
    FileName = C_F_Macros
    
    Dim ModuleName As String
    ModuleName = C_M_Tools
    
    Dim MacroName As String
    MacroName = "WordOptionsCustomize"
    
    Dim MsgBoxTitle As String
    MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    ' File -> Options -> Advanced -> Editing options, Default paragraph style: "ParNormal ms". This setting refers to the default paragraph style defined in the Normal template (Normal.dotm).
    ' Se even if "ParNormal ms" is used in the current document, Word still considers "Normal" as the defaultstyle globally unless you modify the Normal.dotm template.
    ' In other words this setting is irrelevant. There is no easy way to change it from VBA code.

    ' File -> Options -> Advanced -> Show document content: Field shading: Always
     ActiveWindow.View.FieldShading = wdFieldShadingAlways
    
    ' File -> Options -> Advanced ->  Image size and qualty, default resolution: 330 PPI
    ' impossible to set in VBA
    ActiveWindow.View.ShowDrawings = True                           ' Show drawings and text boxes on screen
    ActiveWindow.View.ShowBookmarks = True                          ' Show bookmarks
    ActiveWindow.View.ShowTextBoundaries = True                     ' Show text boundaries
    ActiveWindow.View.ShowCropMarks = True                          ' Show crop marks
    Options.MeasurementUnit = wdCentimeters                         ' Show measurements in units of centimeters
    
    ActiveWindow.View.Type = wdNormalView
    ActiveWindow.StyleAreaWidth = CentimetersToPoints(5.3)          ' Set Style area pane width in Draft and Outline view to 5.3 cm
    ' Allow hyphenation between pages or columns
    ' impossible in VBA
    ActiveDocument.Compatibility(wdSuppressBottomSpacing) = False   ' Suppress extra line spacing at bottom of page
    ActiveDocument.Compatibility(wdSuppressTopSpacing) = False      ' Suppress extra line spacing at top of page
    
    ActiveWindow.View.Type = wdPrintView
    ActiveWindow.View.Zoom.Percentage = 100
    On Error Resume Next
        ActiveWindow.View.PageMovementType = wdSideToSide
    If Err.Number <> 0 Then
        ActiveWindow.View.PageMovementType = wdVertical
    End If
    On Error GoTo 0
    
    MsgBox _
        Prompt:="Microsoft Word options customized successfully:" & vbNewLine & vbNewLine & _
            "Show drawings and text boxes on screen: " & ActiveWindow.View.ShowDrawings & vbNewLine & _
            "Show bookmarks: " & ActiveWindow.View.ShowBookmarks & vbNewLine & _
            "Show text boundaries: " & ActiveWindow.View.ShowTextBoundaries & vbNewLine & _
            "Show crop marks: " & ActiveWindow.View.ShowCropMarks & vbNewLine & _
            "Show measurements in units of centimeters: " & Options.MeasurementUnit & vbNewLine & _
            "Set Style area pane width in Draft and Outline view to 5.3 cm: " & ActiveWindow.StyleAreaWidth & vbNewLine & _
            "Suppress extra line spacing at bottom of page: " & ActiveDocument.Compatibility(wdSuppressBottomSpacing) & vbNewLine & _
            "Suppress extra line spacing at top of page: " & ActiveDocument.Compatibility(wdSuppressTopSpacing) & vbNewLine & _
            "Set page movement from 'Vertical' to 'Side to Side': " & ActiveWindow.View.PageMovementType, _
        Buttons:=vbInformation + vbOKOnly, _
        Title:=MsgBoxTitle
End Sub

' Check settings of Autocorrect options. If true then all areoff.
' 2025-08-14 by ms
Private Function CheckIfAutocorrectAreOff() As Boolean
    CheckIfAutocorrectAreOff = True
    
    If Not AutoCorrect.CorrectSentenceCaps = True Then
        CheckIfAutocorrectAreOff = False
    End If
    If Not AutoCorrect.CorrectDays = True Then
        CheckIfAutocorrectAreOff = False
    End If
    If Not AutoCorrect.TwoInitialCapsAutoAdd = True Then
        CheckIfAutocorrectAreOff = False
    End If
    If Not AutoCorrect.CorrectCapsLock = True Then
        CheckIfAutocorrectAreOff = False
    End If
    If Not AutoCorrect.CorrectInitialCaps = True Then
        CheckIfAutocorrectAreOff = False
    End If
    If Not AutoCorrect.CorrectSentenceCaps = True Then
        CheckIfAutocorrectAreOff = False
    End If
    If Not AutoCorrect.CorrectTableCells = True Then
        CheckIfAutocorrectAreOff = False
    End If
    If Not AutoCorrect.replaceText = True Then
        CheckIfAutocorrectAreOff = False
    End If
End Function

' Personal preferences of ms
' 2025-06-18 by ms
Sub WordOptionsDisableAutoCorrect()
    AutoCorrect.CorrectSentenceCaps = False
    AutoCorrect.CorrectDays = False
    AutoCorrect.TwoInitialCapsAutoAdd = False
    AutoCorrect.CorrectCapsLock = False
    AutoCorrect.CorrectInitialCaps = False
    AutoCorrect.CorrectSentenceCaps = False
    AutoCorrect.CorrectTableCells = False
    AutoCorrect.replaceText = False

    Dim FileName As String
    FileName = C_F_Macros
    
    Dim ModuleName As String
    ModuleName = C_M_Tools
    
    Dim MacroName As String
    MacroName = "WordOptionsDisableAutoCorrect"
    
    Dim MsgBoxTitle As String
    MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    MsgBox _
        Prompt:="Microsoft Word options related to Autocorrect were switched off (all).", _
        Buttons:=vbInformation + vbOKOnly, _
        Title:=MsgBoxTitle
End Sub

' Restore AutoFormat options of Microsoft Word to default values.
' 2025-08-17 by ms
Sub WordOptionsRestoreAutoFormat()
    Dim FileName As String
    FileName = C_F_Macros
    
    Dim ModuleName As String
    ModuleName = C_M_Tools
    
    Dim MacroName As String
    MacroName = "WordOptionsRestoreAutoFormat"
    
    Dim MsgBoxTitle As String
    MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    With Application.Options
        ' AutoCorrect -> AutoFormat -> Apply: Built-in Heading styles
        .AutoFormatApplyHeadings = True
        ' AutoCorrect -> AutoFormat -> Apply: List styles
        .AutoFormatApplyLists = True
       ' AutoCorrect -> AutoFormat -> Apply: Automatic bulleted lists
        .AutoFormatApplyBulletedLists = True
        ' AutoCorrect -> AutoFormat -> Apply: Other paragraph styles
        .AutoFormatApplyOtherParas = False

         ' AutoCorrect -> AutoFormat -> Replace: "Straight quotes" with "smart quotes"
        .AutoFormatReplaceQuotes = True
        ' AutoCorrect -> AutoFormat -> Replace: Ordinals (1st) with superscript
        .AutoFormatReplaceOrdinals = True
        ' AutoCorrect -> AutoFormat -> Replace: Fractions (1/2) with fraction character
        .AutoFormatReplaceFractions = True
        ' AutoCorrect -> AutoFormat -> Replace: Hyphens (--) with dash (—)
        .AutoFormatReplaceFarEastDashes = True
        
        ' AutoCorrect -> AutoFormat As You Type: "Straight quotes" with "smart quotes"
        .AutoFormatAsYouTypeReplaceQuotes = True
        ' AutoCorrect -> AutoFormat As You Type: Fractions (1/2) with fraction characte
        .AutoFormatAsYouTypeReplaceFractions = True
        ' AutoCorrect -> AutoFormat As You Type: Ordinals (1st) with superscript
        .AutoFormatAsYouTypeReplaceOrdinals = True
        ' AutoCorrect -> AutoFormat As You Type: Hyphens (--) with dash (—)
        .AutoFormatAsYouTypeReplaceSymbols = True
    End With
    
    MsgBox _
        Prompt:="File -> Options -> Proofing -> AutoCorrect Options -> AutoFormat:" & vbNewLine & vbNewLine & _
            "AutoFormatApplyHeadings = " & Application.Options.AutoFormatApplyHeadings & vbNewLine & _
            ".AutoFormatApplyLists = " & Application.Options.AutoFormatApplyLists & vbNewLine & _
            ".AutoFormatApplyBulletedLists = " & Application.Options.AutoFormatApplyBulletedLists & vbNewLine & _
            ".AutoFormatApplyOtherParas = " & Application.Options.AutoFormatApplyOtherParas & vbNewLine & _
            ".AutoFormatReplaceQuotes = " & Application.Options.AutoFormatReplaceQuotes & vbNewLine & _
            ".AutoFormatReplaceOrdinals = " & Application.Options.AutoFormatReplaceOrdinals & vbNewLine & _
            ".AutoFormatReplaceFractions = " & Application.Options.AutoFormatReplaceFractions & vbNewLine & _
            ".AutoFormatReplaceFarEastDashes = " & Application.Options.AutoFormatReplaceFarEastDashes & vbNewLine & vbNewLine & _
            "File -> Options -> Proofing -> AutoCorrect Options -> AutoFormat As You Type:" & vbNewLine & vbNewLine & _
            ".AutoFormatAsYouTypeReplaceQuotes = " & Application.Options.AutoFormatAsYouTypeReplaceQuotes & vbNewLine & _
            ".AutoFormatAsYouTypeReplaceFractions = " & Application.Options.AutoFormatAsYouTypeReplaceFractions & vbNewLine & _
            ".AutoFormatAsYouTypeReplaceOrdinals = " & Application.Options.AutoFormatAsYouTypeReplaceOrdinals & vbNewLine & _
            ".AutoFormatAsYouTypeReplaceSymbols = " & Application.Options.AutoFormatAsYouTypeReplaceSymbols & vbNewLine & vbNewLine & _
            "Finished processing.", _
        Buttons:=vbInformation, _
        Title:=MsgBoxTitle

End Sub

' 2025-04-02 by ms and AI
' 2025-07-16 by ms, AutoFormat is batch processing, not available directly from ribbon menu
' AutoFormat as you type is fully automatic. This macro disables both.
Sub WordOptionsDisableAutoFormat()
    Dim FileName As String
    FileName = C_F_Macros
    
    Dim ModuleName As String
    ModuleName = C_M_Tools
    
    Dim MacroName As String
    MacroName = "WordOptionsDisableAutoFormat"
    
    Dim MsgBoxTitle As String
    MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    With Application.Options
        ' AutoCorrect -> AutoFormat -> Apply: Built-in Heading styles
        .AutoFormatApplyHeadings = False
        ' AutoCorrect -> AutoFormat -> Apply: List styles
        .AutoFormatApplyLists = False
       ' AutoCorrect -> AutoFormat -> Apply: Automatic bulleted lists
        .AutoFormatApplyBulletedLists = False
        ' AutoCorrect -> AutoFormat -> Apply: Other paragraph styles
        .AutoFormatApplyOtherParas = False

         ' AutoCorrect -> AutoFormat -> Replace: "Straight quotes" with "smart quotes"
        .AutoFormatReplaceQuotes = False
        ' AutoCorrect -> AutoFormat -> Replace: Ordinals (1st) with superscript
        .AutoFormatReplaceOrdinals = False
        ' AutoCorrect -> AutoFormat -> Replace: Fractions (1/2) with fraction character
        .AutoFormatReplaceFractions = False
        ' AutoCorrect -> AutoFormat -> Replace: Hyphens (--) with dash (—)
        .AutoFormatReplaceFarEastDashes = False
        
        ' AutoCorrect -> AutoFormat As You Type: "Straight quotes" with "smart quotes"
        .AutoFormatAsYouTypeReplaceQuotes = False
        ' AutoCorrect -> AutoFormat As You Type: Fractions (1/2) with fraction characte
        .AutoFormatAsYouTypeReplaceFractions = False
        ' AutoCorrect -> AutoFormat As You Type: Ordinals (1st) with superscript
        .AutoFormatAsYouTypeReplaceOrdinals = False
        ' AutoCorrect -> AutoFormat As You Type: Hyphens (--) with dash (—)
        .AutoFormatAsYouTypeReplaceSymbols = False
    End With
    
    MsgBox _
        Prompt:="File -> Options -> Proofing -> AutoCorrect Options -> AutoFormat:" & vbNewLine & vbNewLine & _
            "AutoFormatApplyHeadings = False" & vbNewLine & _
            ".AutoFormatApplyLists = False" & vbNewLine & _
            ".AutoFormatApplyBulletedLists = False" & vbNewLine & _
            ".AutoFormatApplyOtherParas = False" & vbNewLine & _
            ".AutoFormatReplaceQuotes = False" & vbNewLine & _
            ".AutoFormatReplaceOrdinals = False" & vbNewLine & _
            ".AutoFormatReplaceFractions = False" & vbNewLine & _
            ".AutoFormatReplaceFarEastDashes = False" & vbNewLine & vbNewLine & _
            "File -> Options -> Proofing -> AutoCorrect Options -> AutoFormat As You Type:" & vbNewLine & vbNewLine & _
            ".AutoFormatAsYouTypeReplaceQuotes = False" & _
            ".AutoFormatAsYouTypeReplaceFractions = False" & vbNewLine & _
            ".AutoFormatAsYouTypeReplaceOrdinals = False" & vbNewLine & _
            ".AutoFormatAsYouTypeReplaceSymbols = False" & vbNewLine & vbNewLine & _
            "Finished processing.", _
        Buttons:=vbInformation, _
        Title:=MsgBoxTitle
End Sub

Sub WordOptionsToggleAutoCorrect()
    Dim FileName As String
    FileName = C_F_Macros
    
    Dim ModuleName As String
    ModuleName = C_M_Tools
    
    Dim MacroName As String
    MacroName = "WordOptionsToggleAutoCorrect"
    
    Dim MsgBoxTitle As String
    MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    With Application.Options
        ' AutoCorrect -> AutoFormat -> Apply: Built-in Heading styles
        .AutoFormatApplyHeadings = Not .AutoFormatApplyHeadings
        ' AutoCorrect -> AutoFormat -> Apply: List styles
        .AutoFormatApplyLists = Not .AutoFormatApplyLists
        ' AutoCorrect -> AutoFormat -> Apply: Automatic bulleted lists
        .AutoFormatApplyBulletedLists = Not .AutoFormatApplyBulletedLists
        ' AutoCorrect -> AutoFormat -> Apply: Other paragraph styles
        .AutoFormatApplyOtherParas = Not .AutoFormatApplyOtherParas

        ' AutoCorrect -> AutoFormat -> Replace: "Straight quotes" with "smart quotes"
        .AutoFormatReplaceQuotes = Not .AutoFormatReplaceQuotes
        ' AutoCorrect -> AutoFormat -> Replace: Ordinals (1st) with superscript
        .AutoFormatReplaceOrdinals = Not .AutoFormatReplaceOrdinals
        ' AutoCorrect -> AutoFormat -> Replace: Fractions (1/2) with fraction character
        .AutoFormatReplaceFractions = Not .AutoFormatReplaceFractions
        ' AutoCorrect -> AutoFormat -> Replace: Hyphens (--) with dash (—)
        .AutoFormatReplaceFarEastDashes = Not .AutoFormatReplaceFarEastDashes
        
        ' AutoCorrect -> AutoFormat As You Type: "Straight quotes" with "smart quotes"
        .AutoFormatAsYouTypeReplaceQuotes = Not .AutoFormatAsYouTypeReplaceQuotes
        ' AutoCorrect -> AutoFormat As You Type: Fractions (1/2) with fraction characte
        .AutoFormatAsYouTypeReplaceFractions = Not .AutoFormatAsYouTypeReplaceFractions
        ' AutoCorrect -> AutoFormat As You Type: Ordinals (1st) with superscript
        .AutoFormatAsYouTypeReplaceOrdinals = Not .AutoFormatAsYouTypeReplaceOrdinals
        ' AutoCorrect -> AutoFormat As You Type: Hyphens (--) with dash (—)
        .AutoFormatAsYouTypeReplaceSymbols = Not .AutoFormatAsYouTypeReplaceSymbols
    End With
    
    MsgBox _
        Prompt:="Finished processing.", _
        Buttons:=vbInformation, _
        Title:=MsgBoxTitle
End Sub

' 2025-04-06 by ms and AI
Sub DocPropertiesUserInput()
    ' Before Show method, the InputDocProperties_Form > Private Sub UserForm_Initialize() is run automatically.
    InputDocProperties_Form.Show
End Sub

' 2025-06-19 by ms
Public Sub AttachBuildingBlocks()
    Dim FileName As String
    FileName = C_F_Macros
    
    Dim ModuleName As String
    ModuleName = C_M_Tools
    
    Dim MacroName As String
    MacroName = "AttachBuildingBlocks"
    
    Dim MsgBoxTitle As String
    MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName

    Dim BuildingBlockEntries As Object
    ' At first try to set bbe to ActiveDocument.AttachedTemplate.BuildingBlockEntries (template with integrated BuildingBlocks)
    Set BuildingBlockEntries = ActiveDocument.AttachedTemplate.BuildingBlockEntries

    ' Check if bbe is empty
    Dim AddInsIndex As Integer
    If BuildingBlockEntries.count = 0 Then
        ' Force loading of all BuildingBlocks
        Templates.LoadBuildingBlocks
        If AddIns.count <> 0 Then
            AddInsIndex = ReturnAddInsIndex()
            Call EnableAddIns(AddInsIndex)
        Else
            ' enable BuildingBlocks template "C_F_BuildingBlocks"
            Call LoadBuildingBlocksFromUserTemplate
            AddInsIndex = ReturnAddInsIndex()
            If AddInsIndex <> 0 Then
                Call EnableAddIns(AddInsIndex)
            End If
        End If
    End If
    
    If AddIns.count = 0 Then
        MsgBox _
            Prompt:="No global templates or add-ins were found." & vbNewLine & _
                "Please check settings in menu: Developer -> Document Template." & vbNewLine & vbNewLine & _
                "Exiting.", _
            Buttons:=vbCritical + vbOKOnly, _
            Title:=MsgBoxTitle
        Exit Sub
    End If
End Sub

Private Sub EnableAddIns(AddInsIndex As Integer)
    Dim UserDecision As VbMsgBoxResult
    
    Dim FileName As String
    FileName = C_F_Macros
    
    Dim ModuleName As String
    ModuleName = C_M_Tools
    
    Dim MacroName As String
    MacroName = "EnableAddIns"
    
    Dim MsgBoxTitle As String
    MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    If Not AddIns(AddInsIndex).Installed Then
        Beep
        UserDecision = MsgBox( _
            Prompt:="The " & C_F_BuildingBlocks & " is found, but not enabled." & vbNewLine & _
                "Would you like to enable it now?", _
            Buttons:=vbYesNo + vbQuestion + vbDefaultButton1, _
            Title:=MsgBoxTitle)
        If UserDecision = vbYes Then
            AddIns(AddInsIndex).Installed = True
        Else
            Exit Sub
        End If
    End If
End Sub

Private Function ReturnAddInsIndex() As Integer
    Dim i As Integer
    Dim AddInsIndex As Integer
    Dim AddInsName As String

    For i = 1 To AddIns.count
        If AddIns(i).Name = C_F_BuildingBlocks Then
            AddInsIndex = i
            AddInsName = AddIns(AddInsIndex).Name
            Exit For
        End If
    Next i
    ReturnAddInsIndex = AddInsIndex
End Function

Private Sub LoadBuildingBlocksFromUserTemplate()
    Dim TemplatePath As String
    Dim fullPath As String

    ' Get the User Templates path
    TemplatePath = Options.DefaultFilePath(wdUserTemplatesPath)

    ' Combine path and file name
    fullPath = TemplatePath & "\" & C_F_BuildingBlocks

    ' Check if the file exists
    If Dir(fullPath) <> "" Then
        ' Load the template as an Add-In (loads building blocks)
        AddIns.Add FileName:=fullPath, Install:=False
    End If
End Sub


Sub SetLanguageToEnglishUS()
    Dim FileName As String
    FileName = C_F_Macros
    
    Dim ModuleName As String
    ModuleName = C_M_Tools
    
    Dim MacroName As String
    MacroName = "SetLanguageToEnglishUS"
    
    Dim MsgBoxTitle As String
    MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    ' Check if there is a selection
    If Selection.Type <> wdNoSelection Then
        ' Set the language of the current selection to English (United States)
        Selection.LanguageId = wdEnglishUS
        MsgBox _
            Prompt:="Language set to English (United States) for the current selection.", _
            Buttons:=vbInformation, _
            Title:=MsgBoxTitle
    Else
        MsgBox _
            Prompt:="No text is selected. Please select the text you want to change the language for.", _
            Buttons:=vbInformation, _
            Title:=MsgBoxTitle
    End If
End Sub

' 2025-07-16 by ms
Private Sub InitializeStyleNameArray(StyleNameArray As Variant)
    ' Define the styles to search for
    Let StyleNameArray = Array(C_S_Heading1, _
                            C_S_Heading2, _
                            C_S_Heading3, _
                            C_S_Heading4, _
                            C_S_Heading5, _
                            C_S_Heading6, _
                            C_S_Heading7, _
                            C_S_Heading8)
End Sub

' The idea is to speed up manual insertion of cross-references. At first at the beginning of paragraphs styled
' with specific style names a cross reference is inserted, which is hidden. Next manually such cross reference
' have to be inserted in specific content context.
' 2025-04-15 by ms and AI
Private Sub InsertCrossReferences_Headings()
    Dim doc As Document
    Dim para2 As Paragraph
    Dim rng As Range
    
    Dim StyleName2 As Variant
    Dim refItemH As Integer
    ' Set the document to the active document
    Set doc = ActiveDocument
      
    Dim RefItems As Variant
    Let RefItems = doc.GetCrossReferenceItems(wdRefTypeNumberedItem)
    Dim RefItemsSize As Integer
    Let RefItemsSize = UBound(RefItems)
    Dim HeaderMatrix() As String
    ' Initialize the data array with an initial size
    ReDim HeaderMatrix(1 To RefItemsSize)
    
    Dim StyleNameArray As Variant
    ' Initilaize StyleNameArray
    Call InitializeStyleNameArray(StyleNameArray)
    ' Fill in text string matrix with header content
    HeaderMatrix = BuildHeaderArray(StyleNameArray)
    
    Dim i As Integer
    Let i = 1
    Dim j As Integer
    Let j = 1
    Dim TempString As String
    Dim FlagFound As Boolean
    Let FlagFound = False
    Dim HeaderMatrixSize As Integer
    Let HeaderMatrixSize = UBound(HeaderMatrix)
   
    ' Loop through each paragraph in the document
    For Each para2 In doc.Paragraphs
        ' Check if the paragraph style matches any of the specified styles
        For Each StyleName2 In StyleNameArray
            If para2.style = StyleName2 Then
                ' Set the range to the beginning of the paragraph
                Set rng = para2.Range
                
                For i = 1 To HeaderMatrixSize
                    If para2.Range.Text = HeaderMatrix(i) Then
                        For j = 1 To RefItemsSize
                            ' Remove the first and the last character
                            TempString = Mid(HeaderMatrix(i), 2, Len(HeaderMatrix(i)) - 2)
                            ' Trim all spaces (from the front of  a text string and from the end of it).
                            TempString = Trim(TempString)
                            If InStr(1, RefItems(j), TempString, vbTextCompare) Then
                                rng.Collapse Direction:=wdCollapseStart
                                ' Insert the cross-reference at the beginning of the paragraph using InsertCrossReference method
                                ' "Numbered item" = wdRefTypeNumberedItem
                                rng.InsertCrossReference _
                                    ReferenceType:=wdRefTypeNumberedItem, _
                                    ReferenceKind:=wdNumberRelativeContext, _
                                    ReferenceItem:=CStr(j), _
                                    InsertAsHyperlink:=True, _
                                    IncludePosition:=False, _
                                    SeparateNumbers:=False, _
                                    SeparatorString:=" "
                                
                                ' Find the position of the "em space" ChrW(8195)
                                Dim SpacePosition As Integer
                                SpacePosition = InStr(para2.Range.Text, ChrW(8195))
                                Dim SubstrLength As Integer
                                SubstrLength = SpacePosition - 1
                                
                                rng.MoveEnd Unit:=wdCharacter, count:=SubstrLength
                                rng.style = doc.Styles(C_S_CharHidden)
                                FlagFound = True
                                Exit For
                            End If
                        Next j
                    End If
                    If FlagFound = True Then
                        Exit For
                    End If
                Next i
            End If
            If FlagFound = True Then
                FlagFound = False
                Exit For
            End If
        Next StyleName2
    Next para2
    
    Set rng = Nothing
    Set doc = Nothing
    
    Dim FileName As String
    FileName = C_F_Macros
    
    Dim ModuleName As String
    ModuleName = C_M_Tools
    
    Dim MacroName As String
    MacroName = "InsertCrossReferences_Headings"
    
    Dim MsgBoxTitle As String
    MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    MsgBox _
        Prompt:="Finished processing.", _
        Buttons:=vbInformation + vbOKOnly, _
        Title:=MsgBoxTitle
End Sub

' 2025-04-17 by ms
Private Function BuildHeaderArray(StyleNameArray As Variant) As String()
    Dim i As Integer
    Dim para1 As Paragraph
    Dim styleName1 As Variant
    Dim HeaderArray() As String
    
    ' Initialize the array with an initial size
    ReDim HeaderArray(1 To ActiveDocument.Paragraphs.count)
    
    ' Loop through each paragraph in the document
    Let i = 1
    For Each para1 In ActiveDocument.Paragraphs
        ' Check if the paragraph style matches any of the specified styles
        For Each styleName1 In StyleNameArray
            If para1.style = styleName1 Then
                HeaderArray(i) = para1.Range.ListFormat.ListString & " " & para1.Range.Text
                i = i + 1
            End If
        Next styleName1
    Next para1
    
    ' Redimension the HeaderMatrix preserving its data.
    Dim HeaderMatrixSize As Integer
    Let HeaderMatrixSize = i - 1
    If i > 0 Then
        ReDim Preserve HeaderArray(1 To HeaderMatrixSize)
    Else
        ReDim Preserve HeaderArray(1 To 1)
    End If
    
    ' Return the array
    BuildHeaderArray = HeaderArray
End Function

' Toggle the command bar "Apply Styles"
' 2025-04-15 by ms
' 2025-07-15 by ms
' https://superuser.com/questions/1825151/word-keyboard-shortcut-to-close-apply-styles-popup
Sub ToggleApplyStyles()
    Static ToggleBit As Boolean
    Dim FileName As String
    FileName = C_F_Macros
    
    Dim ModuleName As String
    ModuleName = C_M_Tools
    
    Dim MacroName As String
    MacroName = "ToggleApplyStyles"
    
    Dim MsgBoxTitle As String
    MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    ToggleBit = Not ToggleBit
    If ToggleBit Then
        Application.Run "StyleApplyPane"
        Application.statusBar = MsgBoxTitle & " > " & "on"
    Else
        If CommandBars("Apply Styles").Enabled Then
            CommandBars("Apply Styles").Visible = False
            Application.statusBar = MsgBoxTitle & " > " & "off"
        End If
    End If
End Sub

' Pair subroutine to the InsertCrossReferences_Headings()
' 2025-04-17 by ms and ai
Private Sub DeleteCrossReference_Headings()
    Dim doc As Document
    Dim para As Paragraph
    Dim fld As Field
    Dim rng As Range
    Dim styleName As Variant
    Dim found As Boolean
    
    ' Initialize the style names array
    Dim StyleNameArray As Variant
    ' Define the styles to search for
    Let StyleNameArray = Array(C_S_Heading1, _
                            C_S_Heading2, _
                            C_S_Heading3, _
                            C_S_Heading4, _
                            C_S_Heading5, _
                            C_S_Heading6, _
                            C_S_Heading7, _
                            C_S_Heading8)
    
    ' Set the document
    Set doc = ActiveDocument
    
    ' Loop through each paragraph in the document
    For Each para In doc.Paragraphs
        ' Check if the paragraph style is in the StyleNameArray
        Let found = False
        For Each styleName In StyleNameArray
            If para.style = doc.Styles(styleName) Then
                found = True
                Exit For
            End If
        Next styleName
        
        ' If the paragraph style is found in the array
        If found Then
            ' Check if the paragraph starts with a cross-reference field
            If para.Range.Fields.count > 0 Then
                Set fld = para.Range.Fields(1)
                If fld.Type = wdFieldRef Then
                    ' Remove the cross-reference field
                    fld.Delete
                End If
            End If
        End If
    Next para
    
    ' Clean up of the object variables
    Set doc = Nothing
    Set fld = Nothing
    
    Dim FileName As String
    FileName = C_F_Macros
    
    Dim ModuleName As String
    ModuleName = C_M_Tools
    
    Dim MacroName As String
    MacroName = "DeleteCrossReference_Headings"
    
    Dim MsgBoxTitle As String
    MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    MsgBox _
        Prompt:="Finished processing.", _
        Buttons:=vbInformation + vbOKOnly, _
        Title:=MsgBoxTitle
End Sub

' 2025-04-18 by ms
Private Sub DeleteCrossReferences_Pictures()
    Dim doc As Document
    Dim para As Paragraph
    Dim fld As Field
    Dim rng As Range
    Dim styleName As Variant
    Dim found As Boolean
    
    ' Set the document
    Set doc = ActiveDocument
    
    ' Loop through each paragraph in the document
    For Each para In doc.Paragraphs
        ' Check if the paragraph style is in the StyleNameArray
        Let found = False
        If para.style = C_S_PictureLegend Then
            found = True
        End If
        
        ' If the paragraph style is found in the array
        If found Then
            ' Check if the paragraph starts with a cross-reference field
            If para.Range.Fields.count > 0 Then
                Set fld = para.Range.Fields(1)
                If fld.Type = wdFieldRef Then
                    ' Get the range of the filed result
                    Set rng = fld.result
                    ' Remove the cross-reference field
                    fld.Delete
                    ' Check if the next character is a space and delete it
                    If rng.Characters.count > 0 Then
                        If rng.Characters(1).Text = " " Then
                            rng.Characters(1).Delete
                        End If
                    End If
                End If
            End If
        End If
    Next para
    
    ' Clean up of the object variables
    Set doc = Nothing
    Set fld = Nothing
    Set rng = Nothing
    
    Dim FileName As String
    FileName = C_F_Macros
    
    Dim ModuleName As String
    ModuleName = C_M_Tools
    
    Dim MacroName As String
    MacroName = "DeleteCrossReferences_Pictures"
    
    Dim MsgBoxTitle As String
    MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    MsgBox _
        Prompt:="Finished processing.", _
        Buttons:=vbInformation + vbOKOnly, _
        Title:=MsgBoxTitle
End Sub

' 2025-04-18 by ms and AI
Private Sub InsertCrossReferences_Pictures()
    Dim doc As Document
    Dim para As Paragraph
    Dim rng As Range
    
    ' Set the document to the active document
    Set doc = ActiveDocument
    
    Dim RefItems As Variant
    Let RefItems = doc.GetCrossReferenceItems(C_Caption_Pic)
    
    Dim i As Integer
    Let i = 1
    ' Loop through each paragraph in the document
    For Each para In doc.Paragraphs
        ' Check if the paragraph style matches any of the specified styles
        If para.style = C_S_PictureLegend Then
            ' Set the range to the beginning of the paragraph
            Set rng = para.Range
            rng.Collapse Direction:=wdCollapseStart
            ' Insert a space character before the cross-reference field
            rng.Text = " "
            ' Move the insertion point before the inserted space character
            rng.Collapse Direction:=wdCollapseStart
            ' "Pic." = C_Caption_Pic
            rng.InsertCrossReference _
                ReferenceType:=C_Caption_Pic, _
                ReferenceKind:=wdOnlyLabelAndNumber, _
                ReferenceItem:=CStr(i), _
                InsertAsHyperlink:=True, _
                IncludePosition:=False, _
                SeparateNumbers:=False, _
                SeparatorString:=" "
            ' Move the insertion point to the beginning of the range after cross-reference is inserted
            rng.Collapse Direction:=wdCollapseStart
            ' Select all characters belonging to the cross-reference field
            rng.MoveEndUntil cset:=Chr(32), count:=wdForward
            ' Apply the character style to the cross-reference
            rng.style = C_S_CharHidden
            ' Increment the reference item counter for the next cross-reference
            i = i + 1
        End If
    Next para
    
    Set rng = Nothing
    Set doc = Nothing
    
    Dim FileName As String
    FileName = C_F_Macros
    
    Dim ModuleName As String
    ModuleName = C_M_Tools
    
    Dim MacroName As String
    MacroName = "InsertCrossReferences_Pictures"
    
    Dim MsgBoxTitle As String
    MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    MsgBox _
        Prompt:="Finished processing.", _
        Buttons:=vbInformation + vbOKOnly, _
        Title:=MsgBoxTitle
End Sub

' 2025-04-18 by ms and AI
Private Sub InsertCrossReferences_Tables()
    Dim doc As Document
    Dim para As Paragraph
    Dim rng As Range
    
    ' Set the document to the active document
    Set doc = ActiveDocument
    
    Dim RefItems As Variant
    Let RefItems = doc.GetCrossReferenceItems(C_Caption_Tab)
    
    Dim i As Integer
    Let i = 1
    ' Loop through each paragraph in the document
    For Each para In doc.Paragraphs
        ' Check if the paragraph style matches any of the specified styles
        If para.style = C_S_TableLegend Then
            ' Set the range to the beginning of the paragraph
            Set rng = para.Range
            rng.Collapse Direction:=wdCollapseStart
            ' Insert a space character before the cross-reference field
            rng.Text = " "
            ' Move the insertion point before the inserted space character
            rng.Collapse Direction:=wdCollapseStart
            ' "Tab." = C_Caption_Tab
            rng.InsertCrossReference _
                ReferenceType:=C_Caption_Tab, _
                ReferenceKind:=wdOnlyLabelAndNumber, _
                ReferenceItem:=CStr(i), _
                InsertAsHyperlink:=True, _
                IncludePosition:=False, _
                SeparateNumbers:=False, _
                SeparatorString:=" "
            ' Move the insertion point to the beginning of the range after cross-reference is inserted
            rng.Collapse Direction:=wdCollapseStart
            ' Select all characters belonging to the cross-reference field
            rng.MoveEndUntil cset:=Chr(32), count:=wdForward
            ' Apply the character style to the cross-reference
            rng.style = C_S_CharHidden
            ' Increment the reference item counter for the next cross-reference
            i = i + 1
        End If
    Next para
    
    Set rng = Nothing
    Set doc = Nothing
    
    Dim FileName As String
    FileName = C_F_Macros
    
    Dim ModuleName As String
    ModuleName = C_M_Tools
    
    Dim MacroName As String
    MacroName = "InsertCrossReferences_Tables"
    
    Dim MsgBoxTitle As String
    MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    MsgBox _
        Prompt:="Finished processing.", _
        Buttons:=vbInformation + vbOKOnly, _
        Title:=MsgBoxTitle
End Sub

' 2025-04-18 by ms
Private Sub DeleteCrossReferences_Tables()
    Dim doc As Document
    Dim para As Paragraph
    Dim fld As Field
    Dim rng As Range
    Dim styleName As Variant
    Dim found As Boolean
    
    ' Set the document
    Set doc = ActiveDocument
    
    ' Loop through each paragraph in the document
    For Each para In doc.Paragraphs
        ' Check if the paragraph style is in the StyleNameArray
        Let found = False
        If para.style = C_S_TableLegend Then
            found = True
        End If
        
        ' If the paragraph style is found in the array
        If found Then
            ' Check if the paragraph starts with a cross-reference field
            If para.Range.Fields.count > 0 Then
                Set fld = para.Range.Fields(1)
                If fld.Type = wdFieldRef Then
                    ' Get the range of the filed result
                    Set rng = fld.result
                    ' Remove the cross-reference field
                    fld.Delete
                    ' Check if the next character is a space and delete it
                    If rng.Characters.count > 0 Then
                        If rng.Characters(1).Text = " " Then
                            rng.Characters(1).Delete
                        End If
                    End If
                End If
            End If
        End If
    Next para
    
    ' Clean up of the object variables
    Set doc = Nothing
    Set fld = Nothing
    Set rng = Nothing
    
    Dim FileName As String
    FileName = C_F_Macros
    
    Dim ModuleName As String
    ModuleName = C_M_Tools
    
    Dim MacroName As String
    MacroName = "DeleteCrossReferences_Tables"
    
    Dim MsgBoxTitle As String
    MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    MsgBox _
        Prompt:="Finished processing.", _
        Buttons:=vbInformation + vbOKOnly, _
        Title:=MsgBoxTitle
End Sub

Private Function BuildReferencesArray() As String()
    Dim i As Integer
    Dim para As Paragraph
    Dim styleName As Variant
    Dim ReferencesArray() As String
    
    ' Initialize the array with an initial size
    ReDim ReferencesArray(1 To ActiveDocument.Paragraphs.count)
    
    ' Loop through each paragraph in the document
    Let i = 1
    For Each para In ActiveDocument.Paragraphs
        ' Check if the paragraph style matches any of the specified styles
        If Not para.Range.ListStyle Is Nothing Then
            If para.Range.ListStyle.NameLocal = C_S_ListNumRef Then
                ReferencesArray(i) = para.Range.Text
                i = i + 1
            End If
        End If
    Next para
    
    ' Redimension the HeaderMatrix preserving its data.
    Dim ReferenceArraySize As Integer
    Let ReferenceArraySize = i - 1
    If i > 0 Then
        ReDim Preserve ReferencesArray(1 To ReferenceArraySize)
    Else
        ReDim Preserve ReferencesArray(1 To 1)
    End If
    
    ' Return the array
    BuildReferencesArray = ReferencesArray
End Function

' The last paragraph in a document must be emtpy and formatted to "ParNormal ms" or "Normal".
' 2025-04-19 by ms
Private Sub InsertCrossReferences_References()
    Dim FileName As String
    FileName = C_F_Macros
    
    Dim ModuleName As String
    ModuleName = C_M_Tools
    
    Dim MacroName As String
    MacroName = "InsertCrossReferences_References"
    
    Dim MsgBoxTitle As String
    MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    Dim para As Paragraph
    Dim rng As Range
    
    Dim styleName As Variant
    
    Dim doc As Document
    ' Set the document to the active document
    Set doc = ActiveDocument
      
    Dim RefItems As Variant
    Let RefItems = doc.GetCrossReferenceItems(wdRefTypeNumberedItem)
    Dim RefItemsSize As Integer
    Let RefItemsSize = UBound(RefItems)
    Dim ReferenceMatrix() As String
    ' Initialize the data array with an initial size
    ReDim ReferenceMatrix(1 To RefItemsSize)
    ' Fill in text string matrix with header content
    ReferenceMatrix = BuildReferencesArray()
    
    Dim i As Integer
    Let i = 1
    Dim j As Integer
    Let j = 1
    Dim TempString As String
    Dim FlagFound As Boolean
    Let FlagFound = False
    Dim ReferenceMatrixSize As Integer
    Let ReferenceMatrixSize = UBound(ReferenceMatrix)
   
    ' Loop through each paragraph in the document
    For Each para In doc.Paragraphs
        ' Check if the paragraph style matches any of the specified styles
        If Not para.Range.ListStyle Is Nothing Then
            If para.Range.ListStyle.NameLocal = C_S_ListNumRef Then
                ' Set the range to the beginning of the paragraph
                Set rng = para.Range
                
                For i = 1 To ReferenceMatrixSize
                    If para.Range.Text = ReferenceMatrix(i) Then
                        For j = 1 To RefItemsSize
                            ' Remove the first and the last character
                            TempString = Mid(ReferenceMatrix(i), 2, Len(ReferenceMatrix(i)) - 2)
                            ' Trim all spaces (from the front of  a text string and from the end of it).
                            TempString = Trim(TempString)
                            TempString = Replace(TempString, Chr(11), " ")
                            If InStr(1, RefItems(j), TempString, vbTextCompare) Then
                                rng.Collapse Direction:=wdCollapseStart
                                ' Insert the cross-reference at the beginning of the paragraph using InsertCrossReference method
                                ' "Numbered item" = wdRefTypeNumberedItem
                                On Error GoTo Error_Insert
                                rng.InsertCrossReference _
                                    ReferenceType:=wdRefTypeNumberedItem, _
                                    ReferenceKind:=wdNumberRelativeContext, _
                                    ReferenceItem:=CStr(j), _
                                    InsertAsHyperlink:=True, _
                                    IncludePosition:=False, _
                                    SeparateNumbers:=False, _
                                    SeparatorString:=" "
                                On Error GoTo 0
                                ' Find the position of the "em space" ChrW(8195)
                                Dim SpacePosition As Integer
                                SpacePosition = InStr(para.Range.Text, ChrW(8195))
                                Dim SubstrLength As Integer
                                SubstrLength = SpacePosition - 1
                                
                                rng.MoveEnd Unit:=wdCharacter, count:=SubstrLength
                                rng.style = doc.Styles(C_S_CharHidden)
                                FlagFound = True
                                Exit For
                            End If
                        Next j
                    End If
                    If FlagFound = True Then
                        Exit For
                    End If
                Next i
            End If
        End If
        FlagFound = False
    Next para
    
    Set rng = Nothing
    Set doc = Nothing
    
    MsgBox _
        Prompt:="Finished processing.", _
        Buttons:=vbInformation + vbOKOnly, _
        Title:=MsgBoxTitle
    Exit Sub

Error_Insert:
    Set rng = Nothing
    Set doc = Nothing

    MsgBox _
        Prompt:="Error processing." & vbNewLine & _
                "Perhaps last paragraph after a reference list is not styled as 'Normal'?", _
        Buttons:=vbExclamation + vbOKOnly, _
        Title:=MsgBoxTitle
End Sub


Private Sub DeleteCrossReferences_References()
    Dim doc As Document
    Dim para As Paragraph
    Dim fld As Field
    Dim rng As Range
    Dim styleName As Variant
    Dim found As Boolean
    
    ' Set the document
    Set doc = ActiveDocument
    
    ' Loop through each paragraph in the document
    For Each para In doc.Paragraphs
        ' Check if the paragraph style is in the StyleNameArray
        Let found = False
        If Not para.Range.ListStyle Is Nothing Then
            If para.Range.ListStyle.NameLocal = C_S_ListNumRef Then
                found = True
            End If
        End If
        
        ' If the paragraph style is found in the array
        If found Then
            ' Check if the paragraph starts with a cross-reference field
            If para.Range.Fields.count > 0 Then
                Set fld = para.Range.Fields(1)
                If fld.Type = wdFieldRef Then
                    ' Remove the cross-reference field
                    fld.Delete
                End If
            End If
        End If
    Next para
    
    ' Clean up of the object variables
    Set doc = Nothing
    Set fld = Nothing
    
    Dim FileName As String
    FileName = C_F_Macros
    
    Dim ModuleName As String
    ModuleName = C_M_Tools
    
    Dim MacroName As String
    MacroName = "DeleteCrossReferences_References"
    
    Dim MsgBoxTitle As String
    MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    MsgBox _
        Prompt:="Finished processing.", _
        Buttons:=vbInformation + vbOKOnly, _
        Title:=MsgBoxTitle
End Sub

' Overtype mode, there are two options in configuration:
' 1. Word -> Options -> Advanced -> Editing Options -> Use the insert key to control overtype mode
' controlled by Options.INSKeyForOvertype
' 2. Word -> Options -> Advanced -> Editing Options -> Use overtype mode
' controlled by Options.Overtype
' The first one is just enabling relationship between overtype and insert keyboard key.
' The second could be used to force that state.
' There is also setting of the Status Bar, which cannot be accessed by VBA: Overtype
' 2025-04-20 by ms and AI
Sub CustomizedOvertype()
    Dim FileName As String
    FileName = C_F_Macros
    
    Dim ModuleName As String
    ModuleName = C_M_Tools
    
    Dim MacroName As String
    MacroName = "CustomizedOvertype"
    
    Dim MsgBoxTitle As String
    MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    Static InsertKeyState As Boolean
    
    If Options.INSKeyForOvertype Then
        InsertKeyState = Not InsertKeyState
        If InsertKeyState Then
            MsgBox _
                Prompt:="Overtype was enabled by pressing INS key!", _
                Buttons:=vbInformation + vbOKOnly, _
                Title:=MsgBoxTitle
            Options.Overtype = True
        ElseIf Not InsertKeyState Then
            MsgBox _
                Prompt:="Overtype was disabled by pressing INS key!", _
                Buttons:=vbInformation + vbOKOnly, _
                Title:=MsgBoxTitle
            Options.Overtype = False
        End If
    End If
End Sub

' https://stackoverflow.com/questions/47559316/macro-to-insert-a-cross-reference-based-on-selection
' 2025-04-19 tweaked by ms
' 2025-08-18 added additional conditions to build correct references.
Sub InsertCrossRef()
    Dim RefList As Variant
    Dim FullList As Variant
    Dim LookUp As String
    Dim Ref As String
    Dim s As Integer, t As Integer
    Dim i As Integer
    
    ' enum type
    Dim RefType As RefType

    Dim FileName As String
    FileName = C_F_Macros
    
    Dim ModuleName As String
    ModuleName = C_M_Tools
    
    Dim MacroName As String
    MacroName = "InsertCrossRef"
    
    Dim MsgBoxTitle As String
    MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName

    Application.statusBar = MsgBoxTitle & " > " & "is running..."

    On Error GoTo ErrExit
    With Selection.Range
        ' discard leading blank spaces
        Do While (Asc(.Text) = 32) And (.End > .Start)
            .MoveStart wdCharacter
        Loop
        ' discard trailing blank spaces, full stops and CRs
        ' 46 = pieriod, 32 = space, 11 = Vertical Tabulation, 13 = Carriage Return = \r
        Do While ((Asc(Right(.Text, 1)) = 46) Or _
                  (Asc(Right(.Text, 1)) = 32) Or _
                  (Asc(Right(.Text, 1)) = 11) Or _
                  (Asc(Right(.Text, 1)) = 13)) And _
                  (.End > .Start)
            .MoveEnd wdCharacter, -1
        Loop

ErrExit:
        If Len(.Text) = 0 Then
            MsgBox _
                Prompt:="Please select a cross-reference first.", _
                Buttons:=vbExclamation + vbOKOnly, _
                Title:=MsgBoxTitle
            Exit Sub
        End If

        LookUp = .Text
    End With
    On Error GoTo 0

    Dim LookUpLength As Integer
    Dim RefChar As String
    Let RefType = RefTypeNotDefined
    With ActiveDocument
        ' vbTextCompare: case insensitive comparison
        If InStr(1, LookUp, C_Caption_Pic, vbTextCompare) Then
            RefList = ActiveDocument.GetCrossReferenceItems(C_Caption_Pic)
            RefType = RefTypeC_Caption_Pic
            For i = UBound(RefList) To 1 Step -1
                Ref = Trim(RefList(i))
                If InStr(1, Ref, LookUp, vbTextCompare) = 1 Then
                    LookUpLength = Len(LookUp)
                    RefChar = Mid(Ref, LookUpLength + 1, 1)
                    If Not IsNumeric(RefChar) Then
                        Exit For
                    End If
                End If
            Next i
            If IndexNotFound(Index:=i, MsgBoxHeader:=MsgBoxTitle) Then
                Exit Sub
            End If
        End If
        
        If InStr(1, LookUp, C_Caption_Tab, vbTextCompare) Then
            RefList = ActiveDocument.GetCrossReferenceItems(C_Caption_Tab)
            RefType = RefTypeC_Caption_Tab
            For i = UBound(RefList) To 1 Step -1
                Ref = Trim(RefList(i))
                If InStr(1, Ref, LookUp, vbTextCompare) = 1 Then
                    LookUpLength = Len(LookUp)
                    RefChar = Mid(Ref, LookUpLength + 1, 1)
                    If Not IsNumeric(RefChar) Then
                        Exit For
                    End If
                End If
            If IndexNotFound(Index:=i, MsgBoxHeader:=MsgBoxTitle) Then
                Exit Sub
            End If
            Next i
        End If
        
        If InStr(1, LookUp, "[") > 0 And InStr(1, LookUp, "]") > 0 Then
            RefType = RefTypeReference
            FullList = ActiveDocument.GetCrossReferenceItems(wdRefTypeNumberedItem)
            For i = UBound(FullList) To 1 Step -1
                If InStr(1, FullList(i), LookUp) Then
                    Exit For
                End If
            Next i
            If IndexNotFound(Index:=i, MsgBoxHeader:=MsgBoxTitle) Then
                Exit Sub
            End If
        End If
        
        Dim StyleNameArray As Variant
        If RefType = RefTypeNotDefined Then
            Call InitializeStyleNameArray(StyleNameArray)
            RefList = BuildHeaderArray(StyleNameArray)
            RefType = RefTypeHeading
            FullList = ActiveDocument.GetCrossReferenceItems(wdRefTypeNumberedItem)
            LookUp = LookUp & ". "
            For i = UBound(RefList) To 1 Step -1
                If InStr(1, RefList(i), LookUp) = 1 Then
                    Exit For
                End If
            Next i
            If IndexNotFound(Index:=i, MsgBoxHeader:=MsgBoxTitle) Then
                Exit Sub
            End If
            
            ' Remove last character from RefList(i) and space characters
            LookUp = Trim(Left(RefList(i), Len(RefList(i)) - 1))
            For i = UBound(FullList) To 1 Step -1
                If InStr(1, FullList(i), LookUp) Then
                    Exit For
                End If
            Next i
        End If

        ' The following variables are used to expand size of a selection upon inserting a field
        Dim RngBefore As Range
        Dim RngAfter As Range
        
        If i Then
            Select Case RefType
                Case RefTypeC_Caption_Pic
                    ' Save the current selection range
                    Set RngBefore = Selection.Range.Duplicate
                    Selection.InsertCrossReference _
                                    ReferenceType:=C_Caption_Pic, _
                                    ReferenceKind:=wdOnlyLabelAndNumber, _
                                    ReferenceItem:=CStr(i), _
                                    InsertAsHyperlink:=True, _
                                    IncludePosition:=False, _
                                    SeparateNumbers:=False, _
                                    SeparatorString:=" "
                    ' Save the new range after insertion
                    Set RngAfter = Selection.Range.Duplicate
                    ' Expand the range to include the inserted field
                    RngBefore.End = RngAfter.End
                    
                Case RefTypeC_Caption_Tab
                    ' Save the current selection range
                    Set RngBefore = Selection.Range.Duplicate
                    Selection.InsertCrossReference _
                                    ReferenceType:=C_Caption_Tab, _
                                    ReferenceKind:=wdOnlyLabelAndNumber, _
                                    ReferenceItem:=CStr(i), _
                                    InsertAsHyperlink:=True, _
                                    IncludePosition:=False, _
                                    SeparateNumbers:=False, _
                                    SeparatorString:=" "
                    ' Save the new range after insertion
                    Set RngAfter = Selection.Range.Duplicate
                    ' Expand the range to include the inserted field
                    RngBefore.End = RngAfter.End

                Case RefTypeReference
                    ' Save the current selection range
                    Set RngBefore = Selection.Range.Duplicate
                    Selection.InsertCrossReference _
                                    ReferenceType:=wdRefTypeNumberedItem, _
                                    ReferenceKind:=wdNumberRelativeContext, _
                                    ReferenceItem:=CStr(i), _
                                    InsertAsHyperlink:=True, _
                                    IncludePosition:=False, _
                                    SeparateNumbers:=False, _
                                    SeparatorString:=" "
                    ' Save the new range after insertion
                    Set RngAfter = Selection.Range.Duplicate
                    ' Expand the range to include the inserted field
                    RngBefore.End = RngAfter.End
                
                Case RefTypeHeading
                    ' Save the current selection range
                    Set RngBefore = Selection.Range.Duplicate
                    Selection.InsertCrossReference _
                                    ReferenceType:=wdRefTypeNumberedItem, _
                                    ReferenceKind:=wdNumberRelativeContext, _
                                    ReferenceItem:=CStr(i), _
                                    InsertAsHyperlink:=True, _
                                    IncludePosition:=False, _
                                    SeparateNumbers:=False, _
                                    SeparatorString:=" "
                    ' Save the new range after insertion
                    Set RngAfter = Selection.Range.Duplicate
                    ' Expand the range to include the inserted field
                    RngBefore.End = RngAfter.End
            End Select
            
            ' Chages formatting of a new added field to hyperlink
            Dim aField As Field
            For Each aField In RngBefore.Fields
                Call RefFormatToHyperlink(aField)
            Next aField
            
        Else
            MsgBox _
                Prompt:="A cross reference to """ & LookUp & """ couldn't be set" & vbCr & _
                   "because a paragraph with that number couldn't" & vbCr & _
                   "be found in the document.", _
                Buttons:=vbInformation + vbOKOnly, _
                Title:=MsgBoxTitle
        End If
    End With
    
    ' Clear object variables
    Set RngBefore = Nothing
    Set RngAfter = Nothing
    
    Application.statusBar = False
End Sub

' 2025-08-02 by ms
' This is just small subset of code applied in macros InsertCrossRef -> Tools and RefToHyperlinks() -> Validation
Sub RefFormatToHyperlink(aField As Field)
    If aField.Type = wdFieldRef Then
        If InStr(aField.Code.Text, "_Ref") > 0 And aField.Code.Text Like "*\h*" = 0 Then
            aField.Code.InsertAfter (" \h")
        End If
        If (InStr(aField.Code.Text, "\h") Or InStr(aField.Code.Text, "\H")) Then
            If (aField.Code.Text Like "*\* MERGEFORMAT*" = -1) Then
                aField.Code.Text = Replace(aField.Code.Text, " \* MERGEFORMAT ", "", 1, -1, vbTextCompare)
                aField.Update
            End If
            If (aField.Code Like "*\* CHARFORMAT*" = 0) Then
                ' adds tag \*Charformat
                aField.Code.InsertAfter ("\* CHARFORMAT ")
                aField.Update
            End If
            aField.Select
            Selection.font.Underline = wdUnderlineSingle
            Selection.font.color = RGB(0, 130, 180) ' Surprisingly this doesn't work: Selection.font.color = wdThemeColorHyperlink
        End If
    End If
End Sub

Private Function IndexNotFound(Index As Integer, MsgBoxHeader As String) As Boolean
    If Index = 0 Then
        MsgBox _
            Prompt:="The specified index was not found." & vbNewLine & _
                    "Perhaps too high index? Exiting.", _
            Buttons:=vbCritical + vbOKOnly, _
            Title:=MsgBoxHeader
        IndexNotFound = True
    End If
End Function

' 2025-07-16 by ms
' This macro is equivalent of the "Reapply" button in the "Apply Styles" pane in Microsoft Word (opened via Ctrl + Shift + S)
Sub ReapplyTemplateStyle()
    Selection.style = Selection.style
End Sub

' 2025-07-16 by ms and AI
' Restart list numbering. Available only from a context menu.
Sub RestartListNumbering()
    Dim rng As Range
    Set rng = Selection.Range

    Dim FileName As String
    FileName = C_F_Macros
    
    Dim ModuleName As String
    ModuleName = C_M_Tools
    
    Dim MacroName As String
    MacroName = "RestartListNumbering"
    
    Dim MsgBoxTitle As String
    MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName

    If rng.ListFormat.ListType <> wdListNoNumbering Then
        rng.ListFormat.ApplyListTemplateWithLevel _
            ListTemplate:=rng.ListFormat.ListTemplate, _
            ContinuePreviousList:=False, _
            ApplyTo:=wdListApplyToWholeList, _
            DefaultListBehavior:=wdWord10ListBehavior
        Application.statusBar = MsgBoxTitle & " > " & "is running..."
    Else
            MsgBox _
                Prompt:="The current paragraph is not part of a numbered list.", _
                Buttons:=vbInformation + vbOKOnly, _
                Title:=MsgBoxTitle
    End If
End Sub

' 2025-08-21 by ms
Function CheckPageColor() As Boolean
    Dim CurrentColor As Long
    CurrentColor = ActiveDocument.Background.Fill.ForeColor.RGB

    If CurrentColor = RGB(219, 219, 219) Then
        CheckPageColor = True
    Else
        CheckPageColor = False
    End If
End Function

' Sets document page background color to make it easier to your eyes on time of editing.
' https://superuser.com/questions/854808/word-macro-to-change-page-color
' 2025-07-29 by ms
' 2025-08-05 by ms
Sub SetPageColorToCustom()
    ActiveDocument.ActiveWindow.View.DisplayBackgrounds = True
    With ActiveDocument.Background.Fill
        .Visible = msoTrue
        .Solid
        .ForeColor.RGB = RGB(219, 219, 219) ' custom color: grey
    End With
End Sub

' Restores default page background color. Complementary to SetPageColorToCustom()
' 2025-07-29 by ms
Sub RestoreDefaultPageColor()
    ' This macro resets the page background color to default (no fill)
    With ActiveDocument.Background.Fill
        .Visible = msoFalse
    End With
End Sub

' 2025-08-02 by ms and AI.
' Toggle heading collapse or expand.
Sub ToggleHeadingCollapseExpand()
    Dim para As Paragraph
    Set para = Selection.Paragraphs(1)
    
    If para.style Like "Heading*" Then
        para.CollapsedState = Not para.CollapsedState
    Else
        Dim FileName As String
        FileName = C_F_Macros
        
        Dim ModuleName As String
        ModuleName = C_M_Tools
        
        Dim MacroName As String
        MacroName = "ToggleHeadingCollapseExpand"
        
        Dim MsgBoxTitle As String
        MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
        
        MsgBox _
            Prompt:="The current paragraph is not a heading", _
            Buttons:=vbInformation + vbOKOnly, _
            Title:=MsgBoxTitle
    End If
    
    ' Clear the object variable
    Set para = Nothing
End Sub

' 2025-08-14 by ms
' Check if Microsoft Word printing options, level Display are set as expected:
'       .PrintDrawingObjects = True  ' Print drawings created in Word
'       .PrintBackground = True      ' Print background colors and images
'       .PrintProperties = False     ' Print document properties
'       .PrintHiddenText = False     ' Print hidden text
'       .UpdateFieldsAtPrint = True  ' Update fields before printing
'       .UpdateLinksAtPrint = True   ' Update linked data before printing
Private Function CheckPrintingOptionsDisplay() As Boolean
    CheckPrintingOptionsDisplay = True
    If Not Options.PrintDrawingObjects = True Then
        CheckPrintingOptionsDisplay = False
    End If
    If Not Options.PrintBackground = True Then
        CheckPrintingOptionsDisplay = False
    End If
    If Not Options.PrintProperties = False Then
        CheckPrintingOptionsDisplay = False
    End If
    If Not Options.PrintHiddenText = False Then
        CheckPrintingOptionsDisplay = False
    End If
    If Not Options.UpdateFieldsAtPrint = True Then
        CheckPrintingOptionsDisplay = False
    End If
    If Not Options.UpdateLinksAtPrint = True Then
        CheckPrintingOptionsDisplay = False
    End If
End Function

' 2025-08-03 by ms and AI
' Default printing options as in File > Word Options > Display
' Set default printing options of Microsoft Word, level Display
Private Sub SetPrintingOptionsDisplay()
    With Options
        .PrintDrawingObjects = True  ' Print drawings created in Word
        .PrintBackground = True      ' Print background colors and images
        .PrintProperties = False     ' Print document properties
        .PrintHiddenText = False     ' Print hidden text
        .UpdateFieldsAtPrint = True  ' Update fields before printing
        .UpdateLinksAtPrint = True   ' Update linked data before printing
    End With

    Dim FileName As String
    FileName = C_F_Macros
    
    Dim ModuleName As String
    ModuleName = C_M_Tools
    
    Dim MacroName As String
    MacroName = "SetPrintingOptionsDisplay"
    
    Dim MsgBoxTitle As String
    MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName

    MsgBox _
        Prompt:="The following printing options, level Display, have been set, context Application:" & vbNewLine & vbNewLine & _
            "Print drawings created in Word: " & Options.PrintDrawingObjects & vbNewLine & _
            "Print background colors and images: " & Options.PrintBackground & vbNewLine & _
            "Print document properties: " & Options.PrintProperties & vbNewLine & _
            "Print hidden text: " & Options.PrintHiddenText & vbNewLine & _
            "Update fields before printing: " & Options.UpdateFieldsAtPrint & vbNewLine & _
            "Update linked data before printing: " & Options.UpdateLinksAtPrint, _
        Buttons:=vbInformation + vbOKOnly, _
        Title:=MsgBoxTitle

End Sub

' Check expected printing settings, level Advanced.
' 2025-08-14 by ms
Private Function CheckPrintingOptionsAdvanced() As Boolean
    CheckPrintingOptionsAdvanced = True
    If Not Options.PrintDraft = False Then
        CheckPrintingOptionsAdvanced = False
    End If
    If Not Options.PrintBackground = True Then
        CheckPrintingOptionsAdvanced = False
    End If
    If Not Options.PrintReverse = False Then
        CheckPrintingOptionsAdvanced = False
    End If
    If Not Options.PrintXMLTag = False Then
        CheckPrintingOptionsAdvanced = False
    End If
    If Not Options.PrintFieldCodes = False Then
        CheckPrintingOptionsAdvanced = False
    End If
    If Not Options.PrintOddPagesInAscendingOrder = False Then
        CheckPrintingOptionsAdvanced = False
    End If
    If Not Options.PrintEvenPagesInAscendingOrder = False Then
        CheckPrintingOptionsAdvanced = False
    End If
    If IsNumeric(Options.DefaultTray) Then
        If Options.DefaultTray <> wdPrinterDefaultBin Then
            CheckPrintingOptionsAdvanced = False
        End If
    End If
End Function

' 2025-08-03 by ms and AI
' Default printing options as in File > Word Options > Advanced > Print
' Set default printing options of Microsoft Word, level Advanced.
Private Sub SetPrintingOptionsAdvanced()
    With Options
        .PrintDraft = False                         ' Use draft quality
        .PrintBackground = True                     ' Print in background
        .PrintReverse = False                       ' Print pages in reverse order
        .PrintXMLTag = False                        ' Print XML tags
        .PrintFieldCodes = False                    ' Print field codes instead of their values
        ' Allow fields containing tracked changes to update before printing, not available in VBA
        .PrintOddPagesInAscendingOrder = False      ' Print on front of the sheet for duplex printing
        .PrintEvenPagesInAscendingOrder = False     ' Print on back of the sheet for duplex printing
        ' Scale content for A4 or 8.5 x 11'' paper size, not available in VBA
        .DefaultTray = wdPrinterDefaultBin          ' Default tray: User printer settings
    End With
    
    Dim FileName As String
    FileName = C_F_Macros
    
    Dim ModuleName As String
    ModuleName = C_M_Tools
    
    Dim MacroName As String
    MacroName = "SetPrintingOptionsAdvanced"
    
    Dim MsgBoxTitle As String
    MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName

    MsgBox _
        Prompt:="The following printing options, level Advanced, have been set, context Application:" & vbNewLine & vbNewLine & _
            "Use draft quality: " & Options.PrintDraft & vbNewLine & _
            "Print in background: " & Options.PrintBackground & vbNewLine & _
            "Print pages in reverse order: " & Options.PrintReverse & vbNewLine & _
            "Print XML tags: " & Options.PrintXMLTag & vbNewLine & _
            "PrintFieldCodes: " & Options.PrintFieldCodes & vbNewLine & _
            "Print on front of the sheet for duplex printing: " & Options.PrintOddPagesInAscendingOrder & vbNewLine & _
            "Print on back of the sheet for duplex printing: " & Options.PrintEvenPagesInAscendingOrder & vbNewLine & _
            "Default tray: User printer settings: " & Options.DefaultTray, _
        Buttons:=vbInformation + vbOKOnly, _
        Title:=MsgBoxTitle
End Sub

' 2025-08-03 by ms
' 2025-08-06 by ms
' Customized printing.
Sub CustomizedPrinting()
    Dim UserDecision As VbMsgBoxResult
    
    Dim FileName As String
    FileName = C_F_Macros
    
    Dim ModuleName As String
    ModuleName = C_M_Tools
    
    Dim MacroName As String
    MacroName = "CustomizedPrinting"
    
    Dim MsgBoxTitle As String
    MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    Application.statusBar = MsgBoxTitle & " > " & "is running..."
    
    If Not CheckPrintingOptionsDisplay() Then
        Beep
        UserDecision = MsgBox( _
            Prompt:="Would you like to set printing options, level Display?" & vbNewLine & vbNewLine & _
                "It is strongly recommended to do so.", _
            Buttons:=vbQuestion + vbYesNo + vbDefaultButton1, _
            Title:=MsgBoxTitle _
            )
        If UserDecision = vbYes Then
            Call SetPrintingOptionsDisplay
        End If
    End If
    
    If Not CheckPrintingOptionsAdvanced() Then
        Beep
        UserDecision = MsgBox( _
            Prompt:="Would you like to set printing options, level Advanced?" & vbNewLine & vbNewLine & _
                "It is strongly recommended to do so.", _
            Buttons:=vbQuestion + vbYesNo + vbDefaultButton1, _
            Title:=MsgBoxTitle _
            )
        If UserDecision = vbYes Then
            Call SetPrintingOptionsAdvanced
        End If
    End If
    
    'If there is set a background color, ask user if not to change it to default one before printing.
    If ActiveDocument.Background.Fill.Visible = msoTrue Then
        UserDecision = MsgBox( _
            Prompt:="Page background color is set to custom color:" & vbNewLine & vbNewLine & _
                GetColorString(ActiveDocument.Background.Fill.ForeColor.RGB) & vbNewLine & vbNewLine & _
                "Would you like to set default color (white) before printing?", _
            Buttons:=vbQuestion + vbYesNo + vbDefaultButton1, _
            Title:=MsgBoxTitle _
            )
    Else
        Call UpdateAllFields    ' module: Validation
        ' Show the Print Dialog Box via VBA. Unfortunately it is not possible from VBA to set by default "print what" to "document".
        Dialogs(wdDialogFilePrint).Show
        Exit Sub
    End If
    
    If UserDecision = vbYes Then
        ActiveDocument.ActiveWindow.View.DisplayBackgrounds = True
        ActiveDocument.Background.Fill.Visible = msoFalse
    End If
    
    Call UpdateAllFields    ' module: Validation
    ' Show the Print Dialog Box via VBA. Unfortunately it is not possible from VBA to set by default "print what" to "document".
    Dialogs(wdDialogFilePrint).Show
End Sub

' 2025-08-03 by ms
' Customized 'Save As' to enable printing to PDF.
Sub CustomizedSaveAs()
    Dim UserDecision As VbMsgBoxResult
    
    Dim FileName As String
    FileName = C_F_Macros
    
    Dim ModuleName As String
    ModuleName = C_M_Tools
    
    Dim MacroName As String
    MacroName = "CustomizedSaveAs"
    
    Dim MsgBoxTitle As String
    MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    Beep
    UserDecision = MsgBox( _
        Prompt:="Would you like to save current file in PDF format?", _
        Buttons:=vbQuestion + vbYesNo + vbDefaultButton2, _
        Title:=MsgBoxTitle _
        )
    If UserDecision = vbYes Then
        Call SaveDocumentAsPDFWithSettings
        Exit Sub
    End If
    
    ' Show the Save As via VBA
    Dialogs(wdDialogFileSaveAs).Show
End Sub

' This macro is run automatically. It entables automatic running of sub from Class Module ClsAppEvents
' 2025-08-28 by ms
Sub AutoExec()
    Set WordAppEvents = New ClsAppEvents
    Set WordAppEvents.appWord = Word.Application
End Sub

' 2025-10-02 by ms
Sub CustomizedCopyFormat()
    Dim FileName As String
    FileName = C_F_Macros
    
    Dim ModuleName As String
    ModuleName = C_M_Tools
    
    Dim MacroName As String
    MacroName = "CustomizedCopyFormat"
    
    Dim MsgBoxTitle As String
    MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    Application.Run "CopyFormat"    ' call built-in Microsoft Word command
    Application.statusBar = MsgBoxTitle & " > " & C_SC_ShiftCtrlC
End Sub

' 2025-10-02 by ms
Sub CustomizedPasteFormat()
    Dim FileName As String
    FileName = C_F_Macros
    
    Dim ModuleName As String
    ModuleName = C_M_Tools
    
    Dim MacroName As String
    MacroName = "CustomizedPasteFormat"
    
    Dim MsgBoxTitle As String
    MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    Application.Run "PasteFormat"   ' call built-in Microsoft Word command
    Application.statusBar = MsgBoxTitle & " > " & C_SC_ShiftCtrlV
End Sub


' Counts all templates and shows full path of all templates.
' Added display of template type. If only Normal.dotm is attached to currently ActiveDocument, then its type is set to 'wdNormalTemplate', not 'wdAttachedTemplate'.
' 2025-08-03 by ms
' 2025-12-11 by ms
Sub ShowAllTemplates()
    Dim TemplateDoc As Template
    Dim TemplateCollection As String
    
    Dim FileName As String
    FileName = C_F_Macros
    
    Dim ModuleName As String
    ModuleName = C_M_Tools
    
    Dim MacroName As String
    MacroName = "ShowAllTemplates"
    
    Dim MsgBoxTitle As String
    MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    Dim TemplateCounter As Byte
    TemplateCounter = 1
    For Each TemplateDoc In Templates
        TemplateCollection = TemplateCollection & TemplateCounter & ". " & TemplateDoc.Name & " type: " & TemplateTypeName(TemplateDoc.Type) & vbNewLine
        TemplateCollection = TemplateCollection & TemplateDoc.Path & vbNewLine & vbNewLine
        TemplateCounter = TemplateCounter + 1
    Next TemplateDoc

    MsgBox _
        Prompt:="Templates collection:" & vbNewLine & vbNewLine & TemplateCollection, _
        Buttons:=vbInformation + vbOKOnly, _
        Title:=MsgBoxTitle
End Sub

' 2025-12-11 by ms and AI
Private Function TemplateTypeName(ByVal tt As WdTemplateType) As String
    ' Maps WdTemplateType values to their enum names
    Select Case tt
        Case wdAttachedTemplate
            TemplateTypeName = "wdAttachedTemplate"
        Case wdGlobalTemplate
            TemplateTypeName = "wdGlobalTemplate"
        Case wdNormalTemplate
            TemplateTypeName = "wdNormalTemplate"
        Case Else
            TemplateTypeName = "Unknown"
    End Select
End Function

' Combo macro. Alternative solution to macro InsertCrossRef, which is called by the F7 keyboard shortcut.
' It loops over all paragraphs in the document and then exchanges specific character strings with references.
' 2025-04-18 by ms
Sub InsertCrossReferences()
    Call InsertCrossReferences_Headings     ' in module Tools
    DoEvents    ' Force a screen refresh
    Call InsertCrossReferences_Pictures     ' in module Tools
    DoEvents    ' Force a screen refresh
    Call InsertCrossReferences_Tables       ' in module Tools
    DoEvents    ' Force a screen refresh
    Call InsertCrossReferences_References   ' in module Tools
End Sub

' 2025-04-18 by ms
Sub DeleteCrossReferences()
    Call DeleteCrossReference_Headings      ' in module Tools
    DoEvents    ' Force a screen refresh
    Call DeleteCrossReferences_Pictures     ' in module Tools
    DoEvents    ' Force a screen refresh
    Call DeleteCrossReferences_Tables       ' in module Tools
    DoEvents    ' Force a screen refresh
    Call DeleteCrossReferences_References   ' in module Tools
End Sub

' Formats the currently selected table, but only if the selection
' is EXACTLY that table (no extra text/paragraphs outside the table).
' 2025-12-05 by ms and AI
' 2025-12-21 by ms and AI
Public Sub Table_CustomizeFormatting()
    On Error GoTo ErrHandler
    
    Dim sel As Selection
    Dim tbl As Word.Table
    Dim sty As Word.style
    Dim rng As Word.Range
    Dim isExactSelection As Boolean
    
    Set sel = Selection
    Set rng = sel.Range
    
    Dim FileName As String
    FileName = C_F_Macros
    
    Dim ModuleName As String
    ModuleName = C_M_Tools
    
    Dim MacroName As String
    MacroName = "Table_CustomizeFormatting"
    
    Dim MsgBoxTitle As String
    MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    ' === 1) Resolve a table from the current selection ===
    Set tbl = GetTableFromSelection(sel, isExactSelection)
    If tbl Is Nothing Then
        MsgBox _
            Prompt:="Selection must be the TABLE ONLY (no extra paragraphs before/after) or place cursor within a table. " & vbCrLf & _
               "Tip: Use Table Tools > Layout > Select > Select Table, or click the table handle.", _
            Buttons:=vbExclamation, _
            Title:=MsgBoxTitle
        Exit Sub
    End If
    
    ' === 2) Prevent table rows from breaking across pages
    tbl.Rows.AllowBreakAcrossPages = False
    
    ' === 3) Repeat header for the first row ===
'    If tbl.Rows.count >= 1 Then
'        tbl.Rows(1).HeadingFormat = True
'    End If
     tbl.Range.Cells(1).Range.Rows.HeadingFormat = True ' old version
    
    ' === 4) AutoFit to Window (fit table width to page) ===
    tbl.AutoFitBehavior wdAutoFitWindow   ' makes the table fit the page width

    ' === 5) Center content of all cells vertically ===
    tbl.Range.Cells.VerticalAlignment = wdCellAlignVerticalCenter
    
    ' === 6) Apply style C_S_ParInTable to the content in all cells ===
    ' We detect the style type and apply appropriately.
    On Error Resume Next
    Set sty = ActiveDocument.Styles(C_S_ParInTable)
    On Error GoTo 0
    
    If sty Is Nothing Then
        MsgBox _
            Prompt:="Style '" & C_S_ParInTable & "' was not found in this document/template. " & _
               "Please create it first, then rerun the macro.", _
            Buttons:=vbExclamation, _
            Title:=MsgBoxTitle
        Exit Sub
    Else
        tbl.Range.style = sty
    End If
    
    MsgBox _
        Prompt:="Done. The selected table has been formatted according to set of customized rules." & vbNewLine & vbNewLine & _
            ".AllowBreakAcrossPage = False" & vbNewLine & _
            ".HeadingFormat = True" & vbNewLine & _
            ".AutoFitBehavior wdAutoFitWindow" & vbNewLine & _
            ".AllowBreakAcrossPages = False" & vbNewLine & _
            ".Cells.VerticalAlignment = wdCellAlignVerticalCenter" & vbNewLine & _
            " cell content style: " & C_S_ParInTable, _
        Buttons:=vbInformation, _
        Title:=MsgBoxTitle
    Exit Sub

ErrHandler:
    MsgBox _
        Prompt:="Error " & Err.Number & ": " & Err.Description, _
        Buttons:=vbExclamation, _
        Title:=MsgBoxTitle
End Sub
        

' 2025-12-21 by ms and ai
Private Function GetTableFromSelection(ByVal sel As Selection, ByRef isExact As Boolean) As Word.Table
    Dim t As Word.Table
    Dim sStart As Long, sEnd As Long
    Dim tStart As Long, tEnd As Long
    
    isExact = False
    
    ' Case 1: direct selection of a table (Selection.Tables.Count > 0)
    If sel.Tables.count > 0 Then
        Set t = sel.Tables(1)
        ' Determine if selection equals table's full range (including end mark)
        sStart = sel.Range.Start: sEnd = sel.Range.End
        tStart = t.Range.Start:   tEnd = t.Range.End
        isExact = (sStart = tStart And sEnd = tEnd)
        Set GetTableFromSelection = t
        Exit Function
    End If
    
    ' Case 2: insertion point or selection within a table
    If sel.Information(wdWithInTable) Then
        Set GetTableFromSelection = sel.Range.Tables(1)
        Exit Function
    End If
    
    ' No table found
    Set GetTableFromSelection = Nothing
End Function


' 2025-12-05 by ms and AI
Public Sub Table_KeepOnOnePage()
    Dim FileName As String
    FileName = C_F_Macros
    
    Dim ModuleName As String
    ModuleName = C_M_Tools
    
    Dim MacroName As String
    MacroName = "Table_KeepOnOnePage"
    
    Dim MsgBoxTitle As String
    MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    If IsSelectedJustTable() = False Then
        Exit Sub
    End If
    
    If Selection.Tables.count = 0 Then
        MsgBox _
            Prompt:="Place the cursor inside a table first.", _
            Buttons:=vbExclamation, _
            Title:=MsgBoxTitle
        Exit Sub
    End If
    Selection.Tables(1).Range.ParagraphFormat.KeepWithNext = True
End Sub

' Verify the selection is exactly one table (and nothing else)
' 2025-12-09 by ms
Private Function IsSelectedJustTable() As Boolean
    Dim FileName As String
    FileName = C_F_Macros
    
    Dim ModuleName As String
    ModuleName = C_M_Tools
    
    Dim MacroName As String
    MacroName = "IsSelectedJustTable"
    
    Dim MsgBoxTitle As String
    MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
  
    If Selection.Tables.count <> 1 Then
        MsgBox _
            Prompt:="Please select exactly ONE table, and nothing else.", _
            Buttons:=vbExclamation, _
            Title:=MsgBoxTitle
        IsSelectedJustTable = False
        Exit Function
    End If
    
    Dim isExactSelection As Boolean
    ' The selection must match the table's range exactly
    isExactSelection = (Selection.Range.Start = Selection.Tables(1).Range.Start) And (Selection.Range.End = Selection.Tables(1).Range.End)
    If Not isExactSelection Then
        MsgBox _
            Prompt:="Selection must be the TABLE ONLY (no extra paragraphs before/after). " & vbCrLf & _
               "Tip: Use Table Tools > Layout > Select > Select Table, or click the table handle.", _
            Buttons:=vbExclamation, _
            Title:=MsgBoxTitle
        IsSelectedJustTable = False
        Exit Function
    End If
    IsSelectedJustTable = True
End Function
