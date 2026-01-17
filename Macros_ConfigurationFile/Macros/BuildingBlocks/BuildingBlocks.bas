Attribute VB_Name = "BuildingBlocks"
' VBA Module name: BuildingBlocks.bas
' https://github.com/mslonik/Microsoft-Word-Configuration
'
'   License: MIT License.
'
'+---------+-----------------------------+-------------------+------------------+----------------------------------------+
'| No.     | Sub name                    | Ribbon name       | Ribbon section   | Ribbon button name                     |
'+---------+-----------------------------+-------------------+------------------+----------------------------------------+
'| 1       | BB_ExportAll                | BuildingBlocks_ms | no name (custom) | BB_ExportAll                           |
'| 2       | BB_ExportSelectedCategories | BuildingBlocks_ms | no name (custom) | BB_ExportSelectedCategories            |
'| 3       | BB_OpenBuiltInTemplate      | BuildingBlocks_ms | no name (custom) | BB_OpenBuiltInTemplate                 |
'| 4       | BB_DeleteAll                | BuildingBlocks_ms | no name (custom) | BB_DeleteAll                           |
'| 5       | BB_List                     | BuildingBlocks_ms | no name (custom) | BB_List                                |
'| 6       | BB_Add                      | BuildingBlocks_ms | Edit             | BB_Add                                 |
'| 7       | BB_InsertBBTemplate         | BuildingBlocks_ms | Edit             | BB_InsertBBTemplate                    |
'| 8       | BB_RemoveDefParagraphs      | BuildingBlocks_ms | Edit             | BB_RemoveDefParagraphs                 |
'| 9       | —                           | BuildingBlocks_ms | built-in         | Building Block Gallery Content Control |
'| 10      | —                           | BuildingBlocks_ms | built-in         | Building Block Organizer               |
'+---------+-----------------------------+-------------------+------------------+----------------------------------------+
'
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
' Used to enforce the explicit declaration of all variables in a module. When you include Option Explicit at the beginning of a module, it ensures that you must
' declare all variables using the Dim, Private, Public, ReDim, or Static statements before using them. This helps prevent errors caused by typos or
' undeclared variables.
Option Explicit
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
' 2025-12-18 by ms
Sub BB_OpenBuiltInTemplate()
    Dim FileName As String:     FileName = C_F_Macros
    Dim ModuleName As String:   ModuleName = C_M_BuildingBlocks
    Dim MacroName As String:    MacroName = "BB_OpenBuiltInTemplate"
    Dim MsgBoxTitle As String:  MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName

    Dim UserName As String
    ' Get the currently logged user name
    UserName = Environ("USERNAME")

    Dim CurrentLanguageID As MsoLanguageID
    ' Get the Office authoring languages and proofing language ID
    CurrentLanguageID = Application.LanguageSettings.LanguageId(msoLanguageIDInstall)
    
    ' Local constants, to make code shorter
    Const C_Path1 As String = "C:\Users\"
    Const C_Path2 As String = "\AppData\Roaming\Microsoft\Document Building Blocks\"
    
    Dim StringMajorVersion As String
    Dim IntMajorVersion As Integer
    ' Determine the Office major version
    StringMajorVersion = Application.Version
    IntMajorVersion = CInt(Split(StringMajorVersion, ".")(0))

    Dim FilePath As String
    ' Construct the file path
    FilePath = C_Path1 & UserName & C_Path2 & CStr(CurrentLanguageID) & "\" & CStr(IntMajorVersion) & "\" & C_F_BBB

    ' Ask the user to confirm the file path
    Dim UserDecision As VbMsgBoxResult
    Beep
    UserDecision = MsgBox( _
        Prompt:="The following file has beend determined: " & vbNewLine & vbNewLine & _
            FilePath & _
            vbNewLine & vbNewLine & "Is this correct?", _
        Buttons:=vbYesNo + vbQuestion, _
        Title:=MsgBoxTitle)

    If UserDecision = vbYes Then
        If Len(Dir$(FilePath)) > 0 Then
            Application.Documents.Open _
                FileName:=FilePath, _
                ReadOnly:=False, _
                AddToRecentFiles:=True
        Else
            MsgBox _
                Prompt:="There was a problem opening the file" & vbNewLine & vbNewLine & _
                    FilePath & vbNewLine & vbNewLine & _
                    "Exiting.", _
                Buttons:=vbError, _
                Title:=MsgBoxTitle
            Exit Sub
        End If
    End If

End Sub

' Delete all building blocks from currently opened template file.
' Before running this macro it is recommended to rename the default "Built-In Building Blocks.dotx" into e.g. "MicrosoftWordDefaultBB.dotx", then recreate (restart)
' the Microsoft Word and run this macro. This way you will maintain access to set of default Building Blocks if necessary in future, without need to manipulate additional files and macros.
' 2025-12-18 by ms
Sub BB_DeleteAll()
    Dim FileName As String:     FileName = C_F_Macros
    Dim ModuleName As String:   ModuleName = C_M_BuildingBlocks
    Dim MacroName As String:    MacroName = "BB_DeleteAll"
    Dim MsgBoxTitle As String:  MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    Dim IsTemplate As Boolean
    ' Check if currently ActiveDocument is a template file
    IsTemplate = CheckIfActiveDocumentIsTemplate()
    If Not IsTemplate Then
        MsgBox _
            Prompt:="Currently active document is not a template file. It cannot contain BuildingBlocks." & vbNewLine & vbNewLine & _
                "Exiting.", _
            Buttons:=vbCritical, _
            Title:=MsgBoxTitle
        Exit Sub
    End If

    ' Set the context to the currently opened file
    Application.CustomizationContext = ActiveDocument
    
    ' The following line refreshes all the building blocks. It is essential for proper work of this Sub.
    Templates.LoadBuildingBlocks
    
    Dim UserDecision As VbMsgBoxResult
    Dim MyTemplate As Template
    Beep
    UserDecision = MsgBox( _
        Prompt:="Are you sure you want to delete all BuildingBlocks from the file" & vbNewLine & vbNewLine & _
            ActiveDocument.FullName & "?" & vbNewLine & vbNewLine & _
            "This operation will delete all BuildingBlocks permanently.", _
        Buttons:=vbYesNo + vbQuestion, _
        Title:=MsgBoxTitle)
    
    ' If the user confirms, proceed to load the template and delete building blocks
    If UserDecision = vbYes Then
        On Error Resume Next
        Set MyTemplate = Templates(ActiveDocument.FullName)
        On Error GoTo 0

        ' Check if the template is loaded successfully
        If MyTemplate Is Nothing Then
            MsgBox _
                Prompt:="The" & vbNewLine & vbNewLine & _
                    C_F_BBB & vbNewLine & vbNewLine & _
                    " template could not be loaded." & vbNewLine & _
                    "Was it deleted?", _
                Buttons:=vbExclamation, _
                Title:=MsgBoxTitle
            Exit Sub
        End If

        Dim j As Integer
        Dim bb As BuildingBlock
        Dim TotalNoBB
        TotalNoBB = MyTemplate.BuildingBlockEntries.count
        If TotalNoBB = 0 Then
            MsgBox _
                Prompt:="Nothing to delete. Exiting.", _
                Buttons:=vbCritical, _
                Title:=MsgBoxTitle
            Exit Sub
        End If
        
        ' Loop through all building blocks and delete them
        For j = TotalNoBB To 1 Step -1
            Set bb = MyTemplate.BuildingBlockEntries(j)
            bb.Delete
        Next j

        ActiveDocument.Save
        MsgBox _
            Prompt:="All " & TotalNoBB & " Building Blocks have been deleted from the" & vbNewLine & vbNewLine & _
                C_F_BBB & vbNewLine & vbNewLine & _
                "template file." _
                & vbNewLine & vbNewLine & _
                "The file was saved.", _
             Buttons:=vbInformation, _
             Title:=MsgBoxTitle
    Else
        MsgBox _
            Prompt:="Operation canceled by the user.", _
            Buttons:=vbInformation, _
            Title:=MsgBoxTitle
    End If
    
    ' Clear object variables
    Set MyTemplate = Nothing
    Set bb = Nothing
End Sub

' Function to check if currently opened file is a template file.
' 2025-12-11 by ms and AI
Private Function CheckIfActiveDocumentIsTemplate() As Boolean
    Dim DocName As String
    Dim IsTemplate As Boolean
    
    ' Get the active document name
    DocName = ActiveDocument.Name
    CheckIfActiveDocumentIsTemplate = False
    
    ' Check file extension for template types
    If LCase(Right(DocName, 4)) = ".dot" Or _
       LCase(Right(DocName, 5)) = ".dotx" Or _
       LCase(Right(DocName, 5)) = ".dotm" Then
        CheckIfActiveDocumentIsTemplate = True
    Else
        CheckIfActiveDocumentIsTemplate = False
    End If
End Function

' Function returns indes TemplatesIndex of the argument: templatefilename.
'2025-12-27 by ms
Public Function GetTemplateIndex(TemplateFilename As String) As Integer
    Dim i As Integer
    Dim TemplateIndex As Integer
    Dim UserDecision As VbMsgBoxResult
    
    Dim FileName As String:     FileName = C_F_Macros
    Dim ModuleName As String:   ModuleName = C_M_BuildingBlocks
    Dim MacroName As String:    MacroName = "GetTemplateIndex"
    Dim MsgBoxTitle As String:  MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    ' Instead of referencing the template filename I need to reference through number.
    For i = Application.Templates.count To 1 Step -1
        If Application.Templates(i).Name = TemplateFilename Then
            TemplateIndex = i
            GetTemplateIndex = TemplateIndex
            Exit For
        End If
    Next i
        
End Function

' Function returns index AddInsIndex of the argument: template filename.
' 2025-12-12 by ms
Public Function GetAddinIndex(TemplateFilename As String) As Integer
    Dim i As Integer
    Dim AddInsIndex As Integer
    Dim UserDecision As VbMsgBoxResult
    
    Dim FileName As String:     FileName = C_F_Macros
    Dim ModuleName As String:   ModuleName = C_M_BuildingBlocks
    Dim MacroName As String:    MacroName = "GetAddinIndex"
    Dim MsgBoxTitle As String:  MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    ' Force loading of all BuildingBlocks
    Templates.LoadBuildingBlocks
    ' Instead of referencing the template filename I need to reference through number.
    For i = AddIns.count To 1 Step -1
        If AddIns(i).Name = TemplateFilename Then
            AddInsIndex = i
            GetAddinIndex = AddInsIndex
            Exit For
        End If
    Next i
    
    If AddInsIndex = 0 Then
        GetAddinIndex = AddInsIndex
        Exit Function
    End If

    ' Ask user if to enable TemplateFilename
    If Not AddIns(AddInsIndex).Installed Then
        Beep
        UserDecision = MsgBox( _
            Prompt:="The " & TemplateFilename & " is found, but not enabled." & vbNewLine & _
                "Would you like to enable it now?", _
            Buttons:=vbYesNo + vbQuestion, _
            Title:=MsgBoxTitle)
        If UserDecision = vbYes Then
            AddIns(AddInsIndex).Installed = True
        Else
            GetAddinIndex = -1  ' error
            Exit Function
        End If
    End If
    
End Function

' This macro works only for template files (DOTM / DOTX). If run in any other file, it will exit.
' Export all building blocks from currently attached template to a separate DOCX file.
' Output filename: TemplateName & "_BB_Content.docx".
' Output filepath: the same as for opened DOTM / DOTX file.
' 2025-02-17 by ms
' 2025-12-12 by ms
Sub BB_ExportAll()
    Dim TemplateIndex As Integer
    Dim bb As BuildingBlock
    Dim newDoc As Document
    Dim Rng As Range
    Dim bbe As BuildingBlockEntries
    Dim i As Integer
    Dim CategoryCount As Object
    
    Dim FileName As String:      FileName = C_F_Macros
    Dim ModuleName As String:    ModuleName = C_M_BuildingBlocks
    Dim MacroName As String:     MacroName = "BB_ExportAll"
    Dim MsgBoxTitle As String:   MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    Dim IfActiveDocument As Boolean
    IfActiveDocument = CheckIfActiveDocumentIsTemplate() ' private function in this module
    If Not IfActiveDocument Then
        MsgBox _
            Prompt:="Currently active document " & vbNewLine & vbNewLine & _
                ActiveDocument.Name & vbNewLine & vbNewLine & _
                "is not a template file. It cannot contain BuildingBlocks." & vbNewLine & vbNewLine & _
                "Exiting.", _
            Buttons:=vbCritical, _
            Title:=MsgBoxTitle
        Exit Sub
    End If
    
    Dim MainTemplatePath As String  ' Define variables for the full paths of the attached template
    Dim MainTemplateName As String
    Dim MainTemplate As Template
    MainTemplatePath = ActiveDocument.FullName
    MainTemplateName = ActiveDocument.Name
    Set MainTemplate = ActiveDocument.AttachedTemplate
    
    ' Initialize the category count dictionary
    Set CategoryCount = CreateObject("Scripting.Dictionary")
    
    ' Create a new blank document
    Set newDoc = Documents.Add
    
    ' Attach the default template to the new document
    newDoc.AttachedTemplate = C_F_Normal
    
    Dim NewFilename As String
    Dim SavePath As String
    ' Get the template file name without extension
    NewFilename = Left(MainTemplateName, InStrRev(MainTemplateName, ".") - 1)
    
    ' Get the default local file location
    SavePath = Options.DefaultFilePath(wdDocumentsPath) & "\" & NewFilename & "_BB_Content.docx"
    
    ' Make sure the new document contains all "ms" styles
    Call CreateCustomStyles     ' in module Styles
        
    Set bbe = MainTemplate.BuildingBlockEntries
    If bbe.count = 0 Then
        MsgBox _
            Prompt:="No BuildingBlocks found in the temlate file" & vbNewLine & vbNewLine & _
                MainTemplateName & vbNewLine & vbNewLine & _
                "Exiting.", _
            Buttons:=vbCritical + vbOKOnly, _
            Title:=MsgBoxTitle
        Exit Sub
    End If
    
    For i = 1 To bbe.count
        Set bb = bbe(i)
        ' Insert building block name and gallery name into the new document.
        Set Rng = newDoc.Content
        Rng.Collapse Direction:=wdCollapseEnd
'        rng.style = C_S_ParNormal
'        rng.InsertAfter "Building block name: " & bb.Name & vbCrLf & "Gallery: " & bb.Type.Name & vbCrLf
'        rng.style = C_S_ParNormal
        Rng.InsertParagraphAfter
        Rng.style = C_S_ParNormal
        
        ' Collapse range to end before inserting building block itself to ensure correct cursor position.
        Rng.Collapse Direction:=wdCollapseEnd
        
        ' Insert the building block itself and get the range of the inserted content.
        Set Rng = bb.Insert(Where:=Rng, RichText:=True)
        
        ' Collapse range to end after inserting building block to ensure correct cursor position.
        Rng.Collapse Direction:=wdCollapseEnd
        
        Rng.InsertParagraphAfter
        Rng.style = C_S_ParNormal
        
        ' Collapse range to end after inserting paragraph to ensure correct cursor position.
        Rng.Collapse Direction:=wdCollapseEnd
        
        ' Count the building block category
        If CategoryCount.Exists(bb.Type.Name) Then
            CategoryCount(bb.Type.Name) = CategoryCount(bb.Type.Name) + 1
        Else
            CategoryCount.Add bb.Type.Name, 1
        End If
    Next i
    
    ' Insert summary of building block counts
    Set Rng = newDoc.Content
    Rng.Collapse Direction:=wdCollapseEnd
    Rng.InsertAfter "Summary of Building Blocks by Category:" & vbCrLf
    Rng.style = C_S_ParNormal
    Rng.InsertParagraphAfter

    Dim key As Variant
    For Each key In CategoryCount.Keys
        Rng.InsertAfter key & ": " & CategoryCount(key) & vbCrLf
        Rng.style = C_S_ParNormal
        Rng.InsertParagraphAfter
        Rng.style = C_S_ParNormal
    Next key

    ' Save the new document.
    newDoc.SaveAs2 FileName:=SavePath
    newDoc.Close SaveChanges:=wdSaveChanges
        
    ' Clear object variables
    Set CategoryCount = Nothing
    Set newDoc = Nothing
    Set bbe = Nothing
    Set Rng = Nothing
    Set Rng = Nothing
    
    ' Inform the user.
    MsgBox _
        Prompt:="All building blocks have been exported to the new document and saved at: " & vbNewLine & _
            SavePath, _
        Buttons:=vbInformation + vbOKOnly, _
        Title:=MsgBoxTitle
End Sub

' Export selected building blocks categories from currently attached template to a separate DOCX file.
' wdTypeAutoText, wdTypeTableOfContents, wdTypeQuickParts
' 2025-02-20 by ms and AI
Sub BB_ExportSelectedCategories()
    Dim bb As BuildingBlock
    Dim newDoc As Document
    Dim Rng As Range
    Dim TemplateName As String
    Dim bbe As BuildingBlockEntries
    Dim i As Integer
    Dim CategoryCount As Object

    Dim FileName As String:     FileName = C_F_Macros
    Dim ModuleName As String:   ModuleName = C_M_BuildingBlocks
    Dim MacroName As String:    MacroName = "BB_ExportSelectedCategories"
    Dim MsgBoxTitle As String:  MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName

    Dim IfActiveDocument As Boolean
    IfActiveDocument = CheckIfActiveDocumentIsTemplate() ' private function in this module
    If Not IfActiveDocument Then
        MsgBox _
            Prompt:="Currently active document is not a template file. It cannot contain BuildingBlocks." & vbNewLine & vbNewLine & _
                "Exiting.", _
            Buttons:=vbCritical, _
            Title:=MsgBoxTitle
        Exit Sub
    End If
    
    Dim MainTemplatePath As String  ' Define variables for the full paths of the attached template
    Dim MainTemplateName As String
    Dim MainTemplate As Template
    MainTemplatePath = ActiveDocument.FullName
    MainTemplateName = ActiveDocument.Name
    Set MainTemplate = ActiveDocument.AttachedTemplate
    
    ' Initialize the category count dictionary
    Set CategoryCount = CreateObject("Scripting.Dictionary")
    
    ' Create a new blank document
    Set newDoc = Documents.Add
    
    ' Attach the default template to the new document
    newDoc.AttachedTemplate = C_F_Normal
    
    Dim NewFilename As String
    Dim SavePath As String
    ' Get the template file name without extension
    NewFilename = Left(MainTemplateName, InStrRev(MainTemplateName, ".") - 1)
    
    ' Get the default local file location
    SavePath = Options.DefaultFilePath(wdDocumentsPath) & "\" & NewFilename & "_BB_Content.docx"
    
    ' Make sure the new document contains all "ms" styles
    Call CreateCustomStyles     ' in module Styles

    ' Get the BuildingBlockEntries collection from the template
    Set bbe = ActiveDocument.AttachedTemplate.BuildingBlockEntries

    If bbe.count = 0 Then
        MsgBox _
            Prompt:="No BuildingBlocks found in the temlate file" & vbNewLine & vbNewLine & _
                MainTemplateName & vbNewLine & vbNewLine & _
                "Exiting.", _
            Buttons:=vbCritical + vbOKOnly, _
            Title:=MsgBoxTitle
        Exit Sub
    End If

    For i = 1 To bbe.count
        Set bb = bbe(i)
        ' Check if the building block belongs to the desired categories
        If bb.Type = wdTypeAutoText Or bb.Type = wdTypeTableOfContents Or bb.Type = wdTypeQuickParts Then
            ' Insert building block name and gallery name into the new document.
            Set Rng = newDoc.Content
            Rng.Collapse Direction:=wdCollapseEnd
            Rng.InsertAfter "Building block name: " & bb.Name & vbCrLf & "Gallery: " & bb.Type.Name & vbCrLf
            Rng.InsertParagraphAfter
            
            ' Collapse range to end before inserting building block itself to ensure correct cursor position.
            Rng.Collapse Direction:=wdCollapseEnd
            
            ' Insert the building block itself and get the range of the inserted content.
            Set Rng = bb.Insert(Where:=Rng, RichText:=True)
            
            ' Collapse range to end after inserting building block to ensure correct cursor position.
            Rng.Collapse Direction:=wdCollapseEnd
            
            Rng.InsertParagraphAfter
            
            ' Collapse range to end after inserting paragraph to ensure correct cursor position.
            Rng.Collapse Direction:=wdCollapseEnd
        End If
    Next i
    
    ' Save the new document.
    newDoc.SaveAs2 FileName:=SavePath
    
    ' Clear object variables
    Set newDoc = Nothing
    Set bbe = Nothing
    Set bb = Nothing
    Set Rng = Nothing

    ' Inform the user.
    MsgBox _
        Prompt:="Selected building blocks have been exported to the new document and saved at: " & vbNewLine & _
            SavePath, _
        Buttons:=vbInformation + vbOKOnly, _
        Title:=MsgBoxTitle

End Sub

' Statistics only
' 2025-03-08 by ms
Sub BB_List()
    Dim bbe As BuildingBlockEntries
    Dim i As Integer
    Dim AddInsIndex As Integer
    Dim UserDecision As VbMsgBoxResult
    Dim AddInsName As String
    Dim bb As BuildingBlock ' Corrected to BuildingBlock with capital letters
    Dim AutoTextCounter As Integer
    Dim AutoTextList As String
    Dim QuickPartsCounter As String
    Dim QuickPartsList As String
    Dim HeadersCounter As Integer
    Dim HeadersList As String
    Dim FootersCounter As Integer
    Dim FootersList As String
    Dim TOCCounter As Integer
    Dim TOCList As String
    Dim TotalBB As Integer
    Dim BBList As String
    Dim filenum As Integer
    Dim FilePath As String
    Dim CurrentDate As String
    Dim CurrentTime As String

    Dim FileName As String:     FileName = C_F_Macros
    Dim ModuleName As String:   ModuleName = C_M_BuildingBlocks
    Dim MacroName As String:    MacroName = "BB_List"
    Dim MsgBoxTitle As String:  MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
        
    Dim IfActiveDocument As Boolean
    IfActiveDocument = CheckIfActiveDocumentIsTemplate() ' private function in this module
    If Not IfActiveDocument Then
        MsgBox _
            Prompt:="Currently active document is not a template file. It cannot contain BuildingBlocks." & vbNewLine & vbNewLine & _
                "Exiting.", _
            Buttons:=vbCritical, _
            Title:=MsgBoxTitle
        Exit Sub
    End If
    
    ' Force loading of all BuildingBlocks
    Templates.LoadBuildingBlocks

    ' At first try to set bbe to ActiveDocument.AttachedTemplate.BuildingBlockEntries (template with integrated BuildingBlocks)
    Set bbe = ActiveDocument.AttachedTemplate.BuildingBlockEntries
    
    ' Check if bbe is empty
    If bbe.count = 0 Then
        MsgBox _
            Prompt:="No BuildingBlocks found in the temlate file" & vbNewLine & vbNewLine & _
                ActiveDocument.Name & vbNewLine & vbNewLine & _
                "Exiting.", _
            Buttons:=vbCritical + vbOKOnly, _
            Title:=MsgBoxTitle
        Exit Sub
    End If

    AutoTextCounter = 0
    For i = 1 To bbe.count
        Set bb = bbe(i)
        If bb.Type.Index = wdTypeAutoText Then
            If AutoTextCounter = 0 Then
                AutoTextList = bb.Type.Name & ":" & vbNewLine
            End If
            AutoTextCounter = AutoTextCounter + 1
            AutoTextList = AutoTextList & AutoTextCounter & ". " & bb.Name & " | Category name: " & bb.Category.Name & vbNewLine
        End If
    Next i
    ' Debug.Print AutoTextList
    
    HeadersCounter = 0
    For i = 1 To bbe.count
        Set bb = bbe(i)
        If bb.Type.Index = wdTypeHeaders Then
            If HeadersCounter = 0 Then
                HeadersList = bb.Type.Name & ":" & vbNewLine
            End If
            HeadersCounter = HeadersCounter + 1
            HeadersList = HeadersList & HeadersCounter & ". " & bb.Name & " | Category name: " & bb.Category.Name & vbNewLine
        End If
    Next i
    ' Debug.Print HeadersList
            
    FootersCounter = 0
    For i = 1 To bbe.count
        Set bb = bbe(i)
        If bb.Type.Index = wdTypeFooters Then
            If FootersCounter = 0 Then
                FootersList = bb.Type.Name & ":" & vbNewLine
            End If
            FootersCounter = FootersCounter + 1
            FootersList = FootersList & FootersCounter & ". " & bb.Name & " | Category name: " & bb.Category.Name & vbNewLine
        End If
    Next i
    ' Debug.Print FootersList
            
    TOCCounter = 0
    For i = 1 To bbe.count
        Set bb = bbe(i)
        If bb.Type.Index = wdTypeTableOfContents Then
            If TOCCounter = 0 Then
                TOCList = bb.Type.Name & ":" & vbNewLine
            End If
            TOCCounter = TOCCounter + 1
            TOCList = TOCList & TOCCounter & ". " & bb.Name & " | Category name: " & bb.Category.Name & vbNewLine
        End If
    Next i
    ' Debug.Print TOCList
            
    QuickPartsCounter = 0
    For i = 1 To bbe.count
        Set bb = bbe(i)
        If bb.Type.Index = wdTypeQuickParts Then
            If QuickPartsCounter = 0 Then
                QuickPartsList = bb.Type.Name & ":" & vbNewLine
            End If
            QuickPartsCounter = QuickPartsCounter + 1
            QuickPartsList = QuickPartsList & QuickPartsCounter & ". " & bb.Name & " | Category name: " & bb.Category.Name & vbNewLine
        End If
    Next i
    ' Debug.Print QuickPartsList
       
    ' Calculate total number of BuildingBlocks
    TotalBB = AutoTextCounter + HeadersCounter + FootersCounter + TOCCounter + QuickPartsCounter
    
    ' Combine all lists into one string
    BBList = "Template Name: " & ActiveDocument.Name & vbCrLf & _
                    "Total BuildingBlocks: " & TotalBB & vbCrLf & vbCrLf & _
                    AutoTextList & vbCrLf & _
                    HeadersList & vbCrLf & _
                    FootersList & vbNewLine & _
                    TOCList & vbCrLf & _
                    QuickPartsList & vbCrLf
            
    ' Save the shortcut list to a file
    FilePath = Options.DefaultFilePath(wdDocumentsPath) & "\" & MacroName & ".txt"
    filenum = FreeFile
    CurrentDate = Format(Date, "yyyy-mm-dd")
    CurrentTime = Format(Time, "hh:mm:ss")

    Open FilePath For Output As filenum
    Print #filenum, "Macro Name: " & MacroName
    Print #filenum, "Date: " & CurrentDate
    Print #filenum, "Time: " & CurrentTime
    Print #filenum, BBList
    Close filenum
    
    ' Clear object variables
    Set bbe = Nothing
    Set bb = Nothing
    
    ' Display the summary in a message box
    MsgBox _
        Prompt:="Processing complete." & vbNewLine & vbNewLine & "Information was saved to the file:" _
            & vbNewLine & vbNewLine & FilePath, _
        Buttons:=vbInformation, _
        Title:=MsgBoxTitle
End Sub

' 2025-12-21 by ms and AI
' 2025-12-28 by ms
Sub BB_Add()
    On Error GoTo ErrHandler
    Dim MyRng As Word.Range
    Dim MyTemplate As Template
    Dim TemplatePath As String

    Dim FileName As String:     FileName = C_F_Macros
    Dim ModuleName As String:   ModuleName = C_M_BuildingBlocks
    Dim MacroName As String:    MacroName = "BB_Add"
    
    Dim MsgBoxTitle As String
    MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName

    Set MyRng = Selection.Range
    TemplatePath = Options.DefaultFilePath(wdStartupPath) & "\" & C_F_BuildingBlocks
    Set MyTemplate = Templates(TemplatePath)
    If MyTemplate Is Nothing Then
        Err.Raise vbObjectError + 1001, "BB_Add", "Template not found: " & TemplatePath
    End If
      
    Dim bbName As String:           bbName = ""
    Dim bbType As String:           bbType = ""
    Dim bbCategory As String:       bbCategory = "General"
    Dim bbInsertOptions As String:  bbInsertOptions = ""
    Dim bbDescription As String:    bbDescription = ""
    
    ' Parse selected paragraphs
    Dim p As Paragraph, txt As String
    For Each p In Selection.Paragraphs
        txt = Trim$(Replace(Replace(p.Range.Text, vbCr, ""), vbLf, ""))
        
        If LCase$(Left$(txt, 9)) = C_TTR_name Then
            bbName = Trim$(Mid$(txt, 10))
        
        ElseIf LCase$(Left$(txt, 9)) = C_TTR_type Then
            Select Case Trim$(Mid$(txt, 10))
                Case "wdTypeAutoText": bbType = wdTypeAutoText
                Case "wdTypeFooters": bbType = wdTypeFooters
                Case "wdTypeHeaders": bbType = wdTypeHeaders
                Case "wdTypeTableOfContents": bbType = wdTypeTableOfContents
                Case "wdTypeCustom1": bbType = wdTypeCustom1
                Case Else
                    Err.Raise vbObjectError + 1002, MsgBoxTitle, "Invalid bbType value: " & Mid$(txt, 10)
            End Select
        
        ElseIf LCase$(Left$(txt, 13)) = C_TTR_category Then
            bbCategory = Trim$(Mid$(txt, 14))
            
        ElseIf LCase$(Left$(txt, 18)) = C_TTR_insertoptions Then
            Select Case Trim$(Mid$(txt, 19))
                Case "wdInsertContent": bbInsertOptions = wdInsertContent
                Case "wdInsertParagraph": bbInsertOptions = wdInsertParagraph
                Case "wdInsertPage": bbInsertOptions = wdInsertPage
                Case Else
                    Err.Raise vbObjectError + 10003, MsgBoxTitle, "Invalid bbType value: " & Mid$(txt, 19)
            End Select
            
        ElseIf LCase$(Left$(txt, 16)) = C_TTR_description Then
            bbDescription = Trim$(Mid$(txt, 17))
        End If
    Next p
        
    ' Validate required fields
    If bbName = "" Then Err.Raise vbObjectError + 1004, MsgBoxTitle, "Missing " & C_TTR_name & " paragraph."
    If bbType = 0 Then Err.Raise vbObjectError + 1005, MsgBoxTitle, "Missing or invalid " & C_TTR_type & " paragraph."
    If bbCategory = "" Then Err.Raise vbObjectError + 1006, MsgBoxTitle, "Missing or invalid " & C_TTR_category & "paragraph."
    If (bbInsertOptions <> wdInsertContent) And _
        (bbInsertOptions <> wdInsertParagraph) And _
        (bbInsertOptions <> wdInsertPage) Then Err.Raise vbObjectError + 1007, MsgBoxTitle, "Missing or invalid " & C_TTR_insertoptions & " paragraph."
    
    MyTemplate.BuildingBlockEntries.Add _
        Name:=bbName, _
        Type:=CInt(bbType), _
        Category:=bbCategory, _
        Range:=MyRng, _
        Description:=bbDescription, _
        InsertOptions:=CInt(bbInsertOptions)
    
    MyTemplate.Save
    
    ' Clear object variables
    Set MyRng = Nothing
    Set MyTemplate = Nothing
    
    MsgBox _
        Prompt:=" The Building Block " & vbNewLine & vbNewLine & _
            bbName & vbNewLine & vbNewLine & _
            "was added to " & vbNewLine & C_F_BuildingBlocks & vbNewLine & _
            "in category " & vbNewLine & bbCategory & vbNewLine & vbNewLine & _
            "and saved.", _
        Buttons:=vbInformation, _
        Title:=MsgBoxTitle
    Exit Sub

ErrHandler:
    ' Clear object variables
    Set MyRng = Nothing
    Set MyTemplate = Nothing
    MsgBox _
        Prompt:="Error: " & Err.Number & " - " & Err.Description, _
        Buttons:=vbExclamation, _
        Title:=MsgBoxTitle
End Sub

' 2025-12-21 by ms and AI
Sub BB_InsertBBTemplate()
    Dim InsertText As String
    Dim MyRng As Range
    
    ' Prepare the text with paragraph breaks
    InsertText = C_TTR_name & "MyBuildingBlockTemplate" & vbNewLine & _
                 C_TTR_type & "wdTypeAutoText | wdTypeFooters | wdTypeHeaders | wdTypeTableOfContents | wdTypeCustom1" & vbNewLine & _
                 C_TTR_category & "General | Custom 1" & vbNewLine & _
                 C_TTR_description & "" & vbNewLine & _
                 C_TTR_insertoptions & "wdInsertContent | wdInsertParagraph | wdInsertPage"
    
   ' Take a snapshot of the current Selection as a Range (no UI typing)
    Set MyRng = Selection.Range
    ' Collapse to the insertion point (end of selection/caret)
    MyRng.Collapse Direction:=wdCollapseEnd
    
    ' Insert text after the caret
    MyRng.InsertAfter InsertText
    ' No change to Selection: the user's caret remains where it was
    ' Clear object variables
    Set MyRng = Nothing
End Sub

' 2025-12-21 by ms and AI
' 2025-12-28 by ms and AI
Sub BB_RemoveDefParagraphs()
    Dim TagsToRemove As Variant     ' array
    Dim aStory As Range
    Dim p As Paragraph
    Dim i As Long
    Dim txt As String
    Dim Tag As Variant
    Dim RemovedCount As Long
    
    ' Tags to check at the start of each paragraph
    TagsToRemove = Array( _
        C_TTR_name, _
        C_TTR_type, _
        C_TTR_category, _
        C_TTR_description, _
        C_TTR_insertoptions)
        
    Application.ScreenUpdating = False
    RemovedCount = 0
    
    ' Search all areas of documents (StoryRanges), including headers and footers
    ' Iterate backwards to avoid reindexing issues when deleting
    For Each aStory In ActiveDocument.StoryRanges
        Do
            ' Check if in a Story there are any paragraphs at all
            If aStory.Paragraphs.count > 0 Then
                For i = aStory.Paragraphs.count To 1 Step -1
                    ' If specific section or field is blocked / uneditable
                    On Error Resume Next
                    Set p = aStory.Paragraphs(i)
                    
                    If Err.Number = 0 And Not p Is Nothing Then
                        txt = p.Range.Text
                        txt = Replace(Replace(txt, vbCr, ""), vbLf, "")
                        txt = Trim$(txt)
                        
                        For Each Tag In TagsToRemove
                            If Len(txt) >= Len(Tag) Then
                                If Left$(txt, Len(Tag)) = CStr(Tag) Then
                                    p.Range.Delete
                                    
                                    If Err.Number = 0 Then
                                        RemovedCount = RemovedCount + 1
                                    Else
                                        Err.Clear
                                    End If
                                    
                                    Exit For
                                End If
                            End If
                        Next Tag
                    End If
                    On Error GoTo 0
                    ' Let user to stop it while running with Ctrl + Break
                    DoEvents
                Next i
                
                ' Brute force to remove empty paragraphs
                If aStory.StoryType <> wdMainTextStory Then
                    For i = aStory.Paragraphs.count To 1 Step -1
                        Set p = aStory.Paragraphs(i)
                        ' Check if paragraph is empty (if it contains just paragraph character)
                        ' it doesn't contain any 'anchors' of pictures or shapes
                        If Len(p.Range.Text) <= 1 And _
                            p.Range.InlineShapes.count = 0 And _
                            p.Range.ShapeRange.count = 0 Then
                            ' Simulate pressing of DEL key
                            ' Set range to the end of paragraph
                            ' and expand it by one character
                            Dim delRange As Range
                            Set delRange = p.Range
                            delRange.Collapse Direction:=wdCollapseStart
                            delRange.MoveEnd Unit:=wdCharacter, count:=1

                            On Error Resume Next
                            delRange.Delete
                            On Error GoTo 0
                        End If
                    ' Let user to stop it while running with Ctrl + Break
                    DoEvents
                    Next i
                End If
                
            End If
            
            ' Move to the next Story.
            Set aStory = aStory.NextStoryRange
            ' Let user to stop it while running with Ctrl + Break
            DoEvents
        Loop Until aStory Is Nothing
        ' Let user to stop it while running with Ctrl + Break
        DoEvents
    Next aStory
    Application.ScreenUpdating = True
    
    Dim FileName As String:     FileName = C_F_Macros
    Dim ModuleName As String:   ModuleName = C_M_BuildingBlocks
    Dim MacroName As String:    MacroName = "RemoveBBDefParagraphs"
    Dim MsgBoxTitle As String:  MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    ' Clear object variables
    Set p = Nothing
    Set aStory = Nothing
    
    MsgBox _
        Prompt:=RemovedCount & " paragraph(s) removed.", _
        Buttons:=vbInformation, _
        Title:=MsgBoxTitle
End Sub

