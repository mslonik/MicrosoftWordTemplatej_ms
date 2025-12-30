Attribute VB_Name = "Fonts"
' Module Fonts.bas header:
'
'   1. ShowUsedFonts()
'   2. ShowSubstitutedFonts()
'
' The ribbon menu: Tools_ms -> Fonts.
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
' List in a MsgBox fonts used in this document.
' There is no way to list embedded fonts by VBA in Microsoft Word.
' To check which fonts are actually embedded in your Microsoft Word document, you can follow these steps:
' 1. Save your Word document.
' 2. Change the file extension from .docx to .zip.
' 3. Extract the contents of the .zip file.
' 4. Open the extracted folder and navigate to the word folder.
' 5. Look for a file named fontTable.xml and open it with a text editor.
' 6. In the fontTable.xml file, you will see a list of all the fonts used in the document. Embedded fonts will have an attribute indicating that they are embedded.
' To recognize which fonts are embedded in your document by looking at the fontTable.xml file, you need to look for the <w:embedRegular>, <w:embedBold>, <w:embedItalic>, and similar elements within each <w:font> element. These elements indicate that the font is embedded.
' 2025-03-01 by ms and AI
Sub ShowUsedFonts()
    Dim doc As Document
    Dim fontList As String
    Dim font As Variant
    Dim usedFonts As Collection
    Dim rng As Range
    Dim fontName As String
    
    Dim FileName As String
    FileName = C_F_Macros
    
    Dim ModuleName As String
    ModuleName = C_M_Fonts
    
    Dim MacroName As String
    MacroName = "ShowUsedFonts"
    
    Dim MsgBoxTitle As String
    MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    Set doc = ActiveDocument
    Set usedFonts = New Collection
    
    ' Iterate through each range in the document to collect used fonts
    For Each rng In doc.StoryRanges
        Do
            On Error Resume Next
            fontName = rng.font.Name
            If fontName <> "" Then
                usedFonts.Add fontName, fontName
            End If
            On Error GoTo 0
            Set rng = rng.NextStoryRange
        Loop Until rng Is Nothing
    Next rng
    
    ' Create a list of used fonts
    fontList = "Fonts used in this document:" & vbCrLf & vbNewLine
    For Each font In usedFonts
        fontList = fontList & font & vbCrLf
    Next font
    
    ' Clear object variables
    Set doc = Nothing
    Set usedFonts = Nothing
    Set rng = Nothing
    
    MsgBox _
        Prompt:=fontList, _
        Buttons:=vbInformation + vbOKOnly, _
        Title:=MsgBoxTitle
End Sub

' Run by AppWord_DocumentOpen in module ClsAppEvents.
' When document is opened and it contains substituted fonts, the MsgBox is displayed.
' 2025-08-28 by ms and AI
Sub ShowSubstitutedFonts(ByVal TargetDoc As Word.Document)
    Dim fontList As Collection
    Set fontList = New Collection
    Dim fontName As String
    Dim installedFont As Boolean
    Dim missingFonts As String
    Dim i As Integer
    Dim para As Paragraph
    Dim tbl As Table
    Dim cel As Cell
    Dim shp As Shape
    Dim txtRange As Range
    
    ' Collect fonts from paragraphs
    For Each para In TargetDoc.Paragraphs
        fontName = para.Range.font.Name
        On Error Resume Next
        fontList.Add fontName, fontName
        On Error GoTo 0
    Next para

    ' Collect fonts from tables
    For Each tbl In TargetDoc.Tables
        For Each cel In tbl.Range.Cells
            fontName = cel.Range.font.Name
            On Error Resume Next
            fontList.Add fontName, fontName
            On Error GoTo 0
        Next cel
    Next tbl

    ' Collect fonts from shapes
    For Each shp In TargetDoc.Shapes
        If shp.Type = msoTextBox Then
            Set txtRange = shp.TextFrame.TextRange
            fontName = txtRange.font.Name
            On Error Resume Next
            fontList.Add fontName, fontName
            On Error GoTo 0
        End If
    Next shp

    ' Check if each font is installed
    For i = 1 To fontList.count
        fontName = fontList(i)
        installedFont = False
        Dim f As Variant
        For Each f In Application.FontNames
            If f = fontName Then
                installedFont = True
                Exit For
            End If
        Next f
        If Not installedFont And fontName <> "" Then
            missingFonts = missingFonts & fontName & vbCrLf
        End If
    Next i

    Dim FileName As String
    FileName = C_F_Macros
    
    Dim ModuleName As String
    ModuleName = C_M_Fonts
    
    Dim MacroName As String
    MacroName = "ShowSubstitutedFonts"
    
    Dim MsgBoxTitle As String
    MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName

    ' Show results
    If missingFonts <> "" Then
        MsgBox _
            Prompt:="Missing fonts (possibly substituted):" & vbCrLf & missingFonts, _
            Buttons:=vbExclamation, _
            Title:=MsgBoxTitle
    End If
End Sub
