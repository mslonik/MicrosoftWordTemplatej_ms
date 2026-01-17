Attribute VB_Name = "others"
' Paragraph styles and linked styles follows ListNonBuiltInAndSuffixStylesInTemplate report.
' 23. 36. CreateStyle_ParPictureCanvaMs()   -> SaveIniStyle_ParPictureCanvaMs
' 24. 37. CreateStyle_ParSourceCodeMs()     -> SaveIniStyle_ParSourceCodeMs
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
' Used to enforce the explicit declaration of all variables in a module. When you include Option Explicit at the beginning of a module, it ensures that you must
' declare all variables using the Dim, Private, Public, ReDim, or Static statements before using them. This helps prevent errors caused by typos or
' undeclared variables.
Option Explicit

Dim IniPath As String

' Declare API functions
Private Declare PtrSafe Function GetPrivateProfileString Lib "kernel32" _
    Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, _
    ByVal lpKeyName As String, ByVal lpDefault As String, _
    ByVal lpReturnedString As String, ByVal nSize As Long, _
    ByVal lpFileName As String) As Long

Private Declare PtrSafe Function WritePrivateProfileString Lib "kernel32" _
    Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, _
    ByVal lpKeyName As String, ByVal lpString As String, _
    ByVal lpFileName As String) As Long

Private Sub InitializeConstants()
    ' Initialize the variable in a subroutine
    IniPath = Application.Options.DefaultFilePath(wdStartupPath) & "\" & C_M_Styles & ".ini"
End Sub

' Read from INI
Function ReadIniValue(Section As String, key As String, FilePath As String) As String
    Dim RetVal As String * 255
    Dim Length As Long
    Length = GetPrivateProfileString(Section, key, "", RetVal, 255, FilePath)
    ReadIniValue = Left(RetVal, Length)
End Function

' Write to INI
Sub WriteIniValue(Section As String, key As String, Value As String, FilePath As String)
    WritePrivateProfileString Section, key, Value, FilePath
End Sub

' 2025-11-15 by ms
Sub SaveStylesToIni()
    Call SaveIniStyle_ParPictureCanvaMs
    Call SaveIniStyle_ParSourceCodeMs
End Sub

' 2025-11-15 by ms
Private Sub SaveIniStyle_ParSourceCodeMs()
    Call InitializeConstants
    Call WriteIniValue( _
        Section:=C_S_ParSourceCode, _
        key:="Type", _
        Value:=wdStyleTypeParagraph, _
        FilePath:=IniPath _
        )
    Call WriteIniValue( _
        Section:=C_S_ParSourceCode, _
        key:="BaseStyle", _
        Value:=C_S_ParNormal, _
        FilePath:=IniPath _
        )
    Call WriteIniValue( _
        Section:=C_S_ParSourceCode, _
        key:="NextParagraphStyle", _
        Value:=C_S_ParNormal, _
        FilePath:=IniPath _
        )
    Call WriteIniValue( _
        Section:=C_S_ParSourceCode, _
        key:="AutomaticallyUpdate", _
        Value:=False, _
        FilePath:=IniPath _
        )
    Call WriteIniValue( _
        Section:=C_S_ParSourceCode, _
        key:="QuickStyle", _
        Value:=False, _
        FilePath:=IniPath _
        )
    Call WriteIniValue( _
        Section:=C_S_ParSourceCode, _
        key:="LanguageId", _
        Value:=wdEnglishUS, _
        FilePath:=IniPath _
        )
    Call WriteIniValue( _
        Section:=C_S_ParSourceCode, _
        key:="Font_Name", _
        Value:="Consolas", _
        FilePath:=IniPath _
        )
    Call WriteIniValue( _
        Section:=C_S_ParSourceCode, _
        key:="Font_Size", _
        Value:=11, _
        FilePath:=IniPath _
        )
    Call WriteIniValue( _
        Section:=C_S_ParSourceCode, _
        key:="Font_Bold", _
        Value:=False, _
        FilePath:=IniPath _
        )
    Call WriteIniValue( _
        Section:=C_S_ParSourceCode, _
        key:="Font_Italic", _
        Value:=False, _
        FilePath:=IniPath _
        )
    Call WriteIniValue( _
        Section:=C_S_ParSourceCode, _
        key:="Font_Color", _
        Value:=wdColorBlack, _
        FilePath:=IniPath _
        )
    Call WriteIniValue( _
        Section:=C_S_ParSourceCode, _
        key:="ParagraphFormat_Alignment", _
        Value:=wdAlignParagraphCenter, _
        FilePath:=IniPath _
        )
    Call WriteIniValue( _
        Section:=C_S_ParSourceCode, _
        key:="ParagraphFormat_LeftIndent", _
        Value:=CentimetersToPoints(0.2), _
        FilePath:=IniPath _
        )
    Call WriteIniValue( _
        Section:=C_S_ParSourceCode, _
        key:="ParagraphFormat_RightIndent", _
        Value:=CentimetersToPoints(0.2), _
        FilePath:=IniPath _
        )
    Call WriteIniValue( _
        Section:=C_S_ParSourceCode, _
        key:="ParagraphFormat_FirstLineIndent", _
        Value:=CentimetersToPoints(0.2), _
        FilePath:=IniPath _
        )
    Call WriteIniValue( _
        Section:=C_S_ParSourceCode, _
        key:="ParagraphFormat_SpaceBefore", _
        Value:=0, _
        FilePath:=IniPath _
        )
    Call WriteIniValue( _
        Section:=C_S_ParSourceCode, _
        key:="ParagraphFormat_SpaceAfter", _
        Value:=0, _
        FilePath:=IniPath _
        )
    Call WriteIniValue( _
        Section:=C_S_ParSourceCode, _
        key:="ParagraphFormat_LineSpacing", _
        Value:=11, _
        FilePath:=IniPath _
        )
    Call WriteIniValue( _
        Section:=C_S_ParSourceCode, _
        key:="ParagraphFormat_LineSpacingRule", _
        Value:=wdLineSpaceExactly, _
        FilePath:=IniPath _
        )
    Call WriteIniValue( _
        Section:=C_S_ParSourceCode, _
        key:="ParagraphFormat_WidowControl", _
        Value:=True, _
        FilePath:=IniPath _
        )
    Call WriteIniValue( _
        Section:=C_S_ParSourceCode, _
        key:="ParagraphFormat_KeepWithNext", _
        Value:=False, _
        FilePath:=IniPath _
        )
    Call WriteIniValue( _
        Section:=C_S_ParSourceCode, _
        key:="ParagraphFormat_KeepTogether", _
        Value:=False, _
        FilePath:=IniPath _
        )
    Call WriteIniValue( _
        Section:=C_S_ParSourceCode, _
        key:="ParagraphFormat_PageBreakBefore", _
        Value:=False, _
        FilePath:=IniPath _
        )
'    Call WriteIniValue( _
'        Section:=C_S_ParSourceCode, _
'        Key:="Shortcut_KeyCategory", _
'        Value:=wdKeyCategoryStyle, _
'        FilePath:=IniPath _
'        )
'    Call WriteIniValue( _
'        Section:=C_S_ParSourceCode, _
'        Key:="Shortcut_Command", _
'        Value:=C_S_ParSourceCode, _
'        FilePath:=IniPath _
'        )
'    Call WriteIniValue( _
'        Section:=C_S_ParSourceCode, _
'        Key:="Shortcut_Keycode", _
'        Value:=BuildKeyCode(wdKeyControl, wdKeyAlt, wdKeyS), _
'        FilePath:=IniPath _
'        )
    Dim FileName As String
    FileName = C_F_Macros
    
    Dim ModuleName As String
    ModuleName = C_M_Styles
    
    Dim MacroName As String
    MacroName = "SaveIniStyle_ParSourceCodeMs"
    
    Dim MsgBoxTitle As String
    MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    MsgBox _
        Prompt:="The " & vbNewLine & vbNewLine & IniPath & vbNewLine & vbNewLine & " was saved.", _
        Buttons:=vbInformation, _
        Title:=MsgBoxTitle
End Sub

' 2025-11-15 by ms
Private Sub SaveIniStyle_ParPictureCanvaMs()
    Call InitializeConstants
    Call WriteIniValue( _
        Section:=C_S_ParPictureCanva, _
        key:="Type", _
        Value:=wdStyleTypeParagraph, _
        FilePath:=IniPath _
        )
    Call WriteIniValue( _
        Section:=C_S_ParPictureCanva, _
        key:="BaseStyle", _
        Value:=C_S_ParNormal, _
        FilePath:=IniPath _
        )
    Call WriteIniValue( _
        Section:=C_S_ParPictureCanva, _
        key:="NextParagraphStyle", _
        Value:=C_S_ParNormal, _
        FilePath:=IniPath _
        )
    Call WriteIniValue( _
        Section:=C_S_ParPictureCanva, _
        key:="AutomaticallyUpdate", _
        Value:=False, _
        FilePath:=IniPath _
        )
    Call WriteIniValue( _
        Section:=C_S_ParPictureCanva, _
        key:="QuickStyle", _
        Value:=False, _
        FilePath:=IniPath _
        )
    Call WriteIniValue( _
        Section:=C_S_ParPictureCanva, _
        key:="LanguageId", _
        Value:=wdEnglishUS, _
        FilePath:=IniPath _
        )
    Call WriteIniValue( _
        Section:=C_S_ParPictureCanva, _
        key:="Font_Name", _
        Value:=C_FT_Body, _
        FilePath:=IniPath _
        )
    Call WriteIniValue( _
        Section:=C_S_ParPictureCanva, _
        key:="Font_Size", _
        Value:=11, _
        FilePath:=IniPath _
        )
    Call WriteIniValue( _
        Section:=C_S_ParPictureCanva, _
        key:="Font_Bold", _
        Value:=False, _
        FilePath:=IniPath _
        )
    Call WriteIniValue( _
        Section:=C_S_ParPictureCanva, _
        key:="Font_Italic", _
        Value:=False, _
        FilePath:=IniPath _
        )
    Call WriteIniValue( _
        Section:=C_S_ParPictureCanva, _
        key:="Font_Color", _
        Value:=wdColorBlack, _
        FilePath:=IniPath _
        )
    Call WriteIniValue( _
        Section:=C_S_ParPictureCanva, _
        key:="ParagraphFormat_Alignment", _
        Value:=wdAlignParagraphCenter, _
        FilePath:=IniPath _
        )
    Call WriteIniValue( _
        Section:=C_S_ParPictureCanva, _
        key:="ParagraphFormat_LeftIndent", _
        Value:=CentimetersToPoints(0), _
        FilePath:=IniPath _
        )
    Call WriteIniValue( _
        Section:=C_S_ParPictureCanva, _
        key:="ParagraphFormat_RightIndent", _
        Value:=CentimetersToPoints(0), _
        FilePath:=IniPath _
        )
    Call WriteIniValue( _
        Section:=C_S_ParPictureCanva, _
        key:="ParagraphFormat_FirstLineIndent", _
        Value:=CentimetersToPoints(0), _
        FilePath:=IniPath _
        )
    Call WriteIniValue( _
        Section:=C_S_ParPictureCanva, _
        key:="ParagraphFormat_SpaceBefore", _
        Value:=12, _
        FilePath:=IniPath _
        )
    Call WriteIniValue( _
        Section:=C_S_ParPictureCanva, _
        key:="ParagraphFormat_SpaceAfter", _
        Value:=6, _
        FilePath:=IniPath _
        )
    Call WriteIniValue( _
        Section:=C_S_ParPictureCanva, _
        key:="ParagraphFormat_LineSpacing", _
        Value:=11, _
        FilePath:=IniPath _
        )
    Call WriteIniValue( _
        Section:=C_S_ParPictureCanva, _
        key:="ParagraphFormat_LineSpacingRule", _
        Value:=wdLineSpaceExactly, _
        FilePath:=IniPath _
        )
    Call WriteIniValue( _
        Section:=C_S_ParPictureCanva, _
        key:="ParagraphFormat_WidowControl", _
        Value:=True, _
        FilePath:=IniPath _
        )
    Call WriteIniValue( _
        Section:=C_S_ParPictureCanva, _
        key:="ParagraphFormat_KeepWithNext", _
        Value:=True, _
        FilePath:=IniPath _
        )
    Call WriteIniValue( _
        Section:=C_S_ParPictureCanva, _
        key:="ParagraphFormat_KeepTogether", _
        Value:=False, _
        FilePath:=IniPath _
        )
    Call WriteIniValue( _
        Section:=C_S_ParPictureCanva, _
        key:="ParagraphFormat_PageBreakBefore", _
        Value:=False, _
        FilePath:=IniPath _
        )
'    Call WriteIniValue( _
'        Section:=C_S_ParPictureCanva, _
'        Key:="Shortcut_KeyCategory", _
'        Value:=wdKeyCategoryStyle, _
'        FilePath:=IniPath _
'        )
'    Call WriteIniValue( _
'        Section:=C_S_ParPictureCanva, _
'        Key:="Shortcut_Command", _
'        Value:=C_S_ParSourceCode, _
'        FilePath:=IniPath _
'        )
'    Call WriteIniValue( _
'        Section:=C_S_ParPictureCanva, _
'        Key:="Shortcut_Keycode", _
'        Value:=BuildKeyCode(wdKeyControl, wdKeyAlt, wdKeyS), _
'        FilePath:=IniPath _
'        )
    Dim FileName As String
    FileName = C_F_Macros
    
    Dim ModuleName As String
    ModuleName = C_M_Macros
    
    Dim MacroName As String
    MacroName = "SaveIniStyle_ParPictureCanvaMs"
    
    Dim MsgBoxTitle As String
    MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    MsgBox _
        Prompt:="The " & vbNewLine & vbNewLine & IniPath & vbNewLine & vbNewLine & " was saved.", _
        Buttons:=vbInformation, _
        Title:=MsgBoxTitle
End Sub
' Stores some vital parameters within Document Variables, in the template itself.
' It is not very useful at this moment, as the same information can be stored directly in the macro.
' This is more proof of concept.
' 2025-03-02 by ms and AI
Private Sub StoreMarginsHeadersFooters()
    Dim MarginInside As Double
    Dim MarginOutside As Double
    Dim MirrorMarginsDecision As Boolean
    Dim HFDistance As Double

    Dim FileName As String
    FileName = C_F_Macros
    
    Dim ModuleName As String
    ModuleName = C_M_Template
    
    Dim MacroName As String
    MacroName = "StoreMarginsHeadersFooters"
    
    Dim MsgBoxTitle As String
    MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName

    If ActiveDocument.AttachedTemplate.Name <> ActiveDocument.Name Then
        MsgBox _
            Prompt:="This macro can be run only from within a template file (DOTM).", _
            Buttons:=vbCritical, _
            Title:=MsgBoxTitle
        Exit Sub
    End If

    MarginInside = 1.2
    MarginOutside = 2.2
    HFDistance = 0.5                            ' Header Footer Distance
    MirrorMarginsDecision = True                ' If to set mirror margins; in VBA true = -1, false = 0
    
    ' Store Document Variables in Template
    If Not CheckDocVariableExists("MarginInside") Then
        Call AddDocVariable("MarginInside", MarginInside)
    End If
    If Not CheckDocVariableExists("MarginOutside") Then
        Call AddDocVariable("MarginOutside", MarginOutside)
    End If
    If Not CheckDocVariableExists("HFDistance") Then
        Call AddDocVariable("HFDistance", HFDistance)
    End If
    If Not CheckDocVariableExists("MirrorMarginsDecision") Then
        Call AddDocVariable("MirrorMarginsDecision", MirrorMarginsDecision)   ' true in VBA = -1, false = 0
    End If
    
    MsgBox _
        Prompt:="Finished processing.", _
        Buttons:=vbInformation, _
        Title:=MsgBoxTitle
End Sub

Private Function CheckDocVariableExists(DocVarName As String) As Boolean
    Debug.Print "DocVarName: " & DocVarName
    Dim DocVar As Variable
    
    On Error Resume Next ' 0 = no error occured
        IsEmpty (ActiveDocument.Variables(DocVarName))
    If Err.Number = 0 Then
        CheckDocVariableExists = True
    Else
        CheckDocVariableExists = False
    End If
    On Error GoTo 0
End Function

Private Sub AddDocVariable(DocVarName As Variant, DocVarValue As Variant)
    Debug.Print "DocVarName: " & DocVarName
    Debug.Print "DocVarValue: " & DocVarValue
    Dim DocVar As Variable
        
    ' Add the document variable if it doesn't exist
    ActiveDocument.Variables.Add Name:=DocVarName, Value:=DocVarValue
End Sub


' !!! No longer necessary. I left it as a future example how to do that.
' Ads code to this project branch "Microsoft Word Objects" -> ThisDocument. As a consequence added code will run each time the template runs.
' 2025-02-25 by ms and AI
Private Sub CreateCodeToThisDocument()
    Dim vbProj As VBIDE.VBProject
    Dim vbComp As VBIDE.VBComponent
    Dim vbCodeMod As VBIDE.CodeModule
    Dim lineNum As Long
    Dim found As Boolean
    Dim i As Long

    ' Reference the VBA project of the template
    Set vbProj = ThisDocument.VBProject
    ' Reference the ThisDocument module
    Set vbComp = vbProj.VBComponents("ThisDocument")
    Set vbCodeMod = vbComp.CodeModule

    ' Check if the following line "Sub Document_Open()" already exists
    found = False
    For i = 1 To vbCodeMod.CountOfLines
        If InStr(vbCodeMod.Lines(i, 1), "Sub Document_Open()") > 0 Then
            found = True
            Exit For
        End If
    Next i

    ' Add code to the ThisDocument module if not found
    If Not found Then
        lineNum = vbCodeMod.CountOfLines + 1
        vbCodeMod.InsertLines lineNum, _
            "Private Sub Document_Open()" & vbNewLine & _
            "    If ActiveDocument.AttachedTemplate.Name = ActiveDocument.Name Then" & vbNewLine & _
            "        EnableVBIDEReference" & vbNewLine & _
            "    End If" & vbNewLine & _
            "End Sub"
    End If
    
    ' Clear object variables
    Set vbProj = Nothing
    Set vbComp = Nothing
    Set vbCodeMod = Nothing
End Sub


' Enable reference to Microsoft Visual Basic for Applications Extensibility 5.3 = VBIDE
' Unfortunately this sub cannot be applied in the body of other subs. This is "key in the box" issue: to enable reference, it must be at first enabled. When it is disabled, code won't compile.
' So if I'd like to enabe this reference, I have to enable it manually.
' 2025-02-24 by ms and AI
Sub EnableVBIDEReference()
    Dim vbProj As Object
    Dim vbComp As Object
    Dim vbRef As Object
    Dim refName As String
    Dim refFound As Boolean
    Dim FullName As String
    Dim FileName As String
    FileName = C_F_Macros
    
    Dim ModuleName As String
    ModuleName = C_M_Macros
    
    Dim MacroName As String
    MacroName = "EnableVBIDEReference"
    
    Dim MsgBoxTitle As String
    MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName

    FullName = "Microsoft Visual Basic for Applications Extensibility 5.3" ' = VBIDE
    refName = "VBIDE"
    refFound = False

    ' Get the VBA project of the active document
    Set vbProj = ActiveDocument.VBProject

    ' Check if the reference is already added
    For Each vbRef In vbProj.References
        If vbRef.Name = refName Then
            refFound = True
            Exit For
        End If
    Next vbRef

    ' Add the reference if not found
    If Not refFound Then
        vbProj.References.AddFromGuid "{0002E157-0000-0000-C000-000000000046}", 5, 3
        ' CreateCodeToThisDocument    ' adds content to ThisDocument. No longer necessary is information about set reference is stored in the template file.
        MsgBox _
            Prompt:=FullName & " reference has been added. Save the template file to make it permanent.", _
            Buttons:=vbInformation + vbOKOnly, _
            Title:=MsgBoxTitle
    End If
    
    ' Clear object variables
    Set vbProj = Nothing
End Sub


