Attribute VB_Name = "Macros"
' Module Macros.bas header:
'
'   Special conditions:
'   - enable manually: Tools > References > Microsoft Visual Basic for Applications Extensibility 5.3
'   -run the first four macros only from within template file (DOTM)!
'
'   1. ExportAllVBAModules()
'   2. ImportAllVBAModules()
'   3. DeleteAllVBAModules()
'   4. DeleteAllVBAModulesExceptMacros()
'
'   5. ShowMacrosCounter()
'   6. ListMacros()
'   7. ScanProjectForNonAscii()
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
' Used to enforce the explicit declaration of all variables in a module. When you include Option Explicit at the beginning of a module, it ensures that you must
' declare all variables using the Dim, Private, Public, ReDim, or Static statements before using them. This helps prevent errors caused by typos or
' undeclared variables.
Option Explicit

' Results are saved in the default path for all the files.
' Open the VBA editor by pressing Alt + F11.
' Go to Tools -> References.
' In the References dialog box, scroll down and check the box for "Microsoft Visual Basic for Applications Extensibility 5.3".
' Click OK to close the dialog box.
' 2025-02-21 by ms and AI
' 2025-12-30 by ms and AI
Sub ExportAllVBAModules()
    Dim vbProj As VBIDE.VBProject
    Dim vbComp As VBIDE.VBComponent
    Dim ExportPath As String
    Dim Extension As String
    Dim FullFileName As String
    
    Dim FileName As String:     FileName = C_F_Macros
    Dim ModuleName As String:   ModuleName = C_M_Macros
    Dim MacroName As String:    MacroName = "ExportAllVBAModules"
    Dim MsgBoxTitle As String:  MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    If Not CheckIfActiveDocumentIsMacroTemplate() Then
        MsgBox _
            Prompt:="This macro can be run only from within a macro template file (DOTM)." & vbNewLine & vbNewLine & "Exiting.", _
            Buttons:=vbCritical + vbOKOnly, _
            Title:=MsgBoxTitle
        Exit Sub
    End If
    
    ' Set the export path
    ExportPath = Options.DefaultFilePath(wdDocumentsPath)
        
    ' Ensure the export path ends with a backslash
    If Right(ExportPath, 1) <> "\" Then
        ExportPath = ExportPath & "\"
    End If
    
    ' Get the current VBA project
    Set vbProj = ActiveDocument.VBProject
    
    Dim ModuleCounter As Byte
    ModuleCounter = 0
    ' Loop through each component in the project
    For Each vbComp In vbProj.VBE.ActiveVBProject.VBComponents
        Extension = ""
        Select Case vbComp.Type
            Case vbext_ct_StdModule
                Extension = ".bas"
            Case vbext_ct_ClassModule
                Extension = ".cls"
            Case vbext_ct_MSForm
                Extension = ".frm"
            Case vbext_ct_Document
                Extension = ".cls"
        End Select
    
        If Extension <> "" Then
            FullFileName = ExportPath & vbComp.Name & Extension
        
            On Error Resume Next
            If Dir(FullFileName) <> "" Then Kill FullFileName
            
            ' Component export
            vbComp.Export FullFileName
            ModuleCounter = ModuleCounter + 1
            On Error GoTo 0
        End If
    Next vbComp
    
    ' Clean object variables
    Set vbProj = Nothing
    
    MsgBox _
        Prompt:="Finished processing." & vbNewLine & _
            ModuleCounter & " modules have been exported to " & vbNewLine & ExportPath, _
        Buttons:=vbInformation + vbOKOnly, _
        Title:=MsgBoxTitle
End Sub

' 2025-12-19 by ms
Private Function CheckIfActiveDocumentIsMacroTemplate() As Boolean
    Dim DocName As String
    
    ' Get the active document name
    DocName = ActiveDocument.Name
    
    ' Check file extension for template types
    If LCase(Right(DocName, 5)) = ".dotm" Then
        CheckIfActiveDocumentIsMacroTemplate = True
    Else
        CheckIfActiveDocumentIsMacroTemplate = False
    End If
End Function

' Open the VBA editor by pressing Alt + F11.
' Go to Tools -> References.
' In the References dialog box, scroll down and check the box for "Microsoft Visual Basic for Applications Extensibility 5.3".
' Click OK to close the dialog box.
' 2025-02-21 by ms and AI
' 2025-12-30 by ms and AI
Sub ImportAllVBAModules()
    Dim vbProj As VBIDE.VBProject
    Dim vbComp As VBIDE.VBComponent
    Dim TempComp As VBIDE.VBComponent
    Dim importPath As String
    Dim FileExtension As String
    Dim file As Object
    Dim fso As Object
    Dim FolderDialog As FileDialog
    Dim newModule As Object
    Dim FoundException As Boolean
    Dim BaseName As String

    Dim FileName As String:      FileName = C_F_Macros
    Dim ModuleName As String:    ModuleName = C_M_Macros
    Dim MacroName As String:     MacroName = "ImportAllVBAModules"
    Dim MsgBoxTitle As String:   MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName

    If Not CheckIfActiveDocumentIsMacroTemplate() Then
        MsgBox _
            Prompt:="This macro can be run only from within a macro template file (DOTM)." & vbNewLine & vbNewLine & "Exiting.", _
            Buttons:=vbCritical + vbOKOnly, _
            Title:=MsgBoxTitle
        Exit Sub
    End If

    FoundException = False
    ' Create FileDialog object
    Set FolderDialog = Application.FileDialog(msoFileDialogFolderPicker)
    
    ' Set dialog properties
    With FolderDialog
        .Title = "Select Folder Containing VBA Modules (.bas, .cls, .frm)"
        .AllowMultiSelect = False
        If .Show = -1 Then
            importPath = .SelectedItems(1)
        Else
            MsgBox _
                Prompt:="No folder selected. Operation cancelled.", _
                Buttons:=vbExclamation + vbOKOnly, _
                Title:=MsgBoxTitle
            Exit Sub
        End If
    End With
    
    ' Ensure the import path ends with a backslash
    If Right(importPath, 1) <> "\" Then
        importPath = importPath & "\"
    End If
    
    ' Get the current VBA project of the attached template
    On Error Resume Next
    Set vbProj = ActiveDocument.VBProject
    If Err.Number <> 0 Then
        MsgBox _
            Prompt:="No access to VBA project. Check settings in:" & vbNewLine & _
                "Word Options -> Trust Center -> Trust Center Settings…" & vbNewLine & _
                "-> Macro Settings: Developer Macro Settings Trust access to the VBA project object model", _
            Buttons:=vbCritical, _
            Title:=MsgBoxTitle
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Create FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Loop through each file in the folder
    Dim ModuleCounter As Byte
    ModuleCounter = 0
    For Each file In fso.GetFolder(importPath).Files
        FileExtension = LCase(fso.GetExtensionName(file.Name))
        BaseName = fso.GetBaseName(file.Name)                           ' File name without extension
        
        ' Check if the file is a VBA module
        If FileExtension = "bas" Or FileExtension = "cls" Or FileExtension = "frm" Then
        
            ' Trick: content of the ThisDocument is moved to a temporary file, because this module cannot be directly imported
            If BaseName = "ThisDocument" Then
                Set TempComp = vbProj.VBComponents.Import(file.Path)
                Set vbComp = vbProj.VBComponents("ThisDocument")
                If vbComp.CodeModule.CountOfLines > 0 Then
                    With vbComp.CodeModule
                        .DeleteLines 1, .CountOfLines
                        .AddFromString TempComp.CodeModule.Lines(1, TempComp.CodeModule.CountOfLines)
                    End With
                End If
                vbProj.VBComponents.Remove TempComp
            Else
                On Error Resume Next
                Set vbComp = vbProj.VBComponents(BaseName)
                If Err.Number = 0 Then
                    vbProj.VBComponents.Remove vbComp
                End If
                Err.Clear
                On Error GoTo 0
            
                vbProj.VBComponents.Import file.Path
                ModuleCounter = ModuleCounter + 1
            End If
        End If
    Next file
    
    ' Clear object variables
    Set FolderDialog = Nothing
    Set vbProj = Nothing
    Set fso = Nothing
    Set vbComp = Nothing
    Set TempComp = Nothing
    
    MsgBox _
        Prompt:="Finished processing" & vbNewLine & _
            ModuleCounter & " modules have been imported from " & vbNewLine & importPath, _
        Buttons:=vbInformation, _
        Title:=MsgBoxTitle
End Sub

Sub DeleteAllVBAModules()
    Call DeleteVBAModules
End Sub

Sub DeleteAllVBAModulesExceptMacros()
    Call DeleteVBAModules(Exception:="Macros")
End Sub
    

' No exceptions, including content of the ThisDocument
' 2025-02-27 by ms and AI
Private Sub DeleteVBAModules(Optional Exception As String)
    Dim vbProj As VBIDE.VBProject
    Dim vbComp As VBIDE.VBComponent
    Dim compName As String
    Dim DeletedModules As String
    Dim NotDeletedModules As String
    
    Dim FileName As String:     FileName = C_F_Macros
    Dim ModuleName As String:   ModuleName = C_M_Macros
    Dim MacroName As String:    MacroName = "DeleteVBAModules"
    Dim MsgBoxTitle As String:  MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    DeletedModules = ""
    Dim moduleCount As Integer
    moduleCount = 0
    
    If Not CheckIfActiveDocumentIsMacroTemplate() Then
        MsgBox _
            Prompt:="This macro can be run only from within a macro template file (DOTM)." & vbNewLine & vbNewLine & "Exiting.", _
            Buttons:=vbCritical + vbOKOnly, _
            Title:=MsgBoxTitle
        Exit Sub
    End If
    
    ' Get the current VBA project
    Set vbProj = ActiveDocument.VBProject
    
    ' Loop through each component in the project
    ' used for error handling in your code. When this statement is executed, it tells VBA to continue
    ' with the next line of code after an error occurs, instead of stopping the execution and
    ' displaying an error message.
    On Error Resume Next
    For Each vbComp In vbProj.VBComponents
        compName = vbComp.Name
        If vbComp.Type = vbext_ct_StdModule Or _
            vbComp.Type = vbext_ct_ClassModule Or _
            vbComp.Type = vbext_ct_MSForm And _
            compName <> Exception _
            Then
            vbProj.VBComponents.Remove vbComp
            If Err.Number <> 0 Then
                MsgBox _
                    Prompt:="Error removing component: " & Err.Description & vbNewLine & vbNewLine & _
                        "Module name: " & vbComp.Name, _
                    Buttons:=vbInformation + vbOKOnly, _
                    Title:=MsgBoxTitle
                Err.Clear
                NotDeletedModules = NotDeletedModules & compName & vbNewLine
            Else
                DeletedModules = DeletedModules & compName & vbNewLine
            End If
            moduleCount = moduleCount + 1
        End If
    Next vbComp
    ' resets VBA's error handling to its default behavior, which means that any subsequent errors
    ' will cause the program to stop and display an error message
    On Error GoTo 0
    
    ' Check if any modules were Deleted
    If moduleCount > 0 Then
        MsgBox _
            Prompt:="The following modules have been deleted from the active document:" & vbNewLine & _
                DeletedModules & vbNewLine & "Active Document: " & ActiveDocument.Name _
                & vbNewLine & vbNewLine & _
                "The following modules have not been deleted due to an errors:" & vbNewLine & _
                NotDeletedModules _
                & vbNewLine & vbNewLine & _
                "Exit this document and press save the changes to apply deletion of the modules.", _
            Buttons:=vbInformation + vbOKOnly, _
            Title:=MsgBoxTitle
    Else
        MsgBox _
            Prompt:="No modules were found in the active document.", _
            Buttons:=vbInformation, _
            Title:=MsgBoxTitle
    End If
    
    ' If there is no exception such as "Macros", remove also content from "ThisDocument"
    If Exception = "" Then
        Call RemoveContentFromThisDocument  ' in this file
    End If
    
    ' Clear object variables
    Set vbProj = Nothing
End Sub

' Run this macro upon opening the C_F_Macros file.
' Displays result in the MsgBox.
' 2025-07-31 by ms
Sub ShowMacrosCounter()
    Dim vbComp As VBIDE.VBComponent
    Dim vbProj As VBIDE.VBProject
    Dim doc As Document
    Dim output As String
    
    Dim FileName As String:     FileName = C_F_Macros
    Dim ModuleName As String:   ModuleName = C_M_Macros
    Dim MacroName As String:    MacroName = "ShowMacrosCounter"
    Dim MsgBoxTitle As String:  MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    ' Ensure that the VBA project model is accessible
    Application.VBE.MainWindow.Visible = True
    
    ' Get the active document's VBA project
    Set doc = ActiveDocument
    Set vbProj = doc.VBProject
    
    ' Initialize output string
    output = "Forms and Modules in the active document:" & vbCrLf & vbCrLf
    
    ' Loop through all components in the VBA project
    For Each vbComp In vbProj.VBComponents
        output = output & vbComp.Name & " (" & vbComp.Type & ")" & vbCrLf
    Next vbComp
    
    ' Display the output in a message box
    MsgBox _
        Prompt:=output, _
        Buttons:=vbInformation + vbOKOnly, _
        Title:=MsgBoxTitle
    
    ' Clear object variables
    Set doc = Nothing
    Set vbProj = Nothing
End Sub

' Run this macro upon opening the C_F_Macros file.
' Statistics regarding macros dumped into .txt  file.
' 2025-03-08 by ms and AI
Sub ListMacros()
    Dim vbComp As VBIDE.VBComponent
    Dim vbProj As VBIDE.VBProject
    Dim vbCodeMod As VBIDE.CodeModule
    Dim lineNum As Long
    Dim procName As String
    Dim procKind As VBIDE.vbext_ProcKind
    Dim MyTemplate As Template
    Dim output As String
    Dim ModuleCounter As Long
    Dim subCounter As Long
    Dim funcCounter As Long
    Dim FilePath As String
    Dim filenum As Integer
    Dim CurrentDate As String
    Dim CurrentTime As String
    
    Dim FileName As String:     FileName = C_F_Macros
    Dim ModuleName As String:   ModuleName = C_M_Macros
    Dim MacroName As String:    MacroName = "ListMacros"
    Dim MsgBoxTitle As String:  MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    ' Ensure that the VBA project model is accessible
    Application.VBE.MainWindow.Visible = True
    
    ' Get the attached template's VBA project
    Set MyTemplate = ActiveDocument.AttachedTemplate
    Set vbProj = MyTemplate.VBProject
    
    ' Initialize output string
    output = "Subs and Functions in the attached template:" & vbCrLf & vbCrLf
    ModuleCounter = 0
    
    ' Loop through all components in the VBA project
    For Each vbComp In vbProj.VBComponents
        Set vbCodeMod = vbComp.CodeModule
        ModuleCounter = ModuleCounter + 1
        subCounter = 0
        funcCounter = 0
        output = output & ModuleCounter & ". " & vbComp.Name & " (" & GetComponentType(vbComp.Type) & ", code lines: " & vbCodeMod.CountOfLines & "):" & vbCrLf
        
        ' Loop through all procedures in the component
        lineNum = 1
        Do While lineNum < vbCodeMod.CountOfLines
            procName = vbCodeMod.ProcOfLine(lineNum, procKind)
            If procName <> "" Then
                If procKind = vbext_pk_Proc Then
                    subCounter = subCounter + 1
                    output = output & "    " & subCounter & ". " & procName & " (Sub)" & vbCrLf
                Else
                    funcCounter = funcCounter + 1
                    output = output & "    " & funcCounter & ". " & procName & " (Function)" & vbCrLf
                End If
                lineNum = lineNum + vbCodeMod.ProcCountLines(procName, procKind)
            Else
                lineNum = lineNum + 1
            End If
        Loop
        
        output = output & vbCrLf
    Next vbComp
    
    ' Save the output to a file
    FilePath = Options.DefaultFilePath(wdDocumentsPath) & "\" & MacroName & ".txt"
    filenum = FreeFile
    CurrentDate = Format(Date, "yyyy-mm-dd")
    CurrentTime = Format(Time, "hh:mm:ss")
    
    Open FilePath For Output As filenum
    Print #filenum, "Date: " & CurrentDate
    Print #filenum, "Time: " & CurrentTime
    Print #filenum, "File name: " & FileName
    Print #filenum, "Module name:" & ModuleName
    Print #filenum, "Macro name:" & MacroName
    Print #filenum, output
    Close filenum
    
    ' Display the summary in a message box
    MsgBox _
        Prompt:="Processing complete." & vbNewLine & vbNewLine & "Information was saved to the file:" _
            & vbNewLine & vbNewLine & FilePath, _
        Buttons:=vbInformation + vbOKOnly, _
        Title:=MsgBoxTitle
    
    ' Clear object variables
    Set MyTemplate = Nothing
    Set vbProj = Nothing
    Set vbCodeMod = Nothing
End Sub

Function GetComponentType(compType As VBIDE.vbext_ComponentType) As String
    Select Case compType
        Case vbext_ct_StdModule
            GetComponentType = "Standard Module"
        Case vbext_ct_ClassModule
            GetComponentType = "Class Module"
        Case vbext_ct_MSForm
            GetComponentType = "UserForm"
        Case vbext_ct_Document
            GetComponentType = "Document"
        Case Else
            GetComponentType = "Unknown"
    End Select
End Function

' Prepared after a case where I tried to call non-existent macro in KeyBinding.Add method.
' 2025-03-19 by ms and AI
Public Function MacroExists(MacroName As String) As Boolean
    Dim vbProj As VBIDE.VBProject
    Dim vbComp As VBIDE.VBComponent
    Dim vbCodeMod As VBIDE.CodeModule
    Dim lineNum As Long
    Dim procName As String
    Dim procType As VBIDE.vbext_ProcKind

    Set vbProj = ThisDocument.VBProject

    For Each vbComp In vbProj.VBComponents
        Set vbCodeMod = vbComp.CodeModule
        lineNum = 1
        Do While lineNum < vbCodeMod.CountOfLines
            procName = vbCodeMod.ProcOfLine(lineNum, procType)
            If procName <> "" Then
                If procName = MacroName Then
                    MacroExists = True
                    Exit Function
                End If
                lineNum = lineNum + vbCodeMod.ProcCountLines(procName, procType)
            Else
                lineNum = lineNum + 1
            End If
        Loop
    Next vbComp

    MacroExists = False
    
    ' Clear object variables
    Set vbProj = Nothing
    Set vbCodeMod = Nothing
End Function

' Sanity check of VBA code against unwanted characters in the macro code.
' 2025-12-10 by ms and AI
Sub ScanProjectForNonAscii()
    Dim vbProj As VBIDE.VBProject
    Dim vbComp As VBIDE.VBComponent
    Dim codeMod As VBIDE.CodeModule
    Dim i As Long, lineText As String, j As Long, ch As String
    Dim chCode As Long, hit As Boolean, report As String, flagged As String
    
    Set vbProj = Application.VBE.ActiveVBProject
    
    For Each vbComp In vbProj.VBComponents
        Set codeMod = vbComp.CodeModule
        For i = 1 To codeMod.CountOfLines
            lineText = codeMod.Lines(i, 1)
            hit = False: flagged = ""
            
            For j = 1 To Len(lineText)
                ch = Mid$(lineText, j, 1)
                chCode = AscW(ch)
                
                If chCode > 127 Then
                    hit = True
                    flagged = flagged & " U+" & Right$("0000" & Hex$(chCode), 4)
                End If
                
                Select Case chCode
                    Case ZWSP, ZWNJ, ZWJ, NBSP, BOM, &H2013, &H2014, &H2018, &H2019, &H201C, &H201D
                        If InStr(flagged, "U+" & Right$("0000" & Hex$(chCode), 4)) = 0 Then
                            hit = True
                            flagged = flagged & " U+" & Right$("0000" & Hex$(chCode), 4)
                        End If
                End Select
            Next j
            
            If hit Then
                report = report & vbCrLf & vbComp.Name & ":" & i & _
                         " ? [" & flagged & "] " & Replace(lineText, vbTab, "?TAB ")
            End If
        Next i
    Next vbComp
    
    If Len(report) = 0 Then
        MsgBox "No non-ASCII or suspicious characters found.", vbInformation
    Else
        Debug.Print "=== Suspicious code points report ==="
        Debug.Print report
        MsgBox "Scan complete. See the Immediate Window (Ctrl+G) for details.", vbInformation
    End If
End Sub

Private Sub RemoveContentFromThisDocument()
    Dim vbProj As VBIDE.VBProject
    Dim vbComp As VBIDE.VBComponent
    Dim doc As Document
    Dim contentExists As Boolean
    
    Dim FileName As String:     FileName = C_F_Macros
    Dim ModuleName As String:   ModuleName = C_M_Macros
    Dim MacroName As String:    MacroName = "RemoveContentFromThisDocument"
    Dim MsgBoxTitle As String:  MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    ' Get the current VBA project
    Set vbProj = ThisDocument.VBProject
    
    ' Find the ThisDocument component
    Set vbComp = vbProj.VBComponents("ThisDocument")
    
    ' Check if ThisDocument contains any content
    contentExists = vbComp.CodeModule.CountOfLines > 0
    
    If contentExists Then
        ' Remove the content from ThisDocument
        vbComp.CodeModule.DeleteLines 1, vbComp.CodeModule.CountOfLines
        MsgBox _
            Prompt:="Content removed from ThisDocument.", _
            Buttons:=vbInformation + vbOKOnly, _
            Title:=MsgBoxTitle
    Else
        MsgBox _
            Prompt:="No content found in ThisDocument.", _
            Buttons:=vbInformation + vbOKOnly, _
            Title:=MsgBoxTitle
    End If
    
    ' Clear object variables
    Set vbProj = Nothing
    Set vbComp = Nothing
End Sub
