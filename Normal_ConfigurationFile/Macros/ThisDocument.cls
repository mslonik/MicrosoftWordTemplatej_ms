VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' 2026-03-05 by ms and AI
Private Sub Document_Open()
    Dim DocVarName As String: DocVarName = "DocVarToggleSpecificFormatting"
    Dim SavedMode As Integer
    
    Dim FileName As String:     FileName = "Normal.dotm"
    Dim ModuleName As String:   ModuleName = "ThisDocument"
    Dim MacroName As String:    MacroName = "Document_Open"
    Dim MsgBoxTitle As String:  MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    On Error Resume Next
    ' Check if the variable exists in the document being opened
    SavedMode = CInt(ActiveDocument.Variables(DocVarName).Value)
    
    If Err.Number = 0 Then
        ' The variable exists! Apply the view without incrementing.
        MsgBox _
            Prompt:="Restoring saved view mode.", _
            Buttons:=vbInformation, _
            Title:=MsgBoxTitle
        Call Macros_ms.Tools.ApplyViewMode(ViewMode:=SavedMode)
    End If
    On Error GoTo 0
End Sub
