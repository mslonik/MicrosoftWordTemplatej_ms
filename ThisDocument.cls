VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Document_Open()
    'Prepared by Microsoft Copilot and ms on 2025-02-05
    Dim templatePath1 As String
    Dim templatePath2 As String
    Dim docTemplate As String

    ' Display message box with macro name
    MsgBox "Running Macro: Document_Open," & vbCrLf & "attach empty Building Blocks"

    ' Paths to the global templates
    templatePath1 = "c:\Users\v523580\AppData\Roaming\Microsoft\Document Building Blocks\1045\16\Built-In Building Blocks.dotx"
    templatePath2 = "C:\Users\v523580\OneDrive - voestalpine\Inne dane\2025_Zadania\MicrosoftWordTemplate\ms.dotm"

    ' Get the document template path
    docTemplate = ThisDocument.AttachedTemplate.FullName

    ' Check if the document template matches the specified template
    If docTemplate <> templatePath2 Then
        ' Add and activate the global templates
        AddAndActivateGlobalTemplate templatePath1
        AddAndActivateGlobalTemplate templatePath2
    End If
End Sub

Private Sub AddAndActivateGlobalTemplate(templatePath As String)
    Dim addIn As addIn
    On Error Resume Next
    Set addIn = AddIns(templatePath)
    If addIn Is Nothing Then
        Set addIn = AddIns.Add(templatePath, True)
    End If
    addIn.Installed = True
    On Error GoTo 0
End Sub
