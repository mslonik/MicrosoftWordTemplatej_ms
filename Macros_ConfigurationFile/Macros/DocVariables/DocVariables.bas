Attribute VB_Name = "DocVariables"
' VBA Module name: DocVariables.bas
' https://github.com/mslonik/Microsoft-Word-Configuration
'
'   License: MIT License.
'
'+---------+-----------------------+-------------+----------------+-----------------------+
'| No.     | Sub name              | Ribbon name | Ribbon section | Ribbon button name    |
'+---------+-----------------------+-------------+----------------+-----------------------+
'| 1       | ShowDocVariables      | Tools_ms    | DocVariables   | ShowDocVariables      |
'| 2       | DeleteAllDocVariables | Tools_ms    | DocVariables   | DeleteAllDocVariables |
'+---------+-----------------------+-------------+----------------+-----------------------+
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'

' List document variables (no. 10 set of information stored within template)
' 2025-03-02 by ms and AI
Sub ShowDocVariables()
    Dim oVariable As Variable
    Dim ListOfDocVariables As String
    
    Dim FileName As String:     FileName = C_F_Macros
    Dim ModuleName As String:   ModuleName = C_M_DocVariables
    Dim MacroName As String:    MacroName = "ShowDocVariables"
    Dim MsgBoxTitle As String:  MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    ListOfDocVariables = ""
    
    For Each oVariable In ActiveDocument.Variables
        ListOfDocVariables = ListOfDocVariables & oVariable.Name & ": " & oVariable.Value & vbNewLine
    Next oVariable
    
    If ListOfDocVariables <> "" Then
        MsgBox _
            Prompt:="List of Document Variables:" & vbNewLine & vbNewLine & ListOfDocVariables, _
            Buttons:=vbInformation, _
            Title:=MsgBoxTitle
    Else
        MsgBox _
            Prompt:="There are no document variables in this document!" & vbNewLine & ActiveDocument.Name, _
            Buttons:=vbExclamation, _
            Title:=MsgBoxTitle
    End If
End Sub

' Delete all Document Variables
' 2025-03-02 by ms and AI
Sub DeleteAllDocVariables()
    Dim DocVar As Variable
    Dim docVarCount As Integer
    Dim DocName As String
    
    Dim FileName As String:     FileName = C_F_Macros
    Dim ModuleName As String:   ModuleName = C_M_DocVariables
    Dim MacroName As String:    MacroName = "DeleteAllDocVariables"
    Dim MsgBoxTitle As String:  MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    ' Get the current document name
    DocName = ActiveDocument.Name
    
    ' Count the number of document variables
    docVarCount = ActiveDocument.Variables.count
    
    ' Check if any document variables exist
    If docVarCount = 0 Then
        MsgBox _
            Prompt:="There are no Document Variables nn the current document." & vbNewLine & DocName, _
            Buttons:=vbInformation + vbOKOnly, _
            Title:=MsgBoxTitle
        Exit Sub
    End If
    
    ' Ask user for confirmation to delete all document variables
    Beep
    Dim UserDecision As VbMsgBoxResult
    UserDecision = MsgBox( _
        Prompt:="Are you sure you want to delete all Document Variables in the current document?" _
            & vbNewLine & DocName & vbNewLine & "This operation cannot be undone.", _
        Buttons:=vbYesNo + vbQuestion + vbDefaultButton2, _
        Title:=MsgBoxTitle)
    
    ' If user answers yes, delete all document variables
    If UserDecision = vbYes Then
        For Each DocVar In ActiveDocument.Variables
            DocVar.Delete
        Next DocVar
        MsgBox _
            Prompt:="All Document Variables have been deleted.", _
            Buttons:=vbInformation, _
            Title:=MsgBoxTitle
    Else
        ' If user answers no, exit the sub
        MsgBox _
            Prompt:="No Document Variables were deleted.", _
            Buttons:=vbInformation, _
            Title:=MsgBoxTitle
        Exit Sub
    End If
End Sub


