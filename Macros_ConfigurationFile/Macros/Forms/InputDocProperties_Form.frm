VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} InputDocProperties_Form 
   Caption         =   "Input user document properties"
   ClientHeight    =   5505
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   9380.001
   OleObjectBlob   =   "InputDocProperties_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "InputDocProperties_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' 2025-04-06 by ms and AI
Private Sub UserForm_Initialize()
    Dim doc As Document
    Set doc = ActiveDocument
    Dim FileName As String
    FileName = C_F_Macros
    
    Dim ModuleName As String
    ModuleName = C_M_Forms
    
    Dim MacroName As String
    MacroName = "UserForm_Initialize"
    
    Dim MsgBoxTitle As String
    MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    ' CustomPropertyName 1 ÷ 7 are defined in Tools module.
    ' Check if such properties exist
    If doc.CustomDocumentProperties(C_CPN_1).Name = "" Or _
        doc.CustomDocumentProperties(C_CPN_2).Name = "" Or _
        doc.CustomDocumentProperties(C_CPN_3).Name = "" Or _
        doc.CustomDocumentProperties(C_CPN_4).Name = "" Or _
        doc.CustomDocumentProperties(C_CPN_5).Name = "" Or _
        doc.CustomDocumentProperties(C_CPN_7).Name = "" Then
        MsgBox Prompt:="Specific custom properties were not found." & vbNewLine & _
            "Perhaps you need to run the macro 'UpdateDocProperties'? Exiting", _
            Buttons:=vbCritical, _
            Title:=MsgBoxTitle
    End If
    
    ' Initialize Drop-Down List aka ComboBox
    Me.ComboBox1.AddItem "confidential"
    Me.ComboBox1.AddItem "internal document"
    Me.ComboBox1.AddItem "public document"
    
    ' Initialize Form name
    Me.Caption = "Input user document properties : " & doc.Name
    
    ' Load existing values for custom document properties
    On Error Resume Next
    Me.TextBox1.Text = doc.CustomDocumentProperties(C_CPN_1).Value
    Me.TextBox2.Text = doc.CustomDocumentProperties(C_CPN_2).Value
    Me.TextBox3.Text = doc.CustomDocumentProperties(C_CPN_3).Value
    Me.TextBox4.Text = doc.CustomDocumentProperties(C_CPN_4).Value
    Me.ComboBox1.Text = doc.CustomDocumentProperties(C_CPN_7).Value
    On Error GoTo 0
    
    ' Clear objects
    Set doc = Nothing
    
End Sub

' 2025-04-06 by ms and AI
Private Sub CancelButton1_Click()
    Unload Me
End Sub

' 2025-04-06 by ms and AI
Private Sub OkButton2_Click()
    Dim doc As Document
    Set doc = ActiveDocument
    
    On Error Resume Next
    doc.CustomDocumentProperties(C_CPN_1).Value = Me.TextBox1.Text
    doc.CustomDocumentProperties(C_CPN_2).Value = Me.TextBox2.Text
    doc.CustomDocumentProperties(C_CPN_3).Value = Me.TextBox3.Text
    doc.CustomDocumentProperties(C_CPN_4).Value = Me.TextBox4.Text
    doc.CustomDocumentProperties(C_CPN_7).Value = Me.ComboBox1.Text
    On Error GoTo 0
    
    ' Clear objects
    Set doc = Nothing
    
    Unload Me
    
    Call UpdateAllFields    ' module: Validation
End Sub

