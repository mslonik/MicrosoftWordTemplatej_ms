VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} KeyboardShortcuts 
   Caption         =   "Template Keyboard Shortcuts"
   ClientHeight    =   4005
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7290
   OleObjectBlob   =   "KeyboardShortcuts.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "KeyboardShortcuts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    ' Define the data to be displayed in the ListBox
    Dim data As Variant
    data = Array( _
        Array("strikethrough text", "Ctrl + Shift + X", "CrossoutText ms"), _
        Array("hide text", "Ctrl + Shift + H", "HiddenText ms"), _
        Array("list indentation level 1", "Alt + L, 1", "ListParIdent 1 ms"), _
        Array("list indentation level 2", "Alt + L, 2", "ListParIdent 2 ms"), _
        Array("list indentation level 3", "Alt + L, 3", "ListParIdent 3 ms"), _
        Array("list indentation level 4", "Alt + L, 4", "ListParIdent 4 ms"), _
        Array("list numbering", "Alt + L, N", "NumberingOrdered ms"), _
        Array("list punctuation", "Alt + L, B", "NumberingBullets ms"), _
        Array("list of references", "Alt + L, R", "NumberingReferences ms"), _
        Array("normal style for tables", "Alt + N, T", "Normal in table ms"), _
        Array("default style of text paragraph", "Ctrl + Shift + N", "Normal ms"), _
        Array("text style below list", "Alt + N, B", "Normal below ms"), _
        Array("normal style above list or table", "Alt + N, A", "Normal above ms"), _
        Array("label of a picture", "Alt + P, R", "Legend picture ms"), _
        Array("label of a table", "Alt + P, T", "Legend table ms"), _
        Array("font settings", "Alt + C", ""), _
        Array("paragraph settings", "Alt + A", "") _
    )
    
    ' Set the ColumnCount property
    ListBox1.ColumnCount = 3
    
    ' Populate the ListBox
    Dim i As Integer
    For i = LBound(data) To UBound(data)
        ListBox1.AddItem
        ListBox1.List(i, 0) = data(i)(0)
        ListBox1.List(i, 1) = data(i)(1)
        ListBox1.List(i, 2) = data(i)(2)
    Next i
End Sub
