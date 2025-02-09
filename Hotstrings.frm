VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Hotstrings 
   Caption         =   "Template hotstrings (Autotext)"
   ClientHeight    =   12630
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5775
   OleObjectBlob   =   "Hotstrings.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Hotstrings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Written by Microsoft Copilot M365 and ms on 2025-02-09.
' Displays all building blocks which belong to AutoText category

Private Sub UserForm_Initialize()
    ShowHotstrings
End Sub

Private Sub ShowHotstrings()
    Dim objTemplate As template
    Dim objBBType As BuildingBlockType
    Dim objCategory As Category
    Dim objBB As buildingBlock
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer

    ' Reference the currently attached template
    Set objTemplate = ActiveDocument.AttachedTemplate

    ' Clear the listbox
    Me.ListBox1.Clear

    ' Access the BuildingBlockTypes collection
    Set objBBType = objTemplate.BuildingBlockTypes(wdTypeAutoText)

    ' Loop through the Categories collection using a For loop
    For j = 1 To objBBType.Categories.Count
        Set objCategory = objBBType.Categories(j)
        ' Loop through the BuildingBlocks collection for each Category using a For loop
        For k = 1 To objCategory.BuildingBlocks.Count
            Set objBB = objCategory.BuildingBlocks(k)
            ' Add the building block name to the listbox
            Me.ListBox1.AddItem
            Me.ListBox1.List(i, 0) = objBB.Name
            Me.ListBox1.List(i, 1) = objBB.Name
            i = i + 1
        Next k
    Next j
End Sub

Private Sub UserForm_Click()

End Sub
