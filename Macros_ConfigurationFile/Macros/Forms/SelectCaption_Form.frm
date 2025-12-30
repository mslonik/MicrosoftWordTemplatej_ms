VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SelectCaption_Form 
   Caption         =   "Select caption"
   ClientHeight    =   2490
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   4560
   OleObjectBlob   =   "SelectCaption_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SelectCaption_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Initialize()
    Me.Caption = "Select caption"
    OB_PicSmall.Value = True ' Default selection
End Sub

Private Sub CommandButtonOK_Click()
    Me.Hide
End Sub

Private Sub CommandButtonCancel_Click()
    Me.Hide
End Sub
