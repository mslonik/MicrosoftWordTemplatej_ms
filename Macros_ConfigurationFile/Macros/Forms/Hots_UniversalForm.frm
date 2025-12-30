VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Hots_UniversalForm 
   Caption         =   "UserForm1"
   ClientHeight    =   8415.001
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   7970
   OleObjectBlob   =   "Hots_UniversalForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Hots_UniversalForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
' The QueryClose event is a standard event for UserForms in VBA, and it should be named UserForm_QueryClose regardless of the form's name.
' This code is placed in the code module for the Hots_UniversalForm form. The UserForm_QueryClose event should be triggered when you close the form manually by clicking the "X" button in the top right corner.
Private MyInstanceName As String


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' Debug.Print "QueryClose event triggered"
    If CloseMode = vbFormControlMenu Then
        If Me.InstanceName = "Hotkey" Then
            Unload frmHotkey
            Set frmHotkey = Nothing
        End If
        If Me.InstanceName = "Hotstring" Then
            Unload frmHotstring
            Set frmHotstring = Nothing
        End If
        If Me.InstanceName = "HotMacros" Then
            Unload frmHotMacros
            Set frmHotMacros = Nothing
        End If
    End If
End Sub



Public Property Get InstanceName() As String
    InstanceName = MyInstanceName
End Property

Public Property Let InstanceName(Value As String)
    MyInstanceName = Value
End Property
