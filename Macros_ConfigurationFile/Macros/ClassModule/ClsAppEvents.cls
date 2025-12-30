VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsAppEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' This trick enables checking of substituted fonts in every newly opened file in Microsoft Word.
' 2025-08-28 by ms and AI
Public WithEvents appWord As Word.Application
Attribute appWord.VB_VarHelpID = -1

Private Sub AppWord_DocumentOpen(ByVal doc As Word.Document)
    Call ShowSubstitutedFonts(doc)   ' in module Fonts
End Sub
