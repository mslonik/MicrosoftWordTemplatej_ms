Attribute VB_Name = "RibbonControl"
' VBA Module name: RibbonControl.bas
' https://github.com/mslonik/Microsoft-Word-Configuration
'
'   License: MIT License.
'
' This module serves as middle layer between customUI and the rest of existing macros.
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
' Used to enforce the explicit declaration of all variables in a module. When you include Option Explicit at the beginning of a module, it ensures that you must
' declare all variables using the Dim, Private, Public, ReDim, or Static statements before using them. This helps prevent errors caused by typos or
' undeclared variables.
Option Explicit

' tab: Layout_ms
' group: CG_PageSize
' 2026-03-03 by ms
Sub BttnPgFormats(control As IRibbonControl)
    Select Case control.ID
        Case "Btn_A4_V_12"
            Call Macros_ms.Tools.SetPageLayout_A4_V_1_2
        Case "Btn_A4_H_12"
            Call Macros_ms.Tools.SetPageLayout_A4_H_1_2
        Case "Btn_A4_H_05"
            Call Macros_ms.Tools.SetPageLayout_A4_H_0_5
        Case "Btn_A3_H_05"
            Call Macros_ms.Tools.SetPageLayout_A3_H_0_5
        Case "Btn_A3_V_1_2"
            Call Macros_ms.Tools.SetPageLayout_A3_V_1_2
        Case "Btn_AddBlank"
            Call Macros_ms.Tools.AddBlankPages
        Case "Btn_DelBlank"
            Call Macros_ms.Tools.DeleteTempBlankPages
        Case "Btn_AddSection"
            Call Macros_ms.Tools.AddSectionAndKillLinkToPrevious
        Case "Btn_UnlinkHF"
            Call Macros_ms.Tools.UnlinkAllHeadersFooters
    End Select
End Sub

' tab: Layout_ms
' group: CG_BlankPages
' 2026-03-03 by ms
Sub BttnBlankPages(control As IRibbonControl)
    Select Case control.ID
        Case "Btn_AddBlank"
            Call Macros_ms.Tools.AddBlankPages
        Case "Btn_DelBlank"
            Call Macros_ms.Tools.DeleteTempBlankPages
    End Select
End Sub

' tab: Layout_ms
' group: CG_SectionTools
' 2026-03-03 by ms
Sub BttnSectionTools(control As IRibbonControl)
    Select Case control.ID
        Case "Btn_AddSection"
            Call Macros_ms.Tools.AddSectionAndKillLinkToPrevious
        Case "Btn_UnlinkHF"
            Call Macros_ms.Tools.UnlinkAllHeadersFooters
    End Select
End Sub

' tab: Layout_ms
' group: CG_DocProperties
' 2026-03-04 by ms
Sub BttnDocProps(control As IRibbonControl)
    Select Case control.ID
        Case "Btn_UpdateProps"
            Call Macros_ms.Tools.DocPropertiesAddCustom
        Case "Btn_UserInput"
            Call Macros_ms.Tools.DocPropertiesUICustomEdit
        Case "Btn_DelCust"
            Call Macros_ms.Tools.DocPropertiesDeleteCustom
        Case "Btn_DelAll"
            Call Macros_ms.Tools.DocPropertiesDeleteAll
    End Select
End Sub

' tab: Layout_ms
' group: CG_Document
' 2026-03-04 by ms
Sub BttnDocument(control As IRibbonControl)
    Select Case control.ID
        Case "Btn_Hyphenation"
            Call Macros_ms.Tools.SetHyphenation
        Case "Btn_LangEngilshUS"
            Call Macros_ms.Tools.SetLanguageToEnglishUS
        Case "Btn_BckgrColor"
            Call Macros_ms.Tools.SetPageColorToCustom
        Case "Btn_ShowTemplates"
            Call Macros_ms.Tools.ShowAllTemplates
    End Select
End Sub

' tab: Layout_ms
' group: CG_Theme
' 2026-03-04 by ms
Sub BttnTheme(control As IRibbonControl)
    Select Case control.ID
        Case "Btn_ThemeAttach"
            Call Macros_ms.Theme.AttachTheme
    End Select
End Sub

' tab: Layout_ms
' group: CG_DocVariables
' 2026-03-04 by ms
Sub BttnDocVar(control As IRibbonControl)
    Select Case control.ID
        Case "Btn_ShowDocVariables"
            Call Macros_ms.DocVariables.ShowDocVariables
        Case "Btn_DeleteDocVariables"
            Call Macros_ms.DocVariables.DeleteAllDocVariables
    End Select
End Sub

' tab: Layout_ms
' group: CG_Comments
' 2026-03-04 by ms
Sub BttnComment(control As IRibbonControl)
    Select Case control.ID
        Case "Btn_CommAddNumber"
            Call Macros_ms.Tools.CommentAddNumber
        Case "Btn_CommDelete"
            Call Macros_ms.Tools.CommentDeleteNumber
        Case "Btn_CommCountByUser"
            Call Macros_ms.Tools.CommentCountByUser
    End Select
End Sub

' tab: Layout_ms
' group: CG_Canva
' 2026-03-04 by ms
Sub BttnCanva(control As IRibbonControl)
    Select Case control.ID
        Case "Btn_CanvaToggleBorder"
            Call Macros_ms.Tools.CanvaToggleBorder
        Case "Btn_CanvaInsertPNGfiles"
            Call Macros_ms.Tools.CanvaInsertPNGfiles
        Case "Btn_CanvaFormatTextBoxes"
            Call Macros_ms.Tools.CanvaFormatTextBoxes
    End Select
End Sub

' tab: Layout_ms
' group: CG_Captions
' 2026-03-04 by ms
Sub BttnCaption(control As IRibbonControl)
    Select Case control.ID
        Case "Btn_CaptionShow"
            Call Macros_ms.Tools.CaptionShow
        Case "Btn_CaptionAddCustom"
            Call Macros_ms.Tools.CapationAddCustomized
        Case "Btn_CaptionDelCustom"
            Call Macros_ms.Tools.CaptionLabelDeleteCustomized
    End Select
End Sub

' tab: Layout_ms
' group: CG_WordOptions
' 2026-03-04 by ms
Sub BttnWO(control As IRibbonControl)
    Select Case control.ID
        Case "Btn_WO_Customize"
            Call Macros_ms.Tools.WordOptionsCustomize
        Case "Btn_WO_Restore"
            Call Macros_ms.Tools.WordOptionsRestore
        Case "Btn_WO_DisAutoFormat"
            Call Macros_ms.Tools.WordOptionsDisableAutoFormat
        Case "Btn_WO_ResAutoFormat"
            Call Macros_ms.Tools.WordOptionsRestoreAutoFormat
        Case "Btn_WO_DisAutoCorrect"
            Call Macros_ms.Tools.WordOptionsDisableAutoCorrect
        Case "Btn_WO_ToggAutoCorrect"
            Call Macros_ms.Tools.WordOptionsToggleAutoCorrect
    End Select
End Sub

' tab: Layout_ms
' group: CG_BefPrinting
' 2026-03-04 by ms
Sub BttnBP(control As IRibbonControl)
    Select Case control.ID
        Case "Btn_BP_DelBookmarks"
            Call Macros_ms.Tools.DeleteAllUserBookmarks
        Case "Btn_BP_DistCheck"
            Call Macros_ms.Tools.ParDistAtNewSectionCheck
        Case "Btn_BP_DistReduce"
            Call Macros_ms.Tools.ParDistAtNewSectionReduce
        Case "Btn_BP_DistRestore"
            Call Macros_ms.Tools.ParDistAtNewSectionRestore
    End Select
End Sub

' tab: Layout_ms
' group: CG_Tables
' 2026-03-04 by ms
Sub BttnTables(control As IRibbonControl)
    Select Case control.ID
        Case "Btn_Tab_Cust"
            Call Macros_ms.Tools.Table_CustomizeFormatting
        Case "Btn_Tab_KeepOnPage"
            Call Macros_ms.Tools.Table_KeepOnOnePage
    End Select
End Sub

' tab: Layout_ms
' group: CG_Fonts
' 2026-03-04 by ms
Sub BttnFonts(control As IRibbonControl)
    Select Case control.ID
        Case "Btn_Fonts_ShowUsed"
            Call Macros_ms.Fonts.ShowUsedFonts
    End Select
End Sub

' tab: Layout_ms
' group: CG_Fonts
' 2026-03-04 by ms
Sub BttnTabulators(control As IRibbonControl)
    Select Case control.ID
        Case "Btn_Tabs_Show"
            Call Macros_ms.Tools.TabDefaultShow
        Case "Btn_Tabs_SetCustom"
            Call Macros_ms.Tools.TabDefaultSetCustom
        Case "Btn_Tabs_Restore"
            Call Macros_ms.Tools.TabDefaultRestore
    End Select
End Sub


