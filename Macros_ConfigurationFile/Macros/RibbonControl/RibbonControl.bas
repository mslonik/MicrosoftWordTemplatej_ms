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

' tab: Tools_ms
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
    End Select
End Sub

' tab: Tools_ms
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

' tab: Tools_ms
' group: CG_SectionTools
' 2026-03-03 by ms
Sub BttnSectionTools(control As IRibbonControl)
    Select Case control.ID
        Case "Btn_AddSection"
            Call Macros_ms.Tools.SectionAddNewAndUnlinkHF
        Case "Btn_UnlinkHF"
            Call Macros_ms.Tools.SectionUnlinkAllHF
        Case "Btn_RelinkHF"
            Call Macros_ms.Tools.SectionRelinkHF
    End Select
End Sub

' tab: Tools_ms
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

' tab: Tools_ms
' group: CG_Document
' 2026-03-04 by ms
Sub BttnDocument(control As IRibbonControl)
    Select Case control.ID
        Case "Btn_SetHyphenation"
            Call Macros_ms.Tools.SetHyphenation
        Case "Btn_ResetHyphenation"
            Call Macros_ms.Tools.ResetHyphenation
        Case "Btn_LangEngilshUS"
            Call Macros_ms.Tools.SetLanguageToEnglishUS
        Case "Btn_BckgrColor"
            Call Macros_ms.Tools.SetPageColorToCustom
        Case "Btn_ShowTemplates"
            Call Macros_ms.Tools.ShowAllTemplates
    End Select
End Sub

' tab: Tools_ms
' group: CG_Theme
' 2026-03-04 by ms
Sub BttnTheme(control As IRibbonControl)
    Select Case control.ID
        Case "Btn_ThemeAttach"
            Call Macros_ms.Theme.AttachTheme
    End Select
End Sub

' tab: Tools_ms
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

' tab: Tools_ms
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

' tab: Tools_ms
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

' tab: Tools_ms
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

' tab: Tools_ms
' group: CG_WordOptions
' 2026-03-04 by ms
Sub BttnWO(control As IRibbonControl)
    Select Case control.ID
        Case "Btn_WO_DisAll"
            Call Macros_ms.Tools.WordOptionsAutoCorrectAllDisable
        Case "Btn_WO_EnAll"
            Call Macros_ms.Tools.WordOptionsAutoCorrectAllEnable
        Case "Btn_WO_DisAutoFormatAsYouType"
            Call Macros_ms.Tools.WordOptionsAutoFormatAsYouTypeDisable
        Case "Btn_WO_EnDisAutoFormatAsYouType"
            Call Macros_ms.Tools.WordOptionsAutoFormatAsYouTypeEnable
        Case "Btn_WO_DisAutoCorrect"
            Call Macros_ms.Tools.WordOptionsAutoCorrectDisable
        Case "Btn_WO_EnAutoCorrect"
            Call Macros_ms.Tools.WordOptionsAutoCorrectEnable
        Case "Btn_WO_DisAutoFormat"
            Call Macros_ms.Tools.WordOptionsAutoFormatDisable
        Case "Btn_WO_EnAutoFormat"
            Call Macros_ms.Tools.WordOptionsAutoFormatEnable
    End Select
End Sub

' tab: Tools_ms
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

' tab: Tools_ms
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

' tab: Tools_ms
' group: CG_Fonts
' 2026-03-04 by ms
Sub BttnFonts(control As IRibbonControl)
    Select Case control.ID
        Case "Btn_Fonts_ShowUsed"
            Call Macros_ms.Fonts.ShowUsedFonts
    End Select
End Sub

' tab: Tools_ms
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

' tab: Macros_ms
' 2026-03-09 by ms
Sub BttnMacros(control As IRibbonControl)
    Select Case control.ID
        Case "Btn_MacrosImport"
            Call Macros_ms.Macros.ImportAllVBAModules
        Case "Btn_MacrosExport"
            Call Macros_ms.Macros.ExportAllVBAModules
        Case "Btn_MacrosDelExcept"
            Call Macros_ms.Macros.DeleteAllVBAModulesExceptMacros
        Case "Btn_MacrosDelAll"
            Call Macros_ms.Macros.DeleteAllVBAModules
    
        Case "Btn_MacrosCounter"
            Call Macros_ms.Macros.ShowMacrosCounter
        Case "Btn_MacrosList"
            Call Macros_ms.Macros.ListMacros
        Case "Btn_MacrosNonAscii"
            Call Macros_ms.Macros.ScanProjectForNonAscii
    End Select
End Sub

' tab: BB_Tools_ms
' 2026-03-09 by ms
Sub BttnBB(control As IRibbonControl)
    Select Case control.ID
        Case "Btn_BBExportAll"
            Call Macros_ms.BuildingBlocks.BB_ExportAll
        Case "Btn_BBExportSel"
            Call Macros_ms.BuildingBlocks.BB_ExportSelectedCategories
        Case "Btn_BBOpenBuiltIn"
            Call Macros_ms.BuildingBlocks.BB_OpenBuiltInTemplate
        Case "Btn_BBDelAll"
            Call Macros_ms.BuildingBlocks.BB_DeleteAll
        Case "Btn_BBList"
            Call Macros_ms.BuildingBlocks.BB_List
            
        Case "Btn_BBAdd"
            Call Macros_ms.BuildingBlocks.BB_Add
        Case "Btn_BBTemp"
            Call Macros_ms.BuildingBlocks.BB_InsertBBTemplate
        Case "Btn_BBRemHead"
            Call Macros_ms.BuildingBlocks.BB_DeleteHeaderPar
    End Select
End Sub

' tab: Shortcuts_ms
' 2026-03-09 by ms
Sub BttnShortcuts(control As IRibbonControl)
    Select Case control.ID
        Case "Btn_HotMacros"
            Call Macros_ms.Shortcuts.ShowFormHotMacros
        Case "Btn_HotHotstrings"
            Call Macros_ms.Shortcuts.ShowFormHotstrings
        Case "Btn_HotHotkeys"
            Call Macros_ms.Shortcuts.ShowFormHotkeys
        
        Case "Btn_DelStyleShortcuts"
            Call Macros_ms.Shortcuts.ClearActiveDocumentStyleShortcuts
        Case "Btn_DelMacroShortcuts"
            Call Macros_ms.Shortcuts.ClearActiveDocumentMacroShortcuts
        Case "Btn_DefMacroFullShortcuts"
            Call Macros_ms.Shortcuts.RemoveActiveDocumentMacroShortcuts
        
        Case "Btn_AddMacroShortcuts"
            Call Macros_ms.Shortcuts.CreateActiveDocumentMacroShortcuts
        
        Case "Btn_ListStyles"
            Call Macros_ms.Shortcuts.ListHotkeysToTxt
        Case "Btn_ListMacros"
            Call Macros_ms.Shortcuts.ListHotMacrosToTxt
        Case "Btn_ListHotstrings"
            Call Macros_ms.Shortcuts.ListHotstringsToTxt
        Case "Btn_ListAllHots"
            Call Macros_ms.Shortcuts.ListAllShortcutsToTxt
    
        Case "Btn_ListCommands"
            Call Macros_ms.Shortcuts.ListAllMWCommandsToDOCX
        Case "Btn_ListShortcuts"
            Call Macros_ms.Shortcuts.ListMWShortcutsToDOCX
    End Select
End Sub

' tab: Styles_ms
' 2026-03-10 by ms
Sub BttnStyles(control As IRibbonControl)
    Select Case control.ID
        Case "Btn_DelUnused"
            Call Macros_ms.StylesM.DeleteUnusedStyles
        
        Case "Btn_SwitchOffAuto"
            Call Macros_ms.StylesM.SwitchOffAutoupdate
        
        Case "Btn_ShowNoncompliant"
            Call Macros_ms.StylesM.ShowNonComplientStyling
        Case "Btn_DelNCStylingBookmarks"
            Call Macros_ms.StylesM.DeleteAllNCstylingBookmarks
        Case "Btn_DelNCHighlighting"
            Call Macros_ms.StylesM.DeleteNCHighlighting
        
        Case "Btn_AddCompliant"
            Call Macros_ms.StylesM.AddCompliantStyles
        Case "Btn_DelNonCompliant"
            Call Macros_ms.StylesM.DeleteNonCompliantStyles
        
        Case "Btn_ApplyTheme"
            Call Macros_ms.Theme.AttachTheme
            
        Case "Btn_OutputTxtInUse"
            Call Macros_ms.StylesM.ListStylesCurrentlyInUse
        Case "Btn_OutputTxtCompliant"
            Call Macros_ms.StylesM.ListCompliantStyles
        Case "Btn_OutputTxtBuiltin"
            Call Macros_ms.StylesM.ListBuiltInStyles
        Case "Btn_OutputTxtStylesCustom"
            Call Macros_ms.StylesM.ListCustomStyles
            
        Case "Btn_ReapplySimple"
            Call Macros_ms.StylesM.ReapplyStylesFromTemplateSimple
        Case "Btn_ReapplyFull"
            Call Macros_ms.StylesM.ReapplyStylesFromTemplateFull
    
        Case "Btn_ListTemplatesShowNamed"
            Call Macros_ms.StylesM.ListTemplatesShowNamed
        Case "Btn_ListTemplatesShowAll"
            Call Macros_ms.StylesM.ListTemplatesListAll
        Case "Btn_ListTemplatesResetLists"
            Call Macros_ms.StylesM.ResetAllListGalleries
    
        Case "Btn_TemplateAdd"
            Call Macros_ms.StylesM.ReapplyStylesFromTemplateSimple
        Case "Btn_TemplateDelOther"
            Call Macros_ms.StylesM.ReapplyStylesFromTemplateFull
    End Select
End Sub

' tab: Validation_ms
' 2026-03-10 by ms
Sub BttnVal(control As IRibbonControl)
    Select Case control.ID
        Case "Btn_TablesFormat"
            Call Macros_ms.Validation.Tables_Format
        Case "Btn_TablesCheckNesting"
            Call Macros_ms.Validation.Tables_CheckNestingLevel
        
        Case "Btn_UpdateAllFields"
            Call Macros_ms.Validation.UpdateAllFields
        Case "Btn_NoBrakeSpace"
            Call Macros_ms.Validation.InsertNoBrakeSpace
        Case "Btn_ReplaceUnwanted"
            Call Macros_ms.Validation.ReplaceUnwantedTextstrings
        Case "Btn_ModifyReferences"
            Call Macros_ms.Validation.ModifyReferencesToPicTab
            
        Case "Btn_FindCharStyle"
            Call Macros_ms.Validation.FindCharacterStyling
        Case "Btn_FindParStyle"
            Call Macros_ms.Validation.FindParagraphStyling
    End Select
End Sub

' tab: Content_ms
' 2026-03-10 by ms
Sub BttnContent(control As IRibbonControl)
    Select Case control.ID
        Case "Btn_NewFile"
            Call Macros_ms.Content.NewFileContent
    End Select
End Sub
