Attribute VB_Name = "RibbonControl"
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

' 2026-03-03 by ms
Sub BttnBlankPages(control As IRibbonControl)
    Select Case control.ID
        Case "Btn_AddBlank"
            Call Macros_ms.Tools.AddBlankPages
        Case "Btn_DelBlank"
            Call Macros_ms.Tools.DeleteTempBlankPages
    End Select
End Sub

' 2026-03-03 by ms
Sub BttnSectionTools(control As IRibbonControl)
    Select Case control.ID
        Case "Btn_AddSection"
            Call Macros_ms.Tools.AddSectionAndKillLinkToPrevious
        Case "Btn_UnlinkHF"
            Call Macros_ms.Tools.UnlinkAllHeadersFooters
    End Select
End Sub


