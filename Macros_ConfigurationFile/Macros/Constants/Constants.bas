Attribute VB_Name = "Constants"
' VBA Module name: Constants.bas
' https://github.com/mslonik/Microsoft-Word-Configuration
'
'   License: MIT License.
'
' Contains all customized constants applied in the code.
' VBA code file must be saved as ANSI text format (not UTF). By Default this file is ANSI Windows-1252 (Western).
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
' Used to enforce the explicit declaration of all variables in a module. When you include Option Explicit at the beginning of a module, it ensures that you must
' declare all variables using the Dim, Private, Public, ReDim, or Static statements before using them. This helps prevent errors caused by typos or
' undeclared variables.
Option Explicit

Public Const C_StyleSuffix As String = " ms"

' Font names
Public Const C_FT_Body As String = "+Body"                      ' Body, as in Design -> Fonts -> Customize font; to inheritage from Theme
Public Const C_FT_Headings As String = "+Headings"              ' Heading, as in Design -> Fonts -> Customize font; to inheritage from Theme
Public Const C_FT_AntiHomoglyph As String = "Consolas"

' File names
Public Const C_F_BBB = "Built-In Building Blocks.dotx"          ' BBB = Built-in BuildingBlocks, default file (re)created by Microsoft Word
Public Const C_F_Normal = "Normal.dotm"
Public Const C_F_BuildingBlocks As String = "BB_ms.dotm"        ' Name of the template file containing BuildingBlocks
Public Const C_F_Theme As String = "Theme_ms.thmx"
Public Const C_F_Macros As String = "Macros_ms.dotm"

' Module names from C_F_Macros
Public Const C_M_BuildingBlocks = "BuildingBlocks"
Public Const C_M_Constants = "Constants"
Public Const C_M_DocVariables = "DocVariables"
Public Const C_M_Fonts = "Fonts"
Public Const C_M_Macros = "Macros"
Public Const C_M_Scenarios = "Scenarios"
Public Const C_M_Shortcuts = "Shortcuts"
Public Const C_M_Styles = "StylesM"                             ' It cannot be named as just Styles, because VBA may thinks I'm using a type name that conflicts with something else in my project
Public Const C_M_Template = "Template"
Public Const C_M_Theme = "Theme"
Public Const C_M_Tools = "Tools"
Public Const C_M_Validation = "Validation"
Public Const C_M_ThisDocument = "ThisDocument"

' C_S, C = Constant, S = Style
' paragraph styles:
Public Const C_S_Heading1 As String = "ParHeading 1 ms"
Public Const C_S_Heading2 As String = "ParHeading 2 ms"
Public Const C_S_Heading3 As String = "ParHeading 3 ms"
Public Const C_S_Heading4 As String = "ParHeading 4 ms"
Public Const C_S_Heading5 As String = "ParHeading 5 ms"
Public Const C_S_Heading6 As String = "ParHeading 6 ms"
Public Const C_S_Heading7 As String = "ParHeading 7 ms"
Public Const C_S_Heading8 As String = "ParHeading 8 ms"
Public Const C_S_ParInTable As String = "ParInTable ms"
Public Const C_S_ParLegalNote As String = "ParLegalNote ms"
Public Const C_S_PictureLegend As String = "ParLegendPicture ms"
Public Const C_S_TableLegend As String = "ParLegendTable ms"
Public Const C_S_ListHeading As String = "ParListHeading ms"
Public Const C_S_ListLevel1 As String = "ParListIndent1 ms"
Public Const C_S_ListLevelB1 As String = "ParListIndentB1 ms"
Public Const C_S_ListLevel2 As String = "ParListIndent2 ms"
Public Const C_S_ListLevelB2 As String = "ParListIndentB2 ms"
Public Const C_S_ListLevel3 As String = "ParListIndent3 ms"
Public Const C_S_ListLevelB3 As String = "ParListIndentB3 ms"
Public Const C_S_ListLevel4 As String = "ParListIndent4 ms"
Public Const C_S_ListLevelB4 As String = "ParListIndentB4 ms"
Public Const C_S_ParMinimal As String = "ParMinimal ms"
Public Const C_S_ParNormal As String = "ParNormal ms"
Public Const C_S_ParNormalZero As String = "ParNormalZero ms"
Public Const C_S_ParNormalBelow As String = "ParNormalBelow ms"
Public Const C_S_ParNormalAbove As String = "ParNormalAbove ms"
Public Const C_S_ParNormalAB As String = "ParNormalAB ms"
Public Const C_S_ParPictureCanva As String = "ParPictureCanva ms"
Public Const C_S_ParIcon As String = "ParIcon ms"
Public Const C_S_ParSourceCode As String = "ParSourceCode ms"
Public Const C_S_TextBoxes As String = "ParTextBoxes ms"
Public Const C_S_ListNumRef As String = "ParNumRef ms"
Public Const C_S_ListNumTable As String = "ParListInTable ms"
' character styles:
Public Const C_S_Bold As String = "CharBold ms"
Public Const C_S_CharCrossout As String = "CharCrossout ms"
Public Const C_S_CharDefault As String = "CharDefault ms"
Public Const C_S_CharHidden As String = "CharHidden ms"
Public Const C_S_Italic As String = "CharItalic ms"
Public Const C_S_CharLegalNote As String = "CharLegalNote ms"
Public Const C_S_CharSourceCode As String = "CharSourceCode ms"
Public Const C_S_Underline As String = "CharUnderline ms"
' table styles:
Public Const C_S_TabTable As String = "TabTable ms"
Public Const C_S_TabNoGrid As String = "TabTableNoGrid ms"
Public Const C_S_TabNoPadding As String = "TabTableNoPadding ms"

' C_LT = Constant List Template aka multilevel lists
Public Const C_LT_Headings As String = "MultiLevelList Headings ms"
Public Const C_LT_NumOrd As String = "MultiLevelList Ordered ms"
Public Const C_LT_Bullets As String = "MultiLevelList Bullets ms"
Public Const C_LT_ListNumRef As String = "SingleLevelListNumRef ms"
Public Const C_LT_ListNumTable As String = "SingleLevelListInTable ms"
Public Const C_LT_TOC As String = "MultilevelList Table of Content ms"

' Constants related to styles setting
Public Const C_DistParBAList As Single = 3      ' Distance Before and After paragraph of a list
Public Const C_BaseIndent As Single = 0.3       ' cm

' ===========================
' Character Styles Table
' ===========================
Public CharacterStyles As Variant
' ===========================
' Paragraph Styles Table
' ===========================
Public ParagraphStyles As Variant
' ===========================
' Table Styles Table
' ===========================
Public TableStyles As Variant
' ===========================
' List Templates Table
' ===========================
Public ListTemplates As Variant

' C = Constant
Public Const C_PointsToCm As Single = 0.0352778

' C_BM = Constant BookMarks
Public Const C_BM_LastCursorPosition As String = "ms_LastCursorPosition"
Public Const C_BM_ReducedDistance As String = "ms_ReducedDistance_"
Public Const C_BM_Picture As String = "ms_picture_"
Public Const C_BM_NCstylingP As String = "NCstylingP_"
Public Const C_BM_NCstylingT As String = "NCstylingT_"
Public Const C_BM_SearchedStyle As String = "ms_SearchedStyle_"

Public Const C_Caption_Tab As String = "Tab."
Public Const C_Caption_Pic As String = "Pic."
Public Const C_Caption_PicSmall As String = "pic."
Public Const C_Caption_TabSmall As String = "tab."

' C_BB = Constant Building Block
Public Const C_BB_LegendPicture As String = "LegendPicture"

' C_CPN = Constant _ Custom Property Name
Public Const C_CPN_1 As String = "ms_DocumentID"
Public Const C_CPN_2 As String = "ms_DocumentTitle1"
Public Const C_CPN_3 As String = "ms_DocumentTitle2"
Public Const C_CPN_4 As String = "ms_DocumentCategory"
Public Const C_CPN_5 As String = "ms_SVN_Revision"
Public Const C_CPN_6 As String = "ms_reserved"
Public Const C_CPN_7 As String = "ms_Confidentiality"
Public Const C_CPN_8 As String = "ms_reserved"
Public Const C_CPN_9 As String = "ms_reserved"
Public Const C_CPN_10 As String = "ms_reserved"

' C_CPV = Constant — Custom Property Value
Public Const C_CPV_1 As String = ""
Public Const C_CPV_2 As String = "product / system short name"
Public Const C_CPV_3 As String = "product / system full name"
Public Const C_CPV_4 As String = ""
Public Const C_CPV_5 As String = ""
Public Const C_CPV_6 As String = ""
Public Const C_CPV_7 As String = "internal document"
Public Const C_CPV_8 As String = ""
Public Const C_CPV_9 As String = ""
Public Const C_CPV_10 As String = ""

' C_TTR = Constant Tag To Remove
Public Const C_TTR_name As String = "<bb_name>"
Public Const C_TTR_type As String = "<bb_type>"
Public Const C_TTR_category As String = "<bb_category>"
Public Const C_TTR_description As String = "<bb_description>"
Public Const C_TTR_insertoptions As String = "<bb_insertoptions>"

' C_SC = Constant ShortCut
Public Const C_SC_AltHplusS As String = "Alt + H, S"            ' ShowFormHotstrings
Public Const C_SC_AltHplusK As String = "Alt + H, K"            ' ShowFormHotkeys
Public Const C_SC_AltHplusM As String = "Alt + H, M"            ' ShowFormHotMacros
Public Const C_SC_AltRplusS As String = "Alt + R, S"            ' ReapplyTemplateStyle
Public Const C_SC_AltLplusR As String = "Alt + L, R"            ' RestartListNumbering
Public Const C_SC_AltF2 As String = "Alt + F2"                  ' JumpToNextList
Public Const C_SC_AltF3 As String = "Alt + F3"                  ' JumpToNextTable
Public Const C_SC_AltF5 As String = "Alt + F5"                  ' JumpToNextCanvas
Public Const C_SC_AltF9 As String = "Alt + F9"                  ' CustomizedToggleFieldCodes
Public Const C_SC_AltF As String = "Alt + F"                    ' FormatFont
Public Const C_SC_CtrlB As String = "Ctrl + B"                  ' ToggleCharBoldStyle
Public Const C_SC_CtrlI As String = "Ctrl + I"                  ' ToggleCharItalicStyle
Public Const C_SC_CtrlP As String = "Ctrl + P"                  ' CustomizedPrinting
Public Const C_SC_CtrlU As String = "Ctrl + U"                  ' ToggleCharUnderlineStyle
Public Const C_SC_CtrlS As String = "Ctrl + S"                  ' ApplyDistanceBetweenNumberingAndHeading
Public Const C_SC_CtrlW As String = "Ctrl + W"                  ' UpdateAllFieldsAndCloseFile
Public Const C_SC_CtrlF2 As String = "Ctrl + F2"                ' CustomizedPrintPreviewAndPrint
Public Const C_SC_ShiftCtrlX As String = "Shift + Ctrl + X"     ' ToggleCharCrossoutStyle
Public Const C_SC_ShiftCtrlH As String = "Shift + Ctrl + H"     ' ToggleCharHiddenStyle
Public Const C_SC_ShiftCtrlK As String = "Shift + Ctrl + K"     ' ToggleCharSourceCode
Public Const C_SC_ShiftCtrlS As String = "Shift + Ctrl + S"     ' ToggleApplyStyles
Public Const C_SC_ShiftCtrlC As String = "Shift + Ctrl + C"     ' CopyFormat
Public Const C_SC_ShiftCtrlV As String = "Shift + Ctrl + V"     ' PasteFormat
Public Const C_SC_F4 As String = "F4"                           ' ToggleSpecificFormatting
Public Const C_SC_F7 As String = "F7"                           ' InsertCrossRef
Public Const C_SC_F8 As String = "F8"                           ' SetLanguageToEnglishUS
Public Const C_SC_F12 As String = "F12"                         ' CustomizedSaveAs
Public Const C_SC_Insert As String = "Insert"                   ' CustomizedOvertype
Public Const C_SC_AltCtrlH As String = "Alt + Ctrl + H"         ' NavPane
Public Const C_SC_AltCtrlP As String = "Alt + Ctrl + P"         ' FormatParagraph
Public Const C_SC_AltCtrlSqOpen As String = "Alt + Ctrl + ["    ' ToggleHeading

' C_DV = Constant DocumentVariable
Public Const C_DV_NewFileConfAndContent As String = "NewFileConfAndContent"

' Characters to flag explicitly, used in module Macros ->
Public Const ZWSP   As Long = &H200B
Public Const ZWNJ   As Long = &H200C
Public Const ZWJ    As Long = &H200D
Public Const NBSP   As Long = &HA0     ' U+00A0
Public Const BOM    As Long = &HFEFF  ' Zero-width no-break space


' --- Lowercase Polish diacritic letters ---
' In VBA, a Property is like a special procedure that behaves like a variable but can run code when accessed.
' Property Get is the keyword that defines the getter—the procedure that returns a value when you read the property.
' ChrWstands for Character Wide for Unicode charactes, argument type Long, &H stands for hexadecimal literal, returned value is Unicode character
' E.g. &H107 = U+0107 = latin small letter c with acute

Public Property Get PolishDiacritic_A_Lowercase() As String   ' a
    PolishDiacritic_A_Lowercase = ChrW(&H105)
End Property

Public Property Get PolishDiacritic_C_Lowercase() As String    ' c
    PolishDiacritic_C_Lowercase = ChrW(&H107)
End Property

Public Property Get PolishDiacritic_E_Lowercase() As String   ' e
    PolishDiacritic_E_Lowercase = ChrW(&H119)
End Property

Public Property Get PolishDiacritic_L_Lowercase() As String      ' l
    PolishDiacritic_L_Lowercase = ChrW(&H142)
End Property

Public Property Get PolishDiacritic_N_Lowercase() As String    ' n
    PolishDiacritic_N_Lowercase = ChrW(&H144)
End Property

Public Property Get PolishDiacritic_O_Lowercase() As String    ' ó
    PolishDiacritic_O_Lowercase = ChrW(&HF3)
End Property

Public Property Get PolishDiacritic_S_Lowercase() As String    ' s
    PolishDiacritic_S_Lowercase = ChrW(&H15B)
End Property

Public Property Get PolishDiacritic_Zacute_Lowercase() As String    ' z
    PolishDiacritic_Zacute_Lowercase = ChrW(&H17A)
End Property

Public Property Get PolishDiacritic_Zdot_Lowercase() As String      ' z
    PolishDiacritic_Zdot_Lowercase = ChrW(&H17C)
End Property

' Example, alternative literal for specific Polish style name:
'Public Property Get C_S_Heading1() As String
'    C_S_Heading1 = "Heading 1,Nag" & PolishDiacritic_L_Lowercase & PolishDiacritic_O_Lowercase & "wek 1 ms"
'End Property

' --- Uppercase Polish letters ---
Public Property Get Polish_A_ogonek() As String   ' A
    Polish_A_ogonek = ChrW(&H104)
End Property

Public Property Get Polish_C_acute() As String    ' C
    Polish_C_acute = ChrW(&H106)
End Property

Public Property Get Polish_E_ogonek() As String   ' E
    Polish_E_ogonek = ChrW(&H118)
End Property

Public Property Get Polish_L_bar() As String      ' L
    Polish_L_bar = ChrW(&H141)
End Property

Public Property Get Polish_N_acute() As String    ' N
    Polish_N_acute = ChrW(&H143)
End Property

Public Property Get Polish_O_acute() As String    ' Ó
    Polish_O_acute = ChrW(&HF3) ' Same code point for uppercase Ó
End Property

Public Property Get Polish_S_acute() As String    ' S
    Polish_S_acute = ChrW(&H15A)
End Property

Public Property Get Polish_Z_acute() As String    ' Z
    Polish_Z_acute = ChrW(&H179)
End Property

Public Property Get Polish_Z_dot() As String      ' Z
    Polish_Z_dot = ChrW(&H17B)
End Property


Public Sub InitParagraphStyles()
    ParagraphStyles = Array( _
        Array("C_S_Heading1", "ParHeading 1 ms"), Array("C_S_Heading2", "ParHeading 2 ms"), Array("C_S_Heading3", "ParHeading 3 ms"), Array("C_S_Heading4", "ParHeading 4 ms"), Array("C_S_Heading5", "ParHeading 5 ms"), Array("C_S_Heading6", "ParHeading 6 ms"), Array("C_S_Heading7", "ParHeading 7 ms"), Array("C_S_Heading8", "ParHeading 8 ms"), _
        Array("C_S_ParInTable", "ParInTable ms"), _
        Array("C_S_ParLegalNote", "ParLegalNote ms"), _
        Array("C_S_PictureLegend", "ParLegendPicture ms"), _
        Array("C_S_TableLegend", "ParLegendTable ms"), _
        Array("C_S_ListHeading", "ParListHeading ms"), Array("C_S_ListLevel1", "ParListIndent1 ms"), Array("C_S_ListLevelB1", "ParListIndentB1 ms"), Array("C_S_ListLevel2", "ParListIndent2 ms"), Array("C_S_ListLevelB2", "ParListIndentB2 ms"), Array("C_S_ListLevel3", "ParListIndent3 ms"), Array("C_S_ListLevelB3", "ParListIndentB3 ms"), Array("C_S_ListLevel4", "ParListIndent4 ms"), Array("C_S_ListLevelB4", "ParListIndentB4 ms"), _
        Array("C_S_ParMinimal", "ParMinimal ms"), _
        Array("C_S_ParNormal", "ParNormal ms"), _
        Array("C_S_ParNormalZero", "ParNormalZero ms"), _
        Array("C_S_ParNormalBelow", "ParNormalBelow ms"), _
        Array("C_S_ParNormalAbove", "ParNormalAbove ms"), _
        Array("C_S_ParNormalAB", "ParNormalAB ms"), _
        Array("C_S_ParPictureCanva", "ParPictureCanva ms"), _
        Array("C_S_ParSourceCode", "ParSourceCode ms"), _
        Array("C_S_TOC1", "TOC1"), Array("C_S_TOC2", "TOC2"), Array("C_S_TOC3", "TOC3"), _
        Array("C_S_TextBoxes", "ParTextBoxes ms"), _
        Array("C_S_ListNumRef", "ParNumRef ms"), _
        Array("C_S_ListNumTable", "ParListInTable ms"), _
        Array("C_S_ParIcon", "ParIcon ms") _
    )
End Sub

Public Sub InitCharacterStyles()
    CharacterStyles = Array( _
        Array("C_S_Bold", "CharBold ms"), _
        Array("C_S_CharCrossout", "CharCrossout ms"), _
        Array("C_S_CharDefault", "CharDefault ms"), _
        Array("C_S_CharHidden", "CharHidden ms"), _
        Array("C_S_Italic", "CharItalic ms"), _
        Array("C_S_CharLegalNote", "CharLegalNote ms"), _
        Array("C_S_CharSourceCode", "CharSourceCode ms"), _
        Array("C_S_Underline", "CharUnderline ms") _
    )
End Sub

Public Sub InitTableStyles()
    TableStyles = Array( _
        Array("C_S_TabTable", "TabTable ms"), _
        Array("C_S_TabNoGrid", "TabTableNoGrid ms"), _
        Array("C_S_TabNoPadding", "TabTableNoPadding ms") _
    )
End Sub

Public Sub InitListTemplates()
    ListTemplates = Array( _
        Array("C_LT_Headings", "MultiLevelList Headings ms"), _
        Array("C_LT_NumOrd", "MultiLevelList Ordered ms"), _
        Array("C_LT_Bullets", "MultiLevelList Bullets ms"), _
        Array("C_LT_ListNumRef", "SingleLevelListNumRef ms"), _
        Array("C_LT_ListNumTable", "SingleLevelListInTable ms") _
    )
End Sub


