Attribute VB_Name = "Shortcuts"
' VBA Module name: Shortcuts.bas
' https://github.com/mslonik/Microsoft-Word-Configuration
'
'   License: MIT License.
'
'
'+----+------------------------------------+--------------+----------------+------------------------------------+
'| No | Sub name                           | Ribbon name  | Ribbon section | Ribbon button name                 |
'+----+------------------------------------+--------------+----------------+------------------------------------+
'| 1  | ShowFormHotstrings                 | Shortcuts_ms | Show           | ShowFormHotstrings                 |
'| 2  | ShowFormHotkeys                    | Shortcuts_ms | Show           | ShowFormHotkeys                    |
'| 3  | ShowFormHotMacros                  | Shortcuts_ms | Show           | ShowFormHotMacros                  |
'| 4  | ClearActiveDocumentStyleShortcuts  | Shortcuts_ms | Clear / Create | ClearActiveDocumentStyleShortcuts  |
'| 5  | ClearActiveDocumentMacroShortcuts  | Shortcuts_ms | Clear / Create | ClearActiveDocumentMacroShortcuts  |
'| 6  | RemoveActiveDocumentMacroShortcuts | Shortcuts_ms | Clear / Create | RemoveActiveDocumentMacroShortcuts |
'| 7  | CreateActiveDocumentMacroShortcuts | Shortcuts_ms | Clear / Create | CreateActiveDocumentMacroShortcuts |
'| 8  | ListAllShortcutsToTxt              | Shortcuts_ms | List           | ListAllShortcutsToTxt              |
'| 9  | ListHotkeysToTxt                   | Shortcuts_ms | List           | ListHotkeysToTxt                   |
'| 10 | ListHotstringsToTxt                | Shortcuts_ms | List           | ListHotstringsToTxt                |
'| 11 | ListHotMacrosToTxt                 | Shortcuts_ms | List           | ListHotMacrosToTxt                 |
'| 12 | ListMWShortcutsToDOCX              | Shortcuts_ms | List           | ListMWShortcutsToDOCX              |
'| 13 | ListAllMWCommandsToDOCX            | Shortcuts_ms | List           | ListAllMWCommandsToDOCX            |
'+----+------------------------------------+--------------+----------------+------------------------------------+
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'   13. Set_CommandShortcuts()                      -> Reset_CommandShortcut()
'   17. SetShortcut_CustomizedToggleFieldCodes()    -> ResetShortcut_CustomizedToggleFieldCodes()
'   19. SetShortcut_JumpToNextList()                -> ResetShortcut_JumpToNextList()
'   21. SetShortcut_JumpToNextTable()               -> ResetShortcut_JumpToNextTable()
'   23. SetShortcut_JumpToNextCanvas()              -> ResetShortcut_JumpToNextCanvas()
'   25. SetShortcut_SaveFileApplyNumberingDistance() -> ResetShortcut_SaveFileApplyNumberingDistance()
'   27. SetShortcut_CloseFileApplyUpdateFields()    -> ResetShortcut_CloseFileApplyUpdateFields()
'   29. SetShortcut_CustomizedPrintPreviewAndPrint() -> ResetShortcut_CustomizedPrintPreviewAndPrint()
'   31. SetShortcut_ToggleCharBoldStyle()           -> ResetShortcut_ToggleCharBoldStyle()
'   33. SetShortcut_ToggleCharItalicStyle()         -> ResetShortcut_ToggleCharItalicStyle()
'   35. SetShortcut_ToggleCharUnderlineStyle()      -> ResetShortcut_ToggleCharUnderlineStyle()
'   37. SetShortcut_ToggleCharCrossoutStyle()       -> ResetShortcut_ToggleCharCrossoutStyle()
'   39. SetShortcut_ToggleCharHiddenStyle()         -> ResetShortcut_ToggleCharHiddenStyle()
'   41. SetShortcut_ToggleSpecificFormatting()      -> ResetShortcut_ToggleSpecificFormatting()
'   43. SetShortcut_ToggleCharSourceCode()          -> ResetShortcut_ToggleCharSourceCode()
'   45. SetShortcut_SetLanguageToEnglishUS()        -> ResetShortcut_SetLanguageToEnglishUS()
'   47. SetShortcut_ToggleApplyStyles()             -> ResetShortcut_ToggleApplyStyles()
'   49. SetShortcut_ToggleOvertypeMode()            -> ResetShortcut_ToggleOvertypeMode()
'   51. SetShortcut_ReapplyTemplateStyle()          -> ResetShortcut_ReapplyTemplateStyle()
'   53. SetShortcut_RestartListNumbering()          -> ResetShortcut_RestartListNumbering()
'   55. SetShortcut_HotMacros()                     -> ResetShortcut_HotMacros()
'   57. SetShortcut_HotStrings()                    -> ResetShortcut_Strings()
'   59. SetShortcut_HotKeys()                       -> ResetShortcut_HotKeys()
'   61. SetShortcut_ToggleHeadingCollapseExpand()   -> ResetShortcut_ToggleHeadingCollapseExpand()
'   63. SetShortcut_PrintDocument()                 -> ResetShortcut_PrintDocument()
'   65. SetShortcut_CustomizedSaveAs()              -> ResetShortcut_CustomizedSaveAs()
'   67. SetShortcut_CustomizedCopyFormat()          -> ResetShortcut_CustomizedCopyFormat()
'   69. SetShortcut_CustomizedPasteFormat()         -> ResetShortcut_CustomizedPasteFormat()
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
' Used to enforce the explicit declaration of all variables in a module. When you include Option Explicit at the beginning of a module, it ensures that you must
' declare all variables using the Dim, Private, Public, ReDim, or Static statements before using them. This helps prevent errors caused by typos or
' undeclared variables.
Option Explicit

' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
' Hots_UniversalForm: Height = 450, Width = 410

Public frmHotMacros As Hots_UniversalForm   ' Type class: Hots_UniversalForm
Public frmHotkey As Hots_UniversalForm      ' Type: Hots_UniversalForm. This is the specific "Forms" form named Hots_UniversalForms.
Public frmHotstring As Hots_UniversalForm   ' Type: Hots_UniversalForm.

' The following functions enable minimize button to Form windows
Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" _
    (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
    (ByVal hWnd As Long, ByVal nIndex As Long) As Long

Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
    (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long

Const GWL_STYLE = -16
Const WS_MINIMIZEBOX = &H20000
Const WS_SYSMENU = &H80000

' 2025-04-21 by ms, reworked
Sub CreateActiveDocumentMacroShortcuts()
    ' Set the customization context to the current document
    Application.CustomizationContext = ActiveDocument

    ' Pripare all the shortcuts:
    Call Set_CommandShortcuts(IfMsgBox:="quiet")                       ' module: Shortcuts
    
    Call SetShortcut_CustomizedToggleFieldCodes(IfMsgBox:="quiet")     ' module: Shortcuts: 1
    Call SetShortcut_JumpToNextList(IfMsgBox:="quiet")                 ' module: Shortcuts: 2
    Call SetShortcut_JumpToNextTable(IfMsgBox:="quiet")                ' module: Shortcuts: 3
    Call SetShortcut_JumpToNextCanvas(IfMsgBox:="quiet")               ' module: Shortcuts: 4
    Call SetShortcut_SaveFileApplyNumberingDistance(IfMsgBox:="quiet") ' module: Shortcuts: 5
    Call SetShortcut_CloseFileApplyUpdateFields(IfMsgBox:="quiet")     ' module: Shortcuts: 6
    Call SetShortcut_CustomizedPrintPreviewAndPrint(IfMsgBox:="quiet") ' module: Shortcuts: 7
    Call SetShortcut_ToggleCharBoldStyle(IfMsgBox:="quiet")            ' module: Shortcuts: 8
    Call SetShortcut_ToggleCharItalicStyle(IfMsgBox:="quiet")          ' module: Shortcuts: 9
    Call SetShortcut_ToggleCharCrossoutStyle(IfMsgBox:="quiet")        ' module: Shortcuts: 10
    Call SetShortcut_ToggleCharUnderlineStyle(IfMsgBox:="quiet")       ' module: Shortcuts: 11
    Call SetShortcut_ToggleCharHiddenStyle(IfMsgBox:="quiet")          ' module: Shortcuts: 12
    Call SetShortcut_ToggleSpecificFormatting(IfMsgBox:="quiet")       ' module: Shortcuts: 13
    Call SetShortcut_ToggleCharSourceCode(IfMsgBox:="quiet")           ' module: Shortcuts: 14
    Call SetShortcut_SetLanguageToEnglishUS(IfMsgBox:="quiet")         ' module: Shortcuts: 15
    Call SetShortcut_ToggleApplyStyles(IfMsgBox:="quiet")              ' module: Shortcuts: 16
    Call SetShortcut_ToggleOvertypeMode(IfMsgBox:="quiet")             ' module: Shortcuts: 17
    Call SetShortcut_ReapplyTemplateStyle(IfMsgBox:="quiet")           ' module: Shortcuts: 18
    Call SetShortcut_RestartListNumbering(IfMsgBox:="quiet")           ' module: Shortcuts: 19
    Call SetShortcut_InsertCrossReference(IfMsgBox:="quiet")           ' module: Shortcuts: 20
    Call SetShortcut_ToggleHeadingCollapseExpand(IfMsgBox:="quiet")    ' module: Shortcuts: 21
    Call SetShortcut_PrintDocument(IfMsgBox:="quiet")                  ' module: Shortcuts: 22
    Call SetShortcut_CustomizedCopyFormat(IfMsgBox:="quiet")           ' module: Shortcuts: 23
    Call SetShortcut_CustomizedPasteFormat(IfMsgBox:="quiet")          ' module: Shortcuts: 24
    Call SetShortcut_CustomizedSaveAs(IfMsgBox:="quiet")               ' module: Shortcuts: 25

    Call SetShortcut_HotMacros(IfMsgBox:="quiet")                      ' module: Shortcuts: 26
    Call SetShortcut_HotStrings(IfMsgBox:="quiet")                     ' module: Shortcuts: 27
    Call SetShortcut_HotKeys(IfMsgBox:="quiet")                        ' module: Shortcuts: 28
    
    Dim FileName As String:     FileName = C_F_Macros
    Dim ModuleName As String:   ModuleName = C_M_Template
    Dim MacroName As String:    MacroName = "CreateActiveDocumentMacroShortcuts"
    Dim MsgBoxTitle As String:  MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    Dim userResponse As VbMsgBoxResult
    Beep
    userResponse = MsgBox( _
        Prompt:="Would you like to display all the newly defined keyboard shortcuts?" & vbNewLine & vbNewLine, _
        Buttons:=vbYesNo + vbDefaultButton1 + vbQuestion, _
        Title:=MsgBoxTitle _
        )
    
    If userResponse = vbYes Then
        ' Show all the shortcuts Forms
        Call ShowFormHotstrings                             ' module: Shortcuts
        Call ShowFormHotkeys                                ' module: Shortcuts
        Call ShowFormHotMacros                              ' module: Shortcuts
    End If
End Sub

' 2025-04-21 by ms, created
Sub RemoveActiveDocumentMacroShortcuts()
    Dim UserDecision As VbMsgBoxResult
    
    Dim FileName As String:     FileName = C_F_Macros
    Dim ModuleName As String:   ModuleName = C_M_Shortcuts
    Dim MacroName As String:    MacroName = "RemoveActiveDocumentMacroShortcuts"
    Dim MsgBoxTitle As String:  MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName

    Beep
    UserDecision = MsgBox( _
        Prompt:="This macro will remove all shortcuts from currently active document." & vbNewLine & vbNewLine & _
            "Are you sure?", _
        Buttons:=vbQuestion + vbYesNo, _
        Title:=MsgBoxTitle _
        )
    If UserDecision = vbNo Then
        Exit Sub
    End If

    Call Reset_CommandShortcut                       ' module: Shortcuts

    Call ResetShortcut_CustomizedToggleFieldCodes
    Call ResetShortcut_JumpToNextList
    Call ResetShortcut_JumpToNextTable
    Call ResetShortcut_JumpToNextCanvas
    Call ResetShortcut_SaveFileApplyNumberingDistance
    Call ResetShortcut_CloseFileApplyUpdateFields
    Call ResetShortcut_CustomizedPrintPreviewAndPrint
    Call ResetShortcut_ToggleCharBoldStyle
    Call ResetShortcut_ToggleCharItalicStyle
    Call ResetShortcut_ToggleCharCrossoutStyle
    Call ResetShortcut_ToggleCharUnderlineStyle
    Call ResetShortcut_ToggleCharHiddenStyle
    Call ResetShortcut_ToggleSpecificFormatting
    Call ResetShortcut_ToggleCharSourceCode
    Call ResetShortcut_SetLanguageToEnglishUS
    Call ResetShortcut_ToggleApplyStyles
    Call ResetShortcut_ToggleOvertypeMode
    Call ResetShortcut_ReapplyTemplateStyle
    Call ResetShortcut_RestartListNumbering
    Call ResetShortcut_InsertCrossReference
    Call ResetShortcut_ToggleHeadingCollapseExpand
    Call ResetShortcut_PrintDocument
    Call ResetShortcut_CustomizedCopyFormat
    Call ResetShortcut_CustomizedPasteFormat
    Call ResetShortcut_CustomizedSaveAs
    
    Call ResetShortcut_HotMacros
    Call ResetShortcut_Strings
    Call ResetShortcut_HotKeys
    
    ' Destroy all the shortcuts Forms
    Call DestroyFormHotstrings
    Call DestroyFormHotkeys
    Call DestroyFormHotMacros
End Sub


Private Sub HotMacrosObject_Initialize()
    ' The New modifier in VBA (Visual Basic for Applications) is used to create a new instance of an object. When you declare an object variable with the New keyword, VBA automatically creates a new instance of the object when the variable is first used.
    Set frmHotMacros = New Hots_UniversalForm  ' initialization of variable frmHotstring as an new instance / object of type Hots_UniversalForm
    frmHotMacros.InstanceName = "HotMacros"
End Sub

Private Sub HotkeyObject_Initialize()
    ' The New modifier in VBA (Visual Basic for Applications) is used to create a new instance of an object. When you declare an object variable with the New keyword, VBA automatically creates a new instance of the object when the variable is first used.
    Set frmHotkey = New Hots_UniversalForm  ' initialization of variable frmHotstring as an new instance / object of type Hots_UniversalForm
    frmHotkey.InstanceName = "Hotkey"
End Sub

Private Sub HotstringsObject_Initialize()
    ' The New modifier in VBA (Visual Basic for Applications) is used to create a new instance of an object. When you declare an object variable with the New keyword, VBA automatically creates a new instance of the object when the variable is first used.
    Set frmHotstring = New Hots_UniversalForm   ' initialization of variable frmHotstring as an new instance / object of type Hots_UniversalForm
    frmHotstring.InstanceName = "Hotstring"
End Sub

Sub ShowFormHotMacros()
    ' Check if the frmHotstrings object is initialized
    If frmHotMacros Is Nothing Then
        Call HotMacrosObject_Initialize
        Call HotMacrosUserForm_Initialize(HotkeyHotkey:=C_SC_AltHplusM)
        
        ' The following code enable minimize button in the Form window
        Dim hWnd As Long
        hWnd = FindWindow("ThunderDFrame", frmHotMacros.Caption)
        If hWnd <> 0 Then
            Dim lStyle As Long
            lStyle = GetWindowLong(hWnd, GWL_STYLE)
            lStyle = lStyle Or WS_SYSMENU Or WS_MINIMIZEBOX
            SetWindowLong hWnd, GWL_STYLE, lStyle
            DrawMenuBar hWnd
        End If
    End If

    ' Check if the form is already displayed
    If frmHotMacros.Visible Then
        ' If the form is visible, hide it
        frmHotMacros.Hide
    Else
        ' vbModeless enable display of the subwindow without closing
        frmHotMacros.Show vbModeless
        ' Unfortunately it is not possible to get back focus (cursor) to main body of the document. User has to use a mouse of keyboard to switch back to document body.
    End If

End Sub

' 2025-04-21 by ms
Sub DestroyFormHotMacros()
    If Not frmHotMacros Is Nothing Then
        Unload frmHotMacros
        Set frmHotMacros = Nothing
    End If
End Sub

Sub ShowFormHotstrings()
    ' Check if the frmHotstrings object is initialized
    If frmHotstring Is Nothing Then
        Call HotstringsObject_Initialize
        Call HotstringUserForm_Initialize(HotkeyHotkey:=C_SC_AltHplusS)
        
        ' The following code enable minimize button in the Form window
        Dim hWnd As Long
        hWnd = FindWindow("ThunderDFrame", frmHotstring.Caption)
        If hWnd <> 0 Then
            Dim lStyle As Long
            lStyle = GetWindowLong(hWnd, GWL_STYLE)
            lStyle = lStyle Or WS_SYSMENU Or WS_MINIMIZEBOX
            SetWindowLong hWnd, GWL_STYLE, lStyle
            DrawMenuBar hWnd
        End If
    End If

    ' Check if the form is already displayed
    If frmHotstring.Visible Then
        ' If the form is visible, hide it
        frmHotstring.Hide
    Else
        ' vbModeless enable display of the subwindow without closing
        frmHotstring.Show vbModeless
        ' Unfortunately it is not possible to get back focus (cursor) to main body of the document. User has to use a mouse of keyboard to switch back to document body.
    End If
End Sub

' 2025-04-21 by ms
Sub DestroyFormHotstrings()
    If Not frmHotstring Is Nothing Then
        Unload frmHotstring
        Set frmHotstring = Nothing
    End If
End Sub

Sub ShowFormHotkeys()
    ' Check if the frmHotkey object is initialized
    If frmHotkey Is Nothing Then
        ' If the form is not visible, initialize and show it
        Call HotkeyObject_Initialize
        Call HotkeyUserForm_Initialize(HotkeyHotkey:=C_SC_AltHplusK)
        
        ' The following code enable minimize button in the Form window
        Dim hWnd As Long
        hWnd = FindWindow("ThunderDFrame", frmHotkey.Caption)
        If hWnd <> 0 Then
            Dim lStyle As Long
            lStyle = GetWindowLong(hWnd, GWL_STYLE)
            lStyle = lStyle Or WS_SYSMENU Or WS_MINIMIZEBOX
            SetWindowLong hWnd, GWL_STYLE, lStyle
            DrawMenuBar hWnd
        End If
    End If
    
    ' Check if the form is already displayed
    If frmHotkey.Visible Then
        ' If the form is visible, hide it
        frmHotkey.Hide
    Else
        frmHotkey.Show vbModeless
        ' Unfortunately it is not possible to get back focus (cursor) to main body of the document. User has to use a mouse of keyboard to switch back to document body.
    End If
End Sub

' 2025-04-21 by ms
Sub DestroyFormHotkeys()
    If Not frmHotkey Is Nothing Then
        Unload frmHotkey
        Set frmHotkey = Nothing
    End If
End Sub


Private Function HotstringUserForm_Initialize(HotkeyHotkey As String) As Boolean
    ' true = error, false = no error
    ' Define the arrays to be displayed in the ListBox
    Dim MyBBName() As String
    Dim MyBBDescription() As String
    Dim i As Integer
    Dim tempName As String
    Dim tempDescription As String
    Dim j As Integer
    
    Dim FileName As String
    FileName = C_F_Macros
    
    Dim ModuleName As String
    ModuleName = C_M_Shortcuts
    
    Dim MacroName As String
    MacroName = "HotstringUserForm_Initialize"
    
    Dim MsgBoxTitle As String
    MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    ' Set the Forms properties
    frmHotstring.Caption = "Template keyboard shortcuts (HotStrings): " & HotkeyHotkey
    frmHotstring.font = "Consolas"
    
    ' Get the building blocks and their descriptions
    If GetBBKeyBindings(MyBBName, MyBBDescription, MsgBoxTitle) Then
        HotstringUserForm_Initialize = True
        Exit Function
    End If

    ' Sort MyBBName alphabetically and adjust MyBBDescription accordingly
    For i = LBound(MyBBName) To UBound(MyBBName) - 1
        For j = i + 1 To UBound(MyBBName)
            If MyBBName(i) > MyBBName(j) Then
                ' Swap MyBBName
                tempName = MyBBName(i)
                MyBBName(i) = MyBBName(j)
                MyBBName(j) = tempName
                
                ' Swap MyBBDescription
                tempDescription = MyBBDescription(i)
                MyBBDescription(i) = MyBBDescription(j)
                MyBBDescription(j) = tempDescription
            End If
        Next j
    Next i

    ' Add a listbox
    Dim ListBox1 As MSForms.ListBox
    Set ListBox1 = frmHotstring.Controls.Add("Forms.ListBox.1", "ListBox1", True)
    ListBox1.Width = 335
    ListBox1.Height = 390
    ListBox1.font = "Consolas"
    
    ' Add a label
    Dim Label1 As MSForms.Label
    Set Label1 = frmHotstring.Controls.Add("Forms.Label.1", "Label1", True)
    Label1.Left = 6
    Label1.Top = 395
    Label1.Caption = "After typing a hotstring press F3 to Enter it!"
    Label1.ForeColor = &HFF&    ' red
    Label1.font = "Consolas"
    Label1.Height = 18
    Label1.Width = 300

    ' Set the ColumnCount property
    ListBox1.ColumnCount = 2
    frmHotstring.Caption = "Template keyboard strings (HotStrings): " & HotkeyHotkey
    
    ' Populate the ListBox with sorted MyBBName and adjusted MyBBDescription
    For i = LBound(MyBBName) To UBound(MyBBName)
        ListBox1.AddItem
        ListBox1.List(i, 0) = MyBBName(i)
        ListBox1.List(i, 1) = MyBBDescription(i)
    Next i
    HotstringUserForm_Initialize = False ' no error
End Function

' 2026-01-15 by ms
Public Function ReturnBuildingBlockEntries() As BuildingBlockEntries
    Dim i As Integer
    Dim AddInsName As String
    Dim AddInsIndex As Integer
    Dim UserDecision As VbMsgBoxResult
    Dim TemplateIndex As Integer

    Dim FileName As String:     FileName = C_F_Macros
    Dim ModuleName As String:   ModuleName = C_M_Shortcuts
    Dim MacroName As String:    MacroName = "ReturnBuildingBlockEntries"
    Dim MsgBoxTitle As String:  MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    AddInsIndex = 0
    ' At first try to set bbe to ActiveDocument.AttachedTemplate.BuildingBlockEntries (template with integrated BuildingBlocks)
    Set ReturnBuildingBlockEntries = ActiveDocument.AttachedTemplate.BuildingBlockEntries

    ' Check if bbe is empty
    If ReturnBuildingBlockEntries.count = 0 Then
        ' Force loading of all BuildingBlocks
        Templates.LoadBuildingBlocks
        ' Instead of referencing the "C_F_BuildingBlocks" I need to reference through number.
        For i = 1 To AddIns.count
            If AddIns(i).Name = C_F_BuildingBlocks Then
                AddInsIndex = i
                AddInsName = AddIns(AddInsIndex).Name
                Exit For
            End If
        Next i
        
        If AddInsIndex = 0 Then
            MsgBox _
                Prompt:="The " & C_F_BuildingBlocks & " was not found." & vbNewLine & vbNewLine & _
                    "Exiting.", _
                Buttons:=vbCritical, _
                Title:=MsgBoxTitle
                Exit Function
        End If
        ' Ask user if to enable BuildingBlocks template "C_F_BuildingBlocks"
        If Not AddIns(AddInsIndex).Installed Then
            Beep
            UserDecision = MsgBox( _
                Prompt:="The " & C_F_BuildingBlocks & " is found, but not enabled." & vbNewLine & _
                    "Would you like to enable it now?", _
                Buttons:=vbYesNo + vbQuestion, _
                Title:=MsgBoxTitle _
                )
            If UserDecision = vbYes Then
                AddIns(AddInsIndex).Installed = True
            Else
                Exit Function
            End If
        End If
        
        ' If it doesn't work try to set bbe to AddIns(C_F_BuildingBlocks).InstalledTemplate.BuildingBlockEntries (template without integrated BuildingBlocks)
        For i = 1 To Templates.count
            If Templates(i).Name = AddInsName Then
                TemplateIndex = i
                Exit For
            End If
        Next i
        
        Set ReturnBuildingBlockEntries = Templates(TemplateIndex).BuildingBlockEntries
    
        ' Check if bbe is still empty
        If ReturnBuildingBlockEntries.count = 0 Then
            MsgBox _
                Prompt:="Warning!" & vbNewLine & "No building block entries found in either location.", _
                Buttons:=vbExclamation, _
                Title:=MsgBoxTitle
            Exit Function
        End If
    End If

End Function


' 2025-03-07 by ms
Private Function GetBBKeyBindings(ByRef MyBBName() As String, _
                                    ByRef MyBBDescription() As String, _
                                    Optional MsgBoxHeader As String) _
                                    As Boolean
    ' error = true; no error = false
    Dim bb As BuildingBlock
    Dim bbe As BuildingBlockEntries
    Dim i As Integer
    Dim DataIndex As Integer
    Dim NoOfBuildingBlocks As Long
    
    Set bbe = ReturnBuildingBlockEntries()
    NoOfBuildingBlocks = bbe.count

    ' Initialize the data array with an initial size
    ReDim MyBBName(0 To NoOfBuildingBlocks)          ' Initial size, will resize later
    ReDim MyBBDescription(0 To NoOfBuildingBlocks)   ' Initial size, will resize later
    DataIndex = 0
    
    For i = 1 To NoOfBuildingBlocks
        Set bb = bbe(i)
        If bb.Type.Name = "AutoText" Then
            MyBBName(DataIndex) = bb.Name
            MyBBDescription(DataIndex) = bb.Description
            DataIndex = DataIndex + 1
        End If
    Next i
        
    ' Resize the data array to the actual number of items
    ReDim Preserve MyBBName(0 To DataIndex - 1)
    ReDim Preserve MyBBDescription(0 To DataIndex - 1)
    GetBBKeyBindings = False
    
    ' Clear object variables
    Set bbe = Nothing
End Function

Private Sub HotMacrosUserForm_Initialize(HotkeyHotkey As String)
    ' Define the arrays to be displayed in the ListBox
    Dim MyStyleName() As String
    Dim MyShortcut() As String
    
    ' Get the macros and their key bindings
    Call GetMacrosKeyBindings(MyStyleName, MyShortcut)
    
    ' Set the Forms properties
    frmHotMacros.Caption = "Styles macro shortcuts (HotKeys): " & HotkeyHotkey
    frmHotMacros.font = "Consolas"
    
    ' Add a listbox
    Dim ListBox1 As MSForms.ListBox
    Set ListBox1 = frmHotMacros.Controls.Add("Forms.ListBox.1", "ListBox1", True)
    ListBox1.Width = 400
    ListBox1.Height = 315

    ' Set the ColumnCount property
    ListBox1.ColumnCount = 2
    
    ' Set the ColumnWidths property
    ListBox1.ColumnWidths = "300,95" ' Set the width of the first column to 200 points and the second column to 135 points
    
    ' Populate the ListBox
    Dim i As Integer
    For i = LBound(MyStyleName) To UBound(MyStyleName)
        ListBox1.AddItem
        ListBox1.List(i, 0) = MyStyleName(i)
        ListBox1.List(i, 1) = MyShortcut(i)
    Next i
End Sub

' Prepare arrays containing macro name and macro shortcut.
' 2025-08-02 by ms
' 2025-12-31 by ms
Private Sub GetMacrosKeyBindings(ByRef MyStyleName, ByRef MyShortcut)
    Dim kb As keyBinding
    Dim shortcutText As String
    
    Dim FileName As String:      FileName = C_F_Macros
    Dim ModuleName As String:    ModuleName = C_M_Shortcuts
    Dim MacroName As String:     MacroName = "GetMacrosKeyBindings"
    Dim MsgBoxTitle As String:   MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
   
    CustomizationContext = ActiveDocument
   
    ' Initialize the data array with an initial size
    ReDim MyStyleName(0 To Application.KeyBindings.count - 1)
    ReDim MyShortcut(0 To Application.KeyBindings.count - 1)
   
    ' Initialize the shortcut text
    shortcutText = "User Defined Shortcuts:" & vbCrLf & vbCrLf
        
    ' wdKeyCategoryCommand: This category includes built-in Word commands. Checked, it works.
    ' wdKeyCategoryMacro:   This category includes macros that you have created. Checked, it works.
    ' wdKeyCategoryFont:    This category includes font-related commands. Didn't check.
    ' wdKeyCategoryAutoText: This category includes AutoText entries. Seems not to work.
    ' wdKeyCategoryStyle:   This category includes styles. Checked, it works.
    ' wdKeyCategorySymbol:  This category includes symbols. Didn't check.
    ' wdKeyCategoryPrefix:  This category includes prefix keys. Seems not to work.
    
    ' Loop through all key bindings
    Dim DataIndex As Integer
    DataIndex = 0
    For Each kb In Application.KeyBindings
        ' Check if the key binding belongs to the command category
        If (kb.KeyCategory = wdKeyCategoryCommand) And (kb.Context.Name = CustomizationContext) Then
            MyStyleName(DataIndex) = kb.Command
            MyShortcut(DataIndex) = kb.KeyString
            ' Add the command name and its shortcut to the text
            shortcutText = shortcutText & kb.Command & ": " & kb.KeyString & vbCrLf
            DataIndex = DataIndex + 1
        End If
    Next kb
    
    shortcutText = shortcutText & vbNewLine & "Macros keyboard shortcuts" & vbNewLine & vbNewLine
    
    Dim TextResult As String
    Dim LastDotPos As Long
    
    ' Loop through all key bindings
    For Each kb In Application.KeyBindings
        ' Check if the key binding belongs to the macro category
        ' Shorten macro name, e.g. from default ame TemplateProject.Tools.CustomizedOvertype make just CustomizedOvertype.
        If (kb.KeyCategory = wdKeyCategoryMacro) And (kb.Context.Name = CustomizationContext) Then
            LastDotPos = InStrRev(kb.Command, ".")
            If LastDotPos > 0 Then
                TextResult = Mid(kb.Command, LastDotPos + 1)
            Else
                TextResult = kb.Command
            End If
            
            MyStyleName(DataIndex) = TextResult
            ' MyStyleName(DataIndex) = kb.Command
            MyShortcut(DataIndex) = kb.KeyString
            ' Add the command name and its shortcut to the text
            shortcutText = shortcutText & kb.Command & ": " & kb.KeyString & vbCrLf
            DataIndex = DataIndex + 1
        End If
    Next kb
    
    ' Resize the data array to the actual number of items
    If DataIndex > 0 Then
        ReDim Preserve MyStyleName(0 To DataIndex - 1)
        ReDim Preserve MyShortcut(0 To DataIndex - 1)
    Else
        ReDim MyStyleName(0 To 0)
        ReDim MyShortcut(0 To 0)
    End If

End Sub

Private Sub HotkeyUserForm_Initialize(HotkeyHotkey As String)
    ' Define the arrays to be displayed in the ListBox
    Dim MyStyleName() As String
    Dim MyShortcut() As String
    
    ' Get the styles and their key bindings
    Call GetStylesAndKeyBindings(MyStyleName, MyShortcut)
    
    Dim DocumentName As String
    DocumentName = ActiveDocument.Name
    ' Set the Forms properties
    frmHotkey.Caption = DocumentName & " styles keyboard shortcuts (HotKeys): " & HotkeyHotkey
    frmHotkey.font = "Consolas"
    
    ' Add a listbox
    Dim ListBox1 As MSForms.ListBox
    Set ListBox1 = frmHotkey.Controls.Add("Forms.ListBox.1", "ListBox1", True)
    ListBox1.Width = 335
    ListBox1.Height = 450

    ' Set the ColumnCount property
    ListBox1.ColumnCount = 2
    
    ' Populate the ListBox
    Dim i As Integer
    For i = LBound(MyStyleName) To UBound(MyStyleName)
        ListBox1.AddItem
        ListBox1.List(i, 0) = MyStyleName(i)
        ListBox1.List(i, 1) = MyShortcut(i)
    Next i
End Sub

' It is not possible to get access to AttachedTemplate styles, as styles are property of the file. As long as the (attached) template file isn't opened, but just attached, it is not possible to get access to its Styles property.
' Because of that, the GetStylesAndKeyBindings must process the ActiveDocument.Styles instead of ActiveDocument.AttachedTemplate.Styles.
' 2025-11-15 by ms
Private Sub GetStylesAndKeyBindings(ByRef MyStyleName() As String, ByRef MyShortcut() As String)
    Dim i As Integer
    Dim DataIndex As Integer
    Dim DocumentStyles As Word.Styles
    
    ' Get the styles from the currently opened document
    Set DocumentStyles = ActiveDocument.Styles
    
    ' Initialize the data array with an initial size
    ReDim MyStyleName(0 To DocumentStyles.count - 1)
    ReDim MyShortcut(0 To DocumentStyles.count - 1)
    DataIndex = 0
    
    ' Iterate through each style and get its name and key binding
    For i = 1 To DocumentStyles.count
        Dim keyBinding As String
        keyBinding = GetKeyBinding(DocumentStyles(i).NameLocal)
        
        If keyBinding <> "" Then
            MyStyleName(DataIndex) = DocumentStyles(i).NameLocal
            MyShortcut(DataIndex) = keyBinding
            DataIndex = DataIndex + 1
        End If
    Next i
    
    ' Resize the data array to the actual number of items
    If DataIndex > 0 Then
        ReDim Preserve MyStyleName(0 To DataIndex - 1)
        ReDim Preserve MyShortcut(0 To DataIndex - 1)
    Else
        ReDim MyStyleName(0 To 0)
        ReDim MyShortcut(0 To 0)
    End If
    
    ' Clear object variables
    Set DocumentStyles = Nothing
End Sub
Private Function GetKeyBinding(styleName As String) As String
    Dim Shortcut As String
    Dim key As keyBinding
    
    CustomizationContext = ActiveDocument
    
    ' Check if the style has a key binding
    For Each key In KeyBindings
        If key.Command = styleName Then
            Shortcut = key.KeyString
            Exit For
        End If
    Next key
    
    GetKeyBinding = Shortcut
End Function

' 2025-07-27 by ms
' Because a 'Form' window cannot be easily copied to TXT format.
Sub ListHotkeysToTxt()
    Dim MyStyleName() As String
    Dim MyShortcut() As String
    Dim i As Integer
    Dim FilePath As String
    Dim filenum As Integer
    
    Dim FileName As String
    FileName = C_F_Macros
    
    Dim ModuleName As String
    ModuleName = C_M_Shortcuts
    
    Dim MacroName As String
    MacroName = "ListHotkeysToTxt"
    
    Dim MsgBoxTitle As String
    MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    ' Get the styles and their key bindings
    Call GetStylesAndKeyBindings(MyStyleName, MyShortcut)
    
    ' Set the file path to the default file location
    FilePath = Options.DefaultFilePath(wdDocumentsPath) & "\" & MacroName & ".txt"
    
    ' Open the file for writing
    filenum = FreeFile
    Open FilePath For Output As #filenum
    
    Dim CurrentDate As String
    CurrentDate = Format(Date, "yyyy-mm-dd")
    Dim CurrentTime As String
    CurrentTime = Format(Time, "hh:mm:ss")
    Dim TemplateName As String
    TemplateName = ActiveDocument.AttachedTemplate.Name
    
    ' Add Add Add log file header information
    Print #filenum, CurrentDate & " " & CurrentTime
    Print #filenum, "Template name: " & TemplateName
    Print #filenum, "Template name: " & TemplateName
    Print #filenum, "Template name: " & TemplateName
    Print #filenum, "This file contains hotkeys to styles defined within this template."
    Print #filenum, "It was made by the Microsoft Word macro."
    Print #filenum, "Filename: " & FileName
    Print #filenum, "Module name: " & ModuleName
    Print #filenum, "Macro name: " & MacroName
    Print #filenum, vbCrLf
    
    ' Write the hotkeys to the file
    For i = LBound(MyStyleName) To UBound(MyStyleName)
        Print #filenum, i + 1 & " | " & MyStyleName(i) & " | " & MyShortcut(i) & " |"
    Next i
        
    ' Close the file
    Close #filenum
        
    MsgBox _
        Prompt:="Hotkeys saved to " & vbNewLine & FilePath, _
        Buttons:=vbInformation + vbOKOnly, _
        Title:=MsgBoxTitle
End Sub

' 2025-07-27 by ms
' Because a 'Form' window cannot be easily copied to TXT format.
Sub ListHotstringsToTxt()
    Dim MyBBName() As String
    Dim MyBBDescription() As String
    Dim i As Integer
    Dim FilePath As String
    Dim filenum As Integer
    
    Dim FileName As String
    FileName = C_F_Macros
    
    Dim ModuleName As String
    ModuleName = C_M_Shortcuts
    
    Dim MacroName As String
    MacroName = "ListHotstringsToTxt"
    
    Dim MsgBoxTitle As String
    MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    ' Get the building blocks and their descriptions
    Call GetBBKeyBindings(MyBBName, MyBBDescription, MsgBoxTitle)
    
    ' Set the file path to the default file location
    FilePath = Options.DefaultFilePath(wdDocumentsPath) & "\" & MacroName & ".txt"
    
    ' Open the file for writing
    filenum = FreeFile
    Open FilePath For Output As #filenum
    
    Dim CurrentDate As String
    CurrentDate = Format(Date, "yyyy-mm-dd")
    Dim CurrentTime As String
    CurrentTime = Format(Time, "hh:mm:ss")
    Dim TemplateName As String
    TemplateName = ActiveDocument.AttachedTemplate.Name
        
    ' Add Add log file header information
    Print #filenum, CurrentDate & " " & CurrentTime
    Print #filenum, "Template name: " & TemplateName
    Print #filenum, "Template name: " & TemplateName
    Print #filenum, "This file contains hotstrings to building blocks defined within this template."
    Print #filenum, "It was made by the Microsoft Word macro."
    Print #filenum, "Filename: " & FileName
    Print #filenum, "Module name: " & ModuleName
    Print #filenum, "Macro name: " & MacroName
    Print #filenum, vbCrLf
    
    ' Write the hotstrings to the file
    For i = LBound(MyBBName) To UBound(MyBBName)
        Print #filenum, i + 1 & " | " & MyBBName(i) & " | " & MyBBDescription(i) & " |"
    Next i
    
    ' Close the file
    Close #filenum
    
    MsgBox _
        Prompt:="Hotstrings definitions have been saved to " & vbNewLine & FilePath, _
        Buttons:=vbInformation + vbOKOnly, _
        Title:=MsgBoxTitle
End Sub

' 2025-07-26 by ms
Sub ListHotMacrosToTxt()
    Dim MyStyleName() As String
    Dim MyShortcut() As String
    Dim FilePath As String
    Dim filenum As Integer
    
    Dim FileName As String
    FileName = C_F_Macros
    
    Dim ModuleName As String
    ModuleName = C_M_Shortcuts
    
    Dim MacroName As String
    MacroName = "ListHotMacrosToTxt"
    
    Dim MsgBoxTitle As String
    MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    ' Get the macros and their key bindings
    Call GetMacrosKeyBindings(MyStyleName, MyShortcut)
    
    ' Set the file path to the default file location
    FilePath = Options.DefaultFilePath(wdDocumentsPath) & "\" & MacroName & ".txt"
    
    ' Open the file for writing
    filenum = FreeFile
    Open FilePath For Output As #filenum
    
    Dim CurrentDate As String
    CurrentDate = Format(Date, "yyyy-mm-dd")
    Dim CurrentTime As String
    CurrentTime = Format(Time, "hh:mm:ss")
    Dim TemplateName As String
    TemplateName = ActiveDocument.AttachedTemplate.Name
    
    ' Add log file header information
    Print #filenum, CurrentDate & " " & CurrentTime
    Print #filenum, "Template name: " & TemplateName
    Print #filenum, "This file contains hotkeys to styles defined within this template."
    Print #filenum, "It was made by the Microsoft Word macro."
    Print #filenum, "Filename: " & FileName
    Print #filenum, "Module name: " & ModuleName
    Print #filenum, "Macro name: " & MacroName
    Print #filenum, vbCrLf
    
    Dim i As Integer
    ' Write the hotkeys to the file
    For i = LBound(MyStyleName) To UBound(MyStyleName)
        Print #filenum, i + 1 & " | " & MyStyleName(i) & " | " & MyShortcut(i) & " |"
    Next i
        
    ' Close the file
    Close #filenum
        
    MsgBox _
        Prompt:="Hotkeys related to macros have been saved to " & vbNewLine & FilePath, _
        Buttons:=vbInformation + vbOKOnly, _
        Title:=MsgBoxTitle
End Sub

' 2025-07-16 by ms, review
Sub ListMWShortcutsToDOCX()
    Dim FileName As String
    FileName = C_F_Macros
    
    Dim ModuleName As String
    ModuleName = C_M_Shortcuts
    
    Dim MacroName As String
    MacroName = "ListMWShortcutsToDOCX"
    
    Dim MsgBoxTitle As String
    MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    'ListAllCommands = False
    Call CopyMWShortcuts(Argument:=False)

    ' Inform the user
    MsgBox _
        Prompt:="Microsoft Word shortcuts have been copied to a new document.", _
        Buttons:=vbInformation, _
        Title:=MsgBoxTitle
End Sub

' 2025-07-16 by ms, review
Sub ListAllMWCommandsToDOCX()
    Dim FileName As String
    FileName = C_F_Macros
    
    Dim ModuleName As String
    ModuleName = C_M_Shortcuts
    
    Dim MacroName As String
    MacroName = "ListAllMWCommandsToDOCX"
    
    Dim MsgBoxTitle As String
    MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName

    Call CopyMWShortcuts(Argument:=True)

    ' Inform the user
    MsgBox _
        Prompt:="Microsoft Word shortcuts have been copied to a new document.", _
        Buttons:=vbInformation, _
        Title:=MsgBoxTitle
End Sub


' 2025-02-28 by ms and AI
Private Sub CopyMWShortcuts(Argument As Boolean)
    Dim tempDoc As Document
    
    ' List all commands and their shortcuts, including those defined / redefined by a user
    Application.ListCommands ListAllCommands:=Argument
    
    ' Get the document created by ListCommands
    Set tempDoc = ActiveDocument
    
    ' Save the new document
    If Argument = True Then
        tempDoc.SaveAs2 FileName:="MWAllCommandsList.docx"
    Else
        tempDoc.SaveAs2 FileName:="MWKeyboardShortcutsList.docx"
    End If
    
    ' Clear object variables
    Set tempDoc = Nothing
End Sub

' No longer used, left as legacy.
' 2025-11-20 by ms
' Note that specified keybindings are stored in the attached template
' Future: use Hots_UniversalForm
' 2025-02-28 by ms and AI
Private Sub ShowActiveDocumentMacroShortcuts()
    Dim keyBinding As keyBinding
    Dim shortcutText As String
   
    ' This line is essential to get access to information stored actually in the currently attached template
    CustomizationContext = ActiveDocument
   
    ' Initialize the shortcut text
    shortcutText = "User Defined Shortcuts:" & vbCrLf & vbCrLf
        
    ' wdKeyCategoryCommand: This category includes built-in Word commands. Checked, it works.
    ' wdKeyCategoryMacro:   This category includes macros that you have created. Checked, it works.
    ' wdKeyCategoryFont:    This category includes font-related commands. Didn't check.
    ' wdKeyCategoryAutoText: This category includes AutoText entries. Seems not to work.
    ' wdKeyCategoryStyle:   This category includes styles. Checked, it works.
    ' wdKeyCategorySymbol:  This category includes symbols. Didn't check.
    ' wdKeyCategoryPrefix:  This category includes prefix keys. Seems not to work.
    
    ' Loop through all key bindings
    For Each keyBinding In Application.KeyBindings
        ' Check if the key binding belongs to the menu category
        If keyBinding.KeyCategory = wdKeyCategoryCommand Then
            ' Add the command name and its shortcut to the text
            shortcutText = shortcutText & keyBinding.Command & ": " & keyBinding.KeyString & vbCrLf
        End If
    Next keyBinding
    
    shortcutText = shortcutText & vbNewLine & "Macros keyboard shortcuts" & vbNewLine & vbNewLine
    
    ' Loop through all key bindings
    For Each keyBinding In Application.KeyBindings
        ' Check if the key binding belongs to the menu category
        If keyBinding.KeyCategory = wdKeyCategoryMacro Then
            ' Add the command name and its shortcut to the text
            shortcutText = shortcutText & keyBinding.Command & ": " & keyBinding.KeyString & vbCrLf
        End If
    Next keyBinding
    
    Dim FileName As String
    FileName = C_F_Macros
    
    Dim ModuleName As String
    ModuleName = C_M_Shortcuts
    
    Dim MacroName As String
    MacroName = "ShowActiveDocumentMacroShortcuts"
    
    Dim MsgBoxTitle As String
    MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    ' Display the shortcuts in a message box
    MsgBox _
        Prompt:=shortcutText, _
        Buttons:=vbInformation + vbOKOnly, _
        Title:=MsgBoxTitle
End Sub

' 2025-07-15 by ms
' This macro contains changed or added shortcuts. I don't remember why I keep them in separate sub.
Private Sub Set_CommandShortcuts(ByVal IfMsgBox As String)
    Dim FileName As String
    FileName = C_F_Macros
    
    Dim ModuleName As String
    ModuleName = C_M_Shortcuts
    
    Dim MacroName As String
    MacroName = "Set_CommandShortcuts"
    
    Dim MsgBoxTitle As String
    MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
        
    ' Validate argument
    If LCase(IfMsgBox) <> "quiet" And LCase(IfMsgBox) <> "loud" Then
        MsgBox _
            Prompt:="Invalid argument. Use 'quiet' or 'loud'." & vbNewLine & vbNewLine & "Exiting.", _
            Buttons:=vbCritical + vbOKOnly, _
            Title:=MsgBoxTitle
    End If
    
    ' Assign Alt + Ctrl + H to NavPane
    KeyBindings.Add _
        KeyCategory:=wdKeyCategoryCommand, _
        Command:="NavPane", _
        KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyControl, wdKeyH)
    
    ' Assign Alt + Ctrl + P to FormatParagraph
    KeyBindings.Add _
        KeyCategory:=wdKeyCategoryCommand, _
        Command:="FormatParagraph", _
        KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyControl, wdKeyP)
    
    ' Assign Alt + F to FormatFont
    KeyBindings.Add _
        KeyCategory:=wdKeyCategoryCommand, _
        Command:="FormatFont", _
        KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyF)
            
    If LCase(IfMsgBox) = "loud" Then
        MsgBox _
            Prompt:="The following shortcuts have been added:" & vbNewLine & vbNewLine & _
                C_SC_AltCtrlH & " -> " & "NavPane" & vbNewLine & _
                C_SC_AltCtrlP & " -> " & "FormatParagraph" & vbNewLine & _
                C_SC_AltF & " -> " & "FormatFont" & vbNewLine & _
                vbNewLine & _
            "Processing complete.", _
            Buttons:=vbInformation + vbOKOnly, _
            Title:=MsgBoxTitle
    End If
End Sub

' 2025-07-15 by ms
' Macro complementary to the Set_CommandShortcuts()
Public Sub Reset_CommandShortcut()
    Dim kb As keyBinding
    
    ' Clear Alt + Ctrl + H shortcut for NavPane
    Set kb = Application.FindKey(BuildKeyCode(wdKeyAlt, wdKeyControl, wdKeyH))
    kb.Clear
    
    ' Clear Alt + Ctrl + P shortcut for FormatParagraph
    Set kb = Application.FindKey(BuildKeyCode(wdKeyAlt, wdKeyControl, wdKeyP))
    kb.Clear
    
    ' Clear Alt + F shortcut for FormatFont
    Set kb = Application.FindKey(BuildKeyCode(wdKeyAlt, wdKeyF))
    kb.Clear
    
    Dim FileName As String
    FileName = C_F_Macros
    
    Dim ModuleName As String
    ModuleName = C_M_Shortcuts
    
    Dim MacroName As String
    MacroName = "Reset_CommandShortcut"
    
    Dim MsgBoxTitle As String
    MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    MsgBox _
        Prompt:="The following shortcuts have been cleared:" & vbNewLine & vbNewLine & _
            C_SC_AltCtrlH & " -> " & "NavPane" & vbNewLine & _
            C_SC_AltCtrlP & " -> " & "FormatParagraph" & vbNewLine & _
            vbNewLine & _
            "Processing complete.", _
        Buttons:=vbInformation, _
        Title:=MsgBoxTitle
End Sub

' Logs to the file some statistics regarding shortcuts for current active file.
' 2025-03-07 by ms and AI
' 2025-12-01 by ms
Sub ListAllShortcutsToTxt()
    Dim kb As keyBinding
    Dim TemplateName As String
    
    Dim AutoTextList As String
    Dim CommandList As String
    Dim DisableList As String
    Dim fontList As String
    Dim MacroList As String
    Dim NilList As String   ' null or none
    Dim PrefixList As String
    Dim StyleList As String
    Dim SymbolList As String
    
    Dim AutoTextCounter As Integer
    Dim CommandCounter As Integer
    Dim DisableCounter As Integer
    Dim FontCounter As Integer
    Dim MacroCounter As Integer
    Dim NilCounter As Integer
    Dim PrefixCounter As Integer
    Dim StyleCounter As Integer
    Dim SymbolCounter As Integer
    
    Dim totalShortcuts As Integer
    Dim shortcutList As String
    Dim filenum As Integer
    Dim FilePath As String
    Dim CurrentDate As String
    Dim CurrentTime As String
    
    Dim FileName As String
    FileName = C_F_Macros
    
    Dim ModuleName As String
    ModuleName = C_M_Shortcuts
    
    Dim MacroName As String
    MacroName = "ListAllShortcutsToTxt"
    
    Dim MsgBoxTitle As String
    MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    ' Get the name of the attached template
    TemplateName = ActiveDocument.AttachedTemplate.FullName
    CurrentDate = Format(Date, "yyyy-mm-dd")
    CurrentTime = Format(Time, "hh:mm:ss")
    
    ' Initialize the shortcut lists and counters
    AutoTextList = "Shortcuts category AutoText:" & vbCrLf
    CommandList = "Shortcuts category Command:" & vbCrLf
    DisableList = "Shorcuts category Disable:" & vbNewLine
    fontList = "Shortcuts category Font:" & vbCrLf
    MacroList = "Shortcuts category Macro:" & vbCrLf
    NilList = "Shortcuts category Nil:" & vbCrLf
    PrefixList = "Shortcuts category Prefix:" & vbCrLf
    StyleList = "Shortcuts category Style:" & vbCrLf
    SymbolList = "Shortcuts category Symbol:" & vbCrLf
    
    AutoTextCounter = 0
    CommandCounter = 0
    DisableCounter = 0
    FontCounter = 0
    MacroCounter = 0
    NilCounter = 0
    PrefixCounter = 0
    StyleCounter = 0
    SymbolCounter = 0
    
    Word.CustomizationContext = ActiveDocument
    
    ' 1. Loop through all key bindings for AutoText category
    For Each kb In Application.KeyBindings
        If kb.Context = ActiveDocument And kb.KeyCategory = wdKeyCategoryAutoText Then
            AutoTextCounter = AutoTextCounter + 1
            AutoTextList = AutoTextList & AutoTextCounter & ". " & kb.Command & " - " & kb.KeyString & vbCrLf
        End If
    Next kb
    
    ' 2. Loop through all key bindings for Command category
    For Each kb In Application.KeyBindings
        If kb.Context = ActiveDocument And kb.KeyCategory = wdKeyCategoryCommand Then
            CommandCounter = CommandCounter + 1
            CommandList = CommandList & CommandCounter & ". " & kb.Command & " - " & kb.KeyString & vbCrLf
        End If
    Next kb
    
    ' 3. Loop through all key bindings for Disable category
    For Each kb In Application.KeyBindings
        If kb.Context = ActiveDocument And kb.KeyCategory = wdKeyCategoryDisable Then
            CommandCounter = CommandCounter + 1
            CommandList = CommandList & CommandCounter & ". " & kb.Command & " - " & kb.KeyString & vbCrLf
        End If
    Next kb
    
    ' 4. Loop through all key bindings for Font category
    For Each kb In Application.KeyBindings
        If kb.Context = ActiveDocument And kb.KeyCategory = wdKeyCategoryFont Then
            FontCounter = FontCounter + 1
            fontList = fontList & FontCounter & ". " & kb.Command & " - " & kb.KeyString & vbCrLf
        End If
    Next kb
    
    ' 5. Loop through all key bindings for Macro category
    For Each kb In Application.KeyBindings
        If kb.Context = ActiveDocument And kb.KeyCategory = wdKeyCategoryMacro Then
            MacroCounter = MacroCounter + 1
            MacroList = MacroList & MacroCounter & ". " & kb.Command & " - " & kb.KeyString & vbCrLf
        End If
    Next kb
    
    ' 6. Loop through all key bindings for Nil category
    For Each kb In Application.KeyBindings
        If kb.Context = ActiveDocument And kb.KeyCategory = wdKeyCategoryNil Then
            StyleCounter = StyleCounter + 1
            StyleList = StyleList & StyleCounter & ". " & kb.Command & " - " & kb.KeyString & vbCrLf
        End If
    Next kb
      
    ' 7. Loop through all key bindings for Prefix category
    For Each kb In Application.KeyBindings
        If kb.Context = ActiveDocument And kb.KeyCategory = wdKeyCategoryPrefix Then
            PrefixCounter = PrefixCounter + 1
            PrefixList = PrefixList & PrefixCounter & ". " & kb.Command & " - " & kb.KeyString & vbCrLf
        End If
    Next kb
    
    ' 8. Loop through all key bindings for Style category
    For Each kb In Application.KeyBindings
        If kb.Context = ActiveDocument And kb.KeyCategory = wdKeyCategoryStyle Then
            StyleCounter = StyleCounter + 1
            StyleList = StyleList & StyleCounter & ". " & kb.Command & " - " & kb.KeyString & vbCrLf
        End If
    Next kb
    
    ' 9. Loop through all key bindings for Symbol category
    For Each kb In Application.KeyBindings
        If kb.Context = ActiveDocument And kb.KeyCategory = wdKeyCategorySymbol Then
            SymbolCounter = SymbolCounter + 1
            SymbolList = SymbolList & SymbolCounter & ". " & kb.Command & " - " & kb.KeyString & vbCrLf
        End If
    Next kb
    
    ' Calculate total number of shortcuts
    totalShortcuts = AutoTextCounter + CommandCounter + DisableCounter + FontCounter + MacroCounter + NilCounter + PrefixCounter + StyleCounter + SymbolCounter
    
    ' Combine all lists into one string to display in MsgBox
    shortcutList = "Current filename: " & ActiveDocument.Name & vbCrLf & _
                   "Total Shortcuts: " & totalShortcuts & vbCrLf & vbCrLf & _
                   AutoTextList & vbCrLf & _
                   CommandList & vbCrLf & _
                    DisableList & vbNewLine & _
                   fontList & vbCrLf & _
                   MacroList & vbCrLf & _
                    NilList & vbNewLine & _
                   PrefixList & vbCrLf & _
                   StyleList & vbCrLf & _
                   SymbolList & vbCrLf
      
    ' Save the shortcut list to a file
    FilePath = Options.DefaultFilePath(wdDocumentsPath) & "\" & MacroName & ".txt"
    filenum = FreeFile
    Open FilePath For Output As filenum
    Print #filenum, "Macro Name: " & MacroName
    Print #filenum, "Date: " & CurrentDate
    Print #filenum, "Time: " & CurrentTime
    Print #filenum, shortcutList
    Close filenum
    
    ' Display the summary in a message box
    MsgBox _
        Prompt:="Processing complete." & vbNewLine & vbNewLine & _
            "Information was saved to the file:" & vbNewLine & vbNewLine & FilePath, _
        Buttons:=vbInformation, _
        Title:=MsgBoxTitle
End Sub

' 2025-08-03 by ms
' Update of context.
' Added check if the keybindings are from ActiveDocument.AttachedTemplate.
' The issue is that KeyBindings collection contains all keybindings from all contexts and I need to set correct one.
Private Sub SetKeyBindingMacro(ByVal KeybShortcut As String, _
                                ByVal WhichMacro As String, _
                                ByVal IfMsgBox As String)
    Dim FileName As String
    FileName = C_F_Macros
    
    Dim ModuleName As String
    ModuleName = C_M_Shortcuts
    
    Dim MacroName As String
    MacroName = "SetKeyBindingMacro"
    
    Dim MsgBoxTitle As String
    MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    If LCase(IfMsgBox) <> "quiet" And LCase(IfMsgBox) <> "loud" Then
        MsgBox _
            Prompt:="Invalid argument. Use 'quiet' or 'loud'." & vbNewLine & vbNewLine & "Exiting.", _
            Buttons:=vbCritical + vbOKOnly, _
            Title:=MsgBoxTitle
        Exit Sub
    End If
        
    Dim kb As keyBinding
    Dim MyCode1 As Long
    Dim MyCode2 As Integer
        
    MyCode1 = ParseKeyCode1(KeybShortcut)
    MyCode2 = ParseKeyCode2(KeybShortcut)
    
    ' Loop through all key bindings
    For Each kb In Application.KeyBindings
        ' Check if the key binding belongs to the menu category
        If InStr(1, kb.Command, WhichMacro) And kb.KeyCode = MyCode1 Then
            If kb.KeyCode2 <> 0 And kb.KeyCode2 = MyCode2 Then
                ' Such keybinding already exists
                MsgBox _
                    Prompt:="Such keybindg already exists:" & vbNewLine & _
                    KeybShortcut & " : " & WhichMacro & vbNewLine & vbNewLine & "Exiting.", _
                    Buttons:=vbInformation + vbOKOnly + vbDefaultButton1, _
                    Title:=MsgBoxTitle
                Exit Sub
            Else
                MsgBox _
                    Prompt:="Such keybindg already exists:" & vbNewLine & _
                    KeybShortcut & " : " & WhichMacro & vbNewLine & vbNewLine & "Exiting.", _
                    Buttons:=vbInformation + vbOKOnly + vbDefaultButton1, _
                    Title:=MsgBoxTitle
                Exit Sub
            End If
        End If
    Next kb
    
    If MyCode2 <> 0 Then
        On Error GoTo ShortcutError
        Set kb = KeyBindings.Add( _
                    KeyCategory:=wdKeyCategoryMacro, _
                    Command:=WhichMacro, _
                    KeyCode:=MyCode1, _
                    KeyCode2:=MyCode2)
        On Error GoTo 0
    Else
        On Error GoTo ShortcutError
        Set kb = KeyBindings.Add( _
                    KeyCategory:=wdKeyCategoryMacro, _
                    Command:=WhichMacro, _
                    KeyCode:=MyCode1)
        On Error GoTo 0
    End If
    
    If Not kb Is Nothing Then
        If LCase(IfMsgBox) = "loud" Then
            MsgBox _
                Prompt:="Keybinding for " & KeybShortcut & " has been set to " & WhichMacro & "." & _
                    "Customization context:" & vbNewLine & vbNewLine & _
                    CustomizationContext, _
                Buttons:=vbInformation + vbOKOnly, _
                Title:=MsgBoxTitle
            Exit Sub
        End If
    Else
        MsgBox _
            Prompt:="Keybinding for " & KeybShortcut & " was not set." & vbNewLine & vbNewLine & _
                "Customization context:" & vbNewLine & vbNewLine & _
                CustomizationContext, _
            Buttons:=vbExclamation + vbOKOnly, _
            Title:=MsgBoxTitle
    End If
    
    ' Clear object variables
    Set kb = Nothing
    Exit Sub
    
ShortcutError:
    MsgBox _
        Prompt:="Error assigning shortcut: " & vbNewLine & _
            "Error Number: " & Err.Number & vbNewLine & _
            "Error Description: " & Err.Description & vbNewLine & _
            "Perhaps such command does not exist?", _
        Buttons:=vbCritical, _
        Title:=MsgBoxTitle
End Sub

Private Sub SetKeyBindingStyle(KeybShortcut As String, WhichStyle As String)
    Dim kb As keyBinding
    Dim MyCode1 As Long
    Dim MyCode2 As Integer
    
    Dim FileName As String:     FileName = C_F_Macros
    Dim ModuleName As String:   ModuleName = C_M_Shortcuts
    Dim MacroName As String:    MacroName = "SetKeyBindingStyle"
    Dim MsgBoxTitle As String:  MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    ' Set the customization context to the current template. This is conscious exception.
    CustomizationContext = ActiveDocument.AttachedTemplate
    
    MyCode1 = ParseKeyCode1(KeybShortcut)
    MyCode2 = ParseKeyCode2(KeybShortcut)
    
    ' Loop through all key bindings
    For Each kb In Application.KeyBindings
        ' Check if the key binding belongs to the menu category
        If InStr(1, kb.Command, WhichStyle) And kb.KeyCode = MyCode1 Then
            If kb.KeyCode2 <> 0 And kb.KeyCode2 = MyCode2 Then
                ' Such keybinding already exists
                MsgBox _
                    Prompt:="Such keybindg already exists:" & vbNewLine & _
                    KeybShortcut & " : " & MacroName & vbNewLine & vbNewLine & "Exiting.", _
                    Buttons:=vbInformation + vbOKOnly + vbDefaultButton1, _
                    Title:=MsgBoxTitle
                Exit Sub
            Else
                MsgBox _
                    Prompt:="Such keybindg already exists:" & vbNewLine & _
                    KeybShortcut & " : " & MacroName & vbNewLine & vbNewLine & "Exiting.", _
                    Buttons:=vbInformation + vbOKOnly + vbDefaultButton1, _
                    Title:=MsgBoxTitle
                Exit Sub
            End If
        End If
    Next kb
    
    If MyCode2 <> 0 Then
        Set kb = KeyBindings.Add(KeyCategory:=wdKeyCategoryStyle, _
                    Command:=WhichStyle, _
                    KeyCode:=MyCode1, _
                    KeyCode2:=MyCode2)
    Else
        Set kb = KeyBindings.Add(KeyCategory:=wdKeyCategoryStyle, _
                    Command:=WhichStyle, _
                    KeyCode:=MyCode1)
    End If
    
    If Not kb Is Nothing Then
        MsgBox _
            Prompt:="Keybinding for " & KeybShortcut & " has been set to " & WhichStyle, _
            Buttons:=vbInformation + vbOKOnly, _
            Title:=MsgBoxTitle
    Else
        MsgBox _
            Prompt:="Keybinding for " & KeybShortcut & " not set.", _
            Buttons:=vbInformation + vbOKOnly, _
            Title:=MsgBoxTitle
    End If
    
    ' Clear object variables
    Set kb = Nothing
End Sub

' 2025-08-02 by ms
' Added checking of CustomizationContext on time of deletion.
' The issue is that KeyBindings collection contains all keybindings from all contexts and I need to check if the keybinding comes from the right context.
' Deletes only from ActiveDocument.AttachedTemplate.
Private Sub DeleteKeyBinding(KeybShortcut As String)
    Dim kb As keyBinding
    Dim MyCode1 As Long
    Dim MyCode2 As Integer
    
    Dim FileName As String:     FileName = C_F_Macros
    Dim ModuleName As String:   ModuleName = C_M_Shortcuts
    Dim MacroName As String:    MacroName = "DeleteKeyBinding"
    Dim MsgBoxTitle As String:  MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    MyCode1 = ParseKeyCode1(KeybShortcut)
    MyCode2 = ParseKeyCode2(KeybShortcut)
    
    Application.CustomizationContext = ActiveDocument

    Dim FoundFlag As Boolean
    FoundFlag = False
    If MyCode2 <> 0 Then
        For Each kb In KeyBindings
            If kb.KeyCode = MyCode1 And kb.KeyCode2 = MyCode2 And kb.Context.Name = CustomizationContext Then
                kb.Clear
                FoundFlag = True
            End If
        Next kb
    Else
        For Each kb In KeyBindings
            If kb.KeyCode = MyCode1 And kb.Context.Name = CustomizationContext Then
                kb.Clear
                FoundFlag = True
            End If
        Next kb
    End If
    
    If FoundFlag Then
        MsgBox _
            Prompt:="Keybinding for " & KeybShortcut & " has been deleted." & vbNewLine & _
                "from the current context:" & vbNewLine & vbNewLine & _
                CustomizationContext, _
            Buttons:=vbInformation + vbOKOnly, _
            Title:=MsgBoxTitle
    Else
        MsgBox _
            Prompt:="Keybinding for " & KeybShortcut & " was not found in the current context:" & vbNewLine & _
                CustomizationContext, _
            Buttons:=vbInformation + vbOKOnly, _
            Title:=MsgBoxTitle
    End If

End Sub

' 2025-07-16 by ms
' By default parameters in VBA are handled ByRef.
Private Function ParseKeyCode2(ByVal KeybShortcut As String) As Integer
    ' Remove spaces from the MyShortcuttext string
    KeybShortcut = Replace(KeybShortcut, " ", "")

    ' Check if there is a comma in the shortcut text string
    If InStr(KeybShortcut, ",") > 0 Then
        ParseKeyCode2 = Asc(Right(KeybShortcut, 1))
    Else
        ParseKeyCode2 = 0
    End If

End Function

' 2025-07-16 by ms
' By default parameters in VBA are handled ByRef.
Private Function ParseKeyCode1(ByVal KeybShortcut As String) As Long
    Dim parts() As String
    Dim i As Integer
    Dim key As String
    Dim MacroName As String
    Dim ModuleName As String
    Dim MsgBoxHeader As String
    Dim KeyCode As Long
    
    KeybShortcut = Replace(KeybShortcut, " ", "")
    If InStr(KeybShortcut, ",") > 0 Then
        KeybShortcut = Left(KeybShortcut, Len(KeybShortcut) - 2)
    End If

    ' Split the shortcut text string by "+"
    parts = Split(KeybShortcut, "+")

    ' Initialize keyCode
    KeyCode = 0

    ' Loop through each part of the shortcut
    For i = LBound(parts) To UBound(parts)
        key = parts(i)
        Select Case key
            Case "Alt"
                KeyCode = KeyCode Or BuildKeyCode(wdKeyAlt)
            Case "Shift"
                KeyCode = KeyCode Or BuildKeyCode(wdKeyShift)
            Case "Ctrl"
                KeyCode = KeyCode Or BuildKeyCode(wdKeyControl)
            Case Else
                KeyCode = KeyCode Or BuildKeyCode(KeyStringToKeyCode(key))
        End Select
    Next i

    ParseKeyCode1 = KeyCode
End Function

Private Function KeyStringToKeyCode(key As String) As Long
    Select Case key
        Case "F1": KeyStringToKeyCode = wdKeyF1
        Case "F2": KeyStringToKeyCode = wdKeyF2
        Case "F3": KeyStringToKeyCode = wdKeyF3
        Case "F4": KeyStringToKeyCode = wdKeyF4
        Case "F5": KeyStringToKeyCode = wdKeyF5
        Case "F6": KeyStringToKeyCode = wdKeyF6
        Case "F7": KeyStringToKeyCode = wdKeyF7
        Case "F8": KeyStringToKeyCode = wdKeyF8
        Case "F9": KeyStringToKeyCode = wdKeyF9
        Case "F10": KeyStringToKeyCode = wdKeyF10
        Case "F11": KeyStringToKeyCode = wdKeyF11
        Case "F12": KeyStringToKeyCode = wdKeyF12
        Case "Insert": KeyStringToKeyCode = wdKeyInsert
        Case "[": KeyStringToKeyCode = wdKeyOpenSquareBrace
        Case "]": KeyStringToKeyCode = wdKeyCloseSquareBrace
        ' Add more cases as needed for other keys
        Case Else: KeyStringToKeyCode = Asc(key)    'Asc = ASCII code
    End Select
End Function


Private Sub SetShortcut_CustomizedToggleFieldCodes(ByVal IfMsgBox As String)
    Call SetKeyBindingMacro(KeybShortcut:=C_SC_AltF9, _
                            WhichMacro:="CustomizedToggleFieldCodes", _
                            IfMsgBox:="quiet")
End Sub

' Original Microsoft Word command: ViewFieldCodes
Private Sub ResetShortcut_CustomizedToggleFieldCodes()
    Call DeleteKeyBinding(KeybShortcut:=C_SC_AltF9)
End Sub

' Sets keyboard shortcut to the macro JumpToNextList
' 2025-03-08 by ms
Private Sub SetShortcut_JumpToNextList(ByVal IfMsgBox As String)
    Call SetKeyBindingMacro(KeybShortcut:=C_SC_AltF2, _
                            WhichMacro:="JumpToNextList", _
                            IfMsgBox:=IfMsgBox)
End Sub

Private Sub ResetShortcut_JumpToNextList()
    Call DeleteKeyBinding(KeybShortcut:=C_SC_AltF2)
End Sub

' Sets keyboard shortcut to the macro JumpToNextTable
' 2025-03-08 by ms
Private Sub SetShortcut_JumpToNextTable(ByVal IfMsgBox As String)
    Call SetKeyBindingMacro(KeybShortcut:=C_SC_AltF3, _
                            WhichMacro:="JumpToNextTable", _
                            IfMsgBox:=IfMsgBox)
End Sub

Private Sub ResetShortcut_JumpToNextTable()
    Call DeleteKeyBinding(KeybShortcut:=C_SC_AltF3)
End Sub

' Sets keyboard shortcut to the macro JumpToNextCanvas
' 2025-04-11 by ms
Private Sub SetShortcut_JumpToNextCanvas(ByVal IfMsgBox As String)
    Call SetKeyBindingMacro(KeybShortcut:=C_SC_AltF5, _
                            WhichMacro:="JumpToNextCanvas", _
                            IfMsgBox:=IfMsgBox)
End Sub

Private Sub ResetShortcut_JumpToNextCanvas()
    Call DeleteKeyBinding(KeybShortcut:=C_SC_AltF5)
End Sub


' Sets call-forward to built-in command "Save File" (Ctrl + S), adding macro to apply numbering distance
' 2025-03-12 by ms
Private Sub SetShortcut_SaveFileApplyNumberingDistance(ByVal IfMsgBox As String)
    Call SetKeyBindingMacro(KeybShortcut:=C_SC_CtrlS, _
                            WhichMacro:="ApplyDistanceBetweenNumberingAndHeading", _
                            IfMsgBox:=IfMsgBox)
End Sub

' Original Microsoft Word command: FileSave
' 2025-03-12 by ms
Private Sub ResetShortcut_SaveFileApplyNumberingDistance()
    Call DeleteKeyBinding(KeybShortcut:=C_SC_CtrlS)
End Sub

' Sets call-forward macro to built-in command
' "Close File" / command DocClose (Ctrl + W),
' 2025-03-15 by ms
Private Sub SetShortcut_CloseFileApplyUpdateFields(ByVal IfMsgBox As String)
    Call SetKeyBindingMacro(KeybShortcut:=C_SC_CtrlW, _
                            WhichMacro:="UpdateAllFieldsAndCloseFile", _
                            IfMsgBox:=IfMsgBox)
End Sub

' Original Microsoft Word command: DocClose
' 2025-03-12 by ms
Private Sub ResetShortcut_CloseFileApplyUpdateFields()
    Call DeleteKeyBinding(KeybShortcut:=C_SC_CtrlW)
End Sub

' Sets call-forward macro to built-in command
' "Print Preview" / command PrintPreviewAndPrint (Ctrl + F2),
' 2025-03-15 by ms
Private Sub SetShortcut_CustomizedPrintPreviewAndPrint(ByVal IfMsgBox As String)
    Call SetKeyBindingMacro(KeybShortcut:=C_SC_CtrlF2, _
                            WhichMacro:="CustomizedPrintPreviewAndPrint", _
                            IfMsgBox:=IfMsgBox)
End Sub

' Original Microsoft Word command: PrintPreviewAndPrint
' 2025-03-12 by ms
Private Sub ResetShortcut_CustomizedPrintPreviewAndPrint()
    Call DeleteKeyBinding(KeybShortcut:=C_SC_CtrlF2)
End Sub

' 2025-03-20 by ms
Private Sub SetShortcut_ToggleCharBoldStyle(ByVal IfMsgBox As String)
    ActiveDocument.Variables(C_S_Bold).Value = True
    Call SetKeyBindingMacro(KeybShortcut:=C_SC_CtrlB, _
                            WhichMacro:="ToggleCharBoldStyle", _
                            IfMsgBox:=IfMsgBox)
End Sub

' Original Microsoft Word command: Bold
' 2025-03-20 by ms
Private Sub ResetShortcut_ToggleCharBoldStyle()
    Call DeleteKeyBinding(KeybShortcut:=C_SC_CtrlB)
    ActiveDocument.Variables(C_S_Bold).Value = False
End Sub

' 2025-03-20 by ms
Private Sub SetShortcut_ToggleCharItalicStyle(ByVal IfMsgBox As String)
    ActiveDocument.Variables(C_S_Italic).Value = True
    Call SetKeyBindingMacro(KeybShortcut:=C_SC_CtrlI, _
                            WhichMacro:="ToggleCharItalicStyle", _
                            IfMsgBox:=IfMsgBox)
End Sub

' Original Microsoft Word command: Italic
' 2025-03-20 by ms
Private Sub ResetShortcut_ToggleCharItalicStyle()
    Call DeleteKeyBinding(KeybShortcut:=C_SC_CtrlI)
    ActiveDocument.Variables(C_S_Italic).Value = False
End Sub

' 2025-03-20 by ms
Private Sub SetShortcut_ToggleCharUnderlineStyle(ByVal IfMsgBox As String)
    ActiveDocument.Variables(C_S_Underline).Value = True
    Call SetKeyBindingMacro(KeybShortcut:=C_SC_CtrlU, _
                            WhichMacro:="ToggleCharUnderlineStyle", _
                            IfMsgBox:=IfMsgBox)
End Sub

' Original Microsoft Word command: Underline
' 2025-03-20 by ms
Private Sub ResetShortcut_ToggleCharUnderlineStyle()
    Call DeleteKeyBinding(KeybShortcut:=C_SC_CtrlU)
    ActiveDocument.Variables(C_S_Underline).Value = False
End Sub

' 2025-03-21 by ms
' It is strange and not yet clarified. This function always returns run-time error '5346': Word cannot change the function of the specified key.
' But when shortcut is set manually to the macro "ToggleCharCrossoutStyle", it works. So this is unclear why this error is set.
Private Sub SetShortcut_ToggleCharCrossoutStyle(ByVal IfMsgBox As String)
    ActiveDocument.Variables(C_S_CharCrossout).Value = True
    Call SetKeyBindingMacro(KeybShortcut:=C_SC_ShiftCtrlX, _
                            WhichMacro:="ToggleCharCrossoutStyle", _
                            IfMsgBox:=IfMsgBox)
End Sub

' 2025-03-21 by ms
Private Sub ResetShortcut_ToggleCharCrossoutStyle()
    Call DeleteKeyBinding(KeybShortcut:=C_SC_ShiftCtrlX)
    Call SetKeyBindingStyle(KeybShortcut:=C_SC_ShiftCtrlX, WhichStyle:=C_S_CharCrossout)
    ActiveDocument.Variables(C_S_CharCrossout).Value = False
End Sub

' 2025-03-21 by ms
Private Sub SetShortcut_ToggleCharHiddenStyle(ByVal IfMsgBox As String)
    ActiveDocument.Variables(C_S_CharHidden).Value = True
    Call SetKeyBindingMacro(KeybShortcut:=C_SC_ShiftCtrlH, _
                            WhichMacro:="ToggleCharHiddenStyle", _
                            IfMsgBox:=IfMsgBox)
End Sub

' 2025-03-21 by ms
' This function works only if object was defined by the macro "SetShortcut_ToggleCharHiddenStyle". So it must be run just after.
Private Sub ResetShortcut_ToggleCharHiddenStyle()
    Call DeleteKeyBinding(KeybShortcut:=C_SC_ShiftCtrlH)
    Call SetKeyBindingStyle(KeybShortcut:=C_SC_ShiftCtrlH, WhichStyle:=C_S_CharHidden)
    ActiveDocument.Variables(C_S_CharHidden).Value = False
End Sub

Private Sub SetShortcut_ToggleSpecificFormatting(ByVal IfMsgBox As String)
    Call SetKeyBindingMacro(KeybShortcut:=C_SC_F4, _
                            WhichMacro:="ToggleSpecificFormatting", _
                            IfMsgBox:=IfMsgBox) ' in module: Tools
End Sub

' Original Microsoft Word command: EditRedoOrRepeat
Private Sub ResetShortcut_ToggleSpecificFormatting()
    Call DeleteKeyBinding(KeybShortcut:=C_SC_F4)
End Sub

' 2025-04-14 by ms
Private Sub SetShortcut_ToggleCharSourceCode(ByVal IfMsgBox As String)
    ActiveDocument.Variables(C_S_CharSourceCode).Value = True
    Call SetKeyBindingMacro(KeybShortcut:=C_SC_ShiftCtrlK, _
                            WhichMacro:="ToggleCharSourceCode", _
                            IfMsgBox:=IfMsgBox)
End Sub

' 2025-04-14 by ms
' Original Microsoft Word command: SmallCaps
Private Sub ResetShortcut_ToggleCharSourceCode()
    Call DeleteKeyBinding(KeybShortcut:=C_SC_ShiftCtrlK)
    ActiveDocument.Variables(C_S_CharHidden).Value = False
End Sub

' 2025-04-14 by ms
Private Sub SetShortcut_SetLanguageToEnglishUS(ByVal IfMsgBox As String)
    Call SetKeyBindingMacro(KeybShortcut:=C_SC_F8, _
                            WhichMacro:="SetLanguageToEnglishUS", _
                            IfMsgBox:=IfMsgBox)
End Sub

' 2025-04-14 by ms
' Original Microsoft Word command: ExtendSelection
' This function works only if object was defined by the macro "SetShortcut_SetLanguageToEnglishUS". So it must be run just after.
Private Sub ResetShortcut_SetLanguageToEnglishUS()
    Call DeleteKeyBinding(KeybShortcut:=C_SC_F8)
End Sub

Private Sub SetShortcut_ToggleApplyStyles(ByVal IfMsgBox As String)
    Call SetKeyBindingMacro(KeybShortcut:=C_SC_ShiftCtrlS, _
                            WhichMacro:="ToggleApplyStyles", _
                            IfMsgBox:=IfMsgBox)
End Sub

' Original Microsoft Word command: StyleApplyPane
Private Sub ResetShortcut_ToggleApplyStyles()
    Call DeleteKeyBinding(KeybShortcut:=C_SC_ShiftCtrlS)
End Sub

' 2025-04-20 by ms
Private Sub SetShortcut_ToggleOvertypeMode(ByVal IfMsgBox As String)
    Call SetKeyBindingMacro(KeybShortcut:=C_SC_Insert, _
                            WhichMacro:="CustomizedOvertype", _
                            IfMsgBox:=IfMsgBox)
End Sub

' Original Microsoft Word command: Overtype
' 2025-04-20 by ms
Private Sub ResetShortcut_ToggleOvertypeMode()
    Call DeleteKeyBinding(KeybShortcut:=C_SC_Insert)
End Sub

' 2025-04-20 by ms
Private Sub SetShortcut_InsertCrossReference(ByVal IfMsgBox As String)
    Call SetKeyBindingMacro(KeybShortcut:=C_SC_F7, _
                            WhichMacro:="InsertCrossRef", _
                            IfMsgBox:=IfMsgBox)   ' module: Tools
End Sub

' Original Microsoft Word command: ToolsProofing
' 2025-04-20 by ms
Private Sub ResetShortcut_InsertCrossReference()
    Call DeleteKeyBinding(KeybShortcut:=C_SC_F7)
End Sub

' 2025-07-16 by ms
Private Sub SetShortcut_ReapplyTemplateStyle(ByVal IfMsgBox As String)
    Call SetKeyBindingMacro(KeybShortcut:=C_SC_AltRplusS, _
                            WhichMacro:="ReapplyTemplateStyle", _
                            IfMsgBox:=IfMsgBox) ' module: Tools
End Sub

' 2025-07-16 by ms
Private Sub ResetShortcut_ReapplyTemplateStyle()
    Call DeleteKeyBinding(KeybShortcut:=C_SC_AltRplusS)
End Sub

' 2025-07-16 by ms
Private Sub SetShortcut_RestartListNumbering(ByVal IfMsgBox As String)
    Call SetKeyBindingMacro(KeybShortcut:=C_SC_AltLplusR, _
                            WhichMacro:="RestartListNumbering", _
                            IfMsgBox:=IfMsgBox) ' module: Tools
End Sub

' 2025-07-16 by ms
Private Sub ResetShortcut_RestartListNumbering()
    Call DeleteKeyBinding(KeybShortcut:=C_SC_AltLplusR)
End Sub

' 2025-07-21 by ms
Private Sub SetShortcut_HotMacros(ByVal IfMsgBox As String)
    Call SetKeyBindingMacro(KeybShortcut:=C_SC_AltHplusM, _
                            WhichMacro:="ShowFormHotMacros", _
                            IfMsgBox:=IfMsgBox)
End Sub

' 2025-07-21 by ms
Private Sub ResetShortcut_HotMacros()
    Call DeleteKeyBinding(KeybShortcut:=C_SC_AltHplusM)
End Sub

' 2025-07-21 by ms
Private Sub SetShortcut_HotStrings(ByVal IfMsgBox As String)
    Call SetKeyBindingMacro(KeybShortcut:=C_SC_AltHplusS, _
                            WhichMacro:="ShowFormHotstrings", _
                            IfMsgBox:=IfMsgBox)
End Sub

' 2025-07-21 by ms
Private Sub ResetShortcut_Strings()
    Call DeleteKeyBinding(KeybShortcut:=C_SC_AltHplusS)
End Sub

' 2025-07-21 by ms
Private Sub SetShortcut_HotKeys(ByVal IfMsgBox As String)
    Call SetKeyBindingMacro(KeybShortcut:=C_SC_AltHplusK, _
                            WhichMacro:="ShowFormHotkeys", _
                            IfMsgBox:=IfMsgBox)
End Sub

' 2025-07-21 by ms
Private Sub ResetShortcut_HotKeys()
    Call DeleteKeyBinding(KeybShortcut:=C_SC_AltHplusK)
End Sub

' 2025-07-30 by ms
Private Sub SetShortcut_ToggleHeadingCollapseExpand(ByVal IfMsgBox As String)
    Call SetKeyBindingMacro(KeybShortcut:=C_SC_AltCtrlSqOpen, _
                            WhichMacro:="ToggleHeadingCollapseExpand", _
                            IfMsgBox:=IfMsgBox)
End Sub

' 2025-07-30 by ms
Private Sub ResetShortcut_ToggleHeadingCollapseExpand()
    Call DeleteKeyBinding(KeybShortcut:=C_SC_AltCtrlSqOpen)
End Sub

' 2025-08-03 by ms
Private Sub SetShortcut_PrintDocument(ByVal IfMsgBox As String)
    Call SetKeyBindingMacro(KeybShortcut:=C_SC_CtrlP, _
                            WhichMacro:="CustomizedPrinting", _
                            IfMsgBox:=IfMsgBox) ' module: Tools
End Sub

' 2025-08-03 by ms
Private Sub ResetShortcut_PrintDocument()
    Call DeleteKeyBinding(KeybShortcut:=C_SC_CtrlP)
End Sub

' 2025-08-03 by ms
Private Sub SetShortcut_CustomizedSaveAs(ByVal IfMsgBox As String)
    Call SetKeyBindingMacro(KeybShortcut:=C_SC_F12, _
                            WhichMacro:="CustomizedSaveAs", _
                            IfMsgBox:=IfMsgBox) ' module: Tools
End Sub

' 2025-08-03 by ms
Private Sub ResetShortcut_CustomizedSaveAs()
    Call DeleteKeyBinding(KeybShortcut:=C_SC_F12)
End Sub

' 2025-10-02 by ms
Private Sub SetShortcut_CustomizedCopyFormat(ByVal IfMsgBox As String)
    Call SetKeyBindingMacro(KeybShortcut:=C_SC_ShiftCtrlC, _
                            WhichMacro:="CustomizedCopyFormat", _
                            IfMsgBox:=IfMsgBox) ' module: Tools
End Sub

' 2025-10-02 by ms
Private Sub ResetShortcut_CustomizedCopyFormat()
    Call DeleteKeyBinding(KeybShortcut:=C_SC_ShiftCtrlC)
End Sub

' 2025-10-02 by ms
Private Sub SetShortcut_CustomizedPasteFormat(ByVal IfMsgBox As String)
    Call SetKeyBindingMacro(KeybShortcut:=C_SC_ShiftCtrlV, _
                            WhichMacro:="CustomizedPasteFormat", _
                            IfMsgBox:=IfMsgBox) ' module: Tools
End Sub

' 2025-10-02 by ms
Private Sub ResetShortcut_CustomizedPasteFormat()
    Call DeleteKeyBinding(KeybShortcut:=C_SC_ShiftCtrlV)
End Sub

' 2025-11-20 by ms
Sub ClearActiveDocumentMacroShortcuts()
    Dim kb As keyBinding
    Dim UserDecision As VbMsgBoxResult
    
    Dim FileName As String
    FileName = C_F_Macros
    
    Dim ModuleName As String
    ModuleName = C_M_Macros
    
    Dim MacroName As String
    MacroName = "ClearActiveDocumentMacroShortcuts"
    
    Dim MsgBoxTitle As String
    MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    UserDecision = MsgBox( _
        Prompt:="Remove all shortcuts assigned to macros in the document" & vbNewLine & ActiveDocument.Name & "?", _
        Buttons:=vbYesNo + vbExclamation, _
        Title:=MsgBoxTitle)
    If UserDecision = vbNo Then
        Exit Sub
    End If
    
    CustomizationContext = ActiveDocument
    
    Dim MacroShortcutCounter As Byte
    MacroShortcutCounter = 0
    For Each kb In KeyBindings
        If kb.KeyCategory = wdKeyCategoryMacro Then
            kb.Clear
            MacroShortcutCounter = MacroShortcutCounter + 1
        End If
    Next kb
    
    If MacroShortcutCounter > 0 Then
        MsgBox _
            Prompt:="All " & MacroShortcutCounter & " macro shortcuts cleared from the " & vbNewLine & ActiveDocument.Name, _
            Buttons:=vbInformation, _
            Title:=MsgBoxTitle
    End If
    
    If MacroShortcutCounter = 0 Then
        MsgBox _
            Prompt:="No macro shortcuts found in " & vbNewLine & ActiveDocument.Name, _
            Buttons:=vbInformation, _
            Title:=MsgBoxTitle
    End If
End Sub

' 2025-11-20 by ms
Sub ClearActiveDocumentStyleShortcuts()
    Dim kb As keyBinding
    Dim UserDecision As VbMsgBoxResult
    
    Dim FileName As String
    FileName = C_F_Macros
    
    Dim ModuleName As String
    ModuleName = C_M_Macros
    
    Dim MacroName As String
    MacroName = "ClearActiveDocumentStyleShortcuts"
    
    Dim MsgBoxTitle As String
    MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    UserDecision = MsgBox( _
        Prompt:="Remove all shortcuts assigned to styles in the document" & vbNewLine & ActiveDocument.Name & "?", _
        Buttons:=vbYesNo + vbExclamation, _
        Title:=MsgBoxTitle)
    If UserDecision = vbNo Then
        Exit Sub
    End If
    
    CustomizationContext = ActiveDocument
    
    Dim StyleShortcutCounter As Byte
    StyleShortcutCounter = 0
    For Each kb In KeyBindings
        If kb.KeyCategory = wdKeyCategoryMacro Then
            kb.Clear
            StyleShortcutCounter = StyleShortcutCounter + 1
        End If
    Next kb
    
    If StyleShortcutCounter > 0 Then
        MsgBox _
            Prompt:="All " & StyleShortcutCounter & " styles shortcuts cleared from the " & vbNewLine & ActiveDocument.Name, _
            Buttons:=vbInformation, _
            Title:=MsgBoxTitle
    End If
    
    If StyleShortcutCounter = 0 Then
        MsgBox _
            Prompt:="No style shortcuts found in " & vbNewLine & ActiveDocument.Name, _
            Buttons:=vbInformation, _
            Title:=MsgBoxTitle
    End If
End Sub
