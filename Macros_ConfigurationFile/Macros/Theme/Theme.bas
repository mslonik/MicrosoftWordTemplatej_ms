Attribute VB_Name = "Theme"
' VBA Module name: Theme.bas
' https://github.com/mslonik/Microsoft-Word-Configuration
'
'   License: MIT License.
'
'
'+----+-------------+-------------+----------------+--------------------+
'| No | Sub name    | Ribbon name | Ribbon section | Ribbon button name |
'+----+-------------+-------------+----------------+--------------------+
'| 1  | AttachTheme | Styles_ms   | Theme          | AttachTheme        |
'+----+-------------+-------------+----------------+--------------------+
'
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
' Used to enforce the explicit declaration of all variables in a module. When you include Option Explicit at the beginning of a module, it ensures that you must
' declare all variables using the Dim, Private, Public, ReDim, or Static statements before using them. This helps prevent errors caused by typos or
' undeclared variables.
Option Explicit

Public Enum ColourType
    ColourTypeRGB = &H0             ' 0
    ColourTypeAutomatic = &HFF      ' 255
    ColourTypeSystem = &H80         ' 128
    ColourTypeThemeLow = &HD0       ' 208
    ColourTypeThemeHigh = &HDF      ' 223
End Enum

' 2025-04-02 reworked by ms
Public Type ColourDetails
    ColourValue     As Long
    ColourType      As ColourType
    ThemeColorIndex As WdThemeColorIndex
    ThemeColorText  As String
    TintAndShade    As Double
    TintAndShadeText As String
    RGB             As Long
    RGB_Hex         As String
    Red             As Byte
    Green           As Byte
    Blue            As Byte
End Type

Private Type HSL
    H As Double ' Range 0 - 1
    s As Double ' Range 0 - 1
    L As Double ' Range 0 - 1
End Type

'
' 2025-04-11 by ms and AI
' 2025-07-14 by  ms
Sub AttachTheme()
    Dim FileName As String:     FileName = C_F_Macros
    Dim ModuleName As String:   ModuleName = C_M_Theme
    Dim MacroName As String:    MacroName = "AttachTheme"
    Dim MsgBoxTitle As String:  MsgBoxTitle = FileName & " : " & ModuleName & " : " & MacroName
    
    Dim ThemePathAndFileName As String
    ThemePathAndFileName = Options.DefaultFilePath(wdUserTemplatesPath) & "\Document Themes\" & C_F_Theme
        
    ' Ensure the theme file exists
    If Dir(ThemePathAndFileName) <> "" Then
        ' Apply the theme to the active document
        ActiveDocument.ApplyDocumentTheme ThemePathAndFileName
        MsgBox _
            Prompt:="Theme applied successfully!" & vbNewLine & vbNewLine & _
                C_F_Theme, _
            Buttons:=vbOKOnly + vbInformation, _
            Title:=MsgBoxTitle
    Else
        MsgBox _
            "Theme file not found. Please check the path and try again." & vbNewLine & vbNewLine & _
                C_F_Theme, _
            Buttons:=vbOKOnly + vbInformation, _
            Title:=MsgBoxTitle
    End If
End Sub

' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

' The Style.Font.Color property is oficially undocumented in Microsoft Word VBA documentation
' (https://learn.microsoft.com/en-us/office/vba/api/overview/word)
' Instead it is recommended to use the property style.Font.TextColor.RGB.
' This property provides type Long value which must be converted to RGB values.
' Example: (https://learn.microsoft.com/en-us/office/vba/api/word.wdcolor)
' Long = 15,773,696, #00B0F0, RGB = R:0 G:176 B:240
' https://www.msofficeforums.com/word-vba/18453-converting-color-codes-vba.html
' Conversion of negative long values is a Microsoft's mistery
' Example:
' GUI color: #6DB126 = 7,188,774 (dec) is shown by the property style.Font.TextColor.RGB as
' -738,131,969 (dec) = #D400FFFF
' It doesn't make any sense.

Public Function GetColorString(color As Long) As String
    Dim R As Long
    Dim G As Long
    Dim B As Long
    Dim hexColor As String
    Dim rgbColor As String
    
    ' Handle negative color values by converting to unsigned long
    If color < 0 Then
        GetColorString = color & ", hard to determine hexColor or rgbColor"
        Exit Function
    End If
    
    ' Extract the RGB components from the color value
    R = color Mod 256
    G = (color \ 256) Mod 256
    B = (color \ 256 \ 256) Mod 256
    
    ' Convert to hexadecimal format
    hexColor = "#" & Right("0" & Hex(R), 2) & Right("0" & Hex(G), 2) & Right("0" & Hex(B), 2)
    
    ' Convert to RGB format
    rgbColor = "R:" & R & " G:" & G & " B:" & B
    
    ' Combine both formats into the result string
    GetColorString = hexColor & ", " & rgbColor
End Function

Private Function GetBordersString(borders As borders) As String
    Dim border As border
    Dim bordersString As String
    
    For Each border In borders
        bordersString = bordersString & _
        "Border: " & _
        border.LineStyle & _
        ", Color: " & _
        GetColorString(border.color) & vbCrLf
    Next border
    
    GetBordersString = bordersString
End Function

Private Function GetShadingString(shading As shading) As String
    GetShadingString = "Foreground Color: " & _
    GetColorString(shading.ForegroundPatternColor) & _
    ", Background Color: " & _
    GetColorString(shading.BackgroundPatternColor)
End Function

' 2025-04-02 by ms
Public Sub InitializeColourDetails(ByRef MyVar As ColourDetails)
    MyVar.ColourValue = 0
    MyVar.ColourType = ColourTypeAutomatic
    MyVar.ThemeColorIndex = wdNotThemeColor
    MyVar.ThemeColorText = ""
    MyVar.TintAndShade = 0#
    MyVar.TintAndShadeText = ""
    MyVar.RGB = 0&
    MyVar.RGB_Hex = ""
    MyVar.Red = 0
    MyVar.Green = 0
    MyVar.Blue = 0
End Sub

' 2025-04-05 by ms and AI
Private Function RGBtoHEX(VarRGB As Long) As String
    Dim Red As String
    Dim Green As String
    Dim Blue As String
    Dim HexString As String
    Dim HexValue As String
    
    ' Convert the VarRGB value to a hexadecimal string
    HexValue = Right("000000" & Hex(VarRGB), 6)
    
    ' Extract the red, green, and blue components from the hexadecimal string
    Red = Mid(HexValue, 5, 2)
    Green = Mid(HexValue, 3, 2)
    Blue = Mid(HexValue, 1, 2)
    
    ' Concatenate the components to form the final hexadecimal color code
    HexString = "#" & Red & Green & Blue
    
    RGBtoHEX = HexString
End Function


' 2025-04-02 refactored by ms
Public Function QueryColour(ColourToTest As Long) _
                            As ColourDetails

    Dim ColourToTestHex  As String
    Dim ColourTypeByte   As Byte

    ColourToTestHex = Right$(String$(7, "0") & Hex$(ColourToTest), 8)
    ColourTypeByte = CByte("&H" & Left$(ColourToTestHex, 2))
    
    Select Case ColourTypeByte

        Case ColourTypeRGB
            QueryColour = GetRGBComponents(ColourToTest)
            QueryColour.ColourValue = ColourToTest
            QueryColour.ColourType = ColourTypeRGB
            QueryColour.RGB = ColourToTest
            QueryColour.RGB_Hex = RGBtoHEX(ColourToTest)

        Case ColourTypeAutomatic
            QueryColour.ColourType = ColourTypeAutomatic
            QueryColour.ColourValue = ColourToTest

        Case ColourTypeSystem
            QueryColour.ColourType = ColourTypeSystem
            QueryColour.ColourValue = ColourToTest

        Case ColourTypeThemeLow To ColourTypeThemeHigh
            QueryColour = QueryThemeColor(ColourTypeByte, ColourToTestHex)
            QueryColour.ColourValue = ColourToTest
            QueryColour.RGB_Hex = RGBtoHEX(QueryColour.RGB)
            QueryColour.ThemeColorText = ThemeColorName(QueryColour.ThemeColorIndex)
            QueryColour.TintAndShadeText = TintAndShadeText(QueryColour.TintAndShade)

        Case Else   ' theoretically it should not happen
            QueryColour.ColourType = ColourTypeByte

    End Select
    
End Function
 
' 2025-04-01 by ms
Private Function GetRGBComponents(ColorValue As Long) As ColourDetails
    Dim Red As Integer
    Dim Green As Integer
    Dim Blue As Integer
      
    GetRGBComponents.Red = ColorValue And &HFF
    GetRGBComponents.Green = (ColorValue \ &H100) And &HFF
    GetRGBComponents.Blue = (ColorValue \ &H10000) And &HFF
End Function

' Credits go to:
' https://www.wordarticles.com/Articles/Colours/2007.php
Private Function QueryThemeColor(ColourTypeByte As Byte, _
                                 ColourToTestHex As String) _
                                As ColourDetails

    Const Unchanged      As Byte = &HFF

    Dim LightnessByte    As Byte
    Dim DarknessByte     As Byte
    
    LightnessByte = CByte("&H" & Mid$(ColourToTestHex, 7, 2))   ' Convert to Byte format; Mid$(string, start, length) extracts a substring
    DarknessByte = CByte("&H" & Mid$(ColourToTestHex, 5, 2))    ' Convert to Byte format; Mid$(string, start, length) extracts a substring
    
    QueryThemeColor.ColourType = ColourTypeByte And &HF0
    QueryThemeColor.ThemeColorIndex = ColourTypeByte And &HF
    
    QueryThemeColor.TintAndShade = 1    ' The default value, added by ms
    If DarknessByte <> Unchanged Then
        QueryThemeColor.TintAndShade = Round(-1 + DarknessByte / &HFF, 2)
    End If
    
    If LightnessByte <> Unchanged Then
        QueryThemeColor.TintAndShade = Round(1 - LightnessByte / &HFF, 2)
    End If

    QueryThemeColor.RGB = GetRGB(QueryThemeColor.ThemeColorIndex, _
                                 QueryThemeColor.TintAndShade)
    
    QueryThemeColor.Red = QueryThemeColor.RGB And &HFF
    QueryThemeColor.Green = (QueryThemeColor.RGB \ &H100) And &HFF
    QueryThemeColor.Blue = (QueryThemeColor.RGB \ &H10000) And &HFF
    
End Function
 
' Credits go to:
' https://www.wordarticles.com/Articles/Colours/2007.php
Private Function GetRGB(ThemeColorIndex As WdThemeColorIndex, _
                        TintAndShade As Double) _
                 As String

    Dim ColorSchemeIndex    As MsoThemeColorSchemeIndex
    Dim ColorSchemeRGB      As Long
    Dim ColorSchemeHSL      As HSL
    Dim TintedAndShadedRGB  As Long

    ColorSchemeIndex = ThemeColorSchemeIndex(ThemeColorIndex)
    ColorSchemeRGB = ActiveDocument.DocumentTheme. _
                         ThemeColorScheme(ColorSchemeIndex).RGB

    ColorSchemeHSL = RGBtoHSL(ColorSchemeRGB)
    ColorSchemeHSL.L = (ColorSchemeHSL.L * (Abs(TintAndShade))) + _
                       (Abs(TintAndShade > 0) * (1 - TintAndShade))
    
    TintedAndShadedRGB = HSLtoRGB(ColorSchemeHSL)
    
    GetRGB = TintedAndShadedRGB
        
End Function
 
' Credits go to:
' https://www.wordarticles.com/Articles/Colours/2007.php
Private Function ThemeColorSchemeIndex(ThemeColorIndex As WdThemeColorIndex) _
                 As MsoThemeColorSchemeIndex

    Select Case ThemeColorIndex
        Case wdThemeColorMainDark1:         ThemeColorSchemeIndex = msoThemeDark1
        Case wdThemeColorMainLight1:        ThemeColorSchemeIndex = msoThemeLight1
        Case wdThemeColorMainDark2:         ThemeColorSchemeIndex = msoThemeDark2
        Case wdThemeColorMainLight2:        ThemeColorSchemeIndex = msoThemeLight2
        Case wdThemeColorAccent1:           ThemeColorSchemeIndex = msoThemeAccent1
        Case wdThemeColorAccent2:           ThemeColorSchemeIndex = msoThemeAccent2
        Case wdThemeColorAccent3:           ThemeColorSchemeIndex = msoThemeAccent3
        Case wdThemeColorAccent4:           ThemeColorSchemeIndex = msoThemeAccent4
        Case wdThemeColorAccent5:           ThemeColorSchemeIndex = msoThemeAccent5
        Case wdThemeColorAccent6:           ThemeColorSchemeIndex = msoThemeAccent6
        Case wdThemeColorHyperlink:         ThemeColorSchemeIndex = msoThemeHyperlink
        Case wdThemeColorHyperlinkFollowed: ThemeColorSchemeIndex = msoThemeFollowedHyperlink
        Case wdThemeColorBackground1:       ThemeColorSchemeIndex = msoThemeLight1
        Case wdThemeColorText1:             ThemeColorSchemeIndex = msoThemeDark1
        Case wdThemeColorBackground2:       ThemeColorSchemeIndex = msoThemeLight2
        Case wdThemeColorText2:             ThemeColorSchemeIndex = msoThemeDark2
        Case Else:                          ' This shouldn't really ever happen
 
    End Select
    
End Function
 
' Credits go to:
' https://www.wordarticles.com/Articles/Colours/2007.php
Private Function RGBtoHSL(RGB As Long) As HSL

    Dim R As Double ' Range 0 ? 1
    Dim G As Double ' Range 0 ? 1
    Dim B As Double ' Range 0 ? 1

    Dim RGB_Max  As Double
    Dim RGB_Min  As Double
    Dim RGB_Diff As Double

    Dim HexString As String

    HexString = Right$(String$(7, "0") & Hex$(RGB), 8)
    R = CDbl("&H" & Mid$(HexString, 7, 2)) / 255
    G = CDbl("&H" & Mid$(HexString, 5, 2)) / 255
    B = CDbl("&H" & Mid$(HexString, 3, 2)) / 255

    RGB_Max = R
    If G > RGB_Max Then RGB_Max = G
    If B > RGB_Max Then RGB_Max = B

    RGB_Min = R
    If G < RGB_Min Then RGB_Min = G
    If B < RGB_Min Then RGB_Min = B

    RGB_Diff = RGB_Max - RGB_Min

    With RGBtoHSL
    
        .L = (RGB_Max + RGB_Min) / 2

        If RGB_Diff = 0 Then
    
            .s = 0
            .H = 0
    
        Else

            Select Case RGB_Max
                Case R: .H = (1 / 6) * (G - B) / RGB_Diff - (B > G)
                Case G: .H = (1 / 6) * (B - R) / RGB_Diff + (1 / 3)
                Case B: .H = (1 / 6) * (R - G) / RGB_Diff + (2 / 3)
            End Select
    
            Select Case .L
                Case Is < 0.5: .s = RGB_Diff / (2 * .L)
                Case Else:     .s = RGB_Diff / (2 - (2 * .L))
            End Select
    
        End If

    End With
    
End Function
 
' Credits go to:
' https://www.wordarticles.com/Articles/Colours/2007.php
Private Function HSLtoRGB(HSL As HSL) As Long

    Dim R As Double
    Dim G As Double
    Dim B As Double

    Dim X As Double
    Dim Y As Double

    With HSL
    
        If .s = 0 Then
    
            R = .L
            G = .L
            B = .L
    
        Else
    
            Select Case .L
                Case Is < 0.5: X = .L * (1 + .s)
                Case Else:     X = .L + .s - (.L * .s)
            End Select
    
            Y = 2 * .L - X

            R = H2C(X, Y, IIf(.H > 2 / 3, .H - 2 / 3, .H + 1 / 3))
            G = H2C(X, Y, .H)
            B = H2C(X, Y, IIf(.H < 1 / 3, .H + 2 / 3, .H - 1 / 3))
    
        End If
    
    End With
    
    HSLtoRGB = CLng("&H00" & _
                    Right$("0" & Hex$(Round(B * 255)), 2) & _
                    Right$("0" & Hex$(Round(G * 255)), 2) & _
                    Right$("0" & Hex$(Round(R * 255)), 2))

End Function
 
' Credits go to:
' https://www.wordarticles.com/Articles/Colours/2007.php
Private Function H2C(X As Double, Y As Double, HC As Double) As Double

    Select Case HC
        Case Is < 1 / 6: H2C = Y + ((X - Y) * 6 * HC)
        Case Is < 1 / 2: H2C = X
        Case Is < 2 / 3: H2C = Y + ((X - Y) * ((2 / 3) - HC) * 6)
        Case Else:       H2C = Y
    End Select

End Function
 
' Credits go to:
' https://www.wordarticles.com/Articles/Colours/2007.php
Function ThemeColorName(ThemeColorIndex As WdThemeColorIndex, _
                        Optional LanguageId As MsoLanguageID) _
         As String

    If LanguageId = 0 Then
        LanguageId = LanguageSettings.LanguageId(msoLanguageIDUI)
    End If
    
    'msoLanguageIDEnglishUS
            
    Select Case ThemeColorIndex
        Case wdThemeColorMainDark1:         ThemeColorName = "Dark 1"
        Case wdThemeColorMainLight1:        ThemeColorName = "Light 1"
        Case wdThemeColorMainDark2:         ThemeColorName = "Dark 2"
        Case wdThemeColorMainLight2:        ThemeColorName = "Light 2"
        Case wdThemeColorAccent1:           ThemeColorName = "Accent 1"
        Case wdThemeColorAccent2:           ThemeColorName = "Accent 2"
        Case wdThemeColorAccent3:           ThemeColorName = "Accent 3"
        Case wdThemeColorAccent4:           ThemeColorName = "Accent 4"
        Case wdThemeColorAccent5:           ThemeColorName = "Accent 5"
        Case wdThemeColorAccent6:           ThemeColorName = "Accent 6"
        Case wdThemeColorHyperlink:         ThemeColorName = "Hyperlink"
        Case wdThemeColorHyperlinkFollowed: ThemeColorName = "Followed Hyperlink"
        Case wdThemeColorBackground1:       ThemeColorName = "Background 1"
        Case wdThemeColorText1:             ThemeColorName = "Text 1"
        Case wdThemeColorBackground2:       ThemeColorName = "Background 2"
        Case wdThemeColorText2:             ThemeColorName = "Text 2"
        Case Else:                          ThemeColorName = "Unknown " & ThemeColorIndex
    End Select

End Function
 
' Credits go to:
' https://www.wordarticles.com/Articles/Colours/2007.php
Private Function TintAndShadeText(TintAndShade As Double, _
                          Optional LanguageId As MsoLanguageID) _
         As String

    If LanguageId = 0 Then
        LanguageId = LanguageSettings.LanguageId(msoLanguageIDUI)
    End If
    
    Select Case TintAndShade
    
        Case 0
            TintAndShadeText = ""
            
        Case 1
            TintAndShadeText = "100 %"
            
        Case Is > 0
            TintAndShadeText = "lighter "
            TintAndShadeText = TintAndShadeText & TintAndShade * 100 & "%"
        
        Case Is < 0
            TintAndShadeText = "darker "
            TintAndShadeText = TintAndShadeText & TintAndShade * -100 & "%"
        
    End Select

End Function

' 2025-04-02 by ms
Public Function GetColourTypeName(CurrentVar As ColourType) As String
    Select Case CurrentVar
        Case ColourTypeRGB:         GetColourTypeName = "ColourTypeRGB"
        Case ColourTypeAutomatic:   GetColourTypeName = "ColourTypeAutomatic"
        Case ColourTypeSystem:      GetColourTypeName = "ColourTypeSystem"
        Case ColourTypeThemeLow:    GetColourTypeName = "ColourTypeThemeLow"
        Case ColourTypeThemeHigh:   GetColourTypeName = "ColourTypeThemeHigh"
    End Select
End Function

' Credits go to:
' https://www.wordarticles.com/Articles/Colours/2007BuildSet.php
' 2025-11-18 inserted here by ms
Public Function GetThemeColor(ThemeColorIndex As WdThemeColorIndex, _
                       TintAndShade As Double) As Long

    Const HexadecimalPrefix As String = "&H"
    Const UseThemeColor     As String = "D"
    Const UnusedZeroByte    As String = "00"
    Const Unchanged         As String = "FF"
    Dim ThemeColor          As String
    Dim LightnessOrDarkness As String
    
    ThemeColor = Hex$(ThemeColorIndex)

    If TintAndShade >= 0 Then
        LightnessOrDarkness = Unchanged & Right$("0" & Hex$((1 - TintAndShade) * &HFF), 2)
    Else
        LightnessOrDarkness = Right$("0" & Hex$((1 + TintAndShade) * &HFF), 2) & Unchanged
    End If

    GetThemeColor = CLng(HexadecimalPrefix & _
                         UseThemeColor & _
                         ThemeColor & _
                         UnusedZeroByte & _
                         LightnessOrDarkness)
End Function


