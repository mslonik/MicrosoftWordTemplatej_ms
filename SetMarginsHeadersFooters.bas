Attribute VB_Name = "SetMarginsHeadersFooters"
'Prepared on 2025-02-02 by Microsoft Copilot M365 and ms

Sub SetMarginsHeadersFooters()
     With ActiveDocument.PageSetup
        .TopMargin = CentimetersToPoints(1.2)
        .BottomMargin = CentimetersToPoints(1.2)
        .LeftMargin = CentimetersToPoints(2.2) ' This sets the inside margin
        .RightMargin = CentimetersToPoints(1.2) ' This sets the outside margin
        .Orientation = wdOrientPortrait
        .MirrorMargins = True
        .PaperSize = wdPaperA4
        .HeaderDistance = CentimetersToPoints(1)
        .FooterDistance = CentimetersToPoints(1)
    End With
End Sub
