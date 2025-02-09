Attribute VB_Name = "SaveDocumentAsPDFWithSettings"
' Written by Maciej Slojewski @ 2024-02-26 in order to facilitate saving of DOCX files in PDF files.

Sub SaveDocumentAsPDFWithSettings()

    Dim filePath As String
    Dim currentDoc As Document
    
    Set currentDoc = ActiveDocument
    ' Define the file path and name for the PDF file
    ' This example uses the document's name and saves the PDF in the same directory
    filePath = currentDoc.Path & "\" & Left(currentDoc.Name, InStrRev(currentDoc.Name, ".") - 1) & ".pdf"
    
    ' Check if the document is saved (it needs a file path)
    If currentDoc.Path = "" Then
        MsgBox "Please save the document before exporting as PDF."
        Exit Sub
    End If
    
    ' Export the document as PDF with specified settings
    ' Note: Some specific settings mentioned cannot be set via VBA and might require manual adjustment
    currentDoc.ExportAsFixedFormat OutputFileName:=filePath, _
                                    ExportFormat:=wdExportFormatPDF, _
                                    OpenAfterExport:=False, _
                                    OptimizeFor:=wdExportOptimizeForPrint, _
                                    Range:=wdExportAllDocument, _
                                    From:=1, _
                                    To:=1, _
                                    item:=wdExportDocumentContent, _
                                    IncludeDocProps:=False, _
                                    KeepIRM:=True, _
                                    CreateBookmarks:=wdExportCreateHeadingBookmarks, _
                                    DocStructureTags:=False, _
                                    BitmapMissingFonts:=True, _
                                    UseISO19005_1:=False
    
    MsgBox "Document exported as PDF: " & filePath
End Sub
