Attribute VB_Name = "ExportBuildingBlocks"
'Word VBA reference: https://learn.microsoft.com/en-us/office/vba/api/overview/word

Sub InsertBuildingBlocks()
    Dim bb As buildingBlock ' Corrected to BuildingBlock with capital letters
    Dim newDoc As Document
    Dim templateDoc As Document
    Dim rng As Range
    Dim templateName As String
    Dim savePath As String
    Dim bbe As BuildingBlockEntries
    Dim i As Integer

    ' Get the current template document
    Set templateDoc = ActiveDocument.AttachedTemplate.OpenAsDocument

    ' Create a new blank document
    Set newDoc = Documents.Add

    ' Attach the "ms.dotm" template to the new document
    newDoc.AttachedTemplate = templateDoc.FullName

    ' Get the template file name without extension
    templateName = Left(templateDoc.Name, InStrRev(templateDoc.Name, ".") - 1)

    ' Get the default local file location
    savePath = Options.DefaultFilePath(wdDocumentsPath) & "\" & templateName & "_BuildingBlocks_Content.docx"

    ' Get the BuildingBlockEntries collection from the template
    Set bbe = templateDoc.AttachedTemplate.BuildingBlockEntries

    For i = 1 To bbe.Count
        Set bb = bbe(i)
        ' Insert each building block into the new document
        Set rng = newDoc.Content
        rng.Collapse Direction:=wdCollapseEnd
        bb.Insert Where:=rng, RichText:=True
        rng.InsertParagraphAfter
    Next i

    ' Save the new document
    newDoc.SaveAs2 fileName:=savePath

    ' Close the template document without saving changes
    templateDoc.Close SaveChanges:=wdDoNotSaveChanges

    ' Inform the user
    MsgBox "All building blocks have been inserted into the new document and saved at: " & savePath
End Sub
