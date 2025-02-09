Attribute VB_Name = "InsertPictures"
' Inserts PNG files from the specified folder.
' Microsoft Copilot M365 and ms on 2025-02-07

Sub InsertPNGFilesIntoCanvases()
    Dim folderPath As String
    Dim pngFiles As Collection
    Dim file As Variant
    Dim doc As Document
    Dim canvasShape As shape
    Dim pictureShape As shape
    Dim rng As Range
    Dim totalFiles As Integer
    Dim estimatedTime As Double
    Dim processingTime As Double
    
    ' Initialize
    Set pngFiles = New Collection
    Set doc = ActiveDocument
    
    ' Get the current cursor position
    Set rng = Selection.Range
    
    ' Open folder selection dialog
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select Folder Containing PNG Files"
        If .Show = -1 Then
            folderPath = .SelectedItems(1)
        Else
            MsgBox "No folder selected. Macro terminated.", vbExclamation
            Exit Sub
        End If
    End With
    
    ' Get all PNG files in the folder
    Set pngFiles = GetPNGFilesInFolder(folderPath)
    
    ' Check if any PNG files were found
    totalFiles = pngFiles.Count
    If totalFiles = 0 Then
        MsgBox "No PNG files found in the selected folder.", vbExclamation
        Exit Sub
    End If
    
    ' Calculate estimated time
    estimatedTime = totalFiles * 2
    processingTime = totalFiles * 0.5
    totalTime = estimatedTime + processingTime
    
    ' Show message box with the number of files and estimated time
    MsgBox "Number of PNG files found: " & totalFiles & vbCrLf & _
           "Estimated time of insertion: " & estimatedTime & " seconds" & vbCrLf & _
           "Processing time: " & processingTime & " seconds" & vbCrLf & _
           "Total time: " & totalTime & " seconds", vbInformation, "PNG Files Found"
    
    ' Insert each PNG file into a separate canvas with an empty paragraph in between
    For Each file In pngFiles
        ' Insert empty paragraph and format with style "Normal ms"
        rng.Collapse Direction:=wdCollapseEnd
        rng.InsertParagraphAfter
        rng.style = "Normal ms"
        rng.Collapse Direction:=wdCollapseEnd
        
        ' Insert next empty paragraph and format with style "PictureCanva ms"
        rng.InsertParagraphAfter
        rng.style = "PictureCanva ms"
        rng.Collapse Direction:=wdCollapseEnd
        
        ' Move back (up) to the empty paragraph formatted with style "PictureCanva ms"
        rng.MoveStart Unit:=wdParagraph, Count:=-1
        rng.Select
        
        ' Add a new canvas
        Set canvasShape = doc.Shapes.AddCanvas(0, 0, 500, 500)
        
        ' Set Format Drawing Canvas Fill to "No fill"
        canvasShape.Fill.Transparency = 1#
        
        ' Add the picture to the canvas
        Set pictureShape = canvasShape.CanvasItems.AddPicture(folderPath & "\" & file)
        
        ' Set Layout Option to 'With Text Wrapping' and 'In Line With Text'
        canvasShape.WrapFormat.Type = wdWrapInline
        
        ' Delay to ensure proper insertion
        Dim startTime As Single
        startTime = Timer
        Do While Timer < startTime + 2
            DoEvents
        Loop
    Next file
    
    ' Get macro name and template name
    Dim macroName As String
    Dim templateName As String
    macroName = "InsertPNGFilesIntoCanvases"
    templateName = doc.AttachedTemplate.Name
    
    ' Show message box indicating successful completion
    MsgBox "Macro '" & macroName & "' completed successfully." & vbCrLf & _
           "Template: " & templateName, vbInformation, "Processing Complete"
End Sub

Function GetPNGFilesInFolder(folderPath As String) As Collection
    Dim pngFiles As Collection
    Dim fileName As String
    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    
    ' Initialize
    Set pngFiles = New Collection
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Check if the folder exists
    If Not fso.FolderExists(folderPath) Then
        MsgBox "The specified folder does not exist.", vbExclamation
        Exit Function
    End If
    
    ' Get all PNG files in the folder
    Set folder = fso.GetFolder(folderPath)
    For Each file In folder.Files
        If LCase(fso.GetExtensionName(file.Name)) = "png" Then
            pngFiles.Add file.Name
        End If
    Next file
    
    ' Return the collection of PNG files
    Set GetPNGFilesInFolder = pngFiles
End Function


