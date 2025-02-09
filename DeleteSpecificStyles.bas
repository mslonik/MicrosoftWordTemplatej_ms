Attribute VB_Name = "DeleteSpecificStyles"
' Created by Microsoft Copilot M365 by ms on 2025-01-28.
' Run this macro as the first one, before the DeleteBuiltInStyles.

Sub DeleteSpecificStyles()
    Dim doc As Document
    Dim templateDoc As Document
    Dim templatePath As String
    Dim style As style
    Dim deletedStyles As String
    Dim filePath As String
    Dim fileNum As Integer
    Dim currentDate As String
    Dim currentTime As String
    Dim docName As String
    Dim macroName As String
    Dim templateStyles As Collection
    Dim templateStyleNames As Collection
    Dim styleName As Variant
    Dim counter As Integer
    
    ' Path to the template file
    templatePath = "C:\Users\v523580\temp1\TQ-S440-en_UserDoc-3036temp.dotm"
    
    ' Path to the output text file
    filePath = "C:\Users\v523580\temp1\DeletedStyles.txt"
    
    ' Open the template file
    Set templateDoc = Documents.Open(templatePath)
    
    ' Initialize collections for template styles
    Set templateStyles = New Collection
    Set templateStyleNames = New Collection
    
    ' Loop through all styles in the template and add to collections
    For Each style In templateDoc.Styles
        If Not style.BuiltIn Then
            templateStyles.Add style
            templateStyleNames.Add style.NameLocal
        End If
    Next style
    
    ' Close the template file without saving changes
    templateDoc.Close SaveChanges:=wdDoNotSaveChanges
    
    ' Get the currently opened document
    Set doc = ActiveDocument
    
    ' Initialize the deleted styles string and counter
    deletedStyles = ""
    counter = 1
    
    ' Loop through all styles in the document
    For Each style In doc.Styles
        If Not style.BuiltIn And Not IsInCollection(templateStyleNames, style.NameLocal) Then
            deletedStyles = deletedStyles & counter & ". " & style.NameLocal & vbCrLf
            style.Delete
            counter = counter + 1
        End If
    Next style
    
    ' Get current date and time
    currentDate = Format(Date, "yyyy-mm-dd")
    currentTime = Format(Time, "hh:mm")
    
    ' Get document name and macro name
    docName = doc.Name
    macroName = "DeleteSpecificStyles"
    
    ' Write the deleted styles to the text file
    fileNum = FreeFile
    Open filePath For Output As fileNum
    Print #fileNum, "Document Name: " & docName
    Print #fileNum, "Template Name: " & templatePath
    Print #fileNum, "Macro Name: " & macroName
    Print #fileNum, "Date: " & currentDate
    Print #fileNum, "Time: " & currentTime
    Print #fileNum, vbCrLf & "Deleted Styles:" & vbCrLf & deletedStyles
    Close fileNum
    
    ' Inform the user that the styles have been deleted and logged
    MsgBox "The specified styles have been deleted and logged to " & filePath, vbInformation, "Styles Deleted"
End Sub

Function IsInCollection(coll As Collection, item As Variant) As Boolean
    Dim i As Integer
    IsInCollection = False
    For i = 1 To coll.Count
        If coll(i) = item Then
            IsInCollection = True
            Exit Function
        End If
    Next i
End Function
