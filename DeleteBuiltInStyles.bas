Attribute VB_Name = "DeleteBuiltInStyles"
    ' Created by Microsoft Copilot M365 by ms on 2025-01-28.
    ' Run this macro as the second one, after the DeleteSpecificStyles.
   
Sub DeleteBuiltInStyles()
    Dim doc As Document
    Dim templateDoc As template
    Dim templatePath As String
    Dim style As style
    Dim deletedStyles As String
    Dim notDeletedStyles As String
    Dim filePath As String
    Dim fileNum As Integer
    Dim counter As Integer
    Dim errorCounter As Integer
    Dim currentDate As String
    Dim currentTime As String
    Dim docName As String
    Dim macroName As String
    
    ' Get the currently active document
    Set doc = ActiveDocument
    
    ' Get the template attached to the document
    Set templateDoc = doc.AttachedTemplate
    templatePath = templateDoc.FullName
    
    ' Initialize the deleted styles string and counters
    deletedStyles = "Deleted Styles:" & vbCrLf
    notDeletedStyles = "Not Deleted Styles:" & vbCrLf
    counter = 1
    errorCounter = 1
    
    ' Loop through all styles in the document
    For Each style In doc.Styles
        If style.BuiltIn Then
            On Error Resume Next
            style.Delete
            If Err.Number = 0 Then
                deletedStyles = deletedStyles & counter & ". " & style.NameLocal & vbCrLf
                counter = counter + 1
            Else
                notDeletedStyles = notDeletedStyles & errorCounter & ". " & style.NameLocal & vbCrLf
                errorCounter = errorCounter + 1
            End If
            On Error GoTo 0
        End If
    Next style
    
    ' Get current date and time
    currentDate = Format(Date, "yyyy-mm-dd")
    currentTime = Format(Time, "hh:mm")
    
    ' Get document name and macro name
    docName = doc.Name
    macroName = "DeleteSpecificStyles"
    
    ' Get the path to the folder of the currently active document
    filePath = doc.Path & "\DeletedBuiltinStyles.txt"
    
    ' Write the deleted and not deleted styles to the text file
    fileNum = FreeFile
    Open filePath For Output As fileNum
    Print #fileNum, "Document Name: " & docName
    Print #fileNum, "Template Name: " & templatePath
    Print #fileNum, "Macro Name: " & macroName
    Print #fileNum, "Date: " & currentDate
    Print #fileNum, "Time: " & currentTime
    Print #fileNum, vbCrLf & deletedStyles
    Print #fileNum, vbCrLf & notDeletedStyles
    Close fileNum
    
    ' Inform the user that the built-in styles have been deleted and logged
    MsgBox "The built-in styles have been deleted and logged to " & filePath, vbInformation, "Styles Deleted"
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
