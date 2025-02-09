Attribute VB_Name = "ListNonBuiltInStyles"
' Created by Microsoft Copilot M365 by ms on 2025-01-28.
' Run this macro as the first one, before the DeleteBuiltInStyles.

Sub ListNonBuiltInAndSuffixStylesInTemplate()
    Dim doc As Document
    Dim template As template
    Dim templatePath As String
    Dim style As style
    Dim styleInfo As String
    Dim filePath As String
    Dim fileNum As Integer
    Dim rowNum As Integer
    Dim paragraphStyles As Collection
    Dim characterStyles As Collection
    Dim tableStyles As Collection
    Dim listStyles As Collection
    Dim styleName As Variant
    Dim styleType As Variant
    Dim styleBuiltIn As Variant
    Dim currentDate As String
    Dim currentTime As String
    Dim docName As String
    Dim macroName As String
    
    ' Get the currently active document
    Set doc = ActiveDocument
    
    ' Get the template attached to the document
    Set template = doc.AttachedTemplate
    templatePath = template.FullName
    
    ' Path to the output text file
    filePath = doc.Path & "\TemplateStyleList.txt"
    
    ' Open the template file
    Set doc = Documents.Open(templatePath)
    
    ' Initialize the style information string
    styleInfo = "No. | Style Name | Type | Built-in" & vbCrLf
    rowNum = 1
    
    ' Initialize collections for each style type
    Set paragraphStyles = New Collection
    Set characterStyles = New Collection
    Set tableStyles = New Collection
    Set listStyles = New Collection
    
    ' Loop through all styles in the template
    For Each style In doc.Styles
        If Not style.BuiltIn Or InStr(style.NameLocal, " ms") > 0 Then
            styleInfo = styleInfo & rowNum & " | " & _
                        style.NameLocal & " | " & _
                        StyleTypeName(style.Type) & " | " & _
                        CStr(style.BuiltIn) & vbCrLf
            rowNum = rowNum + 1
            
            ' Add style names to respective collections
            Select Case style.Type
                Case wdStyleTypeParagraph
                    paragraphStyles.Add Array(style.NameLocal, StyleTypeName(style.Type), CStr(style.BuiltIn))
                Case wdStyleTypeCharacter
                    characterStyles.Add Array(style.NameLocal, StyleTypeName(style.Type), CStr(style.BuiltIn))
                Case wdStyleTypeTable
                    tableStyles.Add Array(style.NameLocal, StyleTypeName(style.Type), CStr(style.BuiltIn))
                Case wdStyleTypeList
                    listStyles.Add Array(style.NameLocal, StyleTypeName(style.Type), CStr(style.BuiltIn))
            End Select
        End If
    Next style
    
    ' Close the template file without saving changes
    doc.Close SaveChanges:=wdDoNotSaveChanges
    
    ' Sort and add paragraph styles to styleInfo
    styleInfo = styleInfo & vbCrLf & "Paragraph Styles:" & vbCrLf
    Call SortCollection(paragraphStyles)
    rowNum = 1
    For Each styleName In paragraphStyles
        styleInfo = styleInfo & rowNum & " | " & styleName(0) & " | " & styleName(1) & " | " & styleName(2) & vbCrLf
        rowNum = rowNum + 1
    Next styleName
    
    ' Sort and add character styles to styleInfo
    styleInfo = styleInfo & vbCrLf & "Character Styles:" & vbCrLf
    Call SortCollection(characterStyles)
    rowNum = 1
    For Each styleName In characterStyles
        styleInfo = styleInfo & rowNum & " | " & styleName(0) & " | " & styleName(1) & " | " & styleName(2) & vbCrLf
        rowNum = rowNum + 1
    Next styleName
    
    ' Sort and add table styles to styleInfo
    styleInfo = styleInfo & vbCrLf & "Table Styles:" & vbCrLf
    Call SortCollection(tableStyles)
    rowNum = 1
    For Each styleName In tableStyles
        styleInfo = styleInfo & rowNum & " | " & styleName(0) & " | " & styleName(1) & " | " & styleName(2) & vbCrLf
        rowNum = rowNum + 1
    Next styleName
    
    ' Sort and add list styles to styleInfo
    styleInfo = styleInfo & vbCrLf & "List Styles:" & vbCrLf
    Call SortCollection(listStyles)
    rowNum = 1
    For Each styleName In listStyles
        styleInfo = styleInfo & rowNum & " | " & styleName(0) & " | " & styleName(1) & " | " & styleName(2) & vbCrLf
        rowNum = rowNum + 1
    Next styleName
    
    ' Get current date and time
    currentDate = Format(Date, "yyyy-mm-dd")
    currentTime = Format(Time, "hh:mm")
    
    ' Get document name and macro name
    docName = doc.Name
    macroName = "DeleteSpecificStyles"
    
    ' Write the style information to the text file
    fileNum = FreeFile
    Open filePath For Output As fileNum
    Print #fileNum, "Document Name: " & docName
    Print #fileNum, "Template Name: " & templatePath
    Print #fileNum, "Macro Name: " & macroName
    Print #fileNum, "Date: " & currentDate
    Print #fileNum, "Time: " & currentTime
    Print #fileNum, vbCrLf & styleInfo
    Close fileNum
    
    ' Inform the user that the styles have been logged
    MsgBox "The non-built-in and suffix ' ms' styles have been logged to " & filePath, vbInformation, "Styles Logged"
End Sub

Function StyleTypeName(styleType As WdStyleType) As String
    Select Case styleType
        Case wdStyleTypeParagraph
            StyleTypeName = "Paragraph"
        Case wdStyleTypeCharacter
            StyleTypeName = "Character"
        Case wdStyleTypeTable
            StyleTypeName = "Table"
        Case wdStyleTypeList
            StyleTypeName = "List"
        Case Else
            StyleTypeName = "Unknown"
    End Select
End Function

Sub SortCollection(ByRef coll As Collection)
    Dim i As Integer, j As Integer
    Dim temp As Variant
    
    ' Simple bubble sort
    For i = 1 To coll.Count - 1
        For j = i + 1 To coll.Count
            If coll(i)(0) > coll(j)(0) Then
                temp = coll(i)
                coll.Add coll(j), Before:=i
                coll.Add temp, Before:=j
                coll.Remove i + 1
                coll.Remove j + 1
            End If
        Next j
    Next i
End Sub

