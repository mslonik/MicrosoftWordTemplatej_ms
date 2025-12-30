Attribute VB_Name = "SeqFieldList"
' Source: https://gregmaxey.com/word_tip_pages/seq_field_numbering.html
' Credits and contributions go to Mr. Gregory Maxey.
' License: unknown (not specified by Mr. Gregory Maxey).
'
'+-----+----------------------+-------------------------+
'| No. | Sub name             | Ribbon name             |
'+-----+----------------------+-------------------------+
'| 1   | StartExtendAddInList | Start/Extend/AddIn List |
'| 2   | StartOrRefreshList   | Renumber/Refresh List   |
'| 3   | InsertInList         | Split List Item         |
'| 4   | DeleteListItem       | Delete List #           |
'| 5   | DeleteListNumber     | Delete List #/Text      |
'+-----+----------------------+-------------------------+
'
' Proposed keyboard shortcut: Alt+Enter starts a new list or creates the next sequential number in a list.
' Table made thanks to https://tableconvert.com/ascii-generator
' 2025-12-27 by ms
'
Option Explicit
Public bInsertIn As Boolean
Public bClipEnd As Boolean
Sub StartExtendAddInList()
'Coding and testing by Greg Maxey and Graham Mayor
Dim oRng As Range
Dim oRngDummy As Range
Dim oSelRng As Range
Dim i As Long
Application.ScreenUpdating = False
If Len(Selection) > 1 Then
  If Not Selection.Characters.Last = Chr$(13) Then
    With Selection
      .MoveEndUntil Cset:=Chr(13), Count:=wdForward
      .MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    End With
  End If
End If
Set oSelRng = Selection.Range
On Error GoTo Handler
If InStr(Selection.Fields(1).Code, "SEQ") = 0 Then
  StartOrRefreshList
End If
On Error GoTo 0
If oSelRng.End = ActiveDocument.Range.End Then
  oSelRng.InsertAfter vbCr
  bClipEnd = True
End If
If Len(oSelRng) = 0 Then
  AddToList
  Exit Sub
End If
For i = 1 To oSelRng.Paragraphs.Count
  Set oRng = oSelRng.Paragraphs(i).Range
  If Len(oRng.Text) = 1 Then
    GoTo Skip
  End If
  If Selection.End = ActiveDocument.Range.End - 1 Then
    Selection.InsertAfter vbCr
    bClipEnd = True
  End If
  'Delete existing SEQ field if present
  If oRng.Characters(1).Fields.Count > 0 Then
    If oRng.Fields(1).Type = wdFieldSequence Then
      oRng.Fields(1).Delete
      'Delete the period and tab
      oRng.End = oRng.Characters(2).End
      oRng.Delete
    End If
  End If
  'Insert new SEQ field
  oRng.Collapse
  'Create a dummy range object.  See notes at end of project.
  Set oRngDummy = oRng.Duplicate
  oRngDummy.Move Unit:=wdCharacter, Count:=1
  oRng.Fields.Add Range:=oRng, _
    Type:=wdFieldEmpty, Text:="SEQ numberedlist", _
    PreserveFormatting:=False
  Set oRng = oRngDummy.Duplicate
  oRng.Move Unit:=wdCharacter, Count:=-1
  'Set paragraph indents as desired for list.
  With oRng.ParagraphFormat
    .LeftIndent = InchesToPoints(0.25)
    .FirstLineIndent = InchesToPoints(-0.25)
  End With
  oRng.InsertAfter "." & vbTab
Skip:
Next i
oSelRng.Select
Update
Selection.Collapse Direction:=wdCollapseEnd
If bClipEnd Then ActiveDocument.Characters.Last.Delete
Application.ScreenUpdating = True
Exit Sub
Handler:
Err.Clear
If Selection.Paragraphs.Count > 1 Or Len(Selection.Paragraphs(1).Range) = 1 Then
  StartOrRefreshList
Else
  AddToList
End If
End Sub
Sub StartOrRefreshList()
'Adapted by Greg Maxey from original code provided by Doug Robbins,
'Dave Rado and Bill Coan
Dim oRng As Range
Dim oRngDummy As Range
Dim oSelRng As Range
Dim i As Long
Dim oStartNum As Long
On Error GoTo Handler
oStartNum = CLng(InputBox("Type the starting number", "Start", 1))
If Len(Selection.Text) = 1 Then
  StartNewList oStartNum
  Exit Sub
End If
Application.ScreenUpdating = False
Set oSelRng = Selection.Range
If oSelRng.End = ActiveDocument.Range.End Then
  oSelRng.InsertAfter vbCr
  bClipEnd = True
End If
For i = oStartNum To oSelRng.Paragraphs.Count + oStartNum - 1
  Set oRng = oSelRng.Paragraphs(i - oStartNum + 1).Range
  'Skip empty paragraphs (shame on you!!)
  If Len(oRng.Text) = 1 Then
    GoTo Skip
  End If
  'Delete existing SEQ field if present
  Set oRng = oSelRng.Paragraphs(i - oStartNum + 1).Range
  If oRng.Characters(1).Fields.Count > 0 Then
    If oRng.Fields(1).Type = wdFieldSequence Then
      oRng.Fields(1).Delete
      'Delete the period and tab
      oRng.End = oRng.Characters(2).End
      oRng.Delete
    End If
  End If
  'Insert new SEQ field
  oRng.Collapse
  'Create a dummy range object.  See "SUB Notes" at end of project.
  Set oRngDummy = oRng.Duplicate
  oRngDummy.Move Unit:=wdCharacter, Count:=1
  If i = oStartNum Then
    oRng.Fields.Add Range:=oRng, _
      Type:=wdFieldEmpty, Text:="SEQ numberedlist\r" & i, _
      PreserveFormatting:=False
  Else
    oRng.Fields.Add Range:=oRng, _
      Type:=wdFieldEmpty, Text:="SEQ numberedlist", _
      PreserveFormatting:=False
  End If
  Set oRng = oRngDummy.Duplicate
  oRng.Move Unit:=wdCharacter, Count:=-1
  'Set paragraph indents as desired for list.
  With oRng.ParagraphFormat
    .LeftIndent = InchesToPoints(0.25)
    .FirstLineIndent = InchesToPoints(-0.25)
  End With
  oRng.InsertAfter "." & vbTab
Skip:
Next i
oSelRng.Select
Update
Selection.Collapse Direction:=wdCollapseEnd
If bClipEnd Then ActiveDocument.Characters.Last.Delete
Application.ScreenUpdating = True
Exit Sub
Handler:
Err.Clear
End Sub

' Added "Private".
' 2025-12-30 by ms
Private Sub ListBuilder(ByVal bInsertIn As Boolean)
Dim oSpot As Range
Dim i As Long
Dim j As Long
Dim k As Long
Dim oNumLegth As Long

Selection.Collapse Direction:=wdCollapseStart
Set oSpot = Selection.Range
i = Selection.Range.Start
Selection.Expand Unit:=wdParagraph
On Error GoTo Handler:
If InStr(Selection.Fields(1).Code, "\r1") > 0 Then
 k = 3
Else
 k = 0
End If
oNumLegth = Len(Selection.Fields(1).Result)
j = Selection.Range.Start
oSpot.Select
If j = Selection.Range.Start Then
  Selection.Move Unit:=wdCharacter, Count:=2 + oNumLegth
  Set oSpot = Selection.Range
  GoTo Procede
End If
If j + 21 + oNumLegth + k = Selection.Range.Start Then
  Selection.Move Unit:=wdCharacter, Count:=2
  Set oSpot = Selection.Range
  GoTo Procede
End If
If j + 22 + oNumLegth + k = Selection.Range.Start Then
  Selection.Move Unit:=wdCharacter, Count:=1
  Set oSpot = Selection.Range
End If
Procede:
Selection.InsertBefore vbCr
Selection.Collapse Direction:=wdCollapseEnd
Selection.Fields.Add Range:=Selection.Range, _
     Type:=wdFieldEmpty, Text:="SEQ numberedlist", _
     PreserveFormatting:=False
With Selection.ParagraphFormat
  .LeftIndent = InchesToPoints(0.25)
  .FirstLineIndent = InchesToPoints(-0.25)
End With
Selection.InsertAfter "." & vbTab
Selection.Collapse Direction:=wdCollapseEnd
If bInsertIn Then
  oSpot.Select
End If
Update
Exit Sub
Handler:
StartOrRefreshList
End Sub

' Added "private"
' 2025-12-30 by ms
Private Sub AddToList()
bClipEnd = False
Selection.Expand wdParagraph
Selection.Collapse wdCollapseEnd
If Len(Selection.Paragraphs(1).Range.Text) > 1 And _
  Selection.End = ActiveDocument.Range.End - 1 Then
  Selection.InsertAfter vbCr
  bClipEnd = True
End If
''''
Selection.Move wdCharacter, -1
bInsertIn = False
ListBuilder bInsertIn
If bClipEnd Then ActiveDocument.Characters.Last.Delete
End Sub
Sub InsertInList()
Dim bInsertIn As Boolean
If Len(Selection.Text) > 1 Then Exit Sub
  bInsertIn = True
  ListBuilder bInsertIn
  bInsertIn = False
End Sub

' Added "private"
' 2025-12-30 by ms
Private Sub StartNewList(ByVal oStartNum As Long)
'Code by Greg Maxey
Dim oFld As Field
Selection.Expand Unit:=wdParagraph
For Each oFld In Selection.Fields
  If oFld.Code.Text = " SEQ numberedlist " Then
    oFld.Code.Text = " SEQ numberedlist\r" & oStartNum & " "
    Update
    Exit Sub
  Else
    Selection.Collapse Direction:=wdCollapseStart
    MsgBox "A new list is already started at this location." _
           & " To refresh the list, selected all list items" _
           & " including the final paragraph mark.", _
           vbOKOnly + vbInformation, "Information"
    Exit Sub
  End If
Next
Selection.Collapse Direction:=wdCollapseStart
Selection.Fields.Add Range:=Selection.Range, _
          Type:=wdFieldEmpty, Text:="SEQ numberedlist\r" & oStartNum, _
                PreserveFormatting:=False
With Selection.ParagraphFormat
  .LeftIndent = InchesToPoints(0.25)
  .FirstLineIndent = InchesToPoints(-0.25)
End With
Selection.InsertAfter "." & vbTab
Selection.Collapse Direction:=wdCollapseEnd
Update
End Sub
Sub DeleteListItem()
'Code by Greg Maxey
Dim i As Long
Dim oPara As Paragraph
For i = 1 To Selection.Range.Paragraphs.Count
  Selection.Paragraphs(1).Range.Select
  On Error GoTo Handler
  If InStr(Selection.Fields(1).Code, "\r") Then
    Selection.Fields(1).Next.Code.Text = " SEQ numberedlist\r" _
      & Right(Selection.Fields(1).Code, 2)
ProcessAnyway:
    Selection.Delete
    Update
  Else
    Selection.Delete
    Update
  End If
Next
Exit Sub
Handler:
If Err.Number = 5941 Then
  If MsgBox("The selected paragraph is not a numbered list" _
    & " item. Do you want to delete the paragraph anyway?", _
    vbYesNo + vbInformation, "Attention") = vbYes Then
    Selection.Delete
  Else
    Selection.Collapse Direction:=wdCollapseStart
  End If
ElseIf Err.Number = 91 Then
  Resume ProcessAnyway
End If
End Sub
Sub DeleteListNumber()
'Code by Greg Maxey
Dim oPara As Paragraph
For Each oPara In Selection.Range.Paragraphs
  oPara.Range.Select
  On Error GoTo Handler
  If InStr(Selection.Fields(1).Code, "\r") Then
    Selection.Fields(1).Next.Code.Text = " SEQ numberedlist\r" _
      & Right(Selection.Fields(1).Code, 2)
ProcessAnyway:
    Selection.Collapse wdCollapseStart
    Selection.MoveEndUntil Chr(9)
    Selection.MoveEnd wdCharacter, 1
    Selection.Delete
    Update
  Else
    Selection.Collapse wdCollapseStart
    Selection.MoveEndUntil Chr(9)
    Selection.MoveEnd wdCharacter, 1
    Selection.Delete
    Update
  End If
Next
Exit Sub
Handler:
If Err.Number = 5941 Then
  If MsgBox("The selected paragraph is not a numbered list" _
    & " item. Do you want to delete the paragraph anyway?", _
    vbYesNo + vbInformation, "Attention") = vbYes Then
    Selection.Delete
  Else
    Selection.Collapse Direction:=wdCollapseStart
  End If
ElseIf Err.Number = 91 Then
  Resume ProcessAnyway
End If
End Sub

' Added "private"
' 2025-12-30 by ms
Private Sub Update()
Dim oFld As Field
Dim oRngUpdate As Range
Set oRngUpdate = Selection.Range
On Error Resume Next
oRngUpdate.MoveStartUntil Cset:=Chr(13), Count:=wdBackward
On Error GoTo 0
oRngUpdate.End = Selection.Sections(1).Range.End
oRngUpdate.Fields.Update
End Sub

' Added "private"
' 2025-12-30 by ms
Private Sub Notes()
'A dummy range is created and moved to the right one character so that it will be
'located AFTER the field that is added. This is necessary because when you add a
'field to a range, the range (unlike selection) ends up at the start of the field!!!
'this is one of the unfortunate "features" of ranges.  After the field is added, we
'then set oRng to the oRngDummy regaining our place ;-)
End Sub
