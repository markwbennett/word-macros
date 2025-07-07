Attribute VB_Name = "FootnoteManager"
' =============================================================================
' FOOTNOTE MANAGER MODULE
' Version: 2.0
' Description: Footnote management and formatting macros
' =============================================================================

Option Explicit

' =============================================================================
' FOOTNOTE MANAGEMENT MACROS
' =============================================================================

' New macro to insert a footnote with specific formatting sequence
Sub InsertFormattedFootnote()
    On Error GoTo ErrorHandler
    
    ' Insert a footnote
    Selection.Footnotes.Add Range:=Selection.Range
    
    ' Backspace once
    Selection.TypeBackspace
    
    ' Select one character to the left
    Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    
    ' Turn off superscript
    Selection.Font.Superscript = False
    
    ' Move one left
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    
    ' Insert tab
    Selection.TypeText Text:=vbTab
    
    ' Move one right
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    
    ' Insert period
    Selection.TypeText Text:="."
    
    ' Insert tab
    Selection.TypeText Text:=vbTab
    
    MacroUtilities.LogAction "Inserted formatted footnote with specific sequence"
    
    Exit Sub
    
ErrorHandler:
    MacroUtilities.DisplayError Err.Number, Err.Description
End Sub

' Moves footnote text to the body of the document
Sub FN_MoveFootnotesToBody()
    On Error GoTo ErrorHandler
    
    MacroUtilities.BackupCurrentDocument
    Application.ScreenUpdating = False
    
    Dim f As Footnote
    Dim r As Range
    Dim footnoteCount As Integer
    
    footnoteCount = ActiveDocument.Footnotes.Count
    
    ' Check if there are any footnotes
    If footnoteCount = 0 Then
        MsgBox "No footnotes found in the document.", vbInformation, "No Footnotes"
        Exit Sub
    End If
    
    ' Process footnotes in reverse order to avoid index shifting issues
    For i = footnoteCount To 1 Step -1
        Set f = ActiveDocument.Footnotes(i)
        On Error Resume Next
        Set r = f.Reference
        
        ' Check for valid reference and content
        If Err.Number = 0 And Not r Is Nothing Then
            ' Get footnote text, removing any leading/trailing whitespace
            Dim fnText As String
            fnText = Trim(f.Range.Text)
            
            ' Only process if there's content
            If Len(fnText) > 0 Then
                r.Collapse wdCollapseEnd
                r.InsertAfter " " & fnText
            End If
        End If
        On Error GoTo ErrorHandler
    Next i
    
    ' Then delete all footnote references (again in reverse order)
    For i = footnoteCount To 1 Step -1
        On Error Resume Next
        ActiveDocument.Footnotes(i).Delete
        On Error GoTo ErrorHandler
    Next i
    
    Application.ScreenUpdating = True
    MacroUtilities.LogAction "Moved all footnotes to document body"
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MacroUtilities.DisplayError Err.Number, Err.Description
End Sub

' Legacy names for backward compatibility
Sub MoveFootnoteTextToBody()
    FN_MoveFootnotesToBody
End Sub

' Additional footnote management functions

Sub MoveToFootnote()
' movetofootnote Macro
' move highlighted text to new footnote, or create new footnote if no text is higlighted
' Create new footnote if no text is hightlighted:
    If Selection.Type = wdSelectionIP Then
        With Selection
            With .FootnoteOptions
                .Location = wdBottomOfPage
                .NumberingRule = wdRestartContinuous
                .StartingNumber = 1
                .NumberStyle = wdNoteNumberStyleArabic
                .LayoutColumns = 0
            End With
        .Footnotes.Add Range:=Selection.Range, Reference:=""
        End With
        Selection.MoveLeft Unit:=wdCharacter, Count:=2
        Selection.TypeText Text:=vbTab
        Selection.MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdExtend
        Selection.Font.Superscript = wdToggle
        Selection.MoveRight Unit:=wdCharacter, Count:=1
        Selection.TypeText Text:="."
        Selection.MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdExtend
        Selection.TypeText Text:=vbTab
    End If
  ' Create new footnote and move text, if text is highlighted
    If Selection.Type = wdSelectionNormal Then
        Selection.Cut
        Selection.TypeBackspace
         With Selection
             With .FootnoteOptions
                    .Location = wdBottomOfPage
                    .NumberingRule = wdRestartContinuous
                    .StartingNumber = 1
                    .NumberStyle = wdNoteNumberStyleArabic
                    .LayoutColumns = 0
                End With
                .Footnotes.Add Range:=Selection.Range, Reference:=""
        End With
        Selection.MoveLeft Unit:=wdCharacter, Count:=2
        Selection.TypeText Text:=vbTab
        Selection.MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdExtend
        Selection.Font.Superscript = wdToggle
        Selection.MoveRight Unit:=wdCharacter, Count:=1
        Selection.TypeText Text:="."
        Selection.MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdExtend
        Selection.TypeText Text:=vbTab
        Selection.PasteAndFormat (wdFormatSurroundingFormattingWithEmphasis)
    End If
End Sub

Sub ConvertFootnotes()
' ConvertFootnotes Macro
' Convert footnote references to non-superscript, and add period after.
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^f "
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute
    Selection.Font.Superscript = wdToggle
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.TypeText Text:="."
End Sub

Sub FixFN()
' FixFN Macro
    Selection.TypeText Text:=vbTab
    Selection.MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    Selection.Font.Superscript = wdToggle
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.TypeText Text:="."
    Selection.MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    Selection.TypeText Text:=vbTab
End Sub

Sub CertInsertFootnote()
' Insert footnote, and Convert number to better format.
    If Selection.Type = wdSelectionIP Then
        With Selection
                With .FootnoteOptions
                .Location = wdBottomOfPage
                .NumberStyle = wdNoteNumberStyleArabic
                .LayoutColumns = 0
        End With
        .Footnotes.Add Range:=Selection.Range, Reference:=""
        End With
        Selection.MoveLeft Unit:=wdCharacter, Count:=2, Extend:=wdExtend
        Selection.Font.Superscript = wdToggle
        Selection.MoveRight Unit:=wdCharacter, Count:=1
        Selection.MoveLeft Unit:=wdCharacter, Count:=1
        Selection.TypeText Text:="."
        Selection.MoveRight Unit:=wdCharacter, Count:=1
    End If
  ' Create new footnote and move text, if text is highlighted
    If Selection.Type = wdSelectionNormal Then
        Selection.Cut
        Selection.TypeBackspace
         With Selection
             With .FootnoteOptions
                    .Location = wdBottomOfPage
                    .NumberingRule = wdRestartContinuous
                    .StartingNumber = 1
                    .NumberStyle = wdNoteNumberStyleArabic
                    .LayoutColumns = 0
                End With
                .Footnotes.Add Range:=Selection.Range, Reference:=""
        End With
        Selection.MoveLeft Unit:=wdCharacter, Count:=2, Extend:=wdExtend
        Selection.Font.Superscript = wdToggle
        Selection.MoveRight Unit:=wdCharacter, Count:=1
        Selection.MoveLeft Unit:=wdCharacter, Count:=1
        Selection.TypeText Text:="."
        Selection.MoveRight Unit:=wdCharacter, Count:=1
        Selection.PasteAndFormat (wdFormatSurroundingFormattingWithEmphasis)
    End If
End Sub

Sub Demo()
Application.ScreenUpdating = False
Dim i As Long, RngNt As Range, RngTxt As Range
With ActiveDocument
  For i = .Footnotes.Count To 1 Step -1
    With .Footnotes(i)
      Set RngNt = .Range
      With RngNt
        .End = .End
        .Start = .Start + 2
      End With
      Set RngTxt = .Reference
      With RngTxt
         .Collapse wdCollapseEnd
         .Collapse wdCollapseStart
        .FormattedText = RngNt.FormattedText
      End With
      .Delete
    End With
  Next
End With
Application.ScreenUpdating = True
End Sub 