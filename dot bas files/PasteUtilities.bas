Attribute VB_Name = "PasteUtilities"
' =============================================================================
' PASTE UTILITIES MODULE
' Version: 2.1
' Description: Paste-related macros and utilities
' =============================================================================

Option Explicit

' =============================================================================
' PASTE UTILITIES
' =============================================================================

' Paste plain text while cleaning nonbreaking spaces and multiple newlines
' Usage: Assign to a keyboard shortcut or ribbon button for quick access
Sub PastePlainTextClean()
    On Error GoTo ErrorHandler
    
    ' Check if we're in a footnote or endnote
    Dim isInFootnote As Boolean
    isInFootnote = (Selection.Information(wdInFootnote) Or Selection.Information(wdInEndnote))
    
    ' Declare and assign start position of the pasted content
    Dim rngFrom As Long, rngTo As Long
    rngFrom = Selection.Start
    
    ' Paste as unformatted text using the same method as the working example
    On Error Resume Next
    Selection.PasteAndFormat (wdFormatSurroundingFormattingWithEmphasis)
    
    ' Check if paste was successful
    If Err.Number <> 0 Then
        On Error GoTo ErrorHandler
        MsgBox "No text content in clipboard to paste.", vbInformation, "Paste Clean Text"
        Exit Sub
    End If
    
    On Error GoTo ErrorHandler
    
    ' Get the range of pasted content
    rngTo = Selection.End
    
    ' Handle selection differently for footnotes vs main document
    If isInFootnote Then
        ' In footnotes, just extend the selection to cover pasted content
        Selection.SetRange rngFrom, rngTo
    Else
        ' In main document, use the Range method
        ActiveDocument.Range(rngFrom, rngTo).Select
    End If
    
    ' Convert nonbreaking spaces (Chr(160)) to regular spaces (Chr(32))
    With Selection.Find
        .Text = ChrW(160)
        .Replacement.Text = Chr(32)
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=wdReplaceAll
    End With
    
    ' Convert line feeds (manual line breaks) to paragraph marks
    With Selection.Find
        .Text = "^l"
        .Replacement.Text = "^p"
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=wdReplaceAll
    End With
    
    ' Collapse all series of whitespace that include \n to \n
    ' First remove spaces/tabs around paragraph marks
    With Selection.Find
        .Text = "[ ^t^s]^p"
        .Replacement.Text = "^p"
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=wdReplaceAll
    End With
    
    With Selection.Find
        .Text = "^p[ ^t^s]"
        .Replacement.Text = "^p"
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=wdReplaceAll
    End With
    
    ' Then collapse multiple paragraph marks
    Dim i As Integer
    For i = 1 To 5
        With Selection.Find
            .Text = "^p^p"
            .Replacement.Text = "^p"
            .Forward = True
            .Wrap = wdFindStop
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
            .Execute Replace:=wdReplaceAll
        End With
    Next i
    
    ' Collapse all series of whitespace that do not include \n to single space
    ' This handles: multiple spaces/tabs = single space
    With Selection.Find
        .Text = "[ ^t^s]{2,}"
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = True
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=wdReplaceAll
    End With
    
    ' Collapse selection to end
    Selection.Collapse Direction:=wdCollapseEnd
    
    ' MacroUtilities.LogAction "Pasted plain text with cleaned spaces and newlines"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbExclamation, "Macro Error"
End Sub 