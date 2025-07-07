Attribute VB_Name = "TextFormatting"
' =============================================================================
' TEXT FORMATTING MODULE
' Version: 2.1
' Description: Text formatting and legal paste utilities
' =============================================================================

Option Explicit

' =============================================================================
' TEXT FORMATTING AND CLEANUP MACROS
' =============================================================================

' Searches and formats text with specific formatting rules
Sub TXT_FormatWithRules(Optional formatHyphens As Boolean = True, _
                        Optional formatQuotes As Boolean = True)
    On Error GoTo ErrorHandler
    
    MacroUtilities.BackupCurrentDocument
    
    If formatHyphens Then
        ' Replace hyphens with scaled hyphens
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        Selection.Find.Replacement.Font.Scaling = 156
        
        With Selection.Find
            .Text = "-"
            .Replacement.Text = "-"
            .Forward = True
            .Wrap = wdFindAsk
            .Format = True
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        
        Selection.Find.Execute Replace:=wdReplaceAll
        MacroUtilities.LogAction "Replaced hyphens with scaled hyphens"
    End If
    
    If formatQuotes Then
        ' Replace straight quotes with curly quotes
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        
        With Selection.Find
            .Text = "'"
            .Replacement.Text = "'"
            .Forward = True
            .Wrap = wdFindAsk
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        
        Selection.Find.Execute Replace:=wdReplaceAll
        MacroUtilities.LogAction "Replaced straight quotes with curly quotes"
    End If
    
    Exit Sub
    
ErrorHandler:
    MacroUtilities.DisplayError Err.Number, Err.Description
End Sub

' Legacy name for backward compatibility
Sub SnippetSearchWithFormatAndWithout()
    TXT_FormatWithRules True, True
End Sub

' Pastes text from legal research with extensive cleanup
Sub TXT_PasteLegal()
    On Error GoTo ErrorHandler
    
    ' Disable screen updating for performance
    Application.ScreenUpdating = False
    
    ' Backup before making extensive changes
    MacroUtilities.BackupCurrentDocument
    
    Selection.TypeText Text:=Chr(13) 'carriage return, so that paste does not mess up style.
    Selection.TypeText Text:=ChrW(8220) 'Start quote with quotation mark.
    
    ' Declare and assign start position of the pasted content
    Dim rngFrom As Long, rngTo As Long
    rngFrom = Selection.Start
    Selection.PasteAndFormat (wdFormatSurroundingFormattingWithEmphasis)
    rngTo = Selection.End
    ActiveDocument.Range(rngFrom, rngTo).Select
    
    'Replace all Lexis pincites with " "
    With Selection.Find
        .Text = ChrW(160) & "\[*@\]" & ChrW(160)  ' Search for pattern with nonbreaking spaces and brackets
        .Replacement.Text = ""  ' Specify your replacement text here
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
        .Format = False
    End With
    
    'Replace all apostrophes with right quote.
    With Selection.Find
        .Text = "([0-9a-zA-Z])'([0-9a-zA-Z])"  ' Search for apostrophes
        .Replacement.Text = "\1'\2" ' Replace apostrophe with single quote.
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
        .Format = False
    End With
    
    'Replace all single quotes with placeholder
    MacroUtilities.ReplaceInSelection Chr(34), Chr(39), False
    
    'Replace paragraph patterns with proper formatting
    MacroUtilities.ReplaceInSelection "^p^p", "^p" & Chr(34), False
    MacroUtilities.ReplaceInSelection "^p^l^p", Chr(34) & " ", False
    MacroUtilities.ReplaceInSelection "^l^p", Chr(34) & " ", False
    MacroUtilities.ReplaceInSelection "^p^l", Chr(34) & " ", False
    
    'Replace nonbreaking space with space.
    MacroUtilities.ReplaceInSelection ChrW(160), " ", False
    
    'Replace double dash with em-dash.
    MacroUtilities.ReplaceInSelection "--", "—", False
    
    'Replace space-em-dash-space with em-dash
    MacroUtilities.ReplaceInSelection " — ", "—", False
    
    'Replace "Tex.Crim.App." with "Tex. Crim. App."
    MacroUtilities.ReplaceInSelection "Tex.Crim.App.", "Tex. Crim. App.", False
    
    ' Unlink fields (e.g., hyperlinks) and adjust the cursor position
    Selection.Fields.Unlink
    
    Selection.Collapse Direction:=wdCollapseEnd
    Selection.MoveEnd Unit:=wdCharacter, Count:=-1
    
    ' Check if the last character is a paragraph mark
    If Asc(Selection.Text) = 13 Then  ' ASCII code 13 is for carriage return (paragraph mark)
        ' Delete the paragraph mark
        Selection.Delete
    End If
    
    Selection.Collapse Direction:=wdCollapseEnd
    Selection.TypeText Text:=". "
    
    MacroUtilities.LogAction "Pasted and cleaned legal research text"
    
    ' Re-enable screen updating
    Application.ScreenUpdating = True
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MacroUtilities.DisplayError Err.Number, Err.Description
End Sub

' Legacy name for backward compatibility
Sub PasteLR()
    TXT_PasteLegal
End Sub

' Utility subroutine to handle find and replace in the selection
Private Sub ReplaceInSelection(findText As String, replaceText As String)
    With Selection.Find
        .Text = findText
        .Replacement.Text = replaceText
        .MatchWildcards = False
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
End Sub 