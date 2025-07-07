Attribute VB_Name = "MainMacros"
' =============================================================================
' MAIN MACROS MODULE
' Version: 2.1
' Description: Main coordination, initialization, and batch operations
' =============================================================================

Option Explicit

' =============================================================================
' UTILITY MACROS AND BATCH OPERATIONS
' =============================================================================

' Run a complete formatting cleanup on the document
Sub BATCH_CompleteDocumentCleanup()
    On Error GoTo ErrorHandler
    
    ' Confirm with user
    If MsgBox("This will perform a complete document cleanup including:" & vbCrLf & _
              "- Standardize quotes and dashes" & vbCrLf & _
              "- Fix footnote formatting" & vbCrLf & _
              "- Fix cross-references" & vbCrLf & _
              "- Set table properties" & vbCrLf & vbCrLf & _
              "A backup will be created first. Continue?", _
              vbQuestion + vbYesNo, "Complete Document Cleanup") = vbNo Then
        Exit Sub
    End If
    
    ' Create backup
    MacroUtilities.BackupCurrentDocument
    
    ' Disable screen updating for performance
    Application.ScreenUpdating = False
    
    ' 1. Standardize quotes
    Selection.HomeKey Unit:=wdStory
    Selection.EndKey Unit:=wdStory, Extend:=wdExtend
    
    With Selection.Find
        .Text = """"
        .Replacement.Text = """"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    
    With Selection.Find
        .Text = "'"
        .Replacement.Text = "'"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    
    ' 2. Fix dashes
    With Selection.Find
        .Text = "--"
        .Replacement.Text = "—"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    
    ' 3. Fix table formatting for all tables
    Dim tbl As Table
    For Each tbl In ActiveDocument.Tables
        tbl.Select
        TableFormatting.TBL_FormatTable
    Next tbl
    
    ' 4. Process all footnotes
    Dim fnt As Footnote
    For Each fnt In ActiveDocument.Footnotes
        fnt.Reference.Select
        Selection.MoveRight Unit:=wdCharacter, Count:=1
        Selection.MoveLeft Unit:=wdCharacter, Count:=2, Extend:=wdExtend
        
        ' If not already formatted as non-superscript, format it
        If Selection.Font.Superscript = True Then
            Selection.Font.Superscript = False
        End If
        
        ' Add period after footnote reference if not already there
        Selection.MoveRight Unit:=wdCharacter, Count:=1
        Selection.MoveLeft Unit:=wdCharacter, Count:=1
        
        If Selection.Text <> "." Then
            Selection.TypeText Text:="."
        End If
    Next fnt
    
    ' 5. Set cross reference style
    TableFormatting.XREF_SetHyperlinkStyle
    
    ' Return to beginning of document
    Selection.HomeKey Unit:=wdStory
    
    ' Re-enable screen updating
    Application.ScreenUpdating = True
    
    MsgBox "Document cleanup complete!", vbInformation, "Complete"
    MacroUtilities.LogAction "Performed complete document cleanup"
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MacroUtilities.DisplayError Err.Number, Err.Description
End Sub

' Create a keyboard shortcut for a macro
Sub UTIL_AssignKeyboardShortcut()
    On Error GoTo ErrorHandler
    
    MsgBox "To assign keyboard shortcuts to these macros:" & vbCrLf & vbCrLf & _
           "1. Go to Word Preferences" & vbCrLf & _
           "2. Select 'Keyboard'" & vbCrLf & _
           "3. Under 'Categories', select 'Macros'" & vbCrLf & _
           "4. Select the desired macro" & vbCrLf & _
           "5. Place cursor in 'Press new keyboard shortcut'" & vbCrLf & _
           "6. Press your desired key combination" & vbCrLf & _
           "7. Click 'Assign'", vbInformation, "Keyboard Shortcuts"
    
    Exit Sub
    
ErrorHandler:
    MacroUtilities.DisplayError Err.Number, Err.Description
End Sub

' List all macros in the document with descriptions
Sub UTIL_ListMacrosWithDescriptions()
    On Error GoTo ErrorHandler
    
    Dim newDoc As Document
    Set newDoc = Documents.Add
    
    ' Create macro list documentation
    Dim macroList As String
    
    macroList = "# WORD MACRO COLLECTION - MODULAR VERSION" & vbCrLf & vbCrLf & _
                "## Core Utilities (MacroUtilities)" & vbCrLf & _
                "- **InitializeMacroEnvironment**: Initialize macro environment" & vbCrLf & _
                "- **ConfigureMacroSettings**: Configure macro settings" & vbCrLf & _
                "- **BackupCurrentDocument**: Create document backup" & vbCrLf & vbCrLf & _
                "## Footnote Management (FootnoteManager)" & vbCrLf & _
                "- **InsertFormattedFootnote**: Insert footnote with specific formatting" & vbCrLf & _
                "- **FN_MoveFootnotesToBody**: Move footnote text to document body" & vbCrLf & vbCrLf & _
                "## Special Characters (SpecialCharacters)" & vbCrLf & _
                "- **CHAR_InsertSpecialChar**: Insert special characters with formatting" & vbCrLf & _
                "- **EmDash**: Insert em dash with zero-width spaces" & vbCrLf & _
                "- **CA5enDash156**: Insert scaled en dash for 5th Circuit style" & vbCrLf & _
                "- **nbsp**: Insert non-breaking space" & vbCrLf & _
                "- **SectionMarkNBSP**: Insert section mark with non-breaking space" & vbCrLf & vbCrLf & _
                "## Document Structure (DocumentStructure)" & vbCrLf & _
                "- **TOC_Insert**: Insert customizable table of contents" & vbCrLf & _
                "- **DOC_InsertWordCount**: Insert word count at bookmark" & vbCrLf & _
                "- **DOC_InsertCenteredDate**: Insert centered formatted date" & vbCrLf & _
                "- **DOC_ExportToPDF**: Export document to PDF with options" & vbCrLf & vbCrLf & _
                "## Table Formatting (TableFormatting)" & vbCrLf & _
                "- **TBL_FormatTable**: Format table with consistent settings" & vbCrLf & _
                "- **XREF_SetHyperlinkStyle**: Set cross-references to hyperlink style" & vbCrLf & _
                "- **XREF_AddHyperlink**: Add hyperlink for selected URL" & vbCrLf & vbCrLf & _
                "## Text Formatting (TextFormatting)" & vbCrLf & _
                "- **TXT_FormatWithRules**: Apply formatting rules to text" & vbCrLf & _
                "- **TXT_PasteLegal**: Paste and clean up legal research text" & vbCrLf & vbCrLf & _
                "## Paste Utilities (PasteUtilities)" & vbCrLf & _
                "- **PastePlainTextClean**: Paste plain text with space/newline cleanup" & vbCrLf & vbCrLf & _
                "## Main Operations (MainMacros)" & vbCrLf & _
                "- **BATCH_CompleteDocumentCleanup**: Run full document formatting" & vbCrLf & _
                "- **UTIL_AssignKeyboardShortcut**: Show keyboard shortcut instructions" & vbCrLf & _
                "- **UTIL_ListMacrosWithDescriptions**: Generate this documentation"
    
    ' Add to document
    newDoc.Range.Text = macroList
    
    MacroUtilities.LogAction "Generated modular macro list documentation"
    
    Exit Sub
    
ErrorHandler:
    MacroUtilities.DisplayError Err.Number, Err.Description
End Sub 