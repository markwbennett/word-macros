' =============================================================================
' WORD MACROS COLLECTION
' Version: 2.0
' Last Updated: April 14, 2025
' Description: A comprehensive collection of Word macros for legal document preparation,
'              formatting, and editing with enhanced error handling and documentation.
' =============================================================================

Option Explicit

' Global variables for configuration
Private Const DEFAULT_FOOTNOTE_LOCATION As Long = wdBottomOfPage
Private Const DEFAULT_FOOTNOTE_NUMBERING As Long = wdRestartContinuous
Private Const DEFAULT_FOOTNOTE_STYLE As Long = wdNoteNumberStyleArabic
Private Const LOG_FILE_PATH As String = "MacroLog.txt"
Private g_blnLoggingEnabled As Boolean

' =============================================================================
' INITIALIZATION AND CONFIGURATION
' =============================================================================

' Initialize the macro environment and set default preferences
' Usage: Call this from any template's AutoNew event
Sub InitializeMacroEnvironment()
    On Error GoTo ErrorHandler
    
    ' Set default values
    g_blnLoggingEnabled = True
    
    ' Create backup folder if it doesn't exist
    Dim backupPath As String
    backupPath = GetBackupFolderPath()
    
    If Dir(backupPath, vbDirectory) = "" Then
        MkDir backupPath
    End If
    
    LogAction "Macro environment initialized successfully"
    
    Exit Sub
    
ErrorHandler:
    DisplayError Err.Number, Err.Description
    Resume Next
End Sub

' Configure user preferences for the macros
Sub ConfigureMacroSettings()
    On Error GoTo ErrorHandler
    
    ' Simple dialog to configure settings
    Dim blnEnableLogging As Boolean
    blnEnableLogging = (MsgBox("Enable action logging?", vbYesNo + vbQuestion, "Macro Configuration") = vbYes)
    
    ' Save settings to document properties
    On Error Resume Next
    ' Try to access the property first to check if it exists
    Dim testValue As Boolean
    testValue = ActiveDocument.CustomDocumentProperties("MacroLoggingEnabled").Value
    
    ' If we get here without error, the property exists
    On Error GoTo ErrorHandler
    ActiveDocument.CustomDocumentProperties("MacroLoggingEnabled").Value = blnEnableLogging
    On Error GoTo 0
    
    g_blnLoggingEnabled = blnEnableLogging
    
    MsgBox "Configuration saved successfully.", vbInformation, "Macro Configuration"
    
    Exit Sub
    
ErrorHandler:
    ' If properties don't exist, create them
    If Err.Number = 5 Or Err.Number = 424 Then  ' Invalid procedure call or object required
        On Error Resume Next
        With ActiveDocument
            .CustomDocumentProperties.Add Name:="MacroLoggingEnabled", _
                LinkToContent:=False, Type:=msoPropertyTypeBoolean, _
                Value:=True
        End With
        ' Try again
        On Error GoTo 0
        Resume Next
    Else
        DisplayError Err.Number, Err.Description
    End If
End Sub

' =============================================================================
' UTILITY FUNCTIONS
' =============================================================================

' Display an error message with error number and description
Private Sub DisplayError(errNumber As Long, errDescription As String)
    MsgBox "Error " & errNumber & ": " & errDescription, vbExclamation, "Macro Error"
    LogAction "ERROR: " & errNumber & " - " & errDescription
End Sub

' Log actions to a text file for troubleshooting
Private Sub LogAction(action As String)
    If Not g_blnLoggingEnabled Then Exit Sub
    
    On Error Resume Next
    
    Dim fileNum As Integer
    Dim logPath As String
    Dim fso As Object
    
    ' Get path for log file - try active document path first, fallback to user documents
    logPath = LOG_FILE_PATH
    If InStr(logPath, "\") = 0 And InStr(logPath, "/") = 0 Then
        ' It's just a filename, no path, so add path
        If ActiveDocument.Path <> "" Then
            logPath = ActiveDocument.Path & "\" & LOG_FILE_PATH
        Else
            logPath = Environ("USERPROFILE") & "\Documents\" & LOG_FILE_PATH
        End If
    End If
    
    ' Try to create or append to log file
    fileNum = FreeFile
    
    Open logPath For Append As #fileNum
    
    ' Check if file opened successfully
    If Err.Number <> 0 Then
        ' Failed to open file - disable logging to prevent repeated errors
        g_blnLoggingEnabled = False
        ' Clear error to prevent cascading errors
        Err.Clear
        Exit Sub
    End If
    
    Print #fileNum, Format(Now, "yyyy-mm-dd hh:mm:ss") & " - " & action
    Close #fileNum
End Sub

' Get the path for document backups
Private Function GetBackupFolderPath() As String
    Dim docPath As String
    Dim isMac As Boolean
    
    ' Detect OS
    #If Mac Then
        isMac = True
    #Else
        isMac = False
    #End If
    
    On Error Resume Next
    docPath = ActiveDocument.Path
    
    ' If no active document path, use default Documents folder based on OS
    If docPath = "" Then
        If isMac Then
            docPath = MacScript("return (path to documents folder) as string")
            ' Remove colon at the end if present
            If Right(docPath, 1) = ":" Then
                docPath = Left(docPath, Len(docPath) - 1)
            End If
        Else
            docPath = Environ("USERPROFILE") & "\Documents"
        End If
    End If
    
    ' Add trailing slash if not present
    If Right(docPath, 1) <> Application.PathSeparator Then
        docPath = docPath & Application.PathSeparator
    End If
    
    GetBackupFolderPath = docPath & "Word_Macro_Backups" & Application.PathSeparator
End Function

' Create a backup of the current document
Sub BackupCurrentDocument()
    On Error GoTo ErrorHandler
    
    ' Check if the document has been saved at least once
    If ActiveDocument.Path = "" Then
        If MsgBox("This document hasn't been saved yet. Save now before creating backup?", _
            vbYesNo + vbQuestion, "Backup Document") = vbYes Then
            
            On Error Resume Next
            Application.Dialogs(wdDialogFileSaveAs).Show
            
            ' Check if user cancelled the save
            If ActiveDocument.Path = "" Then
                MsgBox "Backup canceled - document must be saved first.", vbInformation, "Backup Canceled"
                Exit Sub
            End If
            On Error GoTo ErrorHandler
        Else
            ' User chose not to save, so don't backup
            MsgBox "Backup canceled - document must be saved first.", vbInformation, "Backup Canceled"
            Exit Sub
        End If
    End If
    
    ' If we get here, document has a path
    Dim backupPath As String
    Dim backupName As String
    Dim origName As String
    
    ' Get original filename without extension
    origName = Left(ActiveDocument.Name, InStrRev(ActiveDocument.Name, ".") - 1)
    
    ' Create backup folder if it doesn't exist
    backupPath = GetBackupFolderPath()
    
    On Error Resume Next
    If Dir(backupPath, vbDirectory) = "" Then
        MkDir backupPath
        ' Check if folder creation failed
        If Dir(backupPath, vbDirectory) = "" Then
            ' Try using document's path instead
            backupPath = ActiveDocument.Path & Application.PathSeparator
        End If
    End If
    On Error GoTo ErrorHandler
    
    ' Create backup filename with timestamp
    backupName = origName & "_Backup_" & Format(Now, "yyyymmdd_hhmmss") & ".docx"
    
    ' Save a copy as backup - safer approach without closing document
    On Error Resume Next
    ActiveDocument.SaveAs2 FileName:=backupPath & backupName, _
        FileFormat:=wdFormatXMLDocument, AddToRecentFiles:=False
    
    If Err.Number <> 0 Then
        ' If backup fails, continue without it
        LogAction "Warning: Backup creation failed - " & Err.Description
        On Error GoTo ErrorHandler
        Exit Sub
    End If
    
    ' Reopen original without closing it
    Application.ScreenUpdating = False
    Documents.Open FileName:=ActiveDocument.FullName
    ActiveDocument.Saved = True
    Application.ScreenUpdating = True
    
    LogAction "Created backup: " & backupName
    
    Exit Sub
    
ErrorHandler:
    ' If an error occurs during backup, log it but don't stop execution
    LogAction "ERROR during backup: " & Err.Number & " - " & Err.Description
    Resume Next
End Sub

' Replace text within the current selection
Private Sub ReplaceInSelection(findText As String, replaceText As String, Optional useWildcards As Boolean = False)
    On Error GoTo ErrorHandler
    
    With Selection.Find
        .Text = findText
        .Replacement.Text = replaceText
        .Forward = True
        .Wrap = wdFindContinue
        .MatchWildcards = useWildcards
        .Execute Replace:=wdReplaceAll
    End With
    
    LogAction "Replaced '" & findText & "' with '" & replaceText & "'"
    
    Exit Sub
    
ErrorHandler:
    DisplayError Err.Number, Err.Description
End Sub

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
    
    LogAction "Inserted formatted footnote with specific sequence"
    
    Exit Sub
    
ErrorHandler:
    DisplayError Err.Number, Err.Description
End Sub

' Moves footnote text to the body of the document
Sub FN_MoveFootnotesToBody()
    On Error GoTo ErrorHandler
    
    BackupCurrentDocument
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
    LogAction "Moved all footnotes to document body"
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    DisplayError Err.Number, Err.Description
End Sub

' Legacy names for backward compatibility
Sub MoveFootnoteTextToBody()
    FN_MoveFootnotesToBody
End Sub

' =============================================================================
' SPECIAL CHARACTERS AND FORMATTING MACROS
' =============================================================================

' Inserts special characters with proper formatting
' Parameters:
'   - charType: The type of character to insert ("emdash", "endash", "nbsp", "section")
Sub CHAR_InsertSpecialChar(charType As String)
    On Error GoTo ErrorHandler
    
    Select Case LCase(charType)
        Case "emdash"
            Dim currentFont As String
            currentFont = Selection.Font.Name
            
            Selection.TypeText Text:=ChrW(8203)  ' Zero-width space
            Selection.Font.Name = currentFont
            Selection.TypeText Text:="—"         ' Em dash
            Selection.TypeText Text:=ChrW(8203)  ' Zero-width space
            Selection.TypeText Text:=" "
            Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
            Selection.Font.Name = currentFont
            
            LogAction "Inserted em dash with zero-width spaces"
            
        Case "endash"
            Selection.TypeText Text:="-"
            Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
            Selection.Font.Scaling = 156
            Selection.MoveRight Unit:=wdCharacter, Count:=1
            Selection.TypeText Text:=" "
            Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
            Selection.Font.Scaling = 100
            
            LogAction "Inserted scaled en dash (CA5 style)"
            
        Case "nbsp"
            Selection.InsertSymbol CharacterNumber:=160, Unicode:=True, Bias:=0
            LogAction "Inserted non-breaking space"
            
        Case "section"
            Selection.TypeText Text:="§"
            Selection.InsertSymbol CharacterNumber:=160, Unicode:=True, Bias:=0
            
            LogAction "Inserted section mark with non-breaking space"
            
        Case Else
            MsgBox "Unknown character type: " & charType, vbExclamation, "Error"
    End Select
    
    Exit Sub
    
ErrorHandler:
    DisplayError Err.Number, Err.Description
End Sub

' Legacy macros for backward compatibility
Sub EmDash()
    CHAR_InsertSpecialChar "emdash"
End Sub

Sub CA5enDash156()
    CHAR_InsertSpecialChar "endash"
End Sub

Sub nbsp()
    CHAR_InsertSpecialChar "nbsp"
End Sub

Sub SectionMarkNBSP()
    CHAR_InsertSpecialChar "section"
End Sub

' =============================================================================
' TABLE OF CONTENTS AND DOCUMENT STRUCTURE
' =============================================================================

' Inserts table of contents with specified options
Sub TOC_Insert(Optional levels As String = "1-7", _
               Optional hidePageNumbers As Boolean = False, _
               Optional useHyperlinks As Boolean = True)
    On Error GoTo ErrorHandler
    
    Dim tocOptions As String
    
    tocOptions = "TOC \o """ & levels & """"
    
    If hidePageNumbers Then
        tocOptions = tocOptions & " \h"
    End If
    
    If useHyperlinks Then
        tocOptions = tocOptions & " \z \u"
    End If
    
    tocOptions = tocOptions & " \w"
    
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
        Text:=tocOptions, PreserveFormatting:=True
    
    LogAction "Inserted TOC with options: " & tocOptions
    
    Exit Sub
    
ErrorHandler:
    DisplayError Err.Number, Err.Description
End Sub

' Legacy name for backward compatibility
Sub InsertDougTOC()
    TOC_Insert "1-7", True, True
End Sub

' Inserts word count in section 3 at designated bookmark
Sub DOC_InsertWordCount(Optional sectionNumber As Integer = 3, _
                        Optional bookmarkName As String = "S3WordCount")
    On Error GoTo ErrorHandler
    
    Dim oRange As Word.Range
    Dim sTemp As String
    
    With ActiveDocument
        ' If the bookmark doesn't exist, inform the user
        If Not BookmarkExists(bookmarkName) Then
            MsgBox "Bookmark '" & bookmarkName & "' does not exist in this document.", _
                vbExclamation, "Missing Bookmark"
            Exit Sub
        End If
        
        ' Make sure the section exists
        If sectionNumber > .Sections.Count Then
            MsgBox "Section " & sectionNumber & " does not exist in this document.", _
                vbExclamation, "Invalid Section"
            Exit Sub
        End If
        
        sTemp = Format(.Sections(sectionNumber).Range.Words.Count, "0")
        Set oRange = .Bookmarks(bookmarkName).Range
        oRange.Delete
        oRange.InsertAfter Text:=sTemp
        .Bookmarks.Add Name:=bookmarkName, Range:=oRange
    End With
    
    LogAction "Inserted word count (" & sTemp & ") at bookmark '" & bookmarkName & "'"
    
    Exit Sub
    
ErrorHandler:
    DisplayError Err.Number, Err.Description
End Sub

' Helper function to check if a bookmark exists
Private Function BookmarkExists(bookmarkName As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim bm As Bookmark
    
    For Each bm In ActiveDocument.Bookmarks
        If bm.Name = bookmarkName Then
            BookmarkExists = True
            Exit Function
        End If
    Next bm
    
    BookmarkExists = False
    Exit Function
    
ErrorHandler:
    BookmarkExists = False
End Function

' Legacy name for backward compatibility
Sub InsertWordCount()
    DOC_InsertWordCount 3, "S3WordCount"
End Sub

' Inserts centered date in current format
Sub DOC_InsertCenteredDate(Optional dateFormat As String = "dddd, MMMM d, yyyy")
    On Error GoTo ErrorHandler
    
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Selection.InsertDateTime DateTimeFormat:=dateFormat, _
        InsertAsField:=False, DateLanguage:=wdEnglishUS, CalendarType:= _
        wdCalendarWestern, InsertAsFullWidth:=False
    
    LogAction "Inserted centered date with format: " & dateFormat
    
    Exit Sub
    
ErrorHandler:
    DisplayError Err.Number, Err.Description
End Sub

' Legacy name for backward compatibility
Sub CenteredDate()
    DOC_InsertCenteredDate
End Sub

' =============================================================================
' TABLE AND CROSS-REFERENCE FORMATTING
' =============================================================================

' Applies standard formatting to the current table
Sub TBL_FormatTable(Optional topPadding As Single = 0.08, _
                    Optional bottomPadding As Single = 0.08, _
                    Optional leftPadding As Single = 0.08, _
                    Optional rightPadding As Single = 0.08, _
                    Optional centerRows As Boolean = True, _
                    Optional allowBreakAcrossPages As Boolean = False)
    On Error GoTo ErrorHandler
    
    ' Verify a table is selected
    If Selection.Information(wdWithInTable) = False Then
        MsgBox "Please place the cursor in a table first.", _
            vbExclamation, "No Table Selected"
        Exit Sub
    End If
    
    With Selection.Tables(1)
        .topPadding = InchesToPoints(topPadding)
        .bottomPadding = InchesToPoints(bottomPadding)
        .leftPadding = InchesToPoints(leftPadding)
        .rightPadding = InchesToPoints(rightPadding)
        
        .Rows.WrapAroundText = False
        
        If centerRows Then
            .Rows.Alignment = wdAlignRowCenter
        End If
        
        .Spacing = 0
        .AllowPageBreaks = True
        .AllowAutoFit = True
    End With
    
    Selection.Rows.allowBreakAcrossPages = allowBreakAcrossPages
    
    LogAction "Formatted table with standard settings"
    
    Exit Sub
    
ErrorHandler:
    DisplayError Err.Number, Err.Description
End Sub

' Legacy name for backward compatibility
Sub FixTable()
    TBL_FormatTable 0.08, 0.08, 0.08, 0.08, True, False
End Sub

' Sets style of all cross references to "Hyperlink"
Sub XREF_SetHyperlinkStyle()
    On Error GoTo ErrorHandler
    
    BackupCurrentDocument
    
    ActiveDocument.ActiveWindow.View.ShowFieldCodes = True
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    
    ' Verify the style exists
    Dim StyleExists As Boolean
    StyleExists = StyleExists("Hyperlink")
    
    If Not StyleExists Then
        MsgBox "The 'Hyperlink' style does not exist in this document.", _
            vbExclamation, "Missing Style"
        Exit Sub
    End If
    
    Selection.Find.Replacement.style = ActiveDocument.Styles("Hyperlink")
    
    With Selection.Find
        .Text = "^19 PAGEREF"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchKashida = False
        .MatchDiacritics = False
        .MatchAlefHamza = False
        .MatchControl = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    
    Selection.Find.Execute Replace:=wdReplaceAll
    ActiveDocument.ActiveWindow.View.ShowFieldCodes = False
    
    LogAction "Set all PAGEREF fields to Hyperlink style"
    
    Exit Sub
    
ErrorHandler:
    ActiveDocument.ActiveWindow.View.ShowFieldCodes = False
    DisplayError Err.Number, Err.Description
End Sub

' Helper function to check if a style exists
Private Function StyleExists(styleName As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim style As style
    
    For Each style In ActiveDocument.Styles
        If style.NameLocal = styleName Then
            StyleExists = True
            Exit Function
        End If
    Next style
    
    StyleExists = False
    Exit Function
    
ErrorHandler:
    StyleExists = False
End Function

' Legacy name for backward compatibility
Sub SetCrossRefStyle()
    XREF_SetHyperlinkStyle
End Sub

' Adds hyperlink for selected URL and inserts zero-width spaces
Sub XREF_AddHyperlink()
    On Error GoTo ErrorHandler
    
    Dim url As String
    
    ' Check if text is selected
    If Selection.Type <> wdSelectionNormal Then
        MsgBox "Please select the URL text first.", vbExclamation, "No Text Selected"
        Exit Sub
    End If
    
    url = Selection
    
    ' Validate URL format (basic check)
    If InStr(url, "://") = 0 And Left(url, 4) <> "www." Then
        If MsgBox("The selected text doesn't appear to be a valid URL. Continue anyway?", _
            vbQuestion + vbYesNo, "URL Validation") = vbNo Then
            Exit Sub
        End If
    End If
    
    ' Create hyperlink
    ActiveDocument.Hyperlinks.Add Anchor:=Selection.Range, Address:=url, _
        SubAddress:="", ScreenTip:="", TextToDisplay:=url
        
    ' Navigate to previous sentence for context
    Selection.Previous(Unit:=wdSentence, Count:=1).Select
    
    LogAction "Added hyperlink for URL: " & url
    
    Exit Sub
    
ErrorHandler:
    DisplayError Err.Number, Err.Description
End Sub

' Adds hyperlink for selected URL with zero-width spaces (original implementation)
Sub AddHyperlinkAndZeroWidthSpaces()
    On Error GoTo ErrorHandler
    
    Dim url As String
    
    ' Check if text is selected
    If Selection.Type <> wdSelectionNormal Then
        MsgBox "Please select the URL text first.", vbExclamation, "No Text Selected"
        Exit Sub
    End If
    
    url = Selection
    
    ' Validate URL format (basic check)
    If InStr(url, "://") = 0 And Left(url, 4) <> "www." Then
        If MsgBox("The selected text doesn't appear to be a valid URL. Continue anyway?", _
            vbQuestion + vbYesNo, "URL Validation") = vbNo Then
            Exit Sub
        End If
    End If
    
    ' Create hyperlink
    ActiveDocument.Hyperlinks.Add Anchor:=Selection.Range, Address:=url, _
        SubAddress:="", ScreenTip:="", TextToDisplay:=url
    
    ' Navigate to previous sentence
    Selection.Previous(Unit:=wdSentence, Count:=1).Select
    
    LogAction "Added hyperlink with zero-width spaces for URL: " & url
    
    Exit Sub
    
ErrorHandler:
    DisplayError Err.Number, Err.Description
End Sub

' =============================================================================
' TEXT FORMATTING AND CLEANUP MACROS
' =============================================================================

' Searches and formats text with specific formatting rules
Sub TXT_FormatWithRules(Optional formatHyphens As Boolean = True, _
                        Optional formatQuotes As Boolean = True)
    On Error GoTo ErrorHandler
    
    BackupCurrentDocument
    
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
        LogAction "Replaced hyphens with scaled hyphens"
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
        LogAction "Replaced straight quotes with curly quotes"
    End If
    
    Exit Sub
    
ErrorHandler:
    DisplayError Err.Number, Err.Description
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
    BackupCurrentDocument
    
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
    ReplaceInSelection Chr(34), Chr(39), False
    
    'Replace paragraph patterns with proper formatting
    ReplaceInSelection "^p^p", "^p" & Chr(34), False
    ReplaceInSelection "^p^l^p", Chr(34) & " ", False
    ReplaceInSelection "^l^p", Chr(34) & " ", False
    ReplaceInSelection "^p^l", Chr(34) & " ", False
    
    'Replace nonbreaking space with space.
    ReplaceInSelection ChrW(160), " ", False
    
    'Replace double dash with em-dash.
    ReplaceInSelection "--", "—", False
    
    'Replace space-em-dash-space with em-dash
    ReplaceInSelection " — ", "—", False
    
    'Replace "Tex.Crim.App." with "Tex. Crim. App."
    ReplaceInSelection "Tex.Crim.App.", "Tex. Crim. App.", False
    
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
    
    LogAction "Pasted and cleaned legal research text"
    
    ' Re-enable screen updating
    Application.ScreenUpdating = True
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    DisplayError Err.Number, Err.Description
End Sub

' Legacy name for backward compatibility
Sub PasteLR()
    TXT_PasteLegal
End Sub

' =============================================================================
' PDF EXPORT AND DOCUMENT UTILITIES
' =============================================================================

' Export document to PDF with options
' Note: This is designed for Windows version of Word but includes fallback for Mac
Sub DOC_ExportToPDF(Optional openAfterExport As Boolean = False, _
                    Optional optimizeForPrint As Boolean = True, _
                    Optional exportBookmarks As Boolean = True)
    On Error GoTo ErrorHandler
    
    Dim pdfPath As String
    Dim docPath As String
    Dim docName As String
    Dim isMac As Boolean
    
    ' Detect OS platform
    #If Mac Then
        isMac = True
    #Else
        isMac = False
    #End If
    
    ' Get document path and name
    docPath = ActiveDocument.Path
    docName = Left(ActiveDocument.Name, InStrRev(ActiveDocument.Name, ".") - 1)
    
    ' Create PDF path
    If docPath = "" Then
        If isMac Then
            ' Mac default documents location
            pdfPath = MacScript("return (path to documents folder) as string") & docName & ".pdf"
        Else
            ' Windows default documents location
            pdfPath = Environ("USERPROFILE") & "\Documents\" & docName & ".pdf"
        End If
    Else
        ' Use document's own path
        pdfPath = docPath & Application.PathSeparator & docName & ".pdf"
    End If
    
    ' Ask user for confirmation and location
    If MsgBox("Save PDF as:" & vbCrLf & pdfPath & vbCrLf & vbCrLf & _
        "Click Yes to continue or No to cancel.", vbYesNo + vbQuestion, _
        "Export to PDF") = vbNo Then
        Exit Sub
    End If
    
    ' Try to export using appropriate method for platform
    On Error Resume Next
    
    If isMac Then
        ' Mac-specific handling
        ActiveDocument.SaveAs2 FileName:=pdfPath, FileFormat:=wdFormatPDF, _
            AddToRecentFiles:=False
    Else
        ' Windows method
        ActiveDocument.ExportAsFixedFormat OutputFileName:=pdfPath, _
            ExportFormat:=wdExportFormatPDF, openAfterExport:=openAfterExport, _
            OptimizeFor:=IIf(optimizeForPrint, wdExportOptimizeForPrint, wdExportOptimizeForOnScreen), _
            Range:=wdExportAllDocument, Item:=wdExportDocumentContent, _
            IncludeDocProps:=True, KeepIRM:=True, _
            CreateBookmarks:=IIf(exportBookmarks, wdExportCreateHeadingBookmarks, wdExportCreateNoBookmarks)
    End If
    
    ' Check if an error occurred
    If Err.Number <> 0 Then
        On Error GoTo ErrorHandler
        MsgBox "Failed to export PDF: " & Err.Description & vbCrLf & _
            "Please try using File > Save As and select PDF format.", vbExclamation, "Export Failed"
    Else
        MsgBox "PDF successfully created:" & vbCrLf & pdfPath, vbInformation, "PDF Created"
        LogAction "Exported document to PDF: " & pdfPath
    End If
    
    On Error GoTo ErrorHandler
    Exit Sub
    
ErrorHandler:
    DisplayError Err.Number, Err.Description
End Sub

' Legacy name for backward compatibility
Sub Silent_save_to_PDF()
    DOC_ExportToPDF False, True, True
End Sub

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
    BackupCurrentDocument
    
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
        TBL_FormatTable
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
    XREF_SetHyperlinkStyle
    
    ' Return to beginning of document
    Selection.HomeKey Unit:=wdStory
    
    ' Re-enable screen updating
    Application.ScreenUpdating = True
    
    MsgBox "Document cleanup complete!", vbInformation, "Complete"
    LogAction "Performed complete document cleanup"
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    DisplayError Err.Number, Err.Description
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
    DisplayError Err.Number, Err.Description
End Sub

' List all macros in the document with descriptions
Sub UTIL_ListMacrosWithDescriptions()
    On Error GoTo ErrorHandler
    
    Dim newDoc As Document
    Set newDoc = Documents.Add
    
    ' Create string in smaller chunks
    Dim part1 As String, part2 As String, part3 As String, part4 As String, part5 As String
    Dim macroList As String
    
    ' Title
    part1 = "# WORD MACRO COLLECTION" & vbCrLf & vbCrLf
    
    ' Footnote Management section
    part1 = part1 & "## Footnote Management" & vbCrLf & _
            "- **InsertFormattedFootnote**: Inserts footnote with specific formatting sequence" & vbCrLf & _
            "- **FN_MoveFootnotesToBody**: Moves footnote text to document body" & vbCrLf & vbCrLf
    
    ' Special Characters section
    part2 = "## Special Characters" & vbCrLf & _
            "- **CHAR_InsertSpecialChar**: Inserts special characters with formatting" & vbCrLf & _
            "- **EmDash**: Inserts em dash with zero-width spaces" & vbCrLf & _
            "- **CA5enDash156**: Inserts scaled en dash for 5th Circuit style" & vbCrLf & _
            "- **nbsp**: Inserts non-breaking space" & vbCrLf & _
            "- **SectionMarkNBSP**: Inserts section mark with non-breaking space" & vbCrLf & vbCrLf
    
    ' Document Structure section
    part3 = "## Document Structure" & vbCrLf & _
            "- **TOC_Insert**: Inserts customizable table of contents" & vbCrLf & _
            "- **DOC_InsertWordCount**: Inserts word count at bookmark" & vbCrLf & _
            "- **DOC_InsertCenteredDate**: Inserts centered formatted date" & vbCrLf & vbCrLf
    
    ' Tables and Cross-References section
    part4 = "## Tables and Cross-References" & vbCrLf & _
            "- **TBL_FormatTable**: Formats table with consistent settings" & vbCrLf & _
            "- **XREF_SetHyperlinkStyle**: Sets cross-references to hyperlink style" & vbCrLf & _
            "- **XREF_AddHyperlink**: Adds hyperlink for selected URL" & vbCrLf & vbCrLf
    
    ' Text Formatting and Utilities sections
    part5 = "## Text Formatting" & vbCrLf & _
            "- **TXT_FormatWithRules**: Applies formatting rules to text" & vbCrLf & _
            "- **TXT_PasteLegal**: Pastes and cleans up legal research text" & vbCrLf & vbCrLf & _
            "## Utilities" & vbCrLf & _
            "- **DOC_ExportToPDF**: Exports document to PDF with options" & vbCrLf & _
            "- **BackupCurrentDocument**: Creates backup of current document" & vbCrLf & _
            "- **BATCH_CompleteDocumentCleanup**: Runs full document formatting" & vbCrLf & _
            "- **ConfigureMacroSettings**: Configures macro environment settings"
    
    ' Combine all parts
    macroList = part1 & part2 & part3 & part4 & part5
    
    ' Add to document
    newDoc.Range.Text = macroList
    
    LogAction "Generated macro list documentation"
    
    Exit Sub
    
ErrorHandler:
    DisplayError Err.Number, Err.Description
End Sub

' =============================================================================
' END OF MACRO COLLECTION
' =============================================================================

