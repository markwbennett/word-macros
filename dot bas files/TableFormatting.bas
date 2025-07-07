Attribute VB_Name = "TableFormatting"
' =============================================================================
' TABLE FORMATTING MODULE
' Version: 2.0
' Description: Table and cross-reference formatting macros
' =============================================================================

Option Explicit

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
    
    MacroUtilities.LogAction "Formatted table with standard settings"
    
    Exit Sub
    
ErrorHandler:
    MacroUtilities.DisplayError Err.Number, Err.Description
End Sub

' Legacy name for backward compatibility
Sub FixTable()
    TBL_FormatTable 0.08, 0.08, 0.08, 0.08, True, False
End Sub

' Sets style of all cross references to "Hyperlink"
Sub XREF_SetHyperlinkStyle()
    On Error GoTo ErrorHandler
    
    MacroUtilities.BackupCurrentDocument
    
    ActiveDocument.ActiveWindow.View.ShowFieldCodes = True
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    
    ' Verify the style exists
    Dim styleExists As Boolean
    styleExists = MacroUtilities.StyleExists("Hyperlink")
    
    If Not styleExists Then
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
    
    MacroUtilities.LogAction "Set all PAGEREF fields to Hyperlink style"
    
    Exit Sub
    
ErrorHandler:
    ActiveDocument.ActiveWindow.View.ShowFieldCodes = False
    MacroUtilities.DisplayError Err.Number, Err.Description
End Sub

' Legacy name for backward compatibility
Sub SetCrossRefStyle()
    XREF_SetHyperlinkStyle
End Sub

' Adds hyperlink for selected URL text with zero-width spaces
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
    
    ' Navigate to previous sentence
    Selection.Previous(Unit:=wdSentence, Count:=1).Select
    
    MacroUtilities.LogAction "Added hyperlink with zero-width spaces for URL: " & url
    
    Exit Sub
    
ErrorHandler:
    MacroUtilities.DisplayError Err.Number, Err.Description
End Sub

' Legacy name for backward compatibility
Sub AddHyperlinkAndZeroWidthSpaces()
    XREF_AddHyperlink
End Sub

' Additional table creation function

Sub Table1x1()
' Table1x1 Macro
    ActiveDocument.Tables.Add Range:=Selection.Range, NumRows:=1, NumColumns:= _
        1, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:= _
        wdAutoFitFixed
    With Selection.Tables(1)
        If .Style <> "Table Grid" Then
            .Style = "Table Grid"
        End If
    End With
    Selection.Style = ActiveDocument.Styles("Table Text")
    Selection.Style = ActiveDocument.Styles("Table Text")
End Sub 