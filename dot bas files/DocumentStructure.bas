Attribute VB_Name = "DocumentStructure"
' =============================================================================
' DOCUMENT STRUCTURE MODULE
' Version: 2.0
' Description: Table of contents, word count, and date insertion macros
' =============================================================================

Option Explicit

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
    
    MacroUtilities.LogAction "Inserted TOC with options: " & tocOptions
    
    Exit Sub
    
ErrorHandler:
    MacroUtilities.DisplayError Err.Number, Err.Description
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
        If Not MacroUtilities.BookmarkExists(bookmarkName) Then
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
    
    MacroUtilities.LogAction "Inserted word count (" & sTemp & ") at bookmark '" & bookmarkName & "'"
    
    Exit Sub
    
ErrorHandler:
    MacroUtilities.DisplayError Err.Number, Err.Description
End Sub

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
    
    MacroUtilities.LogAction "Inserted centered date with format: " & dateFormat
    
    Exit Sub
    
ErrorHandler:
    MacroUtilities.DisplayError Err.Number, Err.Description
End Sub

' Legacy name for backward compatibility
Sub CenteredDate()
    DOC_InsertCenteredDate
End Sub

' =============================================================================
' PDF EXPORT UTILITIES
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
        MacroUtilities.LogAction "Exported document to PDF: " & pdfPath
    End If
    
    On Error GoTo ErrorHandler
    Exit Sub
    
ErrorHandler:
    MacroUtilities.DisplayError Err.Number, Err.Description
End Sub

' Legacy name for backward compatibility
Sub Silent_save_to_PDF()
    DOC_ExportToPDF False, True, True
End Sub 