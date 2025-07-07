Attribute VB_Name = "MacroUtilities"
' =============================================================================
' MACRO UTILITIES MODULE
' Version: 2.0
' Description: Core utility functions and constants shared across all macro modules
' =============================================================================

Option Explicit

' Global constants for configuration
Public Const DEFAULT_FOOTNOTE_LOCATION As Long = wdBottomOfPage
Public Const DEFAULT_FOOTNOTE_NUMBERING As Long = wdRestartContinuous
Public Const DEFAULT_FOOTNOTE_STYLE As Long = wdNoteNumberStyleArabic
Public Const LOG_FILE_PATH As String = "MacroLog.txt"

' Global variables
Public g_blnLoggingEnabled As Boolean

' =============================================================================
' ERROR HANDLING AND LOGGING
' =============================================================================

' Display an error message with error number and description
Public Sub DisplayError(errNumber As Long, errDescription As String)
    MsgBox "Error " & errNumber & ": " & errDescription, vbExclamation, "Macro Error"
    LogAction "ERROR: " & errNumber & " - " & errDescription
End Sub

' Log actions to a text file for troubleshooting
Public Sub LogAction(action As String)
    If Not g_blnLoggingEnabled Then Exit Sub
    
    On Error Resume Next
    
    Dim fileNum As Integer
    Dim logPath As String
    
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

' =============================================================================
' FILE AND PATH UTILITIES
' =============================================================================

' Get the path for document backups
Public Function GetBackupFolderPath() As String
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
Public Sub BackupCurrentDocument()
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
    
    ' Save a copy as backup
    On Error Resume Next
    ActiveDocument.SaveAs2 FileName:=backupPath & backupName, _
        FileFormat:=wdFormatXMLDocument, AddToRecentFiles:=False
    
    If Err.Number <> 0 Then
        ' If backup fails, continue without it
        LogAction "Warning: Backup creation failed - " & Err.Description
        On Error GoTo ErrorHandler
        Exit Sub
    End If
    
    ' Reopen original
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

' =============================================================================
' TEXT MANIPULATION UTILITIES
' =============================================================================

' Replace text within the current selection
Public Sub ReplaceInSelection(findText As String, replaceText As String, Optional useWildcards As Boolean = False)
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
' VALIDATION UTILITIES
' =============================================================================

' Helper function to check if a bookmark exists
Public Function BookmarkExists(bookmarkName As String) As Boolean
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

' Helper function to check if a style exists
Public Function StyleExists(styleName As String) As Boolean
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

' =============================================================================
' INITIALIZATION
' =============================================================================

' Initialize the macro environment and set default preferences
Public Sub InitializeMacroEnvironment()
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
Public Sub ConfigureMacroSettings()
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