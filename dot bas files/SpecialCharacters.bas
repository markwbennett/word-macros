Attribute VB_Name = "SpecialCharacters"
' =============================================================================
' SPECIAL CHARACTERS MODULE
' Version: 2.1
' Description: Special character insertion and formatting macros
' =============================================================================

Option Explicit

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
            
            MacroUtilities.LogAction "Inserted em dash with zero-width spaces"
            
        Case "endash"
            Selection.TypeText Text:="-"
            Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
            Selection.Font.Scaling = 156
            Selection.MoveRight Unit:=wdCharacter, Count:=1
            Selection.TypeText Text:=" "
            Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
            Selection.Font.Scaling = 100
            
            MacroUtilities.LogAction "Inserted scaled en dash (CA5 style)"
            
        Case "nbsp"
            Selection.InsertSymbol CharacterNumber:=160, Unicode:=True, Bias:=0
            MacroUtilities.LogAction "Inserted non-breaking space"
            
        Case "section"
            Selection.TypeText Text:="§"
            Selection.InsertSymbol CharacterNumber:=160, Unicode:=True, Bias:=0
            
            MacroUtilities.LogAction "Inserted section mark with non-breaking space"
            
        Case Else
            MsgBox "Unknown character type: " & charType, vbExclamation, "Error"
    End Select
    
    Exit Sub
    
ErrorHandler:
    MacroUtilities.DisplayError Err.Number, Err.Description
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