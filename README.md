# Word Macros Collection - Modular Version

A comprehensive collection of Microsoft Word VBA macros for legal document preparation, formatting, and editing. This modular version separates functionality into logical modules for better organization and maintainability.

**Current Version: 2.1**

## Module Structure

### Core Utilities (`MacroUtilities.bas`)
- **InitializeMacroEnvironment**: Initialize macro environment and create backup folders
- **ConfigureMacroSettings**: Configure user preferences and logging settings
- **BackupCurrentDocument**: Create timestamped backup of current document
- **DisplayError**: Centralized error handling and logging
- **LogAction**: Action logging for troubleshooting
- **ReplaceInSelection**: Text replacement utility
- **BookmarkExists**: Check if bookmark exists in document
- **StyleExists**: Check if style exists in document

### Footnote Management (`FootnoteManager.bas`)
- **InsertFormattedFootnote**: Insert footnote with specific formatting sequence
- **FN_MoveFootnotesToBody**: Move footnote text to document body
- **MoveFootnoteTextToBody**: Legacy compatibility wrapper
- **MoveToFootnote**: Move highlighted text to new footnote or create new footnote
- **ConvertFootnotes**: Convert footnote references to non-superscript with period
- **FixFN**: Fix footnote formatting with tab and period
- **CertInsertFootnote**: Insert footnote with certificate formatting
- **Demo**: Move footnote text to document body (alternative implementation)

### Special Characters (`SpecialCharacters.bas`)
- **CHAR_InsertSpecialChar**: Insert special characters with proper formatting
- **EmDash**: Insert em dash with zero-width spaces
- **CA5enDash156**: Insert scaled en dash for 5th Circuit style
- **nbsp**: Insert non-breaking space
- **SectionMarkNBSP**: Insert section mark with non-breaking space

### Document Structure (`DocumentStructure.bas`)
- **TOC_Insert**: Insert customizable table of contents
- **DOC_InsertWordCount**: Insert word count at specified bookmark
- **DOC_InsertCenteredDate**: Insert centered formatted date
- **DOC_ExportToPDF**: Export document to PDF with options
- **InsertDougTOC**: Legacy TOC insertion
- **InsertWordCount**: Legacy word count insertion
- **CenteredDate**: Legacy date insertion
- **Silent_save_to_PDF**: Legacy PDF export

### Table Formatting (`TableFormatting.bas`)
- **TBL_FormatTable**: Apply standard formatting to tables
- **XREF_SetHyperlinkStyle**: Set cross-references to hyperlink style
- **XREF_AddHyperlink**: Add hyperlink for selected URL text
- **Table1x1**: Create a 1x1 table with standard formatting
- **FixTable**: Legacy table formatting
- **SetCrossRefStyle**: Legacy cross-reference styling
- **AddHyperlinkAndZeroWidthSpaces**: Legacy hyperlink addition

### Text Formatting (`TextFormatting.bas`)
- **TXT_FormatWithRules**: Apply formatting rules (hyphens, quotes)
- **TXT_PasteLegal**: Paste and clean up legal research text
- **ReplaceInSelection**: Private utility for find/replace operations
- **SnippetSearchWithFormatAndWithout**: Legacy text formatting
- **PasteLR**: Legacy legal text paste

### Paste Utilities (`PasteUtilities.bas`)
- **PastePlainTextClean**: Paste plain text with space and newline cleanup

### Main Operations (`MainMacros.bas`)
- **BATCH_CompleteDocumentCleanup**: Run comprehensive document formatting
- **UTIL_AssignKeyboardShortcut**: Show keyboard shortcut instructions
- **UTIL_ListMacrosWithDescriptions**: Generate macro documentation

## Installation

### Method 1: Import Individual Modules
1. Open Microsoft Word
2. Press `Alt + F11` to open the VBA Editor
3. In the Project Explorer, right-click on "Normal (Normal.dotm)" or your document
4. Select "Import File..." for each `.bas` file in the `dot bas files` folder:
   - `DocumentStructure.bas`
   - `FootnoteManager.bas`
   - `MacroUtilities.bas`
   - `MainMacros.bas`
   - `PasteUtilities.bas`
   - `SpecialCharacters.bas`
   - `TableFormatting.bas`
   - `TextFormatting.bas`
5. Save your document/template

### Method 2: Copy and Paste
1. Open Microsoft Word
2. Press `Alt + F11` to open the VBA Editor
3. In the Project Explorer, right-click on your document or template
4. Select "Insert" → "Module" for each `.bas` file
5. Copy and paste the content of each module file into the corresponding module
6. Save your document/template

## Usage

### Getting Started
1. Run `MacroUtilities.InitializeMacroEnvironment` to set up the environment
2. Run `MacroUtilities.ConfigureMacroSettings` to configure preferences
3. Assign keyboard shortcuts to frequently used macros

### Common Operations
- **Clean Paste**: Use `PasteUtilities.PastePlainTextClean` for clean text pasting
- **Legal Research**: Use `TextFormatting.TXT_PasteLegal` for legal text cleanup
- **Document Cleanup**: Use `MainMacros.BATCH_CompleteDocumentCleanup` for full formatting
- **Special Characters**: Use `SpecialCharacters.CHAR_InsertSpecialChar` with appropriate parameters
- **Footnote Management**: Use `FootnoteManager.MoveToFootnote` to create footnotes from selected text

### Keyboard Shortcuts
To assign keyboard shortcuts:
1. Go to Word Preferences → Keyboard (Mac) or File → Options → Customize Ribbon → Keyboard Shortcuts (Windows)
2. Under Categories, select "Macros"
3. Select the desired macro
4. Press your desired key combination
5. Click "Assign"

## Features

- **Modular Design**: Organized into logical modules for better maintainability
- **Error Handling**: Comprehensive error handling and logging
- **Backup System**: Automatic document backup before major operations
- **Cross-Platform**: Works on both Windows and Mac versions of Word
- **Legacy Compatibility**: Maintains backward compatibility with existing macro names
- **Complete Coverage**: All functions from original NewMacros module now included

## File Structure

```
word_macros/
├── README.md
├── dot bas files/
│   ├── DocumentStructure.bas
│   ├── FootnoteManager.bas
│   ├── MacroUtilities.bas
│   ├── MainMacros.bas
│   ├── PasteUtilities.bas
│   ├── SpecialCharacters.bas
│   ├── TableFormatting.bas
│   └── TextFormatting.bas
└── PastePlainTextClean_Standalone.bas
```

## Version History

- **Version 2.1**: Complete modularization with all NewMacros functions integrated
  - Added all missing functions from NewMacros module
  - Improved organization and documentation
  - Enhanced cross-platform compatibility
- **Version 2.0**: Modular restructure with improved organization and maintainability
- **Version 1.0**: Original monolithic macro collection

## Migration from NewMacros

If you're migrating from a single NewMacros module, all functions are now available in the modularized structure:

| Original Function | New Location | New Function Name |
|------------------|--------------|-------------------|
| `movetofootnote` | FootnoteManager.bas | `MoveToFootnote` |
| `ConvertFootnotes` | FootnoteManager.bas | `ConvertFootnotes` |
| `FixFN` | FootnoteManager.bas | `FixFN` |
| `CertInsertFootnote` | FootnoteManager.bas | `CertInsertFootnote` |
| `Demo` | FootnoteManager.bas | `Demo` |
| `Table1x1` | TableFormatting.bas | `Table1x1` |
| `ReplaceInSelection` | TextFormatting.bas | `ReplaceInSelection` |

All other functions maintain their original names for backward compatibility.

## License

This macro collection is provided as-is for legal document preparation and formatting.
