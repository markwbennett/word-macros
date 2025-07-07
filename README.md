# Word Macros Collection - Modular Version

A comprehensive collection of Microsoft Word VBA macros for legal document preparation, formatting, and editing. This modular version separates functionality into logical modules for better organization and maintainability.

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
- **FixTable**: Legacy table formatting
- **SetCrossRefStyle**: Legacy cross-reference styling
- **AddHyperlinkAndZeroWidthSpaces**: Legacy hyperlink addition

### Text Formatting (`TextFormatting.bas`)
- **TXT_FormatWithRules**: Apply formatting rules (hyphens, quotes)
- **TXT_PasteLegal**: Paste and clean up legal research text
- **SnippetSearchWithFormatAndWithout**: Legacy text formatting
- **PasteLR**: Legacy legal text paste

### Paste Utilities (`PasteUtilities.bas`)
- **PastePlainTextClean**: Paste plain text with space and newline cleanup

### Main Operations (`MainMacros.bas`)
- **BATCH_CompleteDocumentCleanup**: Run comprehensive document formatting
- **UTIL_AssignKeyboardShortcut**: Show keyboard shortcut instructions
- **UTIL_ListMacrosWithDescriptions**: Generate macro documentation

## Installation

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

### Keyboard Shortcuts
To assign keyboard shortcuts:
1. Go to Word Preferences → Keyboard
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

## Version History

- **Version 2.0**: Modular restructure with improved organization and maintainability
- **Version 1.0**: Original monolithic macro collection

## License

This macro collection is provided as-is for legal document preparation and formatting.
