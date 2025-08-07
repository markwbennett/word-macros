# Installing Word Macros: PastePlainTextClean and MoveToFootnote

## What These Macros Do

**PastePlainTextClean**: Pastes clipboard text while cleaning up formatting issues like extra spaces, line breaks, and converting special characters. Automatically assigned to Cmd+Option+Shift+V.

**MoveToFootnote**: Creates a new footnote and either moves selected text into it, or just creates an empty footnote if no text is selected.

## Installation Instructions

### Step 1: Open the Visual Basic Editor
1. Open Microsoft Word
2. Press **Alt+F11** (Windows) or **Option+F11** (Mac) to open the Visual Basic Editor
3. If that doesn't work, go to **Tools** → **Macro** → **Visual Basic Editor**

### Step 2: Create a New Module
1. In the Visual Basic Editor, look for the **Project** panel (usually on the left)
2. Find your document or **Normal** template
3. Right-click on **Normal** or your document name
4. Select **Insert** → **Module**
5. A new module window will appear

### Step 3: Copy the Macro Code
1. Copy the entire contents of the "Paste and Footnote" file from this repository
2. Paste it into the empty module window that just opened
3. The code should appear with proper syntax highlighting

### Step 4: Save the Macros
1. Press **Ctrl+S** (Windows) or **Cmd+S** (Mac) to save
2. Close the Visual Basic Editor by clicking the X or pressing **Alt+Q**

### Step 5: Test the Macros

#### Test PastePlainTextClean:
1. Copy some text from a website or PDF (text with formatting issues)
2. In Word, press **Cmd+Option+Shift+V** (Mac) or try **Ctrl+Alt+Shift+V** (Windows)
3. The text should paste with cleaned formatting

#### Test MoveToFootnote:
1. Type some text in Word
2. Select the text you want to move to a footnote
3. Go to **Tools** → **Macro** → **Macros**
4. Find "movetofootnote" in the list
5. Click **Run**
6. Your selected text should move to a new footnote

### Step 6: Create Keyboard Shortcuts (Optional)

#### For MoveToFootnote:
1. Go to **Tools** → **Customize Keyboard** (or **Word** → **Preferences** → **Keyboard** on Mac)
2. In **Categories**, select **Macros**
3. Find "movetofootnote" in the **Macros** list
4. Click in the **Press new shortcut key** box
5. Press your desired key combination (e.g., **Ctrl+Shift+F** or **Cmd+Shift+F**)
6. Click **Assign**

## Troubleshooting

**"Macros are disabled"**: 
- Go to **File** → **Options** → **Trust Center** → **Trust Center Settings** → **Macro Settings**
- Select "Enable all macros" (temporarily for installation)

**"Cannot find macro"**:
- Make sure you saved the module in the **Normal** template, not just a specific document
- Try closing and reopening Word

**PastePlainTextClean shortcut doesn't work**:
- The shortcut is set automatically when Word starts. Try restarting Word after installation.

## Making Macros Available in All Documents

To ensure these macros work in all your Word documents:
1. When saving the module, make sure it's in the **Normal** template (not a specific document)
2. The Normal template loads with every new Word document

## Security Note

These macros only perform text formatting and footnote operations. They don't access external files or networks, making them safe to use.
