# Gemini CLI Agent Rules for vbaWord

## 1. VBA Source File Modification (Critical)
- **Method**: When modifying or updating `.bas`, `.cls`, or `.frm` files, **NEVER** use partial replacement (`replace` tool). Always rewrite the entire file.
- **Workflow**: To prevent encoding issues (mojibake) in the VBA editor:
    1. Generate the complete source code.
    2. Write the content into a temporary `.txt` file first.
    3. Overwrite the target VBA file with the content of the `.txt` file (or move/rename the file).
    4. Delete the temporary file.
- **Encoding**: Ensure the final file is compatible with the Windows system's ANSI/GBK encoding for Chinese characters.

## 2. Environment Compatibility
- **Late Binding**: Use `Object` instead of specific Word/Excel types (e.g., `Dim doc As Object` instead of `Dim doc As Word.Document`) to ensure the code runs correctly across different Office versions and environments (Excel-to-Word automation).
- **Protection Handling**: Always check `ProtectionType` before modifying documents. Perform edits in an unprotected state, and apply `Protect` (if needed) only as the final step before saving.
