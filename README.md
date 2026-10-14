# SolidworksMac – Safe Rename macro

`src/SafeRename.bas` is a SOLIDWORKS VBA macro. It renames a part or sub-assembly that is used in the active assembly **without breaking references** in the assemblies, drawings and derived parts that use it.

## Install

1. In SOLIDWORKS: **Tools > Macro > New…** and save it as `SafeRename.swp`.
2. In the VBA editor, delete the default module (`SafeRename1`/`Module1`). Then use **File > Import File…** and pick `src/SafeRename.bas`.
3. Check **Tools > References**. *SldWorks 20xx Type Library* must be ticked (it is by default).
4. Save. You can also assign the macro to a toolbar button with **Tools > Customize > Commands > Macro**.

## Use

1. Open the top-level assembly and save it.
2. Select one or more components, in the FeatureManager tree or in the graphics area. They can be at any level, including inside sub-assemblies.
3. Run the macro (**Tools > Macro > Run**, entry point `main`).
4. Type a new file name for each selected file.
5. Pick the folder to search for files that use it. The default is the top assembly's folder plus its sub-folders. Choose a higher-level project folder if drawings or other assemblies are stored elsewhere.
6. Check the summary and confirm.

## What it does

| Step | Detail |
|------|--------|
| Where-used search | Every `.sldasm`, `.slddrw` and `.sldprt` in the search folder is checked with `GetDocumentDependencies2`. Files are not opened during the scan. |
| Load dependents | Every file that references the target is opened, so SOLIDWORKS updates it in memory. |
| Rename | `ModelDocExtension.RenameDocument` does the rename. This is the same engine as *Rename Component > Rename Document* (SW 2018+). If it fails, the macro falls back to *Save As* (not a copy). |
| Save | The top assembly (with referenced documents), the renamed file, the parent assemblies inside the active assembly, and all dependents are saved. Drawings are saved last. |
| Verify | Every saved file is read back from disk. It must reference the new file name and must no longer reference the old one. |
| Clean up | Files the macro opened are closed again. A drawing with the same name as the old file can be renamed too, if you agree. Only after verification passes are you offered to move the old file into `_RenameBackup`. The default answer is to keep it. |
| Log | `SafeRename_Log_<timestamp>.txt` is written next to the top assembly. |

## Settings

You can change these constants at the top of the module: `ALLOW_SAVEAS_FALLBACK`, `OFFER_DRAWING_RENAME`, `CLOSE_OPENED_DOCS`, `BACKUP_FOLDER_NAME` and `WRITE_LOG_FILE`.

## Limitations

- Only files inside the search folder are found. A file stored elsewhere that uses the renamed file keeps pointing at the old file. That is why the old file is kept unless you choose otherwise.
- Files that are read-only, open on another computer, or checked in to **SOLIDWORKS PDM** cannot be changed. For a PDM vault, use the vault's own *Rename* command, which updates references server-side.
- Virtual components have no file. Use *Rename Component* for them instead. Toolbox parts should not be renamed.
- Changing only upper/lower case (e.g. `bracket` → `Bracket`) is refused. Rename to a temporary name first.
- Suppressed instances are not loaded and may keep the old reference. Verification will report this; unsuppress them and run again.
- Back up your data (or use PDM) before renaming in bulk. The macro saves every affected document.
