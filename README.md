# SolidworksMac

Two SOLIDWORKS macros:

- **[Safe Rename](#safe-rename)** – rename parts and sub-assemblies without breaking references. [`dist/SafeRename.swp`](dist/SafeRename.swp)
- **[Add Part Properties](#add-part-properties)** – add DESCRIPTION, Material, Weight (kg) and Size custom properties to every part in an assembly. [`src/AddPartProperties.bas`](src/AddPartProperties.bas)

# Safe Rename

A SOLIDWORKS macro with a dialog for renaming parts and sub-assemblies used in an assembly. It renames them **without breaking references** in the assemblies, drawings and derived parts that use them.

**Ready to run:** [`dist/SafeRename.swp`](dist/SafeRename.swp)

## Run it

1. Download `dist/SafeRename.swp` and save it somewhere permanent, e.g. `C:\SOLIDWORKS Macros\`.
2. Open your assembly in SOLIDWORKS. It must have been saved at least once.
3. Select the parts or sub-assemblies to rename, in the FeatureManager tree or the graphics area.
4. **Tools > Macro > Run…** and pick `SafeRename.swp`.

To get a toolbar button: **Tools > Customize > Commands > Macro**. Drag *New Macro Button* onto a toolbar and choose `SafeRename.swp` (method `main`).

The macro is stored as source code. On first run, SOLIDWORKS compiles it, which takes a moment.

## The dialog

| Section | What you do |
|---|---|
| **1. Files to rename** | The selected components are listed. Click a row and type the new file name. The name is checked as you type: invalid characters, existing files, duplicates, read-only files. *Add selected components* adds more without closing the dialog. |
| **2. Where-used search folder** | The folder (and, optionally, its sub-folders) searched for assemblies, drawings and parts that use the files. It defaults to the assembly's folder. *Find where-used* shows how many files use each part. |
| **3. Options** | Rename the drawing with the same name · close documents the macro opened · Save As fallback · keep the old file or move it to `_RenameBackup` · write a log file. |
| Progress and log | A live progress bar and a log of every file opened, saved and verified. |
| **Rename and verify** | Confirms, then renames everything and verifies it. |

The dialog is modeless. You can keep selecting components in SOLIDWORKS while it is open.

## What happens on *Rename and verify*

1. **Where-used search.** Every `.sldasm`, `.slddrw` and `.sldprt` in the search folder is checked with `GetDocumentDependencies2`. Files are not opened during the search.
2. **Load dependents.** Every file that uses the part is opened, so SOLIDWORKS updates it in memory.
3. **Rename.** `ModelDocExtension.RenameDocument` does the rename. This is the same engine as *Rename Component > Rename Document* (SW 2018+). If it fails, the macro falls back to *Save As* (not a copy).
4. **Save.** The top assembly, the renamed file, the parent sub-assemblies and every dependent are saved. Drawings are saved last.
5. **Verify.** Every saved file is read back from disk. It must reference the new name and must no longer reference the old one.
6. **Clean up.** Only after verification passes: rename the matching drawing and/or move the old file to `_RenameBackup`, if those options are on. A log file is written next to the assembly.

## Limitations

- Only files inside the search folder are found. Files stored elsewhere keep pointing at the old file. That is why *Keep the old file* is the default.
- Read-only files, files open on another computer, and files checked in to **SOLIDWORKS PDM** cannot be changed. For a PDM vault, use the vault's own *Rename* command.
- Virtual components have no file. Use *Rename Component* for them instead. Do not rename Toolbox parts.
- Changing only upper/lower case is refused. Rename to a temporary name first.
- Suppressed instances are not loaded and may keep the old reference. Verification reports this.
- The macro saves every affected document. Back up your data before renaming in bulk.

## Troubleshooting

- **"Can't find project or library"** on first run: in the VBA editor (**Tools > Macro > Edit**), open **Tools > References**. Untick any entry marked *MISSING*, and make sure *SldWorks Type Library*, *SOLIDWORKS Constant type library* and *Microsoft Forms 2.0 Object Library* are ticked. Then save.
- **"wrong format and cannot be converted"**: build the macro by hand (below). It takes about two minutes.

### Build the macro by hand (fallback)

1. **Tools > Macro > New…** and save it as `SafeRename.swp`.
2. In the VBA editor, delete the default module. Then **File > Import File…** and pick `src/SafeRename.bas`.
3. **Insert > UserForm**. In the Properties window, set its *(Name)* to `frmSafeRename`.
4. Right-click the form, choose **View Code**, and paste the full contents of `src/frmSafeRename.vba`. All controls are created by code, so the form itself stays empty.
5. Check **Tools > References** as described above. Then **Save**.

# Add Part Properties

A macro that adds custom properties to every part used in the open assembly, sub-assemblies included. It is a single module, [`src/AddPartProperties.bas`](src/AddPartProperties.bas), with no UserForm: every question is a standard input box or message box. Full spec: [`docs/AddPartProperties-spec.md`](docs/AddPartProperties-spec.md).

| Property | Value |
|---|---|
| `DESCRIPTION` | Typed for each part. Each part is highlighted and zoomed to while you type. |
| `Material` | Typed once for all parts. Optionally pick exceptions by number (`3, 7, 9-11`) and give them another material. |
| `Weight` | Automatic: `"SW-Mass@Part.SLDPRT"`, converted to kg when the part's mass unit is not kg, and checked against the real mass. |
| `Size` | Typed: one value for all parts (default `-`) or one per part. |

## Install

1. In SOLIDWORKS: **Tools > Macro > New…**, save it as `AddPartProperties.swp` somewhere permanent.
2. The VBA editor opens with a new module. Select all its code (Ctrl+A) and delete it.
3. Paste the whole content of [`src/AddPartProperties.bas`](src/AddPartProperties.bas) (open it, *Raw*, Ctrl+A, Ctrl+C). Or delete the module and use **File > Import File…** on the downloaded `.bas`.
4. **Save** and close the VBA editor.

No references need to be ticked: the code is late bound.

## Run it

Open the assembly (saved at least once), then **Tools > Macro > Run…** and pick `AddPartProperties.swp` (method `main`). The questions:

1. **Which properties** – type `1234` for all four, or e.g. `23` for Material and Weight. The box also shows how many parts were found and which are skipped.
2. **Old properties** – *Yes* deletes all existing custom properties on the Custom tab first (confirmed again), *No* keeps them.
3. **Material** – type it, then *Yes* = all parts, *No* = pick exceptions by number and type their material.
4. **Size** – *Yes* = same value for all (default `-`), *No* = typed per part.
5. **Each part** – DESCRIPTION (and Size). Empty = keep the existing value, `<` = previous part, *Cancel* = skip the remaining parts or stop.
6. **Overwrite?** – only asked if some parts already have the properties.
7. **Summary** – *OK* writes and saves the parts.

Nothing is written before the summary's *OK*. A log `AddPartProperties_<date>_<time>.log` is written next to the assembly, including every deleted property.

Only the **Custom** tab is changed. Each part file is processed once. These are skipped and listed: parts outside the assembly's folder, virtual components, PDM vault files, read-only files, Toolbox parts, suppressed and envelope components.

## Repository layout

| Path | Contents |
|---|---|
| `src/SafeRename.bas` | Engine and entry point (`main`) |
| `src/frmSafeRename.vba` | Dialog code. Controls are built at run time. |
| `src/AddPartProperties.bas` | Add Part Properties – the whole macro, one module (`main`) |
| `docs/AddPartProperties-spec.md` | Add Part Properties specification |
| `tools/build_swp.py` | Generates `dist/SafeRename.swp` from `src/` (Python 3, standard library only) |
| `dist/SafeRename.swp` | The built Safe Rename macro |

Rebuild Safe Rename after changing its sources:

```
python3 tools/build_swp.py
```

The generator writes an OLE compound file containing a source-only VBA project (MS-CFB / MS-OVBA), inside the APC host storage that SOLIDWORKS uses for macros. The output is byte-for-byte reproducible.

It was validated with `olefile`/`oletools` (container, compression, `dir` stream, form designer). It was also validated by importing the same VBA project into LibreOffice, which loaded all modules and the dialog.

The SOLIDWORKS host streams and project references were taken from a SOLIDWORKS-saved macro by [codestack.net](https://github.com/xarial/codestack) (MIT licence).
