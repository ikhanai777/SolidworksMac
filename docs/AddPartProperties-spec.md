# Spec – Add Part Properties macro

**Status:** Draft for review
**Macro name:** `AddPartProperties.swp` (entry point `main`)
**Target:** SOLIDWORKS 2020 or later, VBA macro, same build approach as `SafeRename` (`src/` + `tools/build_swp.py` → `dist/`)

## 1. Purpose

Add up to four custom properties to every part used in the active assembly, so the BOM can show them without editing each part by hand. At the start the user chooses which of the four to add, and whether to delete the parts' old custom properties first.

| # | Property name | Type | Value / Text Expression | Who supplies it | Scope prompt |
|---|---|---|---|---|---|
| 1 | `DESCRIPTION` | Text | Typed by the user | Asked **once per part** | — (always every part) |
| 2 | `Material` | Text | Typed by the user | Asked **once** for all parts; the user can pick **exceptions** and type a different material for them | All parts / all parts with exceptions |
| 3 | `Weight` | Equation | Linked to `SW-Mass`, in **kg** | Automatic, no prompt | — (always every part) |
| 4 | `Size` | Text | Typed by the user, or `-` | Either **one value for all parts** (default `-`) or **typed per part** | Same for all / per part |

Property names are written exactly as shown above, case included: `DESCRIPTION`, `Material`, `Weight`, `Size`.

## 2. Scope

### In scope
- The active document must be an assembly (`.sldasm`) that has been saved at least once.
- All parts in the assembly, **including parts inside sub-assemblies** at any depth.
- **Only local part files in the assembly's folder** (the folder of the top `.sldasm` and its sub-folders). Parts stored anywhere else are skipped.
- Each part file is processed **once**, however many instances it has. Parts are identified by full file path.
- Properties are written **only** to the part file's **Custom** tab (file-level, configuration name `""`). The *Configuration Specific* tab is never read, written or deleted.

### Out of scope
- Sub-assemblies themselves and the top assembly do not get the properties.
- Drawings are not touched.
- Weldment cut-list properties.

### Skipped automatically (listed in the log and the final summary)
| Case | Reason |
|---|---|
| Parts outside the assembly's folder | Only local files in the folder are changed. |
| Virtual components | Not separate files; never changed. |
| Files in a **SOLIDWORKS PDM** vault (checked in or checked out) | Must be changed through PDM, not by this macro. |
| Read-only files / files open by another user | Cannot be saved. |
| Toolbox parts (`IsToolboxPart`) | Shared library files; must not be changed. |
| Suppressed components | No model loaded. |
| Envelope components | Not real parts. |

PDM detection: a file is treated as a PDM file if it is inside a PDM vault view (a folder whose root carries the PDM vault view marker) or if it is read-only. **To verify during implementation:** the exact vault-view marker on the target PDM version. The read-only test alone already catches checked-in files, because PDM keeps them read-only in the local view.

Lightweight components are resolved before processing.

## 3. Workflow

```
Run macro
 ├─ 1. Checks          active doc is a saved assembly; otherwise message and exit
 ├─ 2. Collect parts   walk the component tree, de-duplicate by file path, apply skip rules
 ├─ 3. Start dialog    which properties to add (4 tick boxes)
 │                     + delete old custom properties? (Keep / Delete all)
 ├─ 4. Material        (if ticked) type value once → All parts / All parts with exceptions
 ├─ 5. Size            (if ticked) Same value for all (default "-") / Type per part
 ├─ 6. Per-part dialog (if DESCRIPTION ticked, or Size "Type per part")
 ├─ 7. Review          summary table, user confirms
 ├─ 8. Write & save    delete old properties (if chosen), write properties, save each part
 └─ 9. Report          summary message + log file
```

Dialogs for properties that are not ticked in step 3 are not shown. Weight has no dialog of its own: if ticked, it is added to every part (§5.3).

The user can **Cancel** at any step up to 7. Nothing is written or deleted in any file before step 8.

## 4. Dialogs

### 4.1 Start dialog (step 3)

- Caption: *Add Part Properties*
- Shows: assembly name and the number of parts found / skipped (with a *Show skipped* link that lists them and why).
- **Properties to add** – four tick boxes, all ticked by default:
  - `DESCRIPTION`
  - `Material`
  - `Weight`
  - `Size`
- **Old custom properties** – two option buttons:
  - **Keep existing properties** (default) – other properties on the Custom tab stay as they are.
  - **Delete all existing custom properties first** – every property on each part's Custom tab is deleted, then the ticked properties are added. A warning line is shown in red: *"All properties on the Custom tab of N parts will be deleted. This cannot be undone."*
- Buttons: **Next**, **Cancel**.
- **Next** is disabled when no property is ticked and *Keep existing* is selected (there would be nothing to do). Ticking nothing with *Delete all* is allowed: it only clears the Custom tab.

### 4.2 Material (step 4)

- Caption: *Material*
- Text box **Material for all parts** – empty; the user types the material as plain text.
- Two option buttons:
  - **Apply to all parts** (default) – every part gets this material.
  - **Apply to all parts, with exceptions** – opens the exceptions list below.
- Buttons: **Next**, **Back**, **Cancel**.
- An empty value is not allowed with **Next**.

**Exceptions list** (only with *Apply to all parts, with exceptions*):

- A list of all collected parts with three columns: tick box *Exception*, *Part* (file name, plus sub-assembly path in grey), *Material*.
- Every row starts unticked, showing the common material in grey.
- Ticking a row makes its *Material* cell editable; the user types that part's material.
- To give several parts the same different material at once: tick them, type the value in **Material for ticked parts** and click **Set**.
- Clicking a row selects/highlights that component in the SOLIDWORKS graphics area so the user can see which part it is.
- A ticked row with an empty *Material* is not allowed with **Next**.
- Buttons: **Next**, **Back**, **Cancel**.

Result: unticked parts get the common material; ticked parts get their own value.

### 4.3 Size (step 5)

- Caption: *Size*
- Two option buttons:
  - **Same value for all parts** (default) – text box pre-filled with `-`. The user can keep `-` or type another value.
  - **Type per part** – Size is asked in the per-part dialog (§4.4).
- Buttons: **Next**, **Back**, **Cancel**.

### 4.4 Per-part dialog (step 6)

Shown once for every collected part, in tree order – only if `DESCRIPTION` is ticked, or Size is *Type per part*.

- Caption: *Part `n` of `N`*
- Shows: part file name, full path, configuration in use.
- The part is selected and zoomed to (`ModelDocExtension.SelectByID2` + *Zoom to Selection*) so the user sees it while typing. The view is restored afterwards.
- Text box **DESCRIPTION** – only if `DESCRIPTION` is ticked. Pre-filled with the part's existing `DESCRIPTION` value, if any.
- Text box **Size** – only if Size is *Type per part*. Pre-filled with the part's existing `Size` value, or `-` if there is none.
- Buttons:
  - **Next** – store the values and go to the next part.
  - **Back** – previous part.
  - **Skip** – keep this part's existing DESCRIPTION (and Size) as they are.
  - **Skip remaining** – keep the existing values for all remaining parts and go to Review.
  - **Cancel** – stop the macro, nothing written.
- Enter = Next.

With *Delete all existing custom properties first*, **Skip** still keeps the old value: the macro deletes the Custom tab and then writes the old value back for that property.

### 4.5 Review (step 7)

A table, one row per part, columns: *Part*, *DESCRIPTION*, *Material*, *Weight*, *Size*, *Note*. Columns for unticked properties are hidden.

- Cells show the value that will be written, or `(unchanged)`.
- *Note* shows warnings, e.g. "existing value will be overwritten", "5 other custom properties will be deleted".
- With *Delete all existing custom properties first*, a red line above the table repeats how many properties on how many parts will be deleted.
- Checkbox **Overwrite existing values** (default **on**) – shown only with *Keep existing properties*. When off, a property that already exists on a part is left as it is.
- Checkbox **Save parts after writing** (default **on**).
- Buttons: **Apply**, **Back**, **Cancel**.

## 5. Property details

All properties are written to the Custom tab with
`ModelDocExtension.CustomPropertyManager("").Add3(Name, Type, Value, swCustomPropertyReplaceValue)`
(or `swCustomPropertyOnlyIfNew` when *Overwrite existing values* is off).
After writing, the value is read back with `Get6` and the resolved value is compared with what was expected.

### 5.0 Deleting old properties
- Only with *Delete all existing custom properties first*.
- For each processed part: read all names with `CustomPropertyManager("").GetNames`, delete each with `Delete2`. The *Configuration Specific* tab is not touched.
- Every deleted property (name, type, value/expression) is written to the log first, so it can be re-entered by hand if needed.
- Deletion and writing happen on the same open document, before one save. If writing fails, the part is closed **without saving**, so its old properties stay on disk.

### 5.1 DESCRIPTION
- Type: `swCustomInfoText`.
- Value: exactly what the user typed. Leading/trailing spaces are trimmed.

### 5.2 Material
- Type: `swCustomInfoText`.
- Value: plain text typed by the user – the common material, or the part's own value if it was ticked as an exception.
- It is **not** linked to the SOLIDWORKS material (`SW-Material`) and does **not** change the material used for mass.

### 5.3 Weight
- Added **automatically to every part**, if ticked. No dialog.
- Type: **Equation** (the *Equation* type shown in the Custom Properties tab, SOLIDWORKS 2020+).
  If the installed API does not expose the equation type constant, fall back to `swCustomInfoText` with the same expression – SOLIDWORKS still evaluates a linked text expression.
- Text expression: `"SW-Mass@<PartFileName>.SLDPRT"` — e.g. `"SW-Mass@Bracket.SLDPRT"`.
  This is a live link: the value updates whenever the part changes.
- **Units – must be kg.** `SW-Mass` is reported in the part's *mass-property* unit (Tools > Options > Document Properties > Units > Mass). The macro reads that unit per part (`swUnitsMassPropMass`) and:
  | Part mass unit | Expression written |
  |---|---|
  | kg | `"SW-Mass@Part.SLDPRT"` |
  | g | `"SW-Mass@Part.SLDPRT" / 1000` |
  | lb | `"SW-Mass@Part.SLDPRT" * 0.45359237` |
  | other | warn in Review, write the plain link |

  Alternative (option on the Review screen, default **off**): *Set part mass unit to kg* – changes each part's mass unit to kg and always writes the plain link. This changes a document setting, so it is opt-in.
- Decimal places follow the part's mass-unit decimal setting.
- **To verify during implementation:** that an Equation-type property accepts arithmetic on a linked `SW-Mass` value on the minimum supported SOLIDWORKS version. If it does not, the default becomes the *Set part mass unit to kg* behaviour, with a confirmation prompt.

### 5.4 Size
- Type: `swCustomInfoText`.
- Value: typed by the user – either one value for all parts (default `-`) or one value per part (§4.3, §4.4).
- The macro does not calculate the size and does not check the text. An empty value is written as `-`.

## 6. Writing and saving

- Each part is opened (or taken from memory if already loaded) with `OpenDoc6`, silently.
- Old properties are deleted (if chosen), the new ones written, then the part is saved with `Save3(swSaveAsOptions_Silent)`.
- Parts the macro opened itself are closed afterwards; parts that were already open stay open.
- The top assembly is rebuilt at the end so the BOM refreshes. It is **not** saved automatically (it was not changed).
- A failure on one part (open, delete, write, verify or save) is logged and the macro continues with the next part.

## 7. Report

- Final message: number of parts updated, skipped and failed, with a button to open the log.
- Log file `AddPartProperties_<yyyymmdd_hhmmss>.log` next to the assembly: one line per part per property with the old value, new value and result, plus every deleted property.

## 8. Acceptance tests

| # | Setup | Expected |
|---|---|---|
| 1 | Run with a part document active | Message "Open an assembly first", nothing changed. |
| 2 | Assembly with 3 parts, one used 4 times | Per-part dialog appears 3 times, not 6. |
| 3 | Part inside a sub-assembly 2 levels deep, in the assembly folder | Part is processed; sub-assembly gets no properties. |
| 4 | Part stored outside the assembly folder | Skipped, listed in *Show skipped* and the log, file unchanged. |
| 5 | Virtual component | Skipped, not modified. |
| 6 | Part checked in to PDM | Skipped, not modified. |
| 7 | Start dialog: only `Weight` ticked | No Material, Size or per-part dialog; only `Weight` is written. |
| 8 | Start dialog: nothing ticked, *Keep existing* | **Next** is disabled. |
| 9 | Part with 5 old properties, *Delete all*, all 4 ticked | Custom tab has exactly the 4 properties; the 5 old ones are in the log. Configuration-specific properties unchanged. |
| 10 | *Delete all*, Skip on a part with an existing DESCRIPTION | That DESCRIPTION is kept; other old properties are deleted. |
| 11 | Material = `S355`, Apply to all | All parts show `Material = S355`. |
| 12 | Material = `S355`, with exceptions: 2 of 5 parts ticked, typed `SS304` | 3 parts show `S355`, the 2 ticked show `SS304`. |
| 13 | Part with mass unit kg, mass 2.5 kg | `Weight` evaluates to `2.5`. |
| 14 | Part with mass unit g, mass 2500 g | `Weight` evaluates to `2.5`. |
| 15 | Change part geometry after run, rebuild | `Weight` value updates. |
| 16 | Size: Same value for all, keep default | Every part shows `Size = -`. |
| 17 | Size: Type per part, type `PL 10` for one part | Size box appears in the per-part dialog; that part shows `PL 10`. |
| 18 | Existing `DESCRIPTION`, *Keep existing*, Overwrite off | Existing value kept. |
| 19 | Cancel in the 3rd per-part dialog | No file modified (check file dates). |
| 20 | Toolbox fastener in the assembly | Not listed, not modified. |
| 21 | Suppressed component | Not listed; reported as skipped. |

## 9. Decisions

- Size is typed (or `-` for all); it is not calculated.
- Weight is automatic on every part, in kg, as an equation linked to `SW-Mass`.
- Material is plain typed text: one value for all, with optional exceptions.
- Only local part files in the assembly's folder are changed; virtual components and PDM files are skipped.
- Properties go only to the Custom tab.
- Property names are exactly `DESCRIPTION`, `Material`, `Weight`, `Size`.
- At the start the user picks which properties to add and whether to delete old custom properties.
