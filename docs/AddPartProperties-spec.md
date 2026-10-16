# Spec – Add Part Properties macro

**Status:** Implemented as a single module, `src/AddPartProperties.bas` (no UserForm; questions use `InputBox` / `MsgBox`)
**Macro:** `src/AddPartProperties.bas`, entry point `main`. Pasted or imported into a new SOLIDWORKS macro; late bound, so no extra references are needed.
**Target:** SOLIDWORKS 2020 or later, VBA

## 1. Purpose

Add up to four custom properties to every part used in the active assembly, so the BOM can show them without editing each part by hand. At the start the user chooses which of the four to add, and whether to delete the parts' old custom properties first.

| # | Property name | Type | Value / Text Expression | Who supplies it | Scope prompt |
|---|---|---|---|---|---|
| 1 | `DESCRIPTION` | Text | Typed by the user | Asked **once per part** | — (always every part) |
| 2 | `Material` | Text | Typed by the user | Asked **once** for all parts; the user can pick **exceptions** and type a different material for them | All parts / all parts with exceptions |
| 3 | `Weight` | Equation (text expression) | Linked to `SW-Mass`, in **kg** | Automatic, no prompt | — (always every part) |
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
| Toolbox library parts (`ToolboxPartType` = standard) | Shared library files; must not be changed. |
| Suppressed components | No model loaded. |
| Envelope components | Not real parts. |

PDM detection: a file is treated as a PDM file if it is inside a PDM vault view (a folder whose root carries the PDM vault view marker) or if it is read-only. **To verify during implementation:** the exact vault-view marker on the target PDM version. The read-only test alone already catches checked-in files, because PDM keeps them read-only in the local view.

A file is listed as skipped for an instance-level reason (suppressed, envelope) only when no other instance of the same file can be used.

Lightweight components are resolved before processing.

## 3. Workflow

```
Run macro (main)
 ├─ 1. Checks          active doc is a saved assembly; otherwise message and exit
 ├─ 2. Collect parts   walk the component tree, de-duplicate by file path, apply skip rules
 ├─ 3. Properties      InputBox: which properties to add (digits 1-4)
 ├─ 4. Old properties  MsgBox: keep or delete all (only if any part has properties)
 ├─ 5. Material        (if chosen) InputBox value → MsgBox all parts / with exceptions
 ├─ 6. Size            (if chosen) MsgBox same for all / per part → InputBox value
 ├─ 7. Per part        (if DESCRIPTION, or Size per part) one InputBox per part and property
 ├─ 8. Overwrite       MsgBox, only if some parts already have the properties (Keep mode)
 ├─ 9. Summary         MsgBox OK / Cancel
 ├─ 10. Write & save   delete old properties (if chosen), write, read back, save each part
 └─ 11. Report         MsgBox with counts, log file, option to open the log
```

Questions for properties that are not chosen in step 3 are not asked. Weight has no question of its own (§5.3).

**Cancel** at any question up to step 9 stops the macro. Nothing is written or deleted in any file before step 10.

## 4. Dialogs

All dialogs are standard `InputBox` / `MsgBox` windows, titled *Add Part Properties*.

### 4.1 Properties to add (step 3)

InputBox listing the number of parts found and the skipped ones (with reasons), and:

```
1 = DESCRIPTION  (typed for each part)
2 = Material  (one value, with exceptions)
3 = Weight  (SW-Mass in kg, automatic)
4 = Size  (typed)
1234 = all four.  0 = none (only delete old properties).
```

Default `1234`. Anything other than the digits 0–4 (spaces and commas allowed) is asked again.

### 4.2 Old custom properties (step 4)

Asked only when at least one part already has Custom-tab properties.
MsgBox *Yes* = delete all existing custom properties first, *No* = keep (default), *Cancel* = stop. *Yes* is confirmed by a second warning (OK / Cancel, Cancel is the default). Choosing `0` in step 3 together with *No* means there is nothing to do; the macro says so and stops.

### 4.3 Material (step 5)

1. InputBox *Material for all parts*. Empty is not accepted.
2. MsgBox *Yes* = apply to all parts, *No* = with exceptions, *Cancel* = stop.
3. Exceptions: the parts are listed in numbered pages of 15 (`1  Bracket.SLDPRT`, already picked parts marked `*`). The user types the numbers of the exceptions – e.g. `3, 7, 9-11` – or leaves the box empty for the next page. Invalid entries are reported and the page is asked again.
4. For each picked part (highlighted and zoomed to in the graphics area) an InputBox asks its material. The default is the last exception material typed, so a group of parts with the same material only needs Enter. Empty (or the common material) = no exception.

### 4.4 Size (step 6)

MsgBox *Yes* = same value for all parts, *No* = type per part, *Cancel* = stop. With *Yes* an InputBox asks the value, default `-` (empty = `-`).

### 4.5 Per part (step 7)

For every part, in tree order, the part is selected and zoomed to, then one InputBox per property (DESCRIPTION, then Size when typed per part) shows:

- part number of total, file name, sub-assembly path, configuration;
- the existing value of the property;
- the text box pre-filled with the existing value (Size: `-` when there is none).

Input:
- a value + OK – store it;
- empty + OK – keep the existing value (a part without a Size gets `-`);
- `<` + OK – back to the previous part;
- Cancel – a MsgBox: *Yes* = skip the remaining parts (keep their values) and continue, *No* = stop the macro, *Cancel* = back to this part.

With *Delete all existing custom properties first*, a kept value is written back after the old properties are deleted.

### 4.6 Overwrite (step 8)

Only in Keep mode, and only when some parts already have one of the chosen properties: *Yes* = overwrite existing values, *No* = keep them and only add missing properties, *Cancel* = stop.

### 4.7 Summary (step 9)

MsgBox with what will be done (delete mode in capitals), the chosen properties, and the planned DESCRIPTION / Material / Size of the first 12 parts. The full plan for every part is written to the log. *OK* writes, *Cancel* stops.

Parts are always saved after writing.

## 5. Property details

All properties are written to the Custom tab with
`ModelDocExtension.CustomPropertyManager("").Add3(Name, Type, Value, swCustomPropertyReplaceValue)`
(or `swCustomPropertyOnlyIfNew` when *Overwrite existing values* is off).
After writing, the value is read back with `Get6` and the resolved value is compared with what was expected.

### 5.0 Deleting old properties
- Only with *Delete all existing custom properties first*.
- For each processed part: read all names with `CustomPropertyManager("").GetNames`, delete each with `Delete2`. The *Configuration Specific* tab is not touched.
- Every deleted property (name, type, value/expression) is written to the log first, so it can be re-entered by hand if needed.
- Deletion and writing happen in memory, before one save. If any write fails, the deleted properties are put back and the part is **not saved**, so its file on disk is unchanged.

### 5.1 DESCRIPTION
- Type: `swCustomInfoText`.
- Value: exactly what the user typed. Leading/trailing spaces are trimmed.

### 5.2 Material
- Type: `swCustomInfoText`.
- Value: plain text typed by the user – the common material, or the part's own value if it was picked as an exception.
- It is **not** linked to the SOLIDWORKS material (`SW-Material`) and does **not** change the material used for mass.

### 5.3 Weight
- Added **automatically to every part**, if chosen. No question.
- Type: equation written as a text expression. The SOLIDWORKS API has no separate *Equation* member in `swCustomInfoType_e`, so the expression is written with `swCustomInfoText`; SOLIDWORKS evaluates the linked value.
- Text expression: `"SW-Mass@<PartFileName>.SLDPRT"` — e.g. `"SW-Mass@Bracket.SLDPRT"`.
  This is a live link: the value updates whenever the part changes.
- **Units – must be kg.** `SW-Mass` is shown in the part's *mass-property* unit (Tools > Options > Document Properties > Units > Mass). The macro finds that unit by comparing the resolved `SW-Mass` with the real mass from `MassProperty.Mass` (always kg), then writes:
  | Part mass unit | Expression written |
  |---|---|
  | kg | `"SW-Mass@Part.SLDPRT"` |
  | g | `"SW-Mass@Part.SLDPRT"/1000` |
  | mg | `"SW-Mass@Part.SLDPRT"/1000000` |
  | lb | `"SW-Mass@Part.SLDPRT"*0.45359237` |
  | oz | `"SW-Mass@Part.SLDPRT"*0.028349523125` |
- The written expression is read back and its resolved value is checked against the real mass in kg. If the plain form does not evaluate, the `=` form (`="SW-Mass@Part.SLDPRT"/1000`) is tried.
- If neither form evaluates, the plain link is kept and the log warns that this part shows its own unit and that its mass unit should be set to kilograms. The macro does not change document unit settings.
- If the part's mass is 0 (no material), the plain link is written and the log says the unit could not be checked.
- Decimal places follow the part's mass-unit decimal setting.

### 5.4 Size
- Type: `swCustomInfoText`.
- Value: typed by the user – either one value for all parts (default `-`) or one value per part (§4.3, §4.4).
- The macro does not calculate the size and does not check the text. An empty value is written as `-`.

## 6. Writing and saving

- Each part is taken from the assembly's memory (`Component2.GetModelDoc2`); no extra windows are opened or closed.
- Old properties are deleted (if chosen), the new ones written and read back, then the part is saved with `Save3(swSaveAsOptions_Silent)`.
- The top assembly is rebuilt at the end so the BOM refreshes. It is **not** saved automatically (it was not changed).
- A failure on one part (open, delete, write, verify or save) is logged and the macro continues with the next part.

## 7. Report

- Final message: number of parts updated, skipped and failed, with a button to open the log.
- Log file `AddPartProperties_<yyyymmdd_hhmmss>.log` next to the assembly: one line per part per property with the old value, new value and result, plus every deleted property.

## 8. Acceptance tests

| # | Setup | Expected |
|---|---|---|
| 1 | Run with a part document active | Message "Open an assembly first...", nothing changed. |
| 2 | Assembly with 3 parts, one used 4 times | DESCRIPTION is asked 3 times, not 6. |
| 3 | Part inside a sub-assembly 2 levels deep, in the assembly folder | Part is processed; sub-assembly gets no properties. |
| 4 | Part stored outside the assembly folder | Skipped, listed in *Show skipped* and the log, file unchanged. |
| 5 | Virtual component | Skipped, not modified. |
| 6 | Part checked in to PDM | Skipped, not modified. |
| 7 | Properties: `3` (only Weight) | No Material, Size or per-part questions; only `Weight` is written. |
| 8 | Properties: `0`, old properties: *No* (keep) | "Nothing to do" message, nothing changed. |
| 9 | Part with 5 old properties, *Delete all*, all 4 chosen | Custom tab has exactly the 4 properties; the 5 old ones are in the log. Configuration-specific properties unchanged. |
| 10 | *Delete all*, empty DESCRIPTION for a part that has one | That DESCRIPTION is kept; other old properties are deleted. |
| 11 | Material = `S355`, *Yes* (all parts) | All parts show `Material = S355`. |
| 12 | Material = `S355`, exceptions `2, 4`, typed `SS304` | 3 parts show `S355`, parts 2 and 4 show `SS304`. |
| 13 | Part with mass unit kg, mass 2.5 kg | `Weight` evaluates to `2.5`. |
| 14 | Part with mass unit g, mass 2500 g | `Weight` evaluates to `2.5`. |
| 15 | Change part geometry after run, rebuild | `Weight` value updates. |
| 16 | Size: *Yes* (same for all), keep default | Every part shows `Size = -`. |
| 17 | Size: *No* (per part), type `PL 10` for one part | Size is asked for each part; that part shows `PL 10`. |
| 18 | Existing `DESCRIPTION`, keep old properties, overwrite *No* | Existing value kept. |
| 19 | Cancel in the 3rd part's DESCRIPTION, then *No* (stop) | No file modified (check file dates). |
| 20 | `<` in part 3's DESCRIPTION | Part 2 is asked again. |
| 21 | Exceptions page: type `x, 99` | "Not understood" message, the page is asked again. |
| 22 | Toolbox fastener in the assembly | Not listed, not modified. |
| 23 | Suppressed component | Not listed; reported as skipped. |

## 9. Decisions

- Size is typed (or `-` for all); it is not calculated.
- Weight is automatic on every part, in kg, as an equation linked to `SW-Mass`.
- Material is plain typed text: one value for all, with optional exceptions.
- Only local part files in the assembly's folder are changed; virtual components and PDM files are skipped.
- Properties go only to the Custom tab.
- Property names are exactly `DESCRIPTION`, `Material`, `Weight`, `Size`.
- At the start the user picks which properties to add and whether to delete old custom properties.
