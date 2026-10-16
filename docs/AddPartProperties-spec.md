# Spec – Add Part Properties macro

**Status:** Draft for review
**Macro name:** `AddPartProperties.swp` (entry point `main`)
**Target:** SOLIDWORKS 2020 or later, VBA macro, same build approach as `SafeRename` (`src/` + `tools/build_swp.py` → `dist/`)

## 1. Purpose

Add four custom properties to every part used in the active assembly, so the BOM can show them without editing each part by hand.

| # | Property name | Type | Value / Text Expression | Who supplies it | Scope prompt |
|---|---|---|---|---|---|
| 1 | `DESCRIPTION` | Text | Typed by the user | Asked **once per part** | — (always every part) |
| 2 | `Material` | Text | Typed by the user | Asked **once** for all parts; the user can pick **exceptions** and type a different material for them | All parts / all parts with exceptions |
| 3 | `Weight` | Equation | Linked to `SW-Mass`, in **kg** | Automatic, no prompt | — (always every part) |
| 4 | `Size` | Text | Typed by the user, or `-` | Either **one value for all parts** (default `-`) or **typed per part** | Same for all / per part |

Property names are written exactly as shown above (case included).

## 2. Scope

### In scope
- The active document must be an assembly (`.sldasm`) that has been saved at least once.
- All parts in the assembly, **including parts inside sub-assemblies** at any depth.
- Each part file is processed **once**, however many instances it has. Parts are identified by full file path.
- Properties are written to the part file's **Custom** tab (file-level, configuration name `""`).

### Out of scope
- Sub-assemblies themselves and the top assembly do not get the properties.
- Drawings are not touched.
- Assembly-level cut lists / weldment cut-list items.

### Skipped automatically (listed in the log and the final summary)
| Case | Reason |
|---|---|
| Suppressed components | No model loaded. |
| Toolbox parts (`IsToolboxPart`) | Shared library files; must not be changed. |
| Virtual components | Optional – see Q1. Default: **included** (they are saved inside the assembly). |
| Read-only files / files open by another user / checked-in PDM files | Cannot be saved. |
| Envelope components | Not real parts. |

Lightweight components are resolved before processing.

## 3. Workflow

```
Run macro
 ├─ 1. Checks          active doc is a saved assembly; otherwise message and exit
 ├─ 2. Collect parts   walk the component tree, de-duplicate by file path, apply skip rules
 ├─ 3. Material        ask value once → All parts / All parts with exceptions
 │                     (exceptions: pick parts, type their material)
 ├─ 4. Size            Same value for all (default "-") / Type per part
 ├─ 5. Per-part dialog DESCRIPTION (always) + Size (only if "Type per part")
 ├─ 6. Review          summary table, user confirms
 ├─ 7. Write & save    write properties (Weight added automatically), save each part
 └─ 8. Report          summary message + log file
```

Weight has no dialog: it is always added to every part (§5.3).

The user can **Cancel** at any step up to 6. Nothing is written to any file before step 7.

## 4. Dialogs

### 4.1 Material (step 3)

- Caption: *Material*
- Text box **Material for all parts**. Pre-filled with the most common existing `Material` value among the collected parts, if any; otherwise empty.
- Two option buttons:
  - **Apply to all parts** (default) – every part gets this material.
  - **Apply to all parts, with exceptions** – opens the exceptions list below.
- Buttons: **Next**, **Skip this property**, **Cancel**.
- An empty value is not allowed with **Next** (use *Skip this property* instead).

**Exceptions list** (only with *Apply to all parts, with exceptions*):

- A list of all collected parts with three columns: tick box *Exception*, *Part* (file name, plus sub-assembly path in grey), *Material*.
- Every row starts unticked, showing the common material in grey.
- Ticking a row makes its *Material* cell editable; the user types that part's material.
- To give several parts the same different material at once: tick them, type the value in **Material for ticked parts** and click **Set**.
- Clicking a row selects/highlights that component in the SOLIDWORKS graphics area so the user can see which part it is.
- A ticked row with an empty *Material* is not allowed with **Next**.
- Buttons: **Next**, **Back**, **Cancel**.

Result: unticked parts get the common material; ticked parts get their own value.

### 4.2 Size (step 4)

- Caption: *Size*
- Two option buttons:
  - **Same value for all parts** (default) – text box pre-filled with `-`. The user can keep `-` or type another value.
  - **Type per part** – Size is asked in the per-part dialog (§4.3).
- Buttons: **Next**, **Back**, **Cancel**.

### 4.3 Per-part dialog (step 5)

Shown once for every collected part, in tree order.

- Caption: *Part `n` of `N`*
- Shows: part file name, full path, configuration in use.
- The part is selected and zoomed to (`ModelDocExtension.SelectByID2` + *Zoom to Selection*) so the user sees it while typing. The view is restored afterwards.
- Text box **DESCRIPTION**, pre-filled with the part's existing `DESCRIPTION` value, if any.
- Text box **Size** – only when *Type per part* was chosen in §4.2. Pre-filled with the part's existing `Size` value, or `-` if there is none.
- Buttons:
  - **Next** – store the values and go to the next part.
  - **Back** – previous part.
  - **Skip** – leave this part's DESCRIPTION (and Size) unchanged.
  - **Skip remaining** – leave all remaining parts unchanged and go to Review.
  - **Cancel** – stop the macro, nothing written.
- Enter = Next.

### 4.4 Review (step 6)

A table, one row per part, columns: *Part*, *DESCRIPTION*, *Material*, *Weight*, *Size*, *Note*.

- Cells show the value that will be written, or `(unchanged)`.
- *Note* shows warnings, e.g. "existing value will be overwritten", "read-only – will be skipped".
- Checkbox **Overwrite existing values** (default **on**). When off, a property that already exists on a part is left as it is.
- Checkbox **Save parts after writing** (default **on**).
- Buttons: **Apply**, **Back**, **Cancel**.

## 5. Property details

All properties are written with
`ModelDocExtension.CustomPropertyManager("").Add3(Name, Type, Value, swCustomPropertyReplaceValue)`
(or `swCustomPropertyOnlyIfNew` when *Overwrite existing values* is off).
After writing, the value is read back with `Get6` and the resolved value is compared with what was expected.

### 5.1 DESCRIPTION
- Type: `swCustomInfoText`.
- Value: exactly what the user typed. Leading/trailing spaces are trimmed.

### 5.2 Material
- Type: `swCustomInfoText`.
- Value: the common material typed by the user, or the part's own value if it was ticked as an exception.
- Note: this is a plain text property. It does **not** change the part's SOLIDWORKS material (the one used for mass). If a part already has `Material` linked to `"SW-Material@<part>.SLDPRT"`, the link is replaced by the typed text – the Review screen warns about this.

### 5.3 Weight
- Added **automatically to every part**. No dialog.
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
- Value: typed by the user – either one value for all parts (default `-`) or one value per part (§4.2, §4.3).
- The macro does not calculate the size and does not check the text. An empty value is written as `-`.

## 6. Writing and saving

- Each part is opened (or taken from memory if already loaded) with `OpenDoc6`, silently.
- Properties are written, then the part is saved with `Save3(swSaveAsOptions_Silent)`.
- Parts the macro opened itself are closed afterwards; parts that were already open stay open.
- The top assembly is rebuilt at the end so the BOM refreshes. It is **not** saved automatically (it was not changed).
- A failure on one part (open, write, verify or save) is logged and the macro continues with the next part.

## 7. Report

- Final message: number of parts updated, skipped and failed, with a button to open the log.
- Log file `AddPartProperties_<yyyymmdd_hhmmss>.log` next to the assembly: one line per part per property with the old value, new value and result.

## 8. Acceptance tests

| # | Setup | Expected |
|---|---|---|
| 1 | Run with a part document active | Message "Open an assembly first", nothing changed. |
| 2 | Assembly with 3 parts, one used 4 times | DESCRIPTION dialog appears 3 times, not 6. |
| 3 | Part inside a sub-assembly 2 levels deep | Part is processed; sub-assembly gets no properties. |
| 4 | Material = `S355`, Apply to all | All parts show `Material = S355`. |
| 5 | Material = `S355`, with exceptions: 2 of 5 parts ticked, typed `SS304` | 3 parts show `S355`, the 2 ticked show `SS304`. |
| 6 | Any run | Every part gets `Weight`; no Weight dialog is shown. |
| 7 | Part with mass unit kg, mass 2.5 kg | `Weight` evaluates to `2.5`. |
| 8 | Part with mass unit g, mass 2500 g | `Weight` evaluates to `2.5`. |
| 9 | Change part geometry after run, rebuild | `Weight` value updates. |
| 10 | Size: Same value for all, keep default | Every part shows `Size = -`. |
| 11 | Size: Type per part, type `PL 10` for one part | Size box appears in the per-part dialog; that part shows `PL 10`. |
| 12 | Existing `DESCRIPTION`, Overwrite off | Existing value kept. |
| 13 | Cancel in the 3rd per-part dialog | No file modified (check file dates). |
| 14 | One read-only part | Shown as skipped in Review and the log; others are saved. |
| 15 | Toolbox fastener in the assembly | Not listed, not modified. |
| 16 | Suppressed component | Not listed; reported as skipped. |

## 9. Open questions

Resolved: Size is typed (or `-` for all); Weight is automatic on every part, in kg; Material is one value for all with optional exceptions.

- **Q1 – Virtual components.** Include them (default) or skip them?
- **Q2 – Configurations.** File-level Custom tab (default) or configuration-specific properties? With configuration-specific, should each configuration get its own DESCRIPTION?
- **Q3 – Material source.** Should Material default to the SOLIDWORKS material name (`SW-Material`) instead of a typed value, or only pre-fill the text box with it?
- **Q4 – Property names.** Exact spelling/case: `DESCRIPTION`, `Material`, `Weight`, `Size` – or match an existing BOM template (e.g. `Description`)?
