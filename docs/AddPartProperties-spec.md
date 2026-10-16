# Spec – Add Part Properties macro

**Status:** Draft for review
**Macro name:** `AddPartProperties.swp` (entry point `main`)
**Target:** SOLIDWORKS 2020 or later, VBA macro, same build approach as `SafeRename` (`src/` + `tools/build_swp.py` → `dist/`)

## 1. Purpose

Add four custom properties to every part used in the active assembly, so the BOM can show them without editing each part by hand.

| # | Property name | Type | Value / Text Expression | Who supplies it | Scope prompt |
|---|---|---|---|---|---|
| 1 | `DESCRIPTION` | Text | Typed by the user | Asked **once per part** | — (always every part) |
| 2 | `Material` | Text | Typed by the user | Asked **once**, same value for every part in scope | All parts / pick parts |
| 3 | `Weight` | Equation | Linked to `SW-Mass`, in **kg** | Calculated by SOLIDWORKS | All parts / pick parts |
| 4 | `Size` | Text | Bounding box `L x W x H` (see §5.4) | Calculated by the macro | All parts / pick parts *(see open question Q1)* |

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
| Virtual components | Optional – see Q3. Default: **included** (they are saved inside the assembly). |
| Read-only files / files open by another user / checked-in PDM files | Cannot be saved. |
| Envelope components | Not real parts. |

Lightweight components are resolved before processing.

## 3. Workflow

```
Run macro
 ├─ 1. Checks          active doc is a saved assembly; otherwise message and exit
 ├─ 2. Collect parts   walk the component tree, de-duplicate by file path, apply skip rules
 ├─ 3. Material        ask value once  → ask scope (All / Pick)
 ├─ 4. Weight          ask scope (All / Pick)
 ├─ 5. Size           ask scope (All / Pick)
 ├─ 6. DESCRIPTION     one dialog per part
 ├─ 7. Review          summary table, user confirms
 ├─ 8. Write & save    write properties, save each changed part
 └─ 9. Report          summary message + log file
```

The user can **Cancel** at any step up to 7. Nothing is written to any file before step 8.

## 4. Dialogs

### 4.1 Material value (step 3)

- Caption: *Material – applied to all selected parts*
- One text box: **Material**. Pre-filled with the most common existing `Material` value among the collected parts, if any; otherwise empty.
- Buttons: **Next**, **Skip this property**, **Cancel**.
- An empty value is not allowed with **Next** (use *Skip this property* instead).

### 4.2 Scope dialog (steps 3, 4, 5)

Shown once for each of Material, Weight and Size.

- Caption: *Apply `<property>` to…*
- Two option buttons:
  - **All parts** (default) – every collected part.
  - **Pick individual parts** – shows a checklist of all collected parts (file name, plus sub-assembly path in grey). Buttons *Select all* / *Select none*. All are ticked when the list first opens.
- When a row in the checklist is clicked, that component is selected/highlighted in the SOLIDWORKS graphics area so the user can see which part it is.
- Buttons: **Next**, **Back**, **Cancel**.
- Picking zero parts is allowed and means "do not write this property".

### 4.3 DESCRIPTION, one per part (step 6)

Shown once for every collected part, in tree order.

- Caption: *DESCRIPTION – part `n` of `N`*
- Shows: part file name, full path, configuration in use.
- The part is selected and zoomed to (`ModelDocExtension.SelectByID2` + *Zoom to Selection*) so the user sees it while typing. The view is restored afterwards.
- Text box **DESCRIPTION**, pre-filled with the part's existing `DESCRIPTION` value, if any.
- Buttons:
  - **Next** – store the value and go to the next part.
  - **Back** – previous part.
  - **Skip** – leave this part's DESCRIPTION unchanged.
  - **Skip remaining** – leave all remaining parts unchanged and go to Review.
  - **Cancel** – stop the macro, nothing written.
- Enter = Next.

### 4.4 Review (step 7)

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
- Value: exactly what the user typed, same for every part in scope.
- Note: this is a plain text property. It does **not** change the part's SOLIDWORKS material (the one used for mass). If a part already has `Material` linked to `"SW-Material@<part>.SLDPRT"`, the link is replaced by the typed text – the Review screen warns about this.

### 5.3 Weight
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
- Value: overall bounding box of the part, largest dimension first: `L x W x H`, e.g. `250 x 120 x 6`.
- Calculated with `PartDoc.GetPartBox(True)` (tight box aligned to the part's axes), converted from metres to **mm**, rounded to **1 decimal**, trailing `.0` dropped.
- This is a static value: it does **not** update if the part changes. Re-run the macro to refresh it. (See Q2 for a live alternative.)

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
| 4 | Material = `S355`, scope All | All parts show `Material = S355`. |
| 5 | Weight scope Pick, 2 of 5 parts ticked | Only those 2 parts get `Weight`. |
| 6 | Part with mass unit kg, mass 2.5 kg | `Weight` evaluates to `2.5`. |
| 7 | Part with mass unit g, mass 2500 g | `Weight` evaluates to `2.5`. |
| 8 | Change part geometry after run, rebuild | `Weight` value updates; `Size` does not. |
| 9 | 100 x 50 x 10 mm block | `Size = 100 x 50 x 10`. |
| 10 | Existing `DESCRIPTION`, Overwrite off | Existing value kept. |
| 11 | Cancel in the 3rd DESCRIPTION dialog | No file modified (check file dates). |
| 12 | One read-only part | Shown as skipped in Review and the log; others are saved. |
| 13 | Toolbox fastener in the assembly | Not listed, not modified. |
| 14 | Suppressed component | Not listed; reported as skipped. |

## 9. Open questions

- **Q1 – Size scope.** The request asks for the *apply to all / pick* choice on properties 2–3. This spec also gives it to Size (4). Confirm, or Size goes to all parts always.
- **Q2 – Size definition.** Is "Size" the bounding box `L x W x H` in mm? Alternatives: stock size typed by the user (e.g. `PL 10`), or a live link to the Bounding Box feature properties (needs a Bounding Box feature in every part). Also confirm the order (largest first) and separator.
- **Q3 – Virtual components.** Include them (default) or skip them?
- **Q4 – Configurations.** File-level Custom tab (default) or configuration-specific properties? With configuration-specific, should each configuration get its own DESCRIPTION?
- **Q5 – Material source.** Should Material default to the SOLIDWORKS material name (`SW-Material`) instead of a typed value, or only pre-fill the text box with it?
- **Q6 – Property names.** Exact spelling/case: `DESCRIPTION`, `Material`, `Weight`, `Size` – or match an existing BOM template (e.g. `Description`)?
