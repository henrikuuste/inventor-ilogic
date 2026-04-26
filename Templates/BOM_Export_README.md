<!-- Copyright (c) 2026 Henri Kuuste -->
# Excel BOM template format

## Layout

1. **Rows above the table** — title block / assembly data. Use the same `{{...}}` placeholders as in the rule (e.g. `{{Project}}`, `{{Part Number}}`, `{{Revision Number}}`). In this area, `BOM.*` resolves to an empty value.

2. **Header row** — human-readable column titles (Jrk, Detaili nr, Tk., …) with **no** `{{` in the cells (recommended).

3. **Mapping row** (required) — this is the **last** row in the used range that has **2 or more** cells containing the string `{{`. In each column, one of:
   - **Property** — `{{Part Number}}`, `{{Custom.Thickness}}`, `{{BOM.Qty}}`, `{{Phys.Mass}}`, …
   - **Script expression** — `{{=BOM.Qty * Custom.Thickness}}` (evaluated in the rule, result written as a number)
   - **Excel formula** — a cell that starts with `=` and does **not** contain `{` (e.g. `=D7*F7`). The rule copies `FormulaR1C1` to each data row. Prefer R1C1-style relative logic so the same formula works on every line.

4. The rule **overwrites the mapping row** with the first BOM line and **inserts** extra rows for additional lines.

## BOM data

The exporter first tries **Structured** view so export follows structured BOM ordering and item numbering for the active model state.  
If Structured BOM is unavailable, it automatically falls back to **Parts Only**.

**Prefixes**

| Form | Source |
|------|--------|
| (none) | Design Tracking Properties (e.g. `Part Number`, `Description`, `Project`) |
| `Custom.*` | Inventor User Defined Properties |
| `Phys.*`, `Phy.*`, `Physical.*` | Mass, Area, Volume from component mass properties |
| `BOM.*` | `Item` / `ItemNumber` (as text in cells, numeric in `{{=}}`), `Qty`, `TotalQty`, `UnitQty` |
| `File.*` | File-system metadata of the source document (created/modified dates, path/name) |
| `Drawing.*` | Associated 1:1 drawing metadata (empty if no associated 1:1 drawing found) |
| `Summary.*` | Inventor Summary Information property set |
| `DocSummary.*` | Inventor Document Summary Information property set |
| `Vault.*` | Vault-style aliases + fallback lookup in common iProperty sets |

## Common placeholders

### `BOM.*` placeholders (built in)

These are always available from the BOM row context:

| Placeholder | Meaning |
|-------------|---------|
| `{{BOM.Item}}` | BOM item number |
| `{{BOM.ItemNumber}}` | Alias of `BOM.Item` |
| `{{BOM.Qty}}` | Row quantity (`ItemQuantity`) |
| `{{BOM.Quantity}}` | Alias of `BOM.Qty` |
| `{{BOM.ItemQty}}` | Alias of `BOM.Qty` |
| `{{BOM.TotalQty}}` | Total quantity (`TotalQuantity`) |
| `{{BOM.UnitQty}}` | Unit/base quantity (currently same source as row quantity in this exporter) |

### No-prefix placeholders (Design Tracking Properties)

No prefix means the exporter reads from Inventor's **Design Tracking Properties** set.
Use the property name exactly as it appears in iProperties.

Common standard examples:

| Placeholder | Typical source field |
|-------------|----------------------|
| `{{Part Number}}` | Part Number |
| `{{Description}}` | Description |
| `{{Project}}` | Project |
| `{{Revision Number}}` | Revision Number |
| `{{Stock Number}}` | Stock Number |
| `{{Designer}}` | Designer |
| `{{Engineer}}` | Engineer |
| `{{Authority}}` | Authority |
| `{{Cost Center}}` | Cost Center |

### Common date/author placeholders (file + author workflows)

| Placeholder | Source | Notes |
|-------------|--------|-------|
| `{{File.ModifiedDate}}` | File system | Last write time of the source IPT/IAM/other file |
| `{{File.CreatedDate}}` | File system | File creation time |
| `{{File.Name}}` | File system | File name with extension |
| `{{File.Path}}` | File system | Full file path |
| `{{Summary.Author}}` | Summary iProperties | Standard document author |
| `{{Summary.Last Saved By}}` | Summary iProperties | Last user who saved document |
| `{{Summary.Creation Time}}` | Summary iProperties | Document creation timestamp (if populated) |
| `{{Summary.Last Save Time}}` | Summary iProperties | Last save timestamp (if populated) |
| `{{Designer}}` | Design Tracking | Common standard design-role field |
| `{{Engineer}}` | Design Tracking | Common standard design-role field |

### Associated 1:1 drawing placeholders (`Drawing.*`)

These use the same association logic as the 1:1 drawing workflow (`BB_SourcePartNumber` + `BB_DrawingType = 1:1`).
If no drawing is associated, metadata values resolve to empty and `Drawing.Exists` resolves to `False`.

| Placeholder | Meaning |
|-------------|---------|
| `{{Drawing.Exists}}` | `True` if associated 1:1 drawing found, otherwise `False` |
| `{{Drawing.FileName}}` | Associated drawing file name |
| `{{Drawing.Name}}` | Alias of `Drawing.FileName` |
| `{{Drawing.Path}}` | Full path to associated drawing |
| `{{Drawing.Description}}` | Drawing Design Tracking `Description` |
| `{{Drawing.PartNumber}}` | Drawing Design Tracking `Part Number` |

### Common Vault-style placeholders

These are useful when Vault metadata is synchronized into iProperties.

| Placeholder | Behavior |
|-------------|----------|
| `{{Vault.Revision}}` / `{{Vault.RevisionNumber}}` | Maps to Design Tracking `Revision Number` |
| `{{Vault.CheckedBy}}` | Maps to Design Tracking `Checked By` |
| `{{Vault.Designer}}` | Maps to Design Tracking `Designer` |
| `{{Vault.ModifiedDate}}` | Maps to file modified date (`File.ModifiedDate`) |
| `{{Vault.CreatedDate}}` | Maps to file created date (`File.CreatedDate`) |
| `{{Vault.<AnyName>}}` | Fallback lookup in Design Tracking, Summary, DocSummary, then Custom |

If a field is not available in your current environment/workflow, it resolves to an empty string.

### Quick practical examples

- Header block: `{{Project}}`, `{{Revision Number}}`, `{{Part Number}}`
- Header dates/authors: `{{File.ModifiedDate}}`, `{{Summary.Author}}`, `{{Summary.Last Saved By}}`
- Table columns: `{{BOM.Item}}`, `{{Part Number}}`, `{{Description}}`, `{{BOM.Qty}}`
- With custom/physical: `{{Custom.Thickness}}`, `{{Phys.Mass}}`
- Drawing-aware table columns: `{{Drawing.Exists}}`, `{{Drawing.FileName}}`

## Template selection UX

When running **Ekspordi BOM**:

1. Exporter first looks for templates in `VaultRoot\\Templates` (local workspace root + `Templates`).
2. If `.xlsx` files are found there, a list is shown first.
3. User can choose one from list or click **Sirvi...** (Browse) to pick any file.
4. If nothing is found in the preferred folder, exporter opens file browser directly.
5. In the template selection window, there is a **Detailne logi** checkbox (off by default).

### Default template preselection

Based on active model state (case-insensitive):

- contains `karkass` or `puit` -> preselect first template containing `puiduspets`
- contains `poroloon` -> preselect first template containing `poroloon`

If no match is found, first template in the list is selected by default.

## Output file defaults

- Output save dialog opens in the source assembly folder.
- Default file name is:
  - `<sourceAssemblyBaseName>_<templateBaseName>.xlsx`
  - Example: `000146_puiduspets.xlsx`

## Example (conceptual)

| (assembly block)     |  |
|------------------------|--|
| Project: `{{Project}}` |  |

| Jrk  | P/N              | Nimetus         | Tk.          | Paksus                  |
|------|------------------|-----------------|--------------|-------------------------|
| `{{BOM.Item}}` | `{{Part Number}}` | `{{Description}}` | `{{BOM.Qty}}` | `{{Custom.Thickness}}` |

## Files

Build your template in Excel and save it next to the assembly or in a known folder. Run **Ekspordi BOM** from a top-level assembly with the correct **Model State** active, then pick the template and the output file.

## Batch export

Use `BOMExportLib.ExportBatch` with a list of `ExportConfig` (model state name, template path, output path) — see the commented example at the end of `Koost/Ekspordi BOM.vb`.
