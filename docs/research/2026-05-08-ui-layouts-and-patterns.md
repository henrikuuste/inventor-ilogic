---
date: 2026-05-08T10:38:00+03:00
researcher: Claude
git_commit: 18fe505346853d76c6a17a998ea713fac5999cb9
branch: main
repository: Inventor-Rules
topic: "UI Layouts, Pixel Specifications, and Construction Patterns"
tags: [research, codebase, winforms, ui, layout]
status: complete
last_updated: 2026-05-08
---

# Research: UI Layouts, Pixel Specifications, and Construction Patterns

**Date**: 2026-05-08 10:38 AM (UTC+3)
**Git Commit**: 18fe505346853d76c6a17a998ea713fac5999cb9
**Branch**: main

## Research Question

Current usage of UI layouts, manual pixel size specifications, and any patterns/methods used in UI construction in the codebase.

## Summary

The codebase uses **runtime-composed Windows Forms** with **absolute pixel positioning**. There is **no centralized UI layout library** or shared constants for WinForms dimensions. UI construction is distributed across ~25 rule files and 4 library modules. The dominant pattern is manual coordinate assignment using `.Left`, `.Top`, `.Width`, `.Height` properties with local layout variables (`currentY`, `yPos`, `labelWidth`, etc.) per file.

Key characteristics:
- **Modal `ShowDialog()`** is the default interaction pattern
- **Non-modal forms** with `DoEvents` loop used for Inventor viewport interaction
- **No `System.Drawing` types** — individual properties used instead (per AGENTS.md constraint)
- **`Tag` property** widely used for passing data between controls and handlers
- **Only one file** (`BoundingBoxStockLib.vb`) uses layout containers (`TableLayoutPanel`)

## Detailed Findings

### 1. Form Creation Patterns

#### Modal Dialog (Default Pattern)
Used in most dialogs. Form is created, controls added, `ShowDialog()` blocks until closed.

```vb
Dim frm As New System.Windows.Forms.Form()
frm.Text = "Dialog Title"
frm.Width = 400
frm.Height = 300
frm.StartPosition = FormStartPosition.CenterScreen
frm.FormBorderStyle = FormBorderStyle.FixedDialog
frm.MaximizeBox = False
frm.MinimizeBox = False

' Add controls...
frm.Controls.Add(lbl)
frm.Controls.Add(btnOK)

Dim result As DialogResult = frm.ShowDialog()
```

**Files using this pattern:**
- `Lib\BoundingBoxStockLib.vb`
- `Lib\BOMExportLib.vb`
- `Lib\ExcelReaderLib.vb`
- `Lib\ElementReleaseLib.vb`
- `Komponendid\Pinnalaotuse vaated.vb`
- `Joonised\*.vb` (all drawing-related dialogs)
- `Mõõdud.vb`
- `Loo detailid.vb`
- `Moodulid\Loo alusmoodul.vb`
- `Koost\Muutujad.vb`
- `Mustrid\Kordused keskelt.vb`

#### Non-Modal with DoEvents Loop
Used when Inventor viewport interaction is needed while form is open.

```vb
frm.Show()

Do While frm.Visible
    System.Windows.Forms.Application.DoEvents()
    System.Threading.Thread.Sleep(10)
Loop
```

**Files using this pattern:**
- `Koordinaadid.vb` (lines 444-450)
- `Katsetused\PlaceSupport.vb` (lines 1362-1368)
- `Katsetused\Moodulid\Test10_DisconnectSaveCheckin.vb`

#### Auto-Sized Form (Single Instance)
Only `BoundingBoxStockLib.vb` uses automatic sizing with layout panels.

```vb
frm.AutoSize = True
frm.AutoSizeMode = AutoSizeMode.GrowAndShrink
frm.Padding = New System.Windows.Forms.Padding(10)

Dim layout As New TableLayoutPanel()
layout.AutoSize = True
layout.AutoSizeMode = AutoSizeMode.GrowAndShrink
frm.Controls.Add(layout)
```

### 2. Pixel Size Specifications

#### Form Sizes (Width × Height)

| Size | Usage |
|------|-------|
| 300×350 | Small picker dialog (`PlaceSupport.vb`) |
| 350×180 | Compact mode selection (`ElementReleaseLib.vb`) |
| 380×420 | Medium dialog (`Koordinaadid.vb`) |
| 400×250 to 450×320 | Standard dialogs (`Joonised\*.vb`) |
| 470×500 to 550×580 | Large dialogs (`BOMExportLib.vb`, `Määra materjalide välimus.vb`) |
| 480×580 | Pattern dialog (`Kordused keskelt.vb`) |
| 500×200 to 500×360 | Medium dialogs |
| 600×400 | List selection dialog |
| 800×500 to 800×600 | DataGridView dialogs (`Loo 1-1 joonised.vb`, `Muutujad.vb`) |
| 900×700 to 1000×800 | Test/result dialogs |
| 950×680 | Large multi-section dialog (`Loo detailid.vb`) |

#### Common Control Sizes

| Control Type | Common Width | Common Height |
|--------------|--------------|---------------|
| OK/Cancel Button | 80-100 | 25-32 (commonly 28) |
| Pick/Browse Button | 30-40 | 23-26 |
| Action Button | 100-140 | 25-35 |
| Label | 80-400 | auto (or 20-48 for multi-line) |
| TextBox | 60-370 | auto |
| ComboBox | 80-200 | auto |
| NumericUpDown | 80 | auto |
| ListBox | 260-430 | 180-280 |
| DataGridView | 760-910 | 60-450 |

#### Layout Variables (Local, Per-File)

**`Mustrid\Kordused keskelt.vb` (lines 251-256):**
```vb
Dim yPos As Integer = 15
Dim labelWidth As Integer = 120
Dim controlLeft As Integer = 135
Dim controlWidth As Integer = 200
Dim btnWidth As Integer = 70
Dim rowHeight As Integer = 30
```

**`Koordinaadid.vb` (lines 165-166):**
```vb
Dim labelWidth As Integer = 100
Dim comboWidth As Integer = 90
```

**`Katsetused\PlaceSupport.vb` (lines 774-779):**
```vb
Dim yPos As Integer = 20
Dim leftCol As Integer = 20
Dim rightCol As Integer = 130
Dim labelWidth As Integer = 100
Dim controlWidth As Integer = 200
Dim rowHeight As Integer = 32
```

#### Common Margin/Padding Values

| Value | Usage |
|-------|-------|
| 10 | Left margin, form padding (`BoundingBoxStockLib`) |
| 12 | Left margin (`Pinnalaotuse vaated.vb`) |
| 15 | Left margin (many files) |
| 20 | Left margin (drawing dialogs) |
| +3 | Label vertical offset to align with text boxes |
| -2/-3 | Control negative offset for alignment |
| 35 | ListBox top when label at 10 |

### 3. Vertical Layout Pattern

Most dialogs use a **`currentY` or `yPos` cursor** that increments after each row:

```vb
Dim currentY As Integer = 10

' Header label
Dim lblHeader As New Label()
lblHeader.Top = currentY
frm.Controls.Add(lblHeader)
currentY += 20

' Input row
Dim lbl As New Label()
lbl.Top = currentY + 3  ' +3 aligns with text box baseline
Dim txt As New TextBox()
txt.Top = currentY
frm.Controls.Add(lbl)
frm.Controls.Add(txt)
currentY += 25

' Next row...
```

Files using this pattern:
- `Joonised\Lisa vaated.vb`
- `Joonised\Lisa mõõdud.vb`
- `Joonised\Uuenda 1-1 joonis.vb`
- `Joonised\Uuenda lehe suurus.vb`
- `Joonised\Loo 1-1 joonised.vb`
- `Määra materjalide välimus.vb`
- `Loo detailid.vb`
- `Mõõdud.vb`
- `Koordinaadid.vb`
- `Mustrid\Kordused keskelt.vb`
- `Katsetused\PlaceSupport.vb`

### 4. Event Handler Patterns

#### Lambda with `Sub(s, e)`
Most common pattern for button clicks:

```vb
AddHandler btn.Click, Sub(s, e)
    frm.Tag = "ACTION"
    frm.DialogResult = DialogResult.OK
End Sub
```

**Files:** Most rule files and `ElementReleaseLib.vb`, `BOMExportLib.vb`

#### `AddressOf` with Named Handler
Used when handler needs control access via `Tag`:

```vb
cboAxis.Tag = frm
AddHandler cboAxis.SelectedIndexChanged, AddressOf OnAxisComboChanged

' Handler:
Public Sub OnAxisComboChanged(sender As Object, e As EventArgs)
    Dim cbo As ComboBox = CType(sender, ComboBox)
    Dim frm As Form = CType(cbo.Tag, Form)
    ' Access other controls via frm.Controls("name")
End Sub
```

**Files:** `BoundingBoxStockLib.vb`, `Koordinaadid.vb`

### 5. Tag Property Usage Patterns

#### Form.Tag for Return Value
```vb
frm.Tag = ReleaseMode.Cancelled  ' Initial value
AddHandler btnFull.Click, Sub(s, e)
    frm.Tag = ReleaseMode.FullModule
    frm.DialogResult = DialogResult.OK
End Sub

' After ShowDialog:
If result = DialogResult.OK Then
    Return CType(frm.Tag, ReleaseMode)
End If
```

**Files:** `ElementReleaseLib.vb`, `Mustrid\Kordused keskelt.vb`, `Katsetused\PlaceSupport.vb`

#### Form.Tag for State Object
```vb
frm.Tag = state  ' Custom state object
' Controls read state via:
Dim state As CoordState = CType(frm.Tag, CoordState)
```

**File:** `Koordinaadid.vb`

#### Control.Tag for Form Reference
```vb
btnOK.Tag = frm
txtInput.Tag = frm
' Handler retrieves form:
Dim frm As Form = CType(CType(sender, Control).Tag, Form)
```

**Files:** `Koordinaadid.vb`, `BoundingBoxStockLib.vb`

#### DataGridViewRow.Tag for Row Data
```vb
dgv.Rows(idx).Tag = expParam  ' Store object
' Retrieve:
Dim expParam As ExposedParameter = CType(row.Tag, ExposedParameter)
```

**File:** `Koost\Muutujad.vb`

### 6. UI-Related Library Modules

| Module | UI Components |
|--------|---------------|
| `Lib\BoundingBoxStockLib.vb` | `ShowConfigForm` — TableLayoutPanel, FlowLayoutPanel, AutoSize |
| `Lib\BOMExportLib.vb` | `SelectTemplateFromListOrBrowse`, `BrowseTemplateFile` — ListBox, FileDialogs |
| `Lib\ExcelReaderLib.vb` | `ShowConfigSelectionDialog` — ListBox + OK/Cancel |
| `Lib\ElementReleaseLib.vb` | `ShowModeSelectionDialog`, `ShowPlanConfirmationDialog`, `ShowCompletionSummary` |
| `Lib\SupportPatternLibrary.vb` | MessageBox only (no forms) |

### 7. Constants

**No WinForms pixel constants exist** in the codebase.

The only layout-related constants are in `Lib\CAMDrawingLib.vb` for Inventor drawing dimensions (in cm):

```vb
Public Const DEFAULT_DIMENSION_OFFSET As Double = 2.5   ' 25mm
Public Const DEFAULT_VIEW_GAP As Double = 1.5           ' 15mm
Public Const DEFAULT_SHEET_PADDING As Double = 0.5      ' 50%
Public Const DEFAULT_BORDER_PADDING As Double = 1.5     ' 15mm
Public Const DIMENSION_OFFSET As Double = DEFAULT_DIMENSION_OFFSET
```

## Code References

- `Lib\BoundingBoxStockLib.vb:407-545` - TableLayoutPanel usage (only layout container example)
- `Mustrid\Kordused keskelt.vb:244-564` - Comprehensive manual layout with local variables
- `Koordinaadid.vb:150-450` - Non-modal form with DoEvents, state via Tag
- `Loo detailid.vb:684-1288` - Large dialog with DataGridView
- `Komponendid\Pinnalaotuse vaated.vb:91-180` - Simple fixed dialog pattern
- `Lib\ElementReleaseLib.vb:160-216` - Modal dialog with enum return via Tag
- `Katsetused\PlaceSupport.vb:758-1260` - Configurable form position and modeless operation

## Architecture Documentation

### Current State
- **No shared UI toolkit** — each file constructs UI independently
- **No centralized constants** for WinForms dimensions
- **No layout managers** except one file (`BoundingBoxStockLib.vb`)
- **Absolute pixel positioning** is universal
- **Local variables** (`yPos`, `currentY`, `labelWidth`, etc.) used per-file for consistency within that file

### Patterns in Use
1. **Vertical cursor pattern** — `currentY += N` after each row
2. **Label alignment offset** — `lbl.Top = yPos + 3` next to text boxes
3. **Tag-based data passing** — avoids ByRef lambda closure issues
4. **AcceptButton/CancelButton** — standard OK/Cancel handling
5. **DialogResult on buttons** — `btnOK.DialogResult = DialogResult.OK`

### Constraints (from AGENTS.md)
- Must use fully qualified `System.Windows.Forms.*` types
- Cannot use `System.Drawing` types (`Size`, `Point`, `Color`, `Font`)
- Must use individual properties (`.Left`, `.Top`, `.Width`, `.Height`)
- Cannot use `ContextMenuStrip` or `ToolStripMenuItem`
- Cannot use ByRef parameters in lambda expressions

## Open Questions

1. Would centralizing common UI dimensions as constants improve maintainability?
2. Should a shared UI construction library be created to reduce code duplication?
3. Are the current pixel values appropriate for different DPI settings?
