# Unified UI Library Implementation Plan

## Overview

Create a unified UI management and creation library (`UILib.vb`) with supporting modules to standardize all dialog creation, picker handling, and viewport helpers across the codebase. This will ensure consistent UX, enable non-modal dialogs for workspace interaction, and provide responsive layouts that work well at 1920x1080 with up to 150% DPI scaling.

## Current State Analysis

### UI Construction (from research: `docs/research/2026-05-08-ui-layouts-and-patterns.md`)

- **No centralized UI library** — each of ~25 rule files and 4 library modules constructs UI independently
- **Absolute pixel positioning** is universal with local variables (`yPos`, `currentY`, `labelWidth`, etc.) per file
- **Modal `ShowDialog()`** is the dominant pattern (blocks Inventor interaction)
- **Only 3 files** use non-modal forms with `DoEvents` loop:
  - `Koordinaadid.vb` — live UCS preview with viewport interaction
  - `Katsetused/PlaceSupport.vb` — modeless settings, closes before picks
  - `Katsetused/Moodulid/Test10_DisconnectSaveCheckin.vb` — wait dialog
- **Only 1 file** (`Lib/BoundingBoxStockLib.vb`) uses layout panels (`TableLayoutPanel`)

### Picker Patterns

- `Lib/UtilsLib.vb` has shared helpers: `PickPoint`, `PickAxis`, `PickPlane`
- Other files duplicate picker wrappers: `Mustrid/Kordused keskelt.vb`, `Lib/BoundingBoxStockLib.vb`
- **No `ClientGraphics` or transient overlays** — only real model features used for preview
- Current workaround: close form → pick in Inventor → reopen form

### Estonian Strings

- **Extensive duplication** across 15+ files
- Common duplicated phrases:
  | Phrase | Occurrences |
  |--------|-------------|
  | `"Aktiivne dokument puudub."` | ~17 files |
  | `"Tühista"` (Cancel button) | ~20 files |
  | `"- ESC tühistamiseks"` | ~8 pick prompts |
  | `"Vali kõik"` / `"Tühjenda"` | ~6 files |
  | `"See reegel töötab ainult..."` | ~15 files |
- **Inconsistencies**: `"Loobu"` vs `"Tühista"` for cancel

### Key Discoveries

- `Koordinaadid.vb:444-451` — only example of full viewport interaction while form open
- `BoundingBoxStockLib.vb:407-545` — only example of `TableLayoutPanel` with `AutoSize`
- `UtilsLib.vb:449-477` — shared picker helpers that accept caller prompts
- `PlaceSupport.vb:1362-1369` — `DoEvents` loop pattern for non-modal forms
- No `ClientGraphics`, `InteractiveGraphics`, or `HighlightSet` usage anywhere

## Desired End State

### Architecture

```
Lib/
  UILib.vb              — Core form factory, non-modal management, layout helpers
  StringsLib.vb         — Estonian translations with format string support
  ViewportHelperLib.vb  — ClientGraphics transient overlays, highlights
  UtilsLib.vb           — Extended with unified picker integration
```

### Capabilities

1. **Non-Modal Forms**: All dialogs allow full Inventor interaction including `CommandManager.Pick` while open
2. **Responsive Layouts**: TableLayoutPanel content + Dock/Anchor form resizing, DPI-aware
3. **Centralized Strings**: Single source for all Estonian UI text
4. **Viewport Helpers**: Transient highlights for pick feedback, real features for complex previews
5. **Unified Pickers**: Consistent pick experience with automatic highlighting and ESC handling

### Verification

- [ ] All dialogs are non-modal (user can orbit, pan, pick while dialog open)
- [ ] All dialogs resize correctly at 100%, 125%, 150% DPI scaling
- [ ] All Estonian strings come from `StringsLib.vb`
- [ ] Picker operations show transient highlights during selection
- [ ] No hardcoded pixel values in dialog code (uses UILib helpers)

## What We're NOT Doing

1. **Custom control library** — We use standard WinForms controls, just with consistent construction
2. **MVVM/data binding patterns** — Too complex for iLogic context; keep simple event handlers
3. **Async/await patterns** — iLogic doesn't support; continue with `DoEvents` loop
4. **Multi-language support** — Estonian only; English stays in logs/code comments
5. **Automated UI testing** — Manual verification only
6. **Changing existing dialog functionality** — Migration preserves behavior, only changes construction

## Implementation Approach

Use a **hybrid layout strategy**:
- `TableLayoutPanel` for form content (rows of label+control pairs)
- `FlowLayoutPanel` for button bars
- `Dock` and `Anchor` for form-level resizing
- DPI scaling factors applied to base dimensions

Use a **hybrid viewport helper strategy**:
- `ClientGraphics` for lightweight transient markers (points, lines, highlights)
- Real model features (UCS, work planes) for complex previews that need persistence

---

## Phase 1: Foundation — Core UILib with Non-Modal Forms

### Overview

Create `Lib/UILib.vb` with core form factory, non-modal message loop management, and layout helpers.

### Changes Required

#### 1. Create UILib.vb

**File**: `Lib/UILib.vb`

**Contents**:

```vb
' UILib.vb - Unified UI Management Library
' Provides consistent form creation, non-modal management, and layout helpers

Public Module UILib
    
    ' ============================================================
    ' CONSTANTS - DPI-aware base dimensions
    ' ============================================================
    
    ' Form sizing (base values for 96 DPI / 100% scaling)
    Public Const FORM_WIDTH_SMALL As Integer = 350
    Public Const FORM_WIDTH_MEDIUM As Integer = 480
    Public Const FORM_WIDTH_LARGE As Integer = 800
    Public Const FORM_HEIGHT_SMALL As Integer = 200
    Public Const FORM_HEIGHT_MEDIUM As Integer = 400
    Public Const FORM_HEIGHT_LARGE As Integer = 600
    
    ' Control sizing
    Public Const BUTTON_WIDTH As Integer = 90
    Public Const BUTTON_HEIGHT As Integer = 28
    Public Const ROW_HEIGHT As Integer = 30
    Public Const LABEL_WIDTH As Integer = 120
    Public Const CONTROL_WIDTH As Integer = 200
    
    ' Spacing
    Public Const PADDING As Integer = 10
    Public Const SPACING As Integer = 6
    
    ' ============================================================
    ' NON-MODAL FORM MANAGEMENT
    ' ============================================================
    
    ' Active non-modal forms for cleanup
    Private m_ActiveForms As New List(Of System.Windows.Forms.Form)
    
    ''' <summary>
    ''' Shows a form as non-modal and runs a message loop until closed.
    ''' Allows full Inventor interaction including CommandManager.Pick.
    ''' </summary>
    Public Sub ShowNonModal(frm As System.Windows.Forms.Form)
        frm.TopMost = True
        m_ActiveForms.Add(frm)
        
        AddHandler frm.FormClosed, Sub(s, e)
            m_ActiveForms.Remove(frm)
        End Sub
        
        frm.Show()
        
        Do While frm.Visible
            System.Windows.Forms.Application.DoEvents()
            System.Threading.Thread.Sleep(10)
        Loop
    End Sub
    
    ''' <summary>
    ''' Shows a form as non-modal and returns immediately.
    ''' Caller must manage the message loop or use callbacks.
    ''' </summary>
    Public Sub ShowNonModalAsync(frm As System.Windows.Forms.Form)
        frm.TopMost = True
        m_ActiveForms.Add(frm)
        
        AddHandler frm.FormClosed, Sub(s, e)
            m_ActiveForms.Remove(frm)
        End Sub
        
        frm.Show()
    End Sub
    
    ''' <summary>
    ''' Pumps message queue once. Call repeatedly in picker loops.
    ''' </summary>
    Public Sub PumpMessages()
        System.Windows.Forms.Application.DoEvents()
        System.Threading.Thread.Sleep(5)
    End Sub
    
    ''' <summary>
    ''' Closes all active non-modal forms (cleanup on rule exit).
    ''' </summary>
    Public Sub CloseAllForms()
        For Each frm In m_ActiveForms.ToArray()
            If frm.Visible Then frm.Close()
        Next
        m_ActiveForms.Clear()
    End Sub
    
    ' ============================================================
    ' FORM FACTORY
    ' ============================================================
    
    ''' <summary>
    ''' Creates a standard tool window form with consistent styling.
    ''' </summary>
    Public Function CreateForm(title As String, Optional width As Integer = FORM_WIDTH_MEDIUM, Optional height As Integer = FORM_HEIGHT_MEDIUM) As System.Windows.Forms.Form
        Dim frm As New System.Windows.Forms.Form()
        
        frm.Text = title
        frm.Width = Scale(width)
        frm.Height = Scale(height)
        frm.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        frm.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        frm.MaximizeBox = False
        frm.MinimizeBox = True
        frm.Padding = New System.Windows.Forms.Padding(Scale(PADDING))
        frm.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        
        Return frm
    End Function
    
    ''' <summary>
    ''' Creates a resizable form with minimum size constraints.
    ''' </summary>
    Public Function CreateResizableForm(title As String, Optional width As Integer = FORM_WIDTH_MEDIUM, Optional height As Integer = FORM_HEIGHT_MEDIUM) As System.Windows.Forms.Form
        Dim frm As New System.Windows.Forms.Form()
        
        frm.Text = title
        frm.Width = Scale(width)
        frm.Height = Scale(height)
        frm.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        frm.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable
        frm.MaximizeBox = True
        frm.MinimizeBox = True
        frm.MinimumSize = New System.Drawing.Size(Scale(width * 0.8), Scale(height * 0.6))
        frm.Padding = New System.Windows.Forms.Padding(Scale(PADDING))
        frm.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        
        Return frm
    End Function
    
    ' ============================================================
    ' LAYOUT HELPERS
    ' ============================================================
    
    ''' <summary>
    ''' Creates a TableLayoutPanel for form content with standard column layout.
    ''' Column 0: Labels (auto-width), Column 1: Controls (fill)
    ''' </summary>
    Public Function CreateContentPanel() As System.Windows.Forms.TableLayoutPanel
        Dim panel As New System.Windows.Forms.TableLayoutPanel()
        
        panel.Dock = System.Windows.Forms.DockStyle.Fill
        panel.AutoSize = True
        panel.ColumnCount = 2
        panel.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.AutoSize))
        panel.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100))
        panel.Padding = New System.Windows.Forms.Padding(0)
        
        Return panel
    End Function
    
    ''' <summary>
    ''' Creates a FlowLayoutPanel for button bars (right-aligned).
    ''' </summary>
    Public Function CreateButtonPanel() As System.Windows.Forms.FlowLayoutPanel
        Dim panel As New System.Windows.Forms.FlowLayoutPanel()
        
        panel.Dock = System.Windows.Forms.DockStyle.Bottom
        panel.FlowDirection = System.Windows.Forms.FlowDirection.RightToLeft
        panel.AutoSize = True
        panel.Padding = New System.Windows.Forms.Padding(0, Scale(SPACING), 0, 0)
        
        Return panel
    End Function
    
    ''' <summary>
    ''' Adds a label+control row to a TableLayoutPanel.
    ''' </summary>
    Public Sub AddRow(panel As System.Windows.Forms.TableLayoutPanel, labelText As String, control As System.Windows.Forms.Control, Optional controlName As String = Nothing)
        Dim rowIndex As Integer = panel.RowCount
        panel.RowCount += 1
        panel.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
        
        Dim lbl As New System.Windows.Forms.Label()
        lbl.Text = labelText
        lbl.AutoSize = True
        lbl.Anchor = System.Windows.Forms.AnchorStyles.Left
        lbl.Margin = New System.Windows.Forms.Padding(0, Scale(6), Scale(SPACING), 0)
        
        control.Anchor = System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right
        control.Margin = New System.Windows.Forms.Padding(0, Scale(3), 0, Scale(3))
        If controlName IsNot Nothing Then control.Name = controlName
        
        panel.Controls.Add(lbl, 0, rowIndex)
        panel.Controls.Add(control, 1, rowIndex)
    End Sub
    
    ''' <summary>
    ''' Adds a full-width control row (spans both columns).
    ''' </summary>
    Public Sub AddFullWidthRow(panel As System.Windows.Forms.TableLayoutPanel, control As System.Windows.Forms.Control, Optional controlName As String = Nothing)
        Dim rowIndex As Integer = panel.RowCount
        panel.RowCount += 1
        panel.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
        
        control.Anchor = System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right
        control.Margin = New System.Windows.Forms.Padding(0, Scale(3), 0, Scale(3))
        If controlName IsNot Nothing Then control.Name = controlName
        
        panel.Controls.Add(control, 0, rowIndex)
        panel.SetColumnSpan(control, 2)
    End Sub
    
    ' ============================================================
    ' CONTROL FACTORY
    ' ============================================================
    
    Public Function CreateLabel(text As String) As System.Windows.Forms.Label
        Dim lbl As New System.Windows.Forms.Label()
        lbl.Text = text
        lbl.AutoSize = True
        Return lbl
    End Function
    
    Public Function CreateTextBox(Optional text As String = "") As System.Windows.Forms.TextBox
        Dim txt As New System.Windows.Forms.TextBox()
        txt.Text = text
        Return txt
    End Function
    
    Public Function CreateComboBox(items() As String, Optional selectedIndex As Integer = 0) As System.Windows.Forms.ComboBox
        Dim cbo As New System.Windows.Forms.ComboBox()
        cbo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        cbo.Items.AddRange(items)
        If items.Length > selectedIndex Then cbo.SelectedIndex = selectedIndex
        Return cbo
    End Function
    
    Public Function CreateCheckBox(text As String, Optional checked As Boolean = False) As System.Windows.Forms.CheckBox
        Dim chk As New System.Windows.Forms.CheckBox()
        chk.Text = text
        chk.Checked = checked
        chk.AutoSize = True
        Return chk
    End Function
    
    Public Function CreateNumericUpDown(min As Decimal, max As Decimal, value As Decimal, Optional decimalPlaces As Integer = 0) As System.Windows.Forms.NumericUpDown
        Dim nud As New System.Windows.Forms.NumericUpDown()
        nud.Minimum = min
        nud.Maximum = max
        nud.Value = value
        nud.DecimalPlaces = decimalPlaces
        Return nud
    End Function
    
    Public Function CreateButton(text As String, Optional width As Integer = BUTTON_WIDTH) As System.Windows.Forms.Button
        Dim btn As New System.Windows.Forms.Button()
        btn.Text = text
        btn.Width = Scale(width)
        btn.Height = Scale(BUTTON_HEIGHT)
        Return btn
    End Function
    
    Public Function CreateListBox(Optional height As Integer = 150) As System.Windows.Forms.ListBox
        Dim lst As New System.Windows.Forms.ListBox()
        lst.Height = Scale(height)
        lst.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended
        Return lst
    End Function
    
    Public Function CreateDataGridView() As System.Windows.Forms.DataGridView
        Dim dgv As New System.Windows.Forms.DataGridView()
        dgv.AllowUserToAddRows = False
        dgv.AllowUserToDeleteRows = False
        dgv.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        dgv.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        dgv.RowHeadersVisible = False
        Return dgv
    End Function
    
    ' ============================================================
    ' DPI SCALING
    ' ============================================================
    
    Private m_ScaleFactor As Double = 0
    
    ''' <summary>
    ''' Gets the current DPI scale factor (1.0 = 96 DPI / 100%).
    ''' </summary>
    Public Function GetScaleFactor() As Double
        If m_ScaleFactor = 0 Then
            Using g As System.Drawing.Graphics = System.Drawing.Graphics.FromHwnd(IntPtr.Zero)
                m_ScaleFactor = g.DpiX / 96.0
            End Using
        End If
        Return m_ScaleFactor
    End Function
    
    ''' <summary>
    ''' Scales a pixel value for current DPI.
    ''' </summary>
    Public Function Scale(value As Integer) As Integer
        Return CInt(value * GetScaleFactor())
    End Function
    
    Public Function Scale(value As Double) As Integer
        Return CInt(value * GetScaleFactor())
    End Function
    
End Module
```

### Success Criteria

#### Verification:
- [ ] UILib.vb compiles without errors when included via AddVbFile
- [ ] `CreateForm` produces correctly styled forms
- [ ] `ShowNonModal` allows Inventor viewport interaction while form is open
- [ ] `CreateContentPanel` + `AddRow` produces correct two-column layout
- [ ] `Scale()` returns appropriate values at 100%, 125%, 150% DPI

#### Manual Verification:
- [ ] Test form at different DPI settings in Windows Display Settings
- [ ] Test non-modal form with orbit, pan, zoom while open
- [ ] Test `CloseAllForms` cleanup on rule exit

**Implementation Note**: After completing this phase and all verification passes, pause here for manual confirmation before proceeding to the next phase.

---

## Phase 2: Estonian Strings Module

### Overview

Create `Lib/StringsLib.vb` with all common Estonian translations, eliminating duplication and ensuring consistency.

### Changes Required

#### 1. Create StringsLib.vb

**File**: `Lib/StringsLib.vb`

**Contents**:

```vb
' StringsLib.vb - Centralized Estonian UI Strings
' All user-facing text in one place for consistency

Public Module StringsLib
    
    ' ============================================================
    ' DOCUMENT GUARDS
    ' ============================================================
    
    Public Const MSG_NO_ACTIVE_DOCUMENT As String = "Aktiivne dokument puudub."
    Public Const MSG_REQUIRES_ASSEMBLY As String = "See reegel töötab ainult koostuga (.iam)."
    Public Const MSG_REQUIRES_PART As String = "See reegel töötab ainult detailiga (.ipt)."
    Public Const MSG_REQUIRES_DRAWING As String = "See reegel töötab ainult joonisega (.idw)."
    Public Const MSG_REQUIRES_ASSEMBLY_OR_PART As String = "See reegel töötab koostuga (.iam) või detailiga (.ipt)."
    
    ' ============================================================
    ' BUTTONS
    ' ============================================================
    
    Public Const BTN_OK As String = "OK"
    Public Const BTN_CANCEL As String = "Tühista"
    Public Const BTN_APPLY As String = "Rakenda"
    Public Const BTN_CREATE As String = "Loo"
    Public Const BTN_UPDATE As String = "Uuenda"
    Public Const BTN_DELETE As String = "Kustuta"
    Public Const BTN_RUN As String = "Käivita"
    Public Const BTN_SELECT As String = "Vali..."
    Public Const BTN_BROWSE As String = "Sirvi..."
    Public Const BTN_SELECT_ALL As String = "Vali kõik"
    Public Const BTN_SELECT_NONE As String = "Tühjenda"
    Public Const BTN_CLEAR_SELECTION As String = "Tühista valik"
    Public Const BTN_PICK_SURFACE As String = "Vali pind"
    Public Const BTN_APPLY_ALL As String = "Rakenda kõigile"
    
    ' ============================================================
    ' PICKER PROMPTS
    ' ============================================================
    
    Public Const PICK_CANCEL_SUFFIX As String = " - ESC tühistamiseks"
    
    Public Const PICK_POINT As String = "Vali punkt" & PICK_CANCEL_SUFFIX
    Public Const PICK_AXIS As String = "Vali telg" & PICK_CANCEL_SUFFIX
    Public Const PICK_PLANE As String = "Vali tasand" & PICK_CANCEL_SUFFIX
    Public Const PICK_FACE As String = "Vali pind" & PICK_CANCEL_SUFFIX
    Public Const PICK_EDGE As String = "Vali serv" & PICK_CANCEL_SUFFIX
    Public Const PICK_COMPONENT As String = "Vali komponent" & PICK_CANCEL_SUFFIX
    Public Const PICK_OCCURRENCE As String = "Vali element" & PICK_CANCEL_SUFFIX
    
    ''' <summary>
    ''' Formats a pick prompt with custom description.
    ''' Example: FormatPickPrompt("Vali alguspunkt") returns "Vali alguspunkt - ESC tühistamiseks"
    ''' </summary>
    Public Function FormatPickPrompt(description As String) As String
        Return description & PICK_CANCEL_SUFFIX
    End Function
    
    ' ============================================================
    ' COMMON LABELS
    ' ============================================================
    
    Public Const LBL_NAME As String = "Nimi:"
    Public Const LBL_DESCRIPTION As String = "Kirjeldus:"
    Public Const LBL_VALUE As String = "Väärtus:"
    Public Const LBL_COUNT As String = "Kogus:"
    Public Const LBL_WIDTH As String = "Laius:"
    Public Const LBL_HEIGHT As String = "Kõrgus:"
    Public Const LBL_THICKNESS As String = "Paksus:"
    Public Const LBL_MATERIAL As String = "Materjal:"
    Public Const LBL_ORIENTATION As String = "Orientatsioon:"
    Public Const LBL_OFFSET As String = "Nihe:"
    Public Const LBL_CENTER_POINT As String = "Keskpunkt:"
    Public Const LBL_START_POINT As String = "Alguspunkt:"
    Public Const LBL_END_POINT As String = "Lõpp-punkt:"
    
    ' ============================================================
    ' COMMON MESSAGES
    ' ============================================================
    
    Public Const MSG_NO_VIEWS_ON_SHEET As String = "Lehel puuduvad vaated."
    Public Const MSG_NO_SELECTION As String = "Midagi pole valitud."
    Public Const MSG_OPERATION_CANCELLED As String = "Toiming tühistatud."
    Public Const MSG_OPERATION_COMPLETE As String = "Toiming lõpetatud."
    Public Const MSG_OPERATION_FAILED As String = "Toiming ebaõnnestus. Vaata logi akent."
    Public Const MSG_CONFIRM_DELETE As String = "Kas oled kindel, et soovid kustutada?"
    Public Const MSG_VALIDATION_ERROR As String = "Viga"
    
    ' ============================================================
    ' DRAWING-SPECIFIC
    ' ============================================================
    
    Public Const MSG_NO_PART_REFERENCE As String = "Jooniselt ei leitud viidet detailile."
    Public Const MSG_USE_CREATE_DRAWINGS As String = "Kasuta 'Loo 1:1 joonised' funktsiooni jooniste loomiseks."
    
    ' ============================================================
    ' MODULE RELEASE
    ' ============================================================
    
    Public Const TITLE_MODULE_RELEASE As String = "Moodulite väljastamine"
    Public Const BTN_ALL_MODULES As String = "Kõik moodulid"
    Public Const BTN_FIRST_MODULE As String = "Esimene moodul"
    Public Const MSG_CONFIRM_RELEASE As String = "Kinnita väljastamine"
    Public Const MSG_RELEASE_COMPLETE As String = "Väljastamine lõpetatud"
    
    ' ============================================================
    ' BOM EXPORT
    ' ============================================================
    
    Public Const TITLE_SELECT_TEMPLATE As String = "Vali BOM mall"
    Public Const MSG_EXCEL_LOCKED As String = "Excel fail on avatud teises rakenduses. Palun sulge see enne jätkamist."
    
    ' ============================================================
    ' PATTERN (KORDUSED)
    ' ============================================================
    
    Public Const TITLE_CENTER_PATTERN As String = "Kordused keskelt"
    Public Const MSG_PATTERN_CREATED As String = "Muster loodud."
    Public Const MSG_PATTERN_FAILED As String = "Mustri loomine ebaõnnestus. Vaata logi akent."
    
    ' ============================================================
    ' FORMAT HELPERS
    ' ============================================================
    
    ''' <summary>
    ''' Creates a rule-specific dialog title.
    ''' Example: FormatDialogTitle("Lisa mõõdud") for drawing dimensions dialog
    ''' </summary>
    Public Function FormatDialogTitle(ruleName As String, Optional subtitle As String = Nothing) As String
        If subtitle IsNot Nothing Then
            Return ruleName & " — " & subtitle
        End If
        Return ruleName
    End Function
    
    ''' <summary>
    ''' Creates a document guard message for a specific rule.
    ''' </summary>
    Public Function FormatGuardMessage(ruleName As String, message As String) As String
        Return message  ' Message is self-contained; ruleName goes in MessageBox title
    End Function
    
End Module
```

### Success Criteria

#### Verification:
- [ ] StringsLib.vb compiles without errors
- [ ] All common phrases have corresponding constants
- [ ] Format functions produce expected output

#### Manual Verification:
- [ ] Review string consistency with existing dialogs
- [ ] Verify Estonian spelling and grammar

**Implementation Note**: After completing this phase and all verification passes, pause here for manual confirmation before proceeding to the next phase.

---

## Phase 3: Viewport Helpers — Transient Graphics

### Overview

Create `Lib/ViewportHelperLib.vb` for ClientGraphics-based transient overlays (highlights, markers) and management of real model features for complex previews.

### Changes Required

#### 1. Create ViewportHelperLib.vb

**File**: `Lib/ViewportHelperLib.vb`

**Contents**:

```vb
' ViewportHelperLib.vb - Viewport Helper Display Library
' Provides transient graphics overlays and highlight management

Public Module ViewportHelperLib
    
    ' ============================================================
    ' HIGHLIGHT MANAGEMENT
    ' ============================================================
    
    Private m_HighlightSet As Object  ' HighlightSet
    Private m_App As Inventor.Application
    
    ''' <summary>
    ''' Initializes the viewport helper with the application instance.
    ''' Call once at rule start.
    ''' </summary>
    Public Sub Initialize(app As Inventor.Application)
        m_App = app
        ClearHighlights()
    End Sub
    
    ''' <summary>
    ''' Highlights an object in the viewport.
    ''' </summary>
    Public Sub Highlight(obj As Object)
        If m_App Is Nothing Then Return
        
        Try
            If m_HighlightSet Is Nothing Then
                m_HighlightSet = m_App.ActiveDocument.HighlightSets.Add()
            End If
            m_HighlightSet.AddItem(obj)
        Catch
            ' Object may not support highlighting
        End Try
    End Sub
    
    ''' <summary>
    ''' Highlights multiple objects.
    ''' </summary>
    Public Sub HighlightMany(objects As IEnumerable)
        For Each obj In objects
            Highlight(obj)
        Next
    End Sub
    
    ''' <summary>
    ''' Clears all highlights.
    ''' </summary>
    Public Sub ClearHighlights()
        Try
            If m_HighlightSet IsNot Nothing Then
                m_HighlightSet.Clear()
            End If
        Catch
        End Try
    End Sub
    
    ''' <summary>
    ''' Removes the highlight set entirely.
    ''' Call on rule exit.
    ''' </summary>
    Public Sub Cleanup()
        Try
            If m_HighlightSet IsNot Nothing Then
                m_HighlightSet.Delete()
                m_HighlightSet = Nothing
            End If
        Catch
        End Try
        
        CleanupClientGraphics()
    End Sub
    
    ' ============================================================
    ' CLIENT GRAPHICS - TRANSIENT MARKERS
    ' ============================================================
    
    Private m_GraphicsNode As Object  ' GraphicsNode
    Private Const GRAPHICS_ID As String = "_ViewportHelper_"
    
    ''' <summary>
    ''' Gets or creates the ClientGraphics node for transient display.
    ''' </summary>
    Private Function GetGraphicsNode() As Object
        If m_App Is Nothing Then Return Nothing
        
        Try
            Dim doc As Document = m_App.ActiveDocument
            Dim clientGraphics As ClientGraphics
            
            Try
                clientGraphics = doc.ComponentDefinition.ClientGraphicsCollection.Item(GRAPHICS_ID)
            Catch
                clientGraphics = doc.ComponentDefinition.ClientGraphicsCollection.Add(GRAPHICS_ID)
            End Try
            
            If m_GraphicsNode Is Nothing Then
                m_GraphicsNode = clientGraphics.AddNode(1)
            End If
            
            Return m_GraphicsNode
        Catch
            Return Nothing
        End Try
    End Function
    
    ''' <summary>
    ''' Adds a transient point marker at the specified location.
    ''' </summary>
    Public Sub AddPointMarker(point As Point, Optional color As Object = Nothing)
        Dim node As Object = GetGraphicsNode()
        If node Is Nothing Then Return
        
        Try
            Dim tg As TransientGeometry = m_App.TransientGeometry
            Dim coords As New List(Of Double)
            coords.Add(point.X)
            coords.Add(point.Y)
            coords.Add(point.Z)
            
            Dim coordSet As GraphicsCoordinateSet = node.CoordinateSets.Add(node.CoordinateSets.Count + 1)
            coordSet.PutCoordinates(coords.ToArray())
            
            Dim pointSet As PointGraphics = node.AddPointGraphics()
            pointSet.CoordinateSet = coordSet
            pointSet.PointRenderStyle = PointRenderStyleEnum.kCirclePointStyle
            
            m_App.ActiveView.Update()
        Catch
        End Try
    End Sub
    
    ''' <summary>
    ''' Adds a transient line between two points.
    ''' </summary>
    Public Sub AddLineMarker(startPoint As Point, endPoint As Point, Optional color As Object = Nothing)
        Dim node As Object = GetGraphicsNode()
        If node Is Nothing Then Return
        
        Try
            Dim coords As New List(Of Double)
            coords.Add(startPoint.X)
            coords.Add(startPoint.Y)
            coords.Add(startPoint.Z)
            coords.Add(endPoint.X)
            coords.Add(endPoint.Y)
            coords.Add(endPoint.Z)
            
            Dim coordSet As GraphicsCoordinateSet = node.CoordinateSets.Add(node.CoordinateSets.Count + 1)
            coordSet.PutCoordinates(coords.ToArray())
            
            Dim lineSet As LineGraphics = node.AddLineGraphics()
            lineSet.CoordinateSet = coordSet
            
            m_App.ActiveView.Update()
        Catch
        End Try
    End Sub
    
    ''' <summary>
    ''' Clears all transient markers.
    ''' </summary>
    Public Sub ClearMarkers()
        Try
            Dim doc As Document = m_App.ActiveDocument
            Dim clientGraphics As ClientGraphics
            
            Try
                clientGraphics = doc.ComponentDefinition.ClientGraphicsCollection.Item(GRAPHICS_ID)
                clientGraphics.Delete()
            Catch
            End Try
            
            m_GraphicsNode = Nothing
            m_App.ActiveView.Update()
        Catch
        End Try
    End Sub
    
    Private Sub CleanupClientGraphics()
        ClearMarkers()
    End Sub
    
    ' ============================================================
    ' REAL FEATURE PREVIEW (UCS, Work Planes)
    ' ============================================================
    
    Private m_PreviewFeatures As New List(Of Object)
    
    ''' <summary>
    ''' Creates a preview UCS that can be updated and deleted.
    ''' Based on Koordinaadid.vb pattern.
    ''' </summary>
    Public Function CreatePreviewUCS(asmDoc As AssemblyDocument, name As String, matrix As Matrix) As UserCoordinateSystem
        Try
            Dim ucs As UserCoordinateSystem = asmDoc.ComponentDefinition.UserCoordinateSystems.Add(matrix)
            ucs.Name = name
            m_PreviewFeatures.Add(ucs)
            m_App.ActiveView.Update()
            Return ucs
        Catch
            Return Nothing
        End Try
    End Function
    
    ''' <summary>
    ''' Updates a preview UCS position/orientation.
    ''' </summary>
    Public Sub UpdatePreviewUCS(ucs As UserCoordinateSystem, matrix As Matrix)
        Try
            ucs.Transformation = matrix
            m_App.ActiveView.Update()
        Catch
        End Try
    End Sub
    
    ''' <summary>
    ''' Creates a temporary work plane for preview.
    ''' </summary>
    Public Function CreatePreviewWorkPlane(compDef As Object, planeInput As Object, name As String) As WorkPlane
        Try
            Dim wp As WorkPlane = compDef.WorkPlanes.AddByPlaneAndOffset(planeInput, 0)
            wp.Name = name
            wp.Visible = True
            m_PreviewFeatures.Add(wp)
            m_App.ActiveView.Update()
            Return wp
        Catch
            Return Nothing
        End Try
    End Function
    
    ''' <summary>
    ''' Deletes all preview features (UCS, work planes, etc.).
    ''' Call on cancel or cleanup.
    ''' </summary>
    Public Sub DeletePreviewFeatures()
        For Each feature In m_PreviewFeatures.ToArray()
            Try
                feature.Delete()
            Catch
            End Try
        Next
        m_PreviewFeatures.Clear()
        
        If m_App IsNot Nothing Then
            Try
                m_App.ActiveView.Update()
            Catch
            End Try
        End If
    End Sub
    
    ''' <summary>
    ''' Commits preview features (keeps them, removes from cleanup list).
    ''' Call on OK/confirm.
    ''' </summary>
    Public Sub CommitPreviewFeatures()
        m_PreviewFeatures.Clear()
    End Sub
    
    ' ============================================================
    ' VIEW UPDATES
    ' ============================================================
    
    ''' <summary>
    ''' Forces a viewport refresh.
    ''' </summary>
    Public Sub RefreshView()
        Try
            If m_App IsNot Nothing Then
                m_App.ActiveView.Update()
            End If
        Catch
        End Try
    End Sub
    
End Module
```

### Success Criteria

#### Verification:
- [ ] ViewportHelperLib.vb compiles without errors
- [ ] `Highlight()` visually highlights objects in viewport
- [ ] `AddPointMarker()` / `AddLineMarker()` display transient graphics
- [ ] `CreatePreviewUCS()` creates visible UCS that can be updated/deleted
- [ ] `Cleanup()` removes all transient graphics and highlights

#### Manual Verification:
- [ ] Test highlight visibility with different object types (faces, edges, occurrences)
- [ ] Test transient markers persist while form is open
- [ ] Test cleanup removes all visual artifacts

**Implementation Note**: After completing this phase and all verification passes, pause here for manual confirmation before proceeding to the next phase.

---

## Phase 4: Unified Picker Integration

### Overview

Extend picker functionality to work with non-modal forms and integrate with viewport helpers for visual feedback.

### Changes Required

#### 1. Add Picker Functions to UILib.vb

**File**: `Lib/UILib.vb` (additions)

**Add to UILib.vb**:

```vb
    ' ============================================================
    ' PICKER INTEGRATION (Non-Modal Compatible)
    ' ============================================================
    
    ''' <summary>
    ''' Performs a pick operation while keeping the form visible.
    ''' Highlights picked object. Returns Nothing on cancel/ESC.
    ''' </summary>
    Public Function PickWithForm(app As Inventor.Application, 
                                  frm As System.Windows.Forms.Form, 
                                  filter As SelectionFilterEnum, 
                                  prompt As String) As Object
        ' Ensure form is non-modal
        Dim wasTopMost As Boolean = frm.TopMost
        frm.TopMost = False  ' Allow Inventor to receive focus
        
        ViewportHelperLib.ClearHighlights()
        
        Dim result As Object = Nothing
        
        Try
            ' Pump messages to keep form responsive
            PumpMessages()
            
            ' Perform pick
            result = app.CommandManager.Pick(filter, prompt)
            
            ' Highlight result
            If result IsNot Nothing Then
                ViewportHelperLib.Highlight(result)
            End If
        Catch
            ' User cancelled with ESC
            result = Nothing
        End Try
        
        frm.TopMost = wasTopMost
        frm.Focus()
        
        Return result
    End Function
    
    ''' <summary>
    ''' Performs a multi-pick operation (select multiple objects).
    ''' Returns list of picked objects. Empty list on cancel.
    ''' </summary>
    Public Function MultiPickWithForm(app As Inventor.Application,
                                       frm As System.Windows.Forms.Form,
                                       filter As SelectionFilterEnum,
                                       prompt As String,
                                       donePrompt As String) As List(Of Object)
        Dim results As New List(Of Object)
        Dim wasTopMost As Boolean = frm.TopMost
        frm.TopMost = False
        
        ViewportHelperLib.ClearHighlights()
        
        Dim keepPicking As Boolean = True
        Do While keepPicking
            PumpMessages()
            
            Try
                Dim obj As Object = app.CommandManager.Pick(filter, prompt)
                If obj IsNot Nothing Then
                    results.Add(obj)
                    ViewportHelperLib.Highlight(obj)
                End If
            Catch
                ' ESC pressed - done picking
                keepPicking = False
            End Try
        Loop
        
        frm.TopMost = wasTopMost
        frm.Focus()
        
        Return results
    End Function
    
    ''' <summary>
    ''' Creates a pick button with integrated picker.
    ''' Handles click to perform pick and update display.
    ''' </summary>
    Public Function CreatePickButton(text As String, 
                                      displayControl As System.Windows.Forms.Control,
                                      pickAction As Action(Of Object)) As System.Windows.Forms.Button
        Dim btn As System.Windows.Forms.Button = CreateButton(text, 70)
        btn.Tag = New Object() {displayControl, pickAction}
        Return btn
    End Function
```

#### 2. Extend UtilsLib.vb with Estonian Prompts

**File**: `Lib/UtilsLib.vb` (modifications)

**Add overloads that use StringsLib**:

```vb
    ' Add these Estonian-default picker functions
    
    ''' <summary>
    ''' Picks a work point with Estonian prompt.
    ''' </summary>
    Public Function PickPointET(app As Inventor.Application) As Object
        Return PickPoint(app, StringsLib.PICK_POINT)
    End Function
    
    ''' <summary>
    ''' Picks a work axis with Estonian prompt.
    ''' </summary>
    Public Function PickAxisET(app As Inventor.Application) As Object
        Return PickAxis(app, StringsLib.PICK_AXIS)
    End Function
    
    ''' <summary>
    ''' Picks a work plane with Estonian prompt.
    ''' </summary>
    Public Function PickPlaneET(app As Inventor.Application) As Object
        Return PickPlane(app, StringsLib.PICK_PLANE)
    End Function
    
    ''' <summary>
    ''' Picks a planar face with Estonian prompt.
    ''' </summary>
    Public Function PickFaceET(app As Inventor.Application) As Object
        Try
            Return app.CommandManager.Pick(SelectionFilterEnum.kPartFacePlanarFilter, StringsLib.PICK_FACE)
        Catch
            Return Nothing
        End Try
    End Function
    
    ''' <summary>
    ''' Picks an occurrence with Estonian prompt.
    ''' </summary>
    Public Function PickOccurrenceET(app As Inventor.Application) As Object
        Try
            Return app.CommandManager.Pick(SelectionFilterEnum.kAssemblyOccurrenceFilter, StringsLib.PICK_OCCURRENCE)
        Catch
            Return Nothing
        End Try
    End Function
```

### Success Criteria

#### Verification:
- [ ] `PickWithForm` allows picking while non-modal form stays open
- [ ] Picked object is highlighted after selection
- [ ] ESC cancels pick and returns Nothing
- [ ] Estonian picker functions use StringsLib prompts

#### Manual Verification:
- [ ] Test pick with non-modal form open - form remains visible
- [ ] Test highlight appears on picked object
- [ ] Test multiple picks in sequence

**Implementation Note**: After completing this phase and all verification passes, pause here for manual confirmation before proceeding to the next phase.

---

## Phase 5: Migrate Drawing Dialogs (Joonised/)

### Overview

Convert all drawing-related dialogs in `Joonised/` folder to use UILib and StringsLib.

### Files to Migrate

1. `Joonised/Lisa mõõdud.vb`
2. `Joonised/Lisa vaated.vb`
3. `Joonised/Loo 1-1 joonised.vb`
4. `Joonised/Uuenda 1-1 joonis.vb`
5. `Joonised/Uuenda lehe suurus.vb`

### Migration Pattern

For each file:

1. Add library references at top:
```vb
AddVbFile "Lib/UILib.vb"
AddVbFile "Lib/StringsLib.vb"
```

2. Replace document guards:
```vb
' Before:
MessageBox.Show("Aktiivne dokument puudub.", "Lisa mõõdud")

' After:
MessageBox.Show(StringsLib.MSG_NO_ACTIVE_DOCUMENT, ruleName)
```

3. Replace form creation:
```vb
' Before:
Dim frm As New System.Windows.Forms.Form()
frm.Text = "Lisa mõõdud"
frm.Width = 400
frm.Height = 300
frm.StartPosition = FormStartPosition.CenterScreen
' ... manual control positioning

' After:
Dim frm As System.Windows.Forms.Form = UILib.CreateForm(StringsLib.FormatDialogTitle(ruleName))
Dim content As System.Windows.Forms.TableLayoutPanel = UILib.CreateContentPanel()
Dim buttons As System.Windows.Forms.FlowLayoutPanel = UILib.CreateButtonPanel()

UILib.AddRow(content, StringsLib.LBL_VALUE, UILib.CreateNumericUpDown(...))
' ... add rows

buttons.Controls.Add(UILib.CreateButton(StringsLib.BTN_OK))
buttons.Controls.Add(UILib.CreateButton(StringsLib.BTN_CANCEL))

frm.Controls.Add(content)
frm.Controls.Add(buttons)
```

4. Replace ShowDialog with ShowNonModal:
```vb
' Before:
Dim result As DialogResult = frm.ShowDialog()

' After:
UILib.ShowNonModal(frm)
Dim result As DialogResult = frm.DialogResult
```

### Success Criteria

#### Verification:
- [ ] All 5 drawing dialogs use UILib for form creation
- [ ] All Estonian strings come from StringsLib
- [ ] All dialogs are non-modal
- [ ] Existing functionality preserved

#### Manual Verification:
- [ ] Test each dialog opens and functions correctly
- [ ] Test viewport interaction while dialog open
- [ ] Test at different DPI settings

**Implementation Note**: After completing this phase and all verification passes, pause here for manual confirmation before proceeding to the next phase.

---

## Phase 6: Migrate BOM and Module Dialogs (Koost/, Moodulid/)

### Overview

Convert BOM and module-related dialogs to use UILib and StringsLib.

### Files to Migrate

1. `Koost/Muutujad.vb`
2. `Koost/Ekspordi BOM.vb`
3. `Koost/Nimeta detailid.vb`
4. `Koost/Sorteeri detailid.vb`
5. `Moodulid/Loo alusmoodul.vb`

### Success Criteria

#### Verification:
- [ ] All 5 files use UILib and StringsLib
- [ ] DataGridView dialogs resize correctly
- [ ] Non-modal behavior works for all dialogs

---

## Phase 7: Migrate Library Dialog Functions (Lib/)

### Overview

Convert dialog functions in library modules to use UILib and StringsLib.

### Files to Migrate

1. `Lib/ModuleReleaseLib.vb` - ShowModeSelectionDialog, ShowPlanConfirmationDialog, ShowCompletionSummary
2. `Lib/BOMExportLib.vb` - SelectTemplateFromListOrBrowse, BrowseTemplateFile
3. `Lib/BoundingBoxStockLib.vb` - ShowConfigForm (already uses TableLayoutPanel)
4. `Lib/ExcelReaderLib.vb` - ShowConfigSelectionDialog

### Success Criteria

#### Verification:
- [ ] All library dialog functions use UILib
- [ ] Library files work with StringsLib (passed as parameter or referenced)
- [ ] Non-modal behavior where appropriate

---

## Phase 8: Migrate Remaining Rule Dialogs

### Overview

Convert all remaining rule dialogs to use UILib and StringsLib.

### Files to Migrate

1. `Koordinaadid.vb` - Already non-modal, adapt to UILib
2. `Loo komponendid.vb` - Large dialog with DataGridView
3. `Mustrid/Kordused keskelt.vb` - Complex dialog with pickers
4. `Mõõdud.vb` - Dimension configuration
5. `Määra materjalide välimus.vb` - Material appearance
6. `Komponendid/Pinnalaotuse vaated.vb` - Flat pattern views
7. `Lehtmetall.vb` - Sheet metal (picker only)
8. `Taasta värvid.vb` - Color restore
9. `Kopeeri pinnad.vb` - Surface copy

### Success Criteria

#### Verification:
- [ ] All remaining dialogs use UILib and StringsLib
- [ ] Complex dialogs (Koordinaadid, Kordused keskelt) maintain picker functionality
- [ ] All dialogs are non-modal

---

## Testing Strategy

### Unit Tests

Each phase includes verification steps. Test:
- Library compilation (AddVbFile without errors)
- Form creation (correct styling, dimensions)
- Non-modal behavior (viewport interaction)
- DPI scaling (multiple settings)

### Integration Tests

After each migration phase:
1. Run the migrated rule/function
2. Verify dialog appears correctly
3. Verify viewport interaction while dialog open
4. Verify all buttons and controls function
5. Verify output matches previous behavior

### Manual Testing Steps

1. **DPI Testing**:
   - Set Windows display scaling to 100%, 125%, 150%
   - Open each dialog
   - Verify no clipping, correct proportions

2. **Non-Modal Testing**:
   - Open dialog
   - Orbit, pan, zoom in Inventor viewport
   - Perform picks while dialog open
   - Verify dialog stays visible and responsive

3. **String Verification**:
   - Review all visible Estonian text
   - Verify consistency across dialogs
   - Check spelling and grammar

## Terminology Checklist

Verify all code uses correct domain terms per UBIQUITOUS_LANGUAGE.md:
- [ ] Dialog titles use correct Estonian terminology
- [ ] StringsLib uses "Aluselement" not "Alusmoodul" for parametric designs
- [ ] StringsLib uses "Väljastatud element" not "Moodul" for released units
- [ ] Folder path references match current structure

## References

- Research: `docs/research/2026-05-08-ui-layouts-and-patterns.md`
- Existing non-modal pattern: `Koordinaadid.vb:444-451`
- Existing TableLayoutPanel pattern: `Lib/BoundingBoxStockLib.vb:407-545`
- Existing pickers: `Lib/UtilsLib.vb:449-477`
- Domain terminology: `docs/UBIQUITOUS_LANGUAGE.md`
- Project conventions: `AGENTS.md`
