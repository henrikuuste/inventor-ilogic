# Inventor iLogic Development Guidelines

This document captures constraints and best practices for writing Autodesk Inventor iLogic rules and VB scripts.

## Project Structure and Conventions

### Product Family Projects

We work with product family projects organized in Vault. Each project represents a product line (e.g., "Lume" furniture series).

- **Project location**: `<VaultRoot>/Tooted/<ProjectName>` (e.g., `C:\_SoftcomVault\Tooted\Lume`)
- **Project property**: The `Project` iProperty identifies which project a file belongs to (e.g., "Lume")
- **Scope**: When searching for linked files, limit searches to the project scope using `UtilsLib.GetProjectPath()`

### Standard Folder Structure

```
Tooted/
  <ProjectName>/
    Algmaterjal/           - Source data (STEP files, design drawings, specifications)
    Alusmoodulid/          - Parametric base modules (design masters)
      <ModuleName>/        - Assemblies
        Eskiis/            - Sketches, skeleton parts and concepts
        Karkass/           - Frame/structure components and subassemblies
          Detailid/        - Individual parts
          Joonised/        - Drawings
        Poroloon/          - Foam/upholstery components and subassemblies
          Detailid/        - Individual parts
          Joonised/        - Drawings
    Moodulid/              - Released module versions (production-ready)
```

### File Naming and Properties

- **Auto-numbering**: All parts and assemblies are automatically numbered by Vault on save
- **File name**: Uses the Vault-generated number (e.g., `000123.ipt`)
- **Part Number property**: Same as the Vault number/file name
- **Description property**: Human-readable name of the part/assembly (e.g., "Iste karkass")
- **Project property**: Always set when creating new files to maintain project scope

### Key Utility Functions

| Task | Function |
|------|----------|
| Extract project name from path | `UtilsLib.ExtractProjectName(filePath)` |
| Get project folder path | `UtilsLib.GetProjectPath(filePath)` |
| Create Vault folders recursively | `VaultNumberingLib.EnsureVaultFolderRecursive(conn, vaultPath)` |
| Register auto-update handler | `DocumentUpdateLib.RegisterUpdateHandler(doc, iLogicAuto, uid, codeLines, triggers)` |

## API Documentation References

When writing Inventor API code, **always search the official documentation first**:

- **Inventor 2026 API Reference (Primary)**: https://help.autodesk.com/view/INVNTOR/2026/ENU/
- **iLogic API Reference**: https://help.autodesk.com/view/INVNTOR/2026/ENU/?guid=110f3019-404c-4fc4-8b5d-7a3143f129da

### Key API Pages

| Topic | URL |
|-------|-----|
| DrawingViews.AddBaseView | https://help.autodesk.com/view/INVNTOR/2026/ENU/?guid=DrawingViews_AddBaseView |
| Sheet Metal Flat Pattern | Use `AddBaseView` with `SheetMetalFoldedModel=False` via NameValueMap |

### Using NameValueMap for Optional API Parameters

Many Inventor API methods accept a `NameValueMap` for optional parameters that aren't exposed as direct method arguments. This is the recommended way to pass advanced options.

**Pattern:**
```vb
Dim options As NameValueMap = app.TransientObjects.CreateNameValueMap()
options.Add("OptionName", optionValue)

' Pass using named parameter syntax
result = SomeMethod(requiredParam1, requiredParam2, AdditionalOptions := options)
```

**Common NameValueMap Options:**

| Method | Option Name | Values | Description |
|--------|-------------|--------|-------------|
| `DrawingViews.AddBaseView` | `SheetMetalFoldedModel` | `True`/`False` | `False` = flat pattern view |
| `DrawingViews.AddBaseView` | `DesignViewAssociative` | `True`/`False` | Associative design view |
| `DrawingViews.AddBaseView` | `PositionalRepresentation` | String | Name of positional representation |

**Example - Create Flat Pattern View:**
```vb
Dim viewOptions As NameValueMap = app.TransientObjects.CreateNameValueMap()
viewOptions.Add("SheetMetalFoldedModel", False)

Dim flatView As DrawingView = sheet.DrawingViews.AddBaseView( _
    partDoc, _
    position, _
    1.0, _
    ViewOrientationTypeEnum.kDefaultViewOrientation, _
    DrawingViewStyleEnum.kHiddenLineRemovedDrawingViewStyle, _
    AdditionalOptions := viewOptions)
```

**Reference:** https://forums.autodesk.com/t5/inventor-programming-forum/ilogic-to-create-drawing-with-flat-pattern-view/td-p/13367792

**Example - Create View with Arbitrary Camera (custom orientation):**
```vb
' For non-axis-aligned view directions, use a Camera object
Dim camera As Camera = app.TransientObjects.CreateCamera()
camera.Eye = tg.CreatePoint(eyeX, eyeY, eyeZ)        ' Where you look FROM
camera.Target = tg.CreatePoint(centerX, centerY, centerZ)  ' What you look AT
camera.UpVector = tg.CreateUnitVector(upX, upY, upZ)  ' Which way is "up"
camera.ViewOrientationType = ViewOrientationTypeEnum.kArbitraryViewOrientation
camera.Perspective = False  ' Orthographic for drawings

Dim customView As DrawingView = sheet.DrawingViews.AddBaseView( _
    partDoc, _
    position, _
    1.0, _
    ViewOrientationTypeEnum.kArbitraryViewOrientation, _
    DrawingViewStyleEnum.kHiddenLineRemovedDrawingViewStyle, _
    "", _      ' ModelViewStyle (empty string)
    camera)    ' Pass camera as 7th parameter
```

**Note:** Projected views from an arbitrary camera base view work correctly - they project orthogonally from the custom view direction.

### How to Find API Documentation

1. Base URL: `https://help.autodesk.com/view/INVNTOR/2026/ENU/`
2. Add `?guid=` followed by the class/method name (e.g., `?guid=DrawingViews_AddBaseView`)
3. Search for specific methods by appending class name and method name with underscore

## QUICK REFERENCE - VERIFY BEFORE WRITING CODE

**STOP and check these rules BEFORE writing code, not after compile errors.**

### Critical Constraints Table

| When Writing... | NEVER Do This | ALWAYS Do This Instead |
|-----------------|---------------|------------------------|
| Lambda with ByRef param | `AddHandler btn.Click, Sub() byRefParam = x` | Store in `frm.Tag`, read after `ShowDialog()` |
| Forms controls | `New TextBox()` | `New System.Windows.Forms.TextBox()` |
| Control position/size | `New Point(x,y)`, `New Size(w,h)` | Use `.Left`, `.Top`, `.Width`, `.Height` |
| Context menus | `ContextMenuStrip`, `ToolStripMenuItem` | Use `Button` or `ComboBox` instead |
| Library files | `AddReference` or `AddVbFile` in library | Put ALL in main script only |
| Library with external types | `Function() As ACW.NumSchm` | `Function() As Object` (late binding) |
| Modal dialog + Pick | Pick while dialog open | Close dialog → Pick → Reopen dialog |
| File-level variables | `Dim x` outside Module/Sub | Declare inside `Sub Main()` |
| API optional parameters | Guessing parameter position | Use `NameValueMap` + `AdditionalOptions :=` |
| Drawing view spacing | Model dimensions (`RangeBox`) | Actual view bounds (`view.Width`, `view.Height`) |
| Extent dimension offset | 15mm offset (too close to model) | Use `CAMDrawingLib.DIMENSION_OFFSET` (25mm) |
| Parameter formulas | `max(a, b)` with comma | `max(a; b)` with semicolon |
| Parameter names | `00011_Name` (starts with digit) | `M_00011_Name` (prefix with letter) |
| CommandManager.Pick prompt | `"Select a face"` (no cancel hint) | `"Vali pind - ESC tühistamiseks"` (with ESC hint) |

### Windows Forms Checklist

Before writing ANY Windows Forms code, verify:

- [ ] **No ByRef parameters used in lambda expressions** - use `form.Tag` or `control.Tag` instead
- [ ] **No System.Drawing types** - no `Size`, `Point`, `Color`, `Font`
- [ ] **No ContextMenuStrip/ToolStripMenuItem** - use Buttons or ComboBox
- [ ] **All controls fully qualified** - `System.Windows.Forms.TextBox`, not `TextBox`
- [ ] **Close form before CommandManager.Pick**, reopen after
- [ ] **Read ByRef values from Tag AFTER ShowDialog returns**, not inside lambda

### Library Module Checklist

Before writing ANY library module code, verify:

- [ ] **No `AddReference` statements** - only main script can have these
- [ ] **No `AddVbFile` statements** - only main script can have these
- [ ] **No `Imports` with aliases for external assemblies** - use `Object` type
- [ ] **No `Logger`, `ThisApplication`, `ThisDoc`, `iLogicVb`** - pass as parameters
- [ ] **Use `Object` return type** for any external API types

---

## Language and Logging Conventions

### Language

- **User-facing messages** (prompts, selection instructions) should be in **Estonian**
- **Script/rule file names** should be in **Estonian**
- **Everything else** should be in **English**:
  - Code comments
  - Log messages
  - Variable names
  - Function names

### Logging

- Use `Logger.Info()`, `Logger.Warn()`, `Logger.Error()` to log to the iLogic log window
- **Do NOT use MessageBox.Show()** for informational summaries or progress updates
- Prefix log messages with the rule name for easy filtering, e.g., `Logger.Info("Lehtmetall: Starting conversion...")`
- **Early-exit errors** (where the rule cannot run at all) should use **both** Logger and MessageBox:
  - Logger message in English for the log
  - MessageBox message in Estonian for the user
  - This ensures the user sees immediate feedback when a rule fails to start

```vb
' BAD - blocks user with popup for information
MessageBox.Show("Conversion complete. Thickness: 2.5 mm", "Lehtmetall")

' GOOD - logs to iLogic log without blocking
Logger.Info("Lehtmetall: Conversion complete. Thickness: 2.5 mm")

' GOOD - early-exit error with both Logger and MessageBox
If doc Is Nothing Then
    Logger.Error("Lehtmetall: No active document.")
    MessageBox.Show("Aktiivne dokument puudub.", "Lehtmetall")
    Exit Sub
End If

' GOOD - user-facing prompt in Estonian
aSideFace = app.CommandManager.Pick(SelectionFilterEnum.kPartFacePlanarFilter, _
    "Vali A-külje pind (ülemine pind) - ESC tühistamiseks")
```

## File Structure

### Runnable Rules vs Library Modules

- **Runnable rules** should have a `Sub Main()` at the file level - this is the entry point when the rule is executed
- **Library modules** use `Public Module ModuleName ... End Module` to expose reusable functions
- Library modules are located in the Lib folder
- **Do NOT mix them in the same file** - a `Module` statement nested inside or after `Sub Main()` causes: `'Module' statements can occur only at file or namespace level`
- To use a library from a rule, use `AddVbFile "Lib/LibraryName.vb"` at the top of the rule

### AddReference and AddVbFile Ordering

- **`AddReference` MUST come BEFORE `AddVbFile`** in runnable scripts
- **Library modules CANNOT contain `AddReference` or `AddVbFile`** - these are iLogic directives that only work in `Sub Main()` context
- This causes: `Statement cannot appear outside of a method body` and `Method arguments must be enclosed in parentheses`
- **All references must be declared in the main runnable script**, not in libraries

```vb
' BAD - AddVbFile before AddReference
AddVbFile "Lib/MyLib.vb"
AddReference "SomeAssembly"  ' Error!

' BAD - AddReference inside library module
Public Module MyLib
    AddReference "SomeAssembly"  ' Error! Cannot appear in module
End Module

' GOOD - All AddReference first, then AddVbFile, in main script only
AddReference "Autodesk.Connectivity.WebServices"
AddReference "Autodesk.DataManagement.Client.Framework.Vault"
AddVbFile "Lib/VaultLib.vb"
AddVbFile "Lib/OtherLib.vb"

Sub Main()
    ' ...
End Sub
```

### Imports and Namespace Aliases in Libraries

- **Libraries CANNOT use `Imports` with namespace aliases** that depend on `AddReference`
- The referenced types won't be available because libraries can't add references
- **Use `Object` type for external API types**, or use fully qualified names
- This causes: `Type 'ACW.NumSchm' is not defined` or `'Imports' statements must precede any declarations`

```vb
' BAD - library using Imports alias for referenced assembly
Imports ACW = Autodesk.Connectivity.WebServices
Public Module VaultLib
    Public Function GetScheme() As ACW.NumSchm  ' Error!
    End Function
End Module

' GOOD - library using Object type (late binding)
Public Module VaultLib
    Public Function GetScheme() As Object
        ' Access properties via late binding
        Return scheme  ' scheme.Name, scheme.SchmID work at runtime
    End Function
End Module

' GOOD - library using fully qualified names (if reference is added by caller)
Public Module VaultLib
    Public Function GetConnection() As Object
        Return Connectivity.InventorAddin.EdmAddin.EdmSecurity.Instance.VaultConnection()
    End Function
End Module
```

### Module-Level Variables and Constants

- **Do NOT declare variables or constants at file level outside of Module/Sub/Function**
- Both `Dim` and `Const` at file level cause: `Statement is not valid in a namespace`
- Instead, declare inside `Sub Main()`, pass data via function parameters (`ByRef` for outputs), or use control `Tag` properties

```vb
' BAD - causes error (both Dim and Const)
Dim m_App As Inventor.Application
Const MODE_EDIT As String = "EDIT"

Sub Main()
    m_App = ThisApplication
End Sub

' GOOD - declare inside Sub Main
Sub Main()
    Const MODE_EDIT As String = "EDIT"
    Dim app As Inventor.Application = ThisApplication
    DoSomething(app, MODE_EDIT)
End Sub
```

## Windows Forms in iLogic

### Type Name Conflicts

- `TextBox` is ambiguous between `System.Windows.Forms.TextBox` and `Inventor.TextBox`
- **Always fully qualify Windows.Forms types:**

```vb
Dim txt As New System.Windows.Forms.TextBox()
Dim btn As New System.Windows.Forms.Button()
Dim lbl As New System.Windows.Forms.Label()
Dim frm As New System.Windows.Forms.Form()
```

### Avoid System.Drawing Types

- `Size`, `Point`, `Color`, `Font` from `System.Drawing` may require unavailable assembly references
- This causes: `Reference required to assembly 'System.Drawing.Common'` or `Type 'System.Drawing.X' is not defined`
- **Use individual properties instead:**

```vb
' BAD - may cause missing assembly errors
btn.Location = New Point(10, 20)
btn.Size = New Size(80, 30)
frm.MinimumSize = New System.Drawing.Size(800, 500)
lbl.Font = New System.Drawing.Font(lbl.Font, System.Drawing.FontStyle.Bold)

' GOOD - use individual properties
btn.Left = 10
btn.Top = 20
btn.Width = 80
btn.Height = 30

' For emphasis, use text decoration instead of Font changes
lbl.Text = "--- Section Header ---"  ' Use dashes or symbols for visual separation
```

### Avoid ContextMenuStrip and ToolStripMenuItem

- `ContextMenuStrip` and `ToolStripMenuItem` internally use `System.Drawing.Image`
- This causes: `Reference required to assembly 'System.Drawing.Common' containing the type 'Image'`
- **Use regular Buttons or ComboBox for actions instead:**

```vb
' BAD - ToolStripMenuItem uses System.Drawing.Image internally
Dim ctxMenu As New ContextMenuStrip()
Dim mnuItem As New ToolStripMenuItem("Action")  ' Error!

' GOOD - use Buttons for actions
Dim btnAction As New System.Windows.Forms.Button()
btnAction.Text = "Action"
btnAction.Left = 10
btnAction.Top = 400
btnAction.Width = 100
AddHandler btnAction.Click, Sub(s, e)
    ' Handle action
End Sub
frm.Controls.Add(btnAction)

' GOOD - use ComboBox for selection-based actions
Dim cboOptions As New System.Windows.Forms.ComboBox()
cboOptions.Items.Add("Option 1")
cboOptions.Items.Add("Option 2")
frm.Controls.Add(cboOptions)
```

### Lambda Closures Have Scoping Issues

- Variables declared after a lambda may cause: `Local variable cannot be referred to before it is declared`
- **ByRef parameters cannot be used in lambda expressions** - causes: `'ByRef' parameter 'paramName' cannot be used in a lambda expression`
- **Use `AddressOf` with a separate Sub, or read control values after dialog closes:**

```vb
' BAD - closure issues with variables declared later
AddHandler btn.Click, Sub(sender, e)
    ' code that references variables declared later
End Sub
Dim result As DialogResult = frm.ShowDialog()  ' Error!

' BAD - ByRef parameter in lambda
Function ShowForm(ByRef action As String) As DialogResult
    AddHandler btn.Click, Sub(s, e)
        action = "CLICK"  ' Error: ByRef parameter cannot be used in lambda
    End Sub
End Function

' GOOD - use Tag property and read after dialog closes
Function ShowForm(ByRef action As String) As DialogResult
    frm.Tag = ""
    AddHandler btn.Click, Sub(s, e)
        frm.Tag = "CLICK"
        frm.DialogResult = DialogResult.OK
    End Sub
    Dim result As DialogResult = frm.ShowDialog()
    action = CStr(frm.Tag)  ' Read after dialog closes
    Return result
End Function

' GOOD - separate handler with AddressOf
AddHandler btn.Click, AddressOf OnButtonClick

Sub OnButtonClick(sender As Object, e As EventArgs)
    ' Access form/controls via sender.Parent, Tag properties, or Controls collection
End Sub
```

### Modal Dialogs Block Inventor Input

- A modal dialog (`ShowDialog()`) blocks Inventor from receiving mouse clicks, even when minimized
- **Close the form before doing CommandManager.Pick(), then reopen:**
- **Always make sure the user can close a form from the close button or cancel**

```vb
Function RunPickerLoop(app As Inventor.Application) As String
    Dim baseName As String = ""
    Do
        Dim action As String = ""
        Dim result As DialogResult = ShowForm(baseName, action)
        
        If result = DialogResult.Cancel Then Return ""
        If result = DialogResult.OK Then Return baseName
        
        If action = "PICK" Then
            ' Form is now closed - Inventor can receive clicks
            baseName = DoComponentPick(app)
            ' Loop will reopen form with picked value
        End If
    Loop
End Function
```

### Passing Data Without Closures

- Use the `Tag` property on controls/forms to pass references:

```vb
' Store data
btnPick.Tag = app
frm.Tag = ""

' Retrieve in handler
Sub OnButtonClick(sender As Object, e As EventArgs)
    Dim btn As System.Windows.Forms.Button = CType(sender, System.Windows.Forms.Button)
    Dim frm As System.Windows.Forms.Form = CType(btn.Parent, System.Windows.Forms.Form)
    Dim app As Inventor.Application = CType(btn.Tag, Inventor.Application)
    
    ' Access named controls
    Dim txt As System.Windows.Forms.TextBox = CType(frm.Controls("txtName"), System.Windows.Forms.TextBox)
End Sub
```

## iLogic-Specific Objects

These are available in the iLogic execution context:

- `ThisApplication` - The Inventor.Application instance
- `ThisDoc.Document` - The active document
- `iLogicVb.Automation` - iLogic automation interface for managing rules

### iLogic Objects Not Available Inside Public Module

- `iLogicVb`, `ThisApplication`, `ThisDoc`, `Logger` are **only available in the `Sub Main()` context** of a runnable rule
- They are **NOT accessible inside a `Public Module`** included via `AddVbFile`
- This causes: `'iLogicVb' is not declared. It may be inaccessible due to its protection level.`
- **Solution: Pass these objects as parameters from the calling script**
- **For Logger: Pass a `List(Of String)` to collect log messages, then output them in the caller:**

```vb
' In library module - collect logs
Public Sub DoWork(ByVal logs As System.Collections.Generic.List(Of String))
    logs.Add("MyLib: Starting work...")
    logs.Add("MyLib: Work completed")
End Sub

' In calling script (Sub Main) - output logs
Dim logs As New System.Collections.Generic.List(Of String)
MyLibrary.DoWork(logs)
For Each logMsg As String In logs
    Logger.Info(logMsg)
Next
```

```vb
' In the runnable script (Sub Main)
AddVbFile "Lib/MyLibrary.vb"

Sub Main()
    Dim app As Inventor.Application = ThisApplication
    Dim iLogicAuto As Object = iLogicVb.Automation
    
    ' Pass to library function
    MyLibrary.DoSomething(app, iLogicAuto)
End Sub

' In the library module (MyLibrary.vb)
Public Module MyLibrary
    Public Sub DoSomething(ByVal app As Inventor.Application, ByVal iLogicAuto As Object)
        ' Now we can use app and iLogicAuto here
        iLogicAuto.AddRule(doc, "RuleName", ruleText)
    End Sub
End Module
```

### Inventor API Members Not Available in iLogic

Some Inventor API members that work in standalone applications do not work in iLogic:

| API Member | Error | Use Instead |
|------------|-------|-------------|
| `app.Assets(AssetTypeEnum.kAssetTypeMaterial)` | `Public member 'Assets' on type 'Application' not found` | `partDoc.Materials` |
| `DerivedPartUniformScaleDef.IncludeAllWorkSurfaces` | `Public member not found` | Only set `DeriveStyle` and individual `IncludeEntity` |
| `DerivedPartUniformScaleDef.IncludeAllParameters` | `Public member not found` | (same as above) |
| `DerivedPartUniformScaleDef.LinkFaceColor` | `Public member not found` | (same as above) |
| `DerivedPartUniformScaleDef.Sketches2D` | `Public member not found` | Use `Sketches` instead (wrap in Try/Catch) |
| Save after `smCompDef.Unfold()` | `E_FAIL` error during SaveAs | Call `smCompDef.FlatPattern.ExitEdit()` before save |

### Parameter Formula Syntax

When creating or updating Inventor parameters with formulas via the API, note the following:

**Semicolon as argument separator**: Inventor uses semicolon `;` instead of comma `,` as the function argument separator in parameter formulas. This is locale-independent and applies when setting `Parameter.Expression` or using `UserParameters.AddByExpression`.

```vb
' BAD - comma separator causes E_INVALIDARG error
param.Expression = "max(1, ceil(span / spacing) - 1)"

' GOOD - semicolon separator works
param.Expression = "max(1; ceil(span / spacing) - 1)"
```

**Common formula patterns:**

| Formula | Syntax |
|---------|--------|
| Maximum of two values | `max(a; b)` |
| Minimum of two values | `min(a; b)` |
| Ceiling (round up) | `ceil(value)` |
| Floor (round down) | `floor(value)` |
| Round | `round(value)` or `round(value; decimals)` |

**Example - Creating a parametric count:**

```vb
' Count formula with minimum of 1
Dim countFormula As String = "max(1; ceil(" & spanParam & " / " & maxSpacingParam & ") - 1)"
userParams.AddByExpression("MyCount", countFormula, UnitsTypeEnum.kUnitlessUnits)
```

**Parameter names must start with a letter**: Parameter names cannot begin with a digit. If using numeric identifiers, prefix with a letter (e.g., `M_00011_Ulatus` instead of `00011_Ulatus`).

**Material enumeration example:**

```vb
' BAD - app.Assets doesn't work in iLogic
For Each asset As Asset In app.Assets(AssetTypeEnum.kAssetTypeMaterial)
    materials.Add(asset.DisplayName)  ' Error!
Next

' GOOD - use partDoc.Materials
For Each mat As Material In partDoc.Materials
    materials.Add(mat.Name)  ' Works!
Next
```

**Excluding sketches and work features from derived parts:**

When deriving parts, sketches, work features, and parameters from the master may be included by default. To derive only the solid body, iterate through the entity collections and set `IncludeEntity = False`.

**Important:** Not all collections exist on `DerivedPartUniformScaleDef`:
- `Sketches2D` does NOT exist - use `Sketches` instead
- Wrap each property access in `Try/Catch` because the property access itself throws if the property doesn't exist

```vb
Dim dpDef As DerivedPartUniformScaleDef = dpcs.CreateUniformScaleDef(masterDoc.FullDocumentName)

' Include only target body
For Each dpe As DerivedPartEntity In dpDef.Solids
    If GetBodyName(dpe) = targetBodyName Then
        dpe.IncludeEntity = True
    Else
        dpe.IncludeEntity = False
    End If
Next

' Exclude all sketches, work features, surfaces, parameters
' CRITICAL: Wrap EACH property access in Try/Catch - not just the loop
' The property access itself throws if property doesn't exist on this type
Try
    For Each dpe As DerivedPartEntity In dpDef.Sketches3D : dpe.IncludeEntity = False : Next
Catch : End Try
Try
    For Each dpe As DerivedPartEntity In dpDef.Sketches : dpe.IncludeEntity = False : Next
Catch : End Try
Try
    For Each dpe As DerivedPartEntity In dpDef.WorkFeatures : dpe.IncludeEntity = False : Next
Catch : End Try
Try
    For Each dpe As DerivedPartEntity In dpDef.Surfaces : dpe.IncludeEntity = False : Next
Catch : End Try
Try
    For Each dpe As DerivedPartEntity In dpDef.Parameters : dpe.IncludeEntity = False : Next
Catch : End Try

dpDef.DeriveStyle = DerivedComponentStyleEnum.kDeriveAsSingleBodyWithSeams
dpcs.Add(dpDef)
```

### Creating/Updating Rules Programmatically

```vb
Dim iLogicAuto As Object = iLogicVb.Automation
Dim existingRule As Object = Nothing

Try
    existingRule = iLogicAuto.GetRule(asmDoc, ruleName)
Catch
    existingRule = Nothing
End Try

If existingRule IsNot Nothing Then
    existingRule.Text = ruleText  ' Update
Else
    iLogicAuto.AddRule(asmDoc, ruleName, ruleText)  ' Create
End If
```

## Component Selection

### Using CommandManager.Pick

- **Always include ESC cancel instruction in prompts** - Users must be able to cancel out of pick operations by pressing ESC. Include `" - ESC tühistamiseks"` (or equivalent) at the end of every pick prompt so users know they can exit if there's nothing valid to select.
- **Handle cancelled picks gracefully** - When the user presses ESC, `CommandManager.Pick` throws an exception. Always wrap in Try/Catch and treat the exception as a cancellation (return Nothing, empty string, or exit gracefully).

```vb
Dim selFilter As SelectionFilterEnum = SelectionFilterEnum.kAssemblyOccurrenceFilter
Dim selectedObj As Object = Nothing

Try
    selectedObj = app.CommandManager.Pick(selFilter, "Vali komponent - ESC tühistamiseks")
Catch
    ' User cancelled with ESC - handle gracefully
    Return Nothing  ' or Exit Sub, etc.
End Try

If TypeOf selectedObj Is ComponentOccurrence Then
    Dim occ As ComponentOccurrence = CType(selectedObj, ComponentOccurrence)
    Dim occName As String = occ.Name  ' e.g., "PartName:1"
End If
```

### SelectionFilterEnum Cannot Be Combined with Or

- `SelectionFilterEnum` values are **NOT bit flags** - you cannot combine them with `Or`
- This silently fails: `kWorkPlaneFilter Or kPartFaceFilter`
- **Use aggregate filters instead:**

```vb
' BAD - combining filters doesn't work
Dim filter As SelectionFilterEnum = SelectionFilterEnum.kWorkPlaneFilter Or _
                                    SelectionFilterEnum.kPartFaceFilter

' GOOD - use aggregate filter that covers multiple types
Dim filter As SelectionFilterEnum = SelectionFilterEnum.kAllPlanarEntities  ' planes and faces
Dim filter As SelectionFilterEnum = SelectionFilterEnum.kAllLinearEntities  ' axes and edges
Dim filter As SelectionFilterEnum = SelectionFilterEnum.kAllPointEntities   ' points and vertices
```

### Extracting Base Name from Occurrence

Occurrence names have format `"ComponentName:InstanceNumber"`. To get base name:

```vb
Dim colonPos As Integer = occName.LastIndexOf(":")
If colonPos > 0 Then
    baseName = occName.Substring(0, colonPos)
Else
    baseName = occName
End If
```

### Selecting/Highlighting Objects Programmatically

- Objects like `ComponentOccurrence` do **NOT** have a `Select()` method
- This causes: `Public member 'Select' on type 'ComponentOccurrence' not found.`
- **Use `SelectSet.Select(object)` on the document's SelectSet:**

```vb
' BAD - ComponentOccurrence has no Select method
occ.Select()

' GOOD - use the document's SelectSet
asmDoc.SelectSet.Clear()
asmDoc.SelectSet.Select(occ)
app.ActiveView.Update()  ' Refresh the view to show selection
```

## Common Patterns

### Document Type Checking

```vb
If doc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
    Dim asmDoc As AssemblyDocument = CType(doc, AssemblyDocument)
ElseIf doc.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
    Dim partDoc As PartDocument = CType(doc, PartDocument)
End If
```

### Error Handling

```vb
Try
    ' Inventor API call
Catch ex As Exception
    MessageBox.Show("Error: " & ex.Message, "Title")
    Exit Sub
End Try
```

### Sub Calls Require Parentheses

- VB.NET (and iLogic) requires parentheses around Sub arguments
- Old VB6 syntax without parentheses causes: `Method arguments must be enclosed in parentheses`

```vb
' BAD - VB6 syntax, causes error
SetCustomProp doc, "Name", "Value"

' GOOD - VB.NET syntax with parentheses
SetCustomProp(doc, "Name", "Value")
```

### Library Function Signature Changes

- When changing a function signature in a library (`Lib/*.vb`), **update ALL callers**
- iLogic compiles each rule independently, so callers will fail with `Argument not specified` errors
- Search for the function name across all `.vb` files to find callers
- This causes: `Argument not specified for parameter 'paramName' of 'FunctionName'`

### Document Update Pattern (DocumentUpdateLib)

Use `DocumentUpdateLib` to register code that should run automatically when parameters change or other events occur. The library manages a local "Uuenda" rule with UID-guarded sections.

**When to use:**
- Your script sets up parameters that need recalculation when changed
- You want code to run on parameter changes, save events, or document open
- Multiple features need to share an auto-triggered update rule

**UpdateTrigger Enum (callers never need PropIds):**

| Enum Value | Event |
|------------|-------|
| `ModelParameterChange` | Any model parameter changes |
| `UserParameterChange` | Any user parameter changes |
| `BeforeSave` | Before document is saved |
| `AfterSave` | After document is saved |
| `DocumentOpen` | After document is opened |
| `PartGeometryChange` | Part geometry changes (parts only) |
| `MaterialChange` | Material changes (parts only) |
| `iPropertyChange` | Any iProperty changes |

**Usage Example:**

```vb
AddVbFile "Lib/DocumentUpdateLib.vb"
AddVbFile "Lib/SupportPlacementLib.vb"

Sub Main()
    Dim doc As Document = ThisDoc.Document
    Dim iLogicAuto As Object = iLogicVb.Automation
    
    ' Register update handler with triggers
    Dim updateCode() As String = {
        "SupportPlacementLib.UpdateAllSupportLengths(ThisApplication, ThisDoc.Document)"
    }
    Dim triggers() As DocumentUpdateLib.UpdateTrigger = {
        DocumentUpdateLib.UpdateTrigger.ModelParameterChange,
        DocumentUpdateLib.UpdateTrigger.UserParameterChange
    }
    DocumentUpdateLib.RegisterUpdateHandler(doc, iLogicAuto, "SupportLengths", updateCode, triggers)
End Sub
```

**API Reference:**

| Function | Description |
|----------|-------------|
| `RegisterUpdateHandler(doc, iLogicAuto, uid, codeLines(), triggers())` | Adds/updates a section in the Uuenda rule |
| `RemoveUpdateHandler(doc, iLogicAuto, uid)` | Removes a section from the Uuenda rule |
| `EnsureUpdateRule(doc, iLogicAuto)` | Creates the Uuenda rule if missing |
| `AddTrigger(doc, trigger)` | Adds a trigger (skips if already exists) |

**Section Marker Format:**

The Uuenda rule uses UID-guarded sections that should not be manually edited:

```vb
Sub Main()
    ' === BEGIN: SupportLengths ===
    SupportPlacementLib.UpdateAllSupportLengths(ThisApplication, ThisDoc.Document)
    ' === END: SupportLengths ===
    
    ' === BEGIN: BoundingBoxStock ===
    BoundingBoxStockLib.RecalculateStock(ThisApplication, ThisDoc.Document)
    ' === END: BoundingBoxStock ===
End Sub
```

## Vault Integration

### Detecting Vault Checkout Status

- **`doc.ReservedForWriteByMe`** - Returns `True` if document is checked out by current user (reliable)
- **`doc.IsModifiable`** - Unreliable for Vault; returns `True` even when file is not checked out

```vb
' Correctly detect if checkout is needed
If Not doc.ReservedForWriteByMe Then
    ' Document needs checkout
End If
```

### Methods/Properties That Do NOT Exist

These cause "Public member not found" errors despite appearing in some documentation:
- `doc.CheckOut()` - does not exist on PartDocument/AssemblyDocument
- `doc.Reserve()` - does not exist
- `doc.ReservationStatus` - does not exist
- `app.FileAccessEvents.AutoCheckOut` - does not exist

### Vault Add-in Access

- **Add-in GUID**: `{48B682BC-42E6-4953-84C5-3D253B52E77A}`
- **Add-in name**: "Inventor Vault"
- The `Automation` property returns a COM object but exposes no accessible methods via reflection or late binding

```vb
Dim vaultAddin As ApplicationAddIn = app.ApplicationAddIns.ItemById("{48B682BC-42E6-4953-84C5-3D253B52E77A}")
If vaultAddin IsNot Nothing AndAlso vaultAddin.Activated Then
    Dim vaultAuto As Object = vaultAddin.Automation  ' Returns System.__ComObject, methods not accessible
End If
```

### Vault Commands via CommandManager

Available checkout commands (all show a dialog - no silent checkout possible):

```vb
Dim cmdMgr As CommandManager = app.CommandManager
Dim ctrlDef As ControlDefinition = cmdMgr.ControlDefinitions.Item("VaultCheckout")
ctrlDef.Execute()  ' Shows checkout dialog
```

| Command Internal Name | Display Name |
|-----------------------|--------------|
| `VaultCheckout` | Check Out |
| `VaultCheckoutTop` | Check Out |
| `VaultUndoCheckout` | Undo Check Out... |
| `VaultGetCheckout` | Get Revision... |

### File Attributes

- Vault-managed files have `ReadOnly` attribute set on disk
- Removing `ReadOnly` via `File.SetAttributes` does NOT bypass Vault control

## Summary of Key Constraints

> **IMPORTANT**: These constraints must be checked BEFORE writing code, not after compile errors.
> See the **QUICK REFERENCE** section at the top of this document for checklists.

| Issue | Solution |
|-------|----------|
| **ByRef in lambda** | **NEVER use ByRef params in lambda. Store in form.Tag, read AFTER ShowDialog()** |
| Module-level Dim/Const outside Module | Declare inside Sub Main or pass via parameters |
| Module inside Sub Main | Separate into different files |
| AddReference/AddVbFile in library | Put ALL AddReference/AddVbFile in main script only |
| AddVbFile before AddReference | Put AddReference BEFORE AddVbFile |
| Type from referenced assembly in library | Use `Object` type with late binding |
| TextBox ambiguous | Use `System.Windows.Forms.TextBox` |
| Size/Point/Font not defined | Use Left/Top/Width/Height; use text decoration for emphasis |
| ContextMenuStrip/ToolStripMenuItem | Use Buttons or ComboBox instead |
| Lambda closure errors | Use `AddressOf` with separate Sub, or use Tag property |
| Modal blocks Inventor picks | Close form before pick, reopen after |
| Sub call without parentheses | Always use `SubName(args)` not `SubName args` |
| iLogicVb/ThisApplication in Module | Pass as parameters from Sub Main |
| SelectionFilterEnum combined with Or | Use aggregate filters like kAllPlanarEntities |
| Library function signature changed | Update ALL callers across all .vb files |
| CreateGeometryProxy is a Sub, not Function | Use `occ.CreateGeometryProxy(obj, result)` with ByRef result |
| object.Select() not found | Use `doc.SelectSet.Select(object)` instead |
| Vault checkout status | Use `doc.ReservedForWriteByMe`, not `doc.IsModifiable` |
| Vault silent checkout | Not possible; all commands show dialogs |
| app.Assets() not found | Use `partDoc.Materials` to enumerate materials |
| DerivedPartUniformScaleDef.IncludeAll* | Only use `DeriveStyle` and individual `IncludeEntity` |
| DerivedPartUniformScaleDef.Sketches2D | Use `Sketches` instead (wrap property access in Try/Catch) |
| Save fails after Unfold() | Call `smCompDef.FlatPattern.ExitEdit()` before SaveAs |
| API optional parameters | Use `NameValueMap` with `AdditionalOptions :=` named parameter |
| Sheet metal flat pattern view | `AddBaseView` with `NameValueMap("SheetMetalFoldedModel", False)` |
| Drawing view spacing/positioning | Use `view.Width` and `view.Height`, not model `RangeBox` dimensions |
| Extent dimension spacing | Use `CAMDrawingLib.DIMENSION_OFFSET` (25mm) for spacing from model |
| Sheet resize fails | Move views within new bounds FIRST, then resize sheet |
| Component patterns to browser folder | `BrowserFolder.Add()` fails with E_FAIL for all pattern types (Mirror, Rectangular, Circular). Patterns must be moved manually. |
| Mirror Component Pattern suppression | Mirror Component Patterns (Inventor 2026 associative) cannot be suppressed via API. `NativeObject` throws E_NOTIMPL, and suppressing individual occurrences breaks/flips the pattern. User must manually configure model states for Mirror patterns. |
| Parameter formula `max(a, b)` fails | Use semicolon: `max(a; b)` - Inventor uses `;` as argument separator |
| Parameter name starts with digit | Prefix with letter: `M_00011_Name` instead of `00011_Name` |
| CommandManager.Pick with no cancel hint | Always include `" - ESC tühistamiseks"` in prompt; wrap in Try/Catch |

