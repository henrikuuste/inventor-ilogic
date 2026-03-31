# Inventor iLogic Development Guidelines

This document captures constraints and best practices for writing Autodesk Inventor iLogic rules and VB scripts.

- Use the 2026 iLogic API for reference https://help.autodesk.com/view/INVNTOR/2026/ENU/?guid=110f3019-404c-4fc4-8b5d-7a3143f129da

## File Structure

### Runnable Rules vs Library Modules

- **Runnable rules** should have a `Sub Main()` at the file level - this is the entry point when the rule is executed
- **Library modules** use `Public Module ModuleName ... End Module` to expose reusable functions
- Library modules are located in the Lib folder
- **Do NOT mix them in the same file** - a `Module` statement nested inside or after `Sub Main()` causes: `'Module' statements can occur only at file or namespace level`
- To use a library from a rule, use `AddVbFile "Lib/LibraryName.vb"` at the top of the rule

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
lbl.Font = New System.Drawing.Font(lbl.Font, System.Drawing.FontStyle.Bold)

' GOOD - use individual properties
btn.Left = 10
btn.Top = 20
btn.Width = 80
btn.Height = 30

' For emphasis, use text decoration instead of Font changes
lbl.Text = "--- Section Header ---"  ' Use dashes or symbols for visual separation
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

- `iLogicVb`, `ThisApplication`, `ThisDoc` are **only available in the `Sub Main()` context** of a runnable rule
- They are **NOT accessible inside a `Public Module`** included via `AddVbFile`
- This causes: `'iLogicVb' is not declared. It may be inaccessible due to its protection level.`
- **Solution: Pass these objects as parameters from the calling script:**

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

```vb
Dim selFilter As SelectionFilterEnum = SelectionFilterEnum.kAssemblyOccurrenceFilter
Dim selectedObj As Object = Nothing

Try
    selectedObj = app.CommandManager.Pick(selFilter, "Select a component:")
Catch
    ' User cancelled
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

## Summary of Key Constraints

| Issue | Solution |
|-------|----------|
| Module-level Dim/Const outside Module | Declare inside Sub Main or pass via parameters |
| Module inside Sub Main | Separate into different files |
| TextBox ambiguous | Use `System.Windows.Forms.TextBox` |
| Size/Point/Font not defined | Use Left/Top/Width/Height; use text decoration for emphasis |
| Lambda closure errors | Use `AddressOf` with separate Sub, or use Tag property |
| ByRef in lambda | Read control values after dialog closes, use form.Tag for action |
| Modal blocks Inventor picks | Close form before pick, reopen after |
| Sub call without parentheses | Always use `SubName(args)` not `SubName args` |
| iLogicVb/ThisApplication in Module | Pass as parameters from Sub Main |
| SelectionFilterEnum combined with Or | Use aggregate filters like kAllPlanarEntities |
| Library function signature changed | Update ALL callers across all .vb files |
| CreateGeometryProxy is a Sub, not Function | Use `occ.CreateGeometryProxy(obj, result)` with ByRef result |

