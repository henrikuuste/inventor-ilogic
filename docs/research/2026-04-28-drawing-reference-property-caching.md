# Drawing Reference Property Caching Issue

**Date**: 2026-04-28  
**Status**: Resolved  
**Related**: Module Release System (`ModuleReleaseLib.vb`)

## Problem Statement

When releasing modules, drawings that reference assemblies were displaying the **old Part Number** in title blocks even though:
- The assembly file was created with correct Part Number and Title properties
- Opening the assembly directly showed correct properties
- The drawing's file descriptors were updated with `ReplaceReference`

## Root Cause Analysis

### The Ancestry Requirement

The `FileDescriptor.ReplaceReference()` method has a documented requirement:

> "The file being replaced and the replacement file must share ancestry (i.e. they must have the same InternalName). Documents have the same internal name if they are copied using 'Save Copy As' or a file explorer copy."

**Source**: [Autodesk API Documentation](https://help.autodesk.com/cloudhelp/2021/ENU/Inventor-API/files/FileDescriptor.htm)

### Our Workflow Conflict

1. **We use `Document.SaveAs()`** to create new assemblies - this generates a **new InternalName (GUID)**
2. **We need unique GUIDs** so source and target files can coexist without document conflicts
3. **But `ReplaceReference` requires same GUIDs** (ancestry) to work correctly

When files have different InternalNames, `ReplaceReference` may update the file path but Inventor's internal property caching doesn't properly resolve the new document's properties.

### Evidence from Logs

```
INFO|  Opened fresh: 00016.iam PN=00016      <- Direct document has CORRECT PN
INFO|  Drawing updated after loading fresh refs
INFO|  Pre-SaveAs ref check: 00016.iam PN=00003   <- But ReferencedDocuments shows WRONG PN
```

The `drawDoc.ReferencedDocuments` collection returns Document objects with stale cached property values, even after:
- Closing old referenced documents
- Opening new referenced documents fresh from disk
- Calling `drawDoc.Update()`

## Solution

### Key Insight: ReferencedDocuments vs DrawingView.ReferencedDocument

The `drawDoc.ReferencedDocuments` collection returns **stale cached Document objects** that don't reflect the actual file properties on disk.

However, **`DrawingView.ReferencedDocumentDescriptor.ReferencedDocument`** may return the correct model with correct properties.

### The Two-Pronged Approach

1. **Save/Close/Reopen the drawing** - Forces Inventor to re-resolve references from disk
2. **Explicitly set drawing iProperties from model** - Don't rely on automatic copying

From Clint Brown's blog (https://clintbrown.co.uk/2019/02/02/ilogic-title-block/):
> "When Inventor drawings are copied, particularly during a 'File > Save As > Save Copy As' operation, the title block of the drawing will show the original save date of the initial drawing. If you were to then replace the reference models, the title block would again show the original part/assembly name."

His solution: Read Part Number from model and explicitly write to drawing:
```vb
oGetTheName = ActiveSheet.View("VIEW1").ModelDocument.DisplayName
iProperties.Value("Project", "Part Number") = oName
```

### Implementation

```vb
' PHASE 1: Open source, replace references, save to temp file
Dim drawDoc = app.Documents.Open(sourceDrawingPath, True)
' ... ReplaceReference calls ...
drawDoc.SaveAs(tempPath, False)
drawDoc.Close(True)  ' CRITICAL: Clear all internal caches

' PHASE 2: Reopen - forces reference re-resolution from disk
drawDoc = app.Documents.Open(tempPath, True)

' PHASE 3: Get Part Number from drawing view's model (most reliable method)
Dim actualModelPN As String = ""
For Each sheet As Sheet In drawDoc.Sheets
    For Each view As DrawingView In sheet.DrawingViews
        If view.ReferencedDocumentDescriptor?.ReferencedDocument IsNot Nothing Then
            Dim modelDoc = view.ReferencedDocumentDescriptor.ReferencedDocument
            actualModelPN = modelDoc.PropertySets.Item("Design Tracking Properties").Item("Part Number").Value
            Exit For
        End If
    Next
    If actualModelPN <> "" Then Exit For
Next

' PHASE 4: Explicitly set drawing properties (don't rely on auto-copy)
drawDoc.PropertySets.Item("Design Tracking Properties").Item("Part Number").Value = actualModelPN
drawDoc.PropertySets.Item("Inventor Summary Information").Item("Title").Value = actualModelPN

' PHASE 5: Force update and save
app.CommandManager.ControlDefinitions.Item("UpdateCopiedModeliPropertiesCmd").Execute2(True)
drawDoc.Update()
drawDoc.SaveAs(targetPath, False)
drawDoc.Close(True)
System.IO.File.Delete(tempPath)  ' Cleanup
```

### Additional Refresh Command

Also execute the built-in command to force iProperty sync:

```vb
Dim oControlDef As ControlDefinition = app.CommandManager.ControlDefinitions.Item("UpdateCopiedModeliPropertiesCmd")
oControlDef.Execute2(True)
```

## Alternative Approaches (Not Used)

### 1. Use File.Copy Instead of SaveAs

Would preserve the InternalName so `ReplaceReference` works correctly, but creates the original GUID conflict problem where source and target files can't coexist.

### 2. ComponentOccurrence.Replace()

Works for assemblies (no ancestry requirement) but there's no equivalent for drawing references.

### 3. Recreate Drawing Views

Some forum posts suggest deleting and recreating views. Too complex and risks losing annotations.

## References

- [FileDescriptor ReplaceReference "must share ancestry"](https://forums.autodesk.com/t5/inventor-programming-forum/filedescriptor-replacereference-method-quot-must-share-ancestry/td-p/13723228) (July 2025)
- [Replacing Referenced Document in Drawing does not update until reopened](https://forums.autodesk.com/t5/inventor-programming-forum/replacing-referenced-document-in-drawing-does-not-update-until/td-p/11585801)
- [Update drawing properties from model iProperties](https://forums.autodesk.com/t5/inventor-programming-forum/update-drawing-properties-from-model-iproperties-automatically/td-p/3078328)
- [Manufacturing DevBlog: Replace File Reference](https://adndevblog.typepad.com/manufacturing/2012/08/replace-the-file-reference-by-inventor-api.html)
