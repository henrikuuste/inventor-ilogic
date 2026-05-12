<!-- Copyright (c) 2026 Henri Kuuste -->
---
date: 2026-04-25T06:26:00+03:00
researcher: Claude
topic: "Loo detailid - Save Errors and Flat Pattern Failures Analysis"
tags: [research, debugging, sheet-metal, vault, loo-detailid]
status: complete
last_updated: 2026-05-12
---

> **TERMINOLOGY NOTE (2026-05-12)**: Script renamed from "Loo komponendid" to "Loo detailid" per `docs/UBIQUITOUS_LANGUAGE.md`. "Detail" is the Estonian term for "Part".

# Research: Loo detailid - Save Errors and Flat Pattern Failures

**Date**: 2026-04-25
**Topic**: Debugging failures in multi-body component creation script

## Research Question

Two issues were observed when running `Loo detailid.vb`:
1. Assembly save error mid-process with Vault popup interruption
2. Flat pattern creation failed for most sheet metal parts (all but 2), though manual creation worked

## Summary

### Issue 1: Assembly Save Error with Vault

The script performs multiple sequential save operations that each can trigger Vault's "Add to Vault" dialog. These dialogs are modal and block script execution, causing timing issues and potential state corruption.

### Issue 2: Flat Pattern Failures

The flat pattern creation fails because **face references become invalid** after converting to sheet metal. The `FindFaceByNormal` function searches for faces *after* the part subtype has been changed and updated, but the original face geometry has been rebuilt. The working manual `Lehtmetall.vb` script avoids this by having the user select the face *before* conversion.

## Detailed Findings

### Issue 1: Save Operations and Vault Interaction

#### Current Save Sequence in `Loo detailid.vb`

```
For each body:
  1. newPart.SaveAs(filePath, False)     → Triggers Vault "Add to Vault" dialog
  2. (Part closed)
  
After all parts:
  3. masterDoc.Save()                     → May trigger Vault dialog
  4. asmDoc.SaveAs() or asmDoc.Save()    → Triggers Vault dialog  
  5. masterDoc.Save() (again)            → May trigger Vault dialog
```

#### Problem Areas

**Line 394-412 (`Loo detailid.vb`)** - Part save:
```vb
Try
    newPart.SaveAs(filePath, False)
    ' Vault may rename file - read actual path
    Dim actualPath As String = newPart.FullDocumentName
```

**Lines 447-468** - Assembly save:
```vb
If assemblyAction = "CREATE" AndAlso Not String.IsNullOrEmpty(assemblyPath) Then
    asmDoc.SaveAs(assemblyPath, False)  ' <-- Vault intercepts this
```

#### Root Causes

1. **Modal Vault Dialogs**: Each `SaveAs` call triggers Vault's "Add to Vault" dialog which is modal and blocks script execution
2. **File Renaming**: Vault may rename files (e.g., assigning part numbers), causing the script's expected `filePath` to differ from `FullDocumentName`
3. **Folder Relocation**: Vault may move files to different folders based on numbering scheme rules
4. **No Dialog Handling**: The script has no mechanism to wait for or suppress Vault dialogs
5. **Sequential Blocking**: Multiple save operations queue up dialogs, confusing users

### Issue 2: Flat Pattern Creation Failure

#### The Critical Difference

**Working approach (`Lehtmetall.vb` lines 54-57, 92)**:
```vb
' Face selected BEFORE conversion
Dim aSideFace As Face = PickASideFace(app)  ' User picks face manually
' ... later, same face reference used ...
partDoc.SubType = SHEET_METAL_GUID
partDoc.Update()
CreateFlatPattern(smCompDef, aSideFace)  ' Uses original face reference
```

**Failing approach (`SheetMetalLib.vb` lines 146-180)**:
```vb
' 1. Convert to sheet metal first
partDoc.SubType = SHEET_METAL_GUID
partDoc.Update()  ' Geometry rebuilt here!

' 2. THEN try to find face - but geometry changed!
Dim aSideFace As Face = FindFaceByNormal(partDoc, thicknessVector)
If aSideFace IsNot Nothing Then
    CreateFlatPattern(smCompDef, aSideFace)
End If
```

#### Why Face References Become Invalid

1. **Geometry Rebuild**: When `partDoc.SubType` changes to sheet metal GUID and `Update()` is called, Inventor rebuilds the part's geometry as sheet metal features
2. **New Face Topology**: The faces in the rebuilt sheet metal model are *new objects*, not the same as the original solid body faces
3. **Normal Matching Failure**: `FindFaceByNormal` tries to match faces by their normal vectors, but:
   - The coordinate system may have shifted
   - Faces may have been split or merged
   - Internal representation changes

#### `FindFaceByNormal` Function Analysis

```vb
Public Function FindFaceByNormal(partDoc As PartDocument, thicknessVector As String) As Face
    ' Iterates through ComponentDefinition.SurfaceBodies
    For Each body As SurfaceBody In partDoc.ComponentDefinition.SurfaceBodies
        For Each face As Face In body.Faces
            ' Checks if face normal matches stored vector
            If Math.Abs(Math.Abs(dot) - 1) < 0.001 Then
                Return face  ' Returns first matching face
```

The function searches the *current* model state, but after sheet metal conversion the body has been restructured.

#### Why 2 Parts Worked (Hypothesis)

Parts that succeeded likely had:
- Simpler geometry where face normals remained stable through conversion
- Faces aligned exactly with principal axes (X, Y, Z) which survive regeneration
- Less complex topology that didn't cause face splitting

#### `CreateFlatPattern` Silent Failure

```vb
Public Sub CreateFlatPattern(smCompDef As SheetMetalComponentDefinition, aSideFace As Face)
    Try
        smCompDef.ASideFace = aSideFace
        smCompDef.Unfold()
    Catch ex As Exception
        UtilsLib.LogWarn("SheetMetalLib: Could not create flat pattern: " & ex.Message)
    End Try
End Sub
```

The exception is caught and logged as warning, but `ConvertToSheetMetal` still returns `True`:
```vb
' Line 178-179 in SheetMetalLib.vb
partDoc.Update()
Return True  ' Returns True even if flat pattern failed!
```

## Other Potential Failure Modes

### 1. Timing/State Issues
- **Missing Update() calls**: Multiple property changes without regeneration
- **Derivation not complete**: Derived part may not be fully computed before sheet metal conversion
- Sequence: `DeriveBodyAsNewPart` → `SetPartProperties` → `AssignMaterial` → `ConvertToSheetMetal`

### 2. Vault-Related
- **File Locking**: Vault may lock files during check-in, preventing subsequent operations
- **Concurrent Dialog Popups**: Multiple Vault dialogs stacking
- **File Rename Race**: Script references old path while Vault has renamed file

### 3. Sheet Metal Conversion
- **Unsupported Geometry**: Some solid bodies may not be suitable for sheet metal conversion
- **Style Not Found**: If "Default_mm" style doesn't exist
- **Thickness Parameter Issues**: If thickness value doesn't match model expectations

### 4. Exception Swallowing
Multiple Try/Catch blocks swallow exceptions:
- `CreateFlatPattern` - line 237 of `SheetMetalLib.vb`
- `SetThickness` - line 198 of `SheetMetalLib.vb`
- Various property setters throughout

## Code References

### Main Script
- `Loo detailid.vb:378-391` - Sheet metal conversion call
- `Loo detailid.vb:394-415` - Part save with Vault interaction
- `Loo detailid.vb:434-468` - Assembly save operations

### Sheet Metal Library  
- `Lib/SheetMetalLib.vb:146-180` - `ConvertToSheetMetal` function
- `Lib/SheetMetalLib.vb:85-112` - `FindFaceByNormal` function
- `Lib/SheetMetalLib.vb:226-239` - `CreateFlatPattern` function

### Working Reference
- `Lehtmetall.vb:54-57` - Manual face selection BEFORE conversion
- `Lehtmetall.vb:92` - Using pre-selected face for flat pattern

## Implemented Fixes

### Issue 1 (Save Errors) - Restructured Save Flow

Changes to `Loo detailid.vb`:

1. **Keep parts open during creation**: Parts are no longer closed immediately after SaveAs
2. **Batch save approach**: 
   - If assembly exists: Save parts first, place in assembly, save assembly once
   - If no assembly: Save parts at the end in sequence
3. **Single master save**: Removed intermediate master document save, only save once at end
4. **Cleaner error handling**: Parts stay open until assembly is saved

The new flow:
```
1. Create all parts (derive, set properties, convert to sheet metal)
2. If assembly:
   a. Save all parts (Vault dialogs here, but grouped)
   b. Place all parts in assembly
   c. Save assembly
   d. Close all parts
3. If no assembly:
   a. Save and close parts one by one at end
4. Save master document once
```

### Issue 2 (Flat Patterns) - Face Matching by Center Point

Changes to `SheetMetalLib.vb`:

1. **New `FindASideFace` function**: Finds A-side face using center point matching
   - Stores face center point before conversion
   - After conversion, finds the face at the same location
   - Falls back to choosing face with highest projection along thickness vector

2. **New `GetFaceCenterPoint` function**: Gets the center point of a planar face

3. **Updated `ConvertToSheetMetal`**:
   - Step 1: Find A-side face and store center point BEFORE conversion
   - Step 2: Convert to sheet metal subtype
   - Step 3: Set style, thickness, properties
   - Step 4: Call `Update()` to apply changes
   - Step 5: Re-find face using stored center point
   - Step 6: Create flat pattern

4. **Added missing `Update()` call**: After setting style/thickness, before finding face

## Open Questions (Remaining)

1. Does center point matching work reliably across all geometry types?
2. Should we add explicit face area matching as additional validation?
3. Consider user notification if flat pattern creation fails despite face being found
