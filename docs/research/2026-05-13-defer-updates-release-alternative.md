<!-- Copyright (c) 2026 Henri Kuuste -->
---
date: 2026-05-13T08:45:00+03:00
researcher: Claude
git_commit: 20c023d5080552ab81d268827b3cc090c1e0bd9f
branch: main
repository: Inventor Rules
topic: "Alternative Release Strategy Using Defer Updates"
tags: [research, release, defer-updates, simplification]
status: complete
last_updated: 2026-05-13
---

# Research: Alternative Release Strategy Using Defer Updates

**Date**: 2026-05-13 08:45 EET
**Git Commit**: 20c023d5080552ab81d268827b3cc090c1e0bd9f
**Branch**: main

## Research Question

Is there an easier way to create frozen/fixed releases by using Inventor's "Defer Updates" feature instead of creating standalone file copies?

## Summary

Inventor's `DeferUpdates` property can freeze assemblies and drawings at a specific geometry state without creating standalone copies. This could significantly simplify the release process, but has important tradeoffs. The key insight is that **production only needs unique drawings with unique numbers** - the part files themselves don't necessarily need separate copies if the sharing classification is documented.

## Current Approach (Standalone Copies)

The current `ElementReleaseLib.vb` (~3100 lines) creates complete standalone copies:

1. **Fingerprint Analysis**: Cycle through all element parameters to compute geometry fingerprints
2. **Part Copies**: Create standalone copies of all parts (break derivation links)
3. **Assembly Copies**: Create assembly copies with reference replacements
4. **Drawing Copies**: Create drawing copies with reference replacements
5. **Vault Numbers**: Each released file gets a unique Vault number

### Pros
- Files are completely independent (no dependency on masters)
- No risk of accidental updates
- Can archive/move files without breaking references
- Each file has its own Vault lifecycle

### Cons
- Many Vault numbers consumed
- Complex reference replacement logic
- ~3100 lines of code to maintain
- Slow: must cycle parameters, update geometry, create copies

## Alternative: Defer Updates Approach

### Core Concept

Use Inventor's `DeferUpdates` property to "freeze" drawings at specific geometry states while still referencing the parametric masters:

```vb
' For drawings
DrawingDocument.DrawingSettings.DeferUpdates = True

' For assemblies
AssemblyDocument.AssemblyOptions.DeferUpdate = True

' When opening with defer enabled
Dim nvm As NameValueMap = app.TransientObjects.CreateNameValueMap()
nvm.Add("DeferUpdates", True)
app.SilentOperation = True
Dim doc As Document = app.Documents.OpenWithOptions(path, nvm)
app.SilentOperation = False
```

### Proposed Workflow

1. **Fingerprint Analysis** (still needed for production planning)
   - Cycle through parameters to determine shared vs unique geometry
   - Write classification to manifest for bulk production planning

2. **For Each Released Element**:
   - Set master parameters to element values
   - Update assembly/drawings to reflect geometry
   - Save copies with `DeferUpdates = True`
   - Assign Vault numbers to drawings only

3. **Output Structure**:
   ```
   Elemendid/
     _manifest.json              (sharing classification for production)
     Element_A/
       Joonised/
         00001.idw               (deferred drawing, references masters)
         00002.idw
     Element_B/
       Joonised/
         00003.idw
         00004.idw
   ```

4. **What's NOT created**:
   - No standalone part copies
   - No assembly copies (or minimal, deferred)
   - Parts in `Ühine/` folder not needed

### Pros
- Much simpler code (no reference replacement logic)
- Fewer Vault numbers consumed
- Faster execution (no derivation breaking, no reference updates)
- Drawings still have unique numbers for production tracking

### Cons
- **Risk of accidental unfreeze**: If someone turns off Defer Updates, drawing updates to current master state
- **Archive complexity**: Can't archive drawings without also archiving masters
- **Master changes affect ALL releases**: If masters are modified, all deferred drawings depend on them being stable
- **Part numbers in drawings**: Show master part numbers, not released numbers (may or may not be acceptable)

## Hybrid Approach (Recommended)

Combine both strategies for optimal results:

### Parts
- **Shared parts**: Keep as masters (no copies needed)
- **Unique parts**: Still create standalone copies (geometry differs per element)
- **Manifest**: Document which parts are shared for bulk production

### Assemblies
- **Option A**: Skip entirely (production doesn't need assembly files, only drawings)
- **Option B**: Create deferred copies for visualization purposes only

### Drawings
- **Always create copies** with unique Vault numbers
- **Set DeferUpdates = True** to freeze at release state
- **Update iProperties** to show released element name (not master name)

### Code Simplification

| Step | Current | Hybrid |
|------|---------|--------|
| Fingerprinting | Full cycle | Full cycle (needed for classification) |
| Part copies | All (shared + unique) | Unique only |
| Reference replacement | Complex logic | None (drawings keep master refs) |
| Assembly copies | Yes | Optional/Skip |
| Drawing copies | Yes, with ref updates | Yes, but simpler (just copy + defer) |

### API for Simpler Drawing Release

```vb
''' <summary>
''' Create a deferred drawing copy at the current geometry state.
''' Much simpler than the full reference replacement approach.
''' </summary>
Public Function CreateDeferredDrawing(app As Inventor.Application, _
                                       sourcePath As String, _
                                       targetPath As String, _
                                       vaultNumber As String, _
                                       description As String) As Boolean
    ' Ensure source is updated to current parameters
    Dim sourceDoc As DrawingDocument = CType(app.Documents.Open(sourcePath, False), DrawingDocument)
    sourceDoc.Update()
    For Each sheet As Sheet In sourceDoc.Sheets
        sheet.Update()
    Next
    
    ' Copy file to target
    System.IO.File.Copy(sourcePath, targetPath, True)
    sourceDoc.Close(False)
    
    ' Open copy with defer enabled
    Dim nvm As NameValueMap = app.TransientObjects.CreateNameValueMap()
    nvm.Add("DeferUpdates", True)
    app.SilentOperation = True
    Dim targetDoc As DrawingDocument = CType(app.Documents.OpenWithOptions(targetPath, nvm), DrawingDocument)
    app.SilentOperation = False
    
    ' Set DeferUpdates property
    targetDoc.DrawingSettings.DeferUpdates = True
    
    ' Update iProperties
    targetDoc.PropertySets.Item("Design Tracking Properties").Item("Part Number").Value = vaultNumber
    targetDoc.PropertySets.Item("Inventor Summary Information").Item("Title").Value = description
    
    ' Save and close
    targetDoc.Save()
    targetDoc.Close()
    
    Return True
End Function
```

## Key Decision Points

### Q1: Are standalone part files required for production?

If YES (need independent files):
- Current approach is correct
- Or hybrid: standalone for unique parts only

If NO (masters are sufficient):
- Defer Updates approach works
- Just document sharing classification in manifest

### Q2: Do assembly files need to be released?

If YES (for customer delivery, BOM, etc.):
- Need either standalone copies or deferred copies

If NO (only drawings needed for production):
- Can skip assemblies entirely
- Major simplification

### Q3: What part numbers should drawings show?

If MASTER numbers are OK:
- Defer Updates approach works as-is

If RELEASED numbers needed:
- Must update title block parameters/iProperties
- Drawing views would still show master geometry (acceptable)

## Recommendation

Based on the user's statement that "they really need unique drawings with unique names/numbers, not necessarily the part files themselves":

1. **Keep fingerprint analysis** - needed for production planning
2. **Skip standalone part copies** - or only for truly unique geometry
3. **Skip assembly copies** - unless required for other purposes
4. **Create deferred drawings** - with unique Vault numbers
5. **Update manifest** - document sharing classification

This could reduce `ElementReleaseLib.vb` from ~3100 lines to perhaps ~500-800 lines.

## Open Questions

1. **Is the production workflow currently dependent on part file copies?**
   - If CAM programs reference part copies, we need them
   - If CAM programs reference masters, we don't

2. **What happens if masters are reorganized?**
   - Deferred drawings would break (can't find referenced parts)
   - Need stable master file paths

3. **How is sharing information used in production?**
   - Do they need a separate "batch production" report?
   - Is the manifest format adequate?

## Related Research

- `docs/plans/2026-04-26-module-release-cycle.md` - Original implementation plan
- Current implementation in `Lib/ElementReleaseLib.vb`
