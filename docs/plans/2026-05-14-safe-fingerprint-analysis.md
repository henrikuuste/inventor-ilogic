# Safe Fingerprint Analysis via Temporary Copies

## Overview

Modify the fingerprint analysis system to use temporary file copies instead of modifying source files in memory. This prevents issues with read-only files, Vault-managed files, and eliminates risk of accidental source file modification during preview.

## Current State Analysis

### How Fingerprinting Currently Works

**File**: `Lib/ElementReleaseLib.vb` - `BuildElementMatrix()` (lines 1396-1600)

1. **Step 1**: Open all parts from the assembly tree
2. **Step 2**: Snapshot master parameters (save current state)
3. **Step 3**: For each element variant:
   - Apply variant parameters to masters
   - Call `doc.Update()` to recompute geometry
   - Compute fingerprints for all parts
4. **Step 4**: Restore original master parameters
5. **Step 5**: Close documents without saving

### Key Discoveries

- **Parameters ARE modified**: `ApplyParameters(doc, variantCfg.Parameters)` changes parameter values in memory
- **Update triggers side effects**: `doc.Update()` triggers `[DimensionUpdate]` hooks that modify iProperties
- **Safeguards exist**: Parameters are restored in `Finally` block, documents closed with `Close(True)` (discard changes)
- **But risks remain**:
  - Read-only files may fail on `Update()`
  - External processes (Vault) may detect "dirty" state
  - Crash before restore leaves files modified in memory
  - External masters from other Aluselemendid folders are also modified

### Files Affected by Current Approach

During fingerprint analysis, these files are modified in memory:
- All master files (`000131.ipt`, `000130.ipt`, `000129.ipt`, etc.)
- External masters from other elements
- Derived parts (indirectly via Update propagation)

## Desired End State

1. **Source files are NEVER modified** - not even in memory
2. **Fingerprint accuracy preserved** - same results as current implementation
3. **Read-only files supported** - works regardless of file permission state
4. **Vault-safe** - no dirty file detection, no checkout prompts
5. **Crash-safe** - if script fails, no cleanup needed

### Verification Criteria

- [ ] Fingerprints match current implementation for identical inputs
- [ ] No source file modification (verify via file timestamps or Vault state)
- [ ] Works with read-only files
- [ ] Performance acceptable (< 2x current duration)
- [ ] Temporary files cleaned up after analysis

## What We're NOT Doing

- Caching fingerprints (separate optimization, different plan)
- Changing how fingerprints are computed (just where)
- Modifying the actual release execution (only the preview/analysis phase)
- Supporting incremental fingerprint updates

## Implementation Approach

**Strategy**: Copy-Analyze-Delete

1. Create temporary folder structure mirroring the assembly
2. Copy all necessary files (masters, parts, assemblies) to temp location
3. Open and analyze temp copies (parameters can be modified freely)
4. Compute fingerprints from temp files
5. Close and delete all temp files
6. Return fingerprint data without touching originals

---

## Phase 1: Create Temporary Copy Infrastructure

### Overview

Add utility functions to create and manage temporary file copies while preserving assembly reference structure.

### Changes Required

#### 1. New Utility Functions
**File**: `Lib/ElementReleaseLib.vb`

Add new functions:

```vb
''' <summary>
''' Create temporary copies of all files needed for fingerprint analysis.
''' Returns a mapping from original paths to temp paths.
''' </summary>
Public Function CreateTempCopiesForAnalysis(app As Inventor.Application, _
                                            tree As AssemblyTree, _
                                            masterPaths As List(Of String)) As Dictionary(Of String, String)
    ' Creates: %TEMP%\ElementReleaseAnalysis\{GUID}\
    ' Copies:
    '   - All masters (including external)
    '   - All parts from tree.Parts
    '   - Root assembly
    ' Returns: Map of original path -> temp path
End Function

''' <summary>
''' Update references in temp files to point to other temp files.
''' Must be called after all files are copied.
''' </summary>
Public Sub UpdateTempFileReferences(app As Inventor.Application, _
                                    pathMap As Dictionary(Of String, String))
    ' For each temp file:
    '   - Open document
    '   - Replace all references using pathMap
    '   - Save
End Sub

''' <summary>
''' Clean up temporary analysis folder.
''' </summary>
Public Sub CleanupTempAnalysisFolder(tempRoot As String)
    ' Close any open temp documents
    ' Delete temp folder recursively
End Sub
```

#### 2. Reference Replacement Helper
**File**: `Lib/ElementReleaseLib.vb`

```vb
''' <summary>
''' Replace all references in a document using a path mapping.
''' </summary>
Private Sub ReplaceAllReferences(doc As Document, pathMap As Dictionary(Of String, String))
    ' Handle both:
    '   - ReferencedDocumentDescriptors (for parts)
    '   - ComponentOccurrences (for assemblies)
End Sub
```

### Success Criteria

#### Verification:
- [ ] `CreateTempCopiesForAnalysis` creates correct folder structure
- [ ] All masters (internal + external) are copied
- [ ] All parts from tree are copied
- [ ] `UpdateTempFileReferences` correctly rewires references
- [ ] `CleanupTempAnalysisFolder` removes all temp files

#### Manual Verification:
- [ ] Temp folder created in expected location
- [ ] Temp assembly opens and shows all parts correctly
- [ ] No broken references in temp files

**Implementation Note**: After completing this phase and all verification passes, pause here for manual confirmation before proceeding to the next phase.

---

## Phase 2: Modify BuildElementMatrix to Use Temp Copies

### Overview

Refactor `BuildElementMatrix` to perform all analysis on temporary copies instead of source files.

### Changes Required

#### 1. Refactor BuildElementMatrix
**File**: `Lib/ElementReleaseLib.vb`

Modify `BuildElementMatrix()`:

**Before** (current flow):
```
1. Open source parts
2. Snapshot parameters
3. Loop variants:
   - Apply parameters to SOURCE masters
   - Update SOURCE assembly
   - Compute fingerprints from SOURCE parts
4. Restore parameters
5. Close parts
```

**After** (new flow):
```
1. Create temp copies (Phase 1 functions)
2. Update temp file references
3. Open temp parts and assembly
4. Loop variants:
   - Apply parameters to TEMP masters
   - Update TEMP assembly
   - Compute fingerprints from TEMP parts
5. Close temp documents
6. Delete temp folder
7. Return fingerprints
```

#### 2. Key Implementation Details

**Opening temp files**:
```vb
' Open temp assembly as active document
Dim tempAsmPath As String = pathMap(tree.RootAssemblyPath)
Dim tempAsmDoc As AssemblyDocument = CType(app.Documents.Open(tempAsmPath, True), AssemblyDocument)

' Parts will be opened automatically by assembly
' Or can be opened explicitly from pathMap
```

**Applying parameters**:
```vb
' Apply to temp masters (no snapshot needed - we'll delete these)
For Each masterPath In masterPaths
    Dim tempMasterPath As String = pathMap(masterPath)
    Dim tempDoc As Document = app.Documents.Open(tempMasterPath, False)
    ApplyParameters(tempDoc, variantCfg.Parameters)
    tempDoc.Update()
Next
```

**Computing fingerprints**:
```vb
' Compute from temp parts
For Each partPath In matrix.PartPaths
    Dim tempPartPath As String = pathMap(partPath)
    ' Find open document or open it
    Dim tempPartDoc As PartDocument = FindOrOpen(app, tempPartPath)
    fp = ComputeGeometryFingerprint(tempPartDoc)
Next
```

**Cleanup (in Finally block)**:
```vb
Finally
    ' Close all temp documents
    For Each kvp In pathMap
        CloseDocumentByPath(app, kvp.Value)
    Next
    
    ' Delete temp folder
    CleanupTempAnalysisFolder(tempRoot)
End Try
```

### Success Criteria

#### Verification:
- [ ] Source files have unchanged timestamps after analysis
- [ ] Fingerprints match previous implementation
- [ ] No temp files remain after completion
- [ ] No temp files remain after error/cancellation

#### Manual Verification:
- [ ] Run with read-only source files - no errors
- [ ] Verify source assembly unchanged (check in Inventor)
- [ ] Performance within acceptable range

**Implementation Note**: After completing this phase and all verification passes, pause here for manual confirmation before proceeding to the next phase.

---

## Phase 3: Handle Edge Cases and Optimization

### Overview

Address edge cases and optimize for performance.

### Changes Required

#### 1. Handle File Copy Failures
- Check disk space before copying
- Handle locked files gracefully
- Report which files failed to copy

#### 2. Optimize Copy Performance
- Only copy files that will be analyzed (skip drawings)
- Use file system copy (not Inventor SaveAs) where possible
- Consider parallel copying for large assemblies

#### 3. Handle Assembly Constraints
- Constraints may fail after reference replacement
- Suppress constraint errors during temp analysis
- Log warnings but don't fail

#### 4. External File Dependencies
- Some parts may reference files outside the assembly tree
- Library parts, standard parts, etc.
- These should NOT be copied - just leave references as-is

### Success Criteria

#### Verification:
- [ ] Large assemblies (50+ parts) complete in reasonable time
- [ ] Failed file copies are reported clearly
- [ ] External references (library parts) don't cause errors

#### Manual Verification:
- [ ] Test with assembly containing library parts
- [ ] Test with very large assembly

---

## Testing Strategy

### Unit Tests

- `CreateTempCopiesForAnalysis` with various assembly structures
- `UpdateTempFileReferences` with different reference types
- `CleanupTempAnalysisFolder` removes all files

### Integration Tests

- Full fingerprint analysis on test assembly
- Compare fingerprints with previous implementation
- Verify source files unchanged

### Manual Testing Steps

1. Run `Loo elemendid.vb` on test assembly
2. Verify fingerprints computed correctly
3. Check source file timestamps - should be unchanged
4. Check temp folder deleted
5. Test with read-only files (set read-only attribute)
6. Test cancellation mid-analysis - verify cleanup

## Performance Considerations

**Current approach**: ~30 seconds for 16 parts, 2 variants
**Expected with temp copies**: ~45-60 seconds (file copy overhead)

Acceptable trade-off for safety.

## Terminology Checklist

Verify all code uses correct domain terms per UBIQUITOUS_LANGUAGE.md:
- [ ] "Aluselement" not "Alusmoodul" for parametric designs
- [ ] "Väljastatud element" not "Moodul" for released units
- [ ] "Detail" not "Component" for parts

## References

- Current implementation: `Lib/ElementReleaseLib.vb:BuildElementMatrix()` (lines 1396-1600)
- Related: `docs/plans/2026-05-14-multi-master-external-references.md`
- Domain terminology: `docs/UBIQUITOUS_LANGUAGE.md`
