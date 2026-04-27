<!-- Copyright (c) 2026 Henri Kuuste -->
---
date: 2026-04-26T07:57:00+03:00
researcher: Claude
git_commit: 455e30425b6ad57a2e7fe980feebe7fa0d5faa38
branch: main
repository: Inventor-Rules
topic: "Moodulid Smart Variant Release - API Features and Existing Code"
tags: [research, codebase, moodulid, variant-release, fingerprinting, derived-parts]
status: complete
last_updated: 2026-04-26T08:15:00+03:00
---

# Research: Moodulid Smart Variant Release - API Features and Existing Code

**Date**: 2026-04-26T08:15:00+03:00 (updated with web research)
**Git Commit**: 455e30425b6ad57a2e7fe980feebe7fa0d5faa38
**Branch**: main

## Research Question

What Inventor API features and existing codebase patterns are needed to implement the Moodulid step-by-step plan for smart variant release with geometry fingerprinting, assembly tree discovery, and minimal file creation?

## Summary

The codebase already contains most of the foundational patterns needed for the Moodulid implementation. Key existing code includes:

- **Fingerprinting**: `MakeComponentsLib.ComputeBodySignature` already implements body fingerprinting using Volume/SurfaceArea/FaceCount
- **Assembly tree discovery**: `VariantReleaseLib.GetAllReferencedFiles` uses `AllReferencedDocuments`
- **Derived part detection**: `MakeComponentsLib.DeriveBodyAsNewPart` and test files show `DerivedPartComponents` usage
- **Drawing discovery**: `VariantReleaseLib.FindAllDrawings` scans folders and checks `ReferencedDocuments`
- **Reference replacement**: `VariantReleaseLib` has both API (`PutLogicalFileName`) and binary patching approaches
- **Excel reading**: `ExcelReaderLib.ReleaseConfig` already exists for variant tables
- **Parameter manipulation**: `CenterPatternLib` and `SupportPlacementLib` show get/set patterns
- **Transactions**: `Taasta värvid.vb` demonstrates `TransactionManager.StartTransaction/End/Abort`

New code needed primarily for: variant matrix building, shared/unique part classification, assembly snapshot creation, and manifest tracking.

---

## Detailed Findings

### 1. Part Fingerprinting (Step 1)

#### API Syntax

```vb
' SurfaceBody methods - NOTE: Use method with tolerance, not bare property
Dim tol As Double = 0.001   ' linear tolerance in database units (cm)
Dim vol As Double = body.Volume(tol)
Dim area As Double = body.SurfaceArea(tol)  ' NOT body.Area
Dim box As Box = body.RangeBox
Dim dx As Double = box.MaxPoint.X - box.MinPoint.X
Dim dy As Double = box.MaxPoint.Y - box.MinPoint.Y
Dim dz As Double = box.MaxPoint.Z - box.MinPoint.Z

' Iterate all bodies in a part
Dim compDef As PartComponentDefinition = partDoc.ComponentDefinition
For Each body As SurfaceBody In compDef.SurfaceBodies
    If body.IsSolid Then
        ' Process solid bodies only
    End If
Next
```

#### Existing Implementation

**`Lib/MakeComponentsLib.vb:62-72`** already has `ComputeBodySignature`:

```vb
Public Function ComputeBodySignature(body As SurfaceBody) As String
    Try
        Dim volume As Double = 0
        Dim area As Double = 0
        Dim faceCount As Integer = body.Faces.Count
        
        Try : volume = body.Volume(0.001) : Catch : End Try
        Try : area = body.SurfaceArea(0.001) : Catch : End Try
        
        Return String.Format("V:{0:F4};F:{1};A:{2:F4}", volume * 1000000, faceCount, area * 10000)
    Catch
        Return ""
    End Try
End Function
```

#### What's Missing for Step 1

- The current signature doesn't include bounding box dimensions (needed for orientation-independent fingerprinting)
- Need a new `ComputePartFingerprint(partDoc)` that aggregates all bodies
- Consider sorting bounding box dimensions for orientation independence

#### Sheet Metal Handling

For sheet metal parts, use `SheetMetalComponentDefinition.FlatPattern.RangeBox` when `HasFlatPattern` is true:

```vb
' Lib/Mõõdud.vb:663-681
Dim smCompDef As SheetMetalComponentDefinition = CType(partDoc.ComponentDefinition, SheetMetalComponentDefinition)
If smCompDef.HasFlatPattern Then
    Dim fpBox As Box = smCompDef.FlatPattern.RangeBox
    ' Use flat pattern dimensions for sheet metal
End If
```

---

### 2. Assembly Tree Discovery (Step 2)

#### API Syntax

```vb
' Get ALL referenced documents (transitive closure - full tree)
For Each refDoc As Document In asmDoc.AllReferencedDocuments
    ' refDoc.FullFileName gives full path
Next

' Get DIRECT references only (one hop)
For Each refDoc As Document In asmDoc.ReferencedDocuments
    ' Direct children only
Next

' Iterate occurrences
For Each occ As ComponentOccurrence In asmDoc.ComponentDefinition.Occurrences
    If occ.DefinitionDocumentType = DocumentTypeEnum.kPartDocumentObject Then
        Dim partDoc As PartDocument = CType(occ.Definition.Document, PartDocument)
    ElseIf occ.DefinitionDocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
        ' Recurse into sub-assembly
        Dim subAsmDef As AssemblyComponentDefinition = CType(occ.Definition, AssemblyComponentDefinition)
        ' Use subAsmDef.Occurrences for recursive walk
    End If
Next
```

#### Existing Implementation

**`Lib/VariantReleaseLib.vb:36-47`** - `GetAllReferencedFiles`:

```vb
Public Function GetAllReferencedFiles(asmDoc As AssemblyDocument) As List(Of String)
    Dim files As New List(Of String)
    files.Add(asmDoc.FullFileName)
    For Each refDoc As Document In asmDoc.AllReferencedDocuments
        If Not files.Contains(refDoc.FullFileName) Then
            files.Add(refDoc.FullFileName)
        End If
    Next
    Return files
End Function
```

**`Lib/VariantReleaseLib.vb:420-431`** - Recursive occurrence collection:

```vb
Private Sub CollectAllOccurrences(occs As ComponentOccurrences, ByRef result As List(Of ComponentOccurrence))
    For Each occ As ComponentOccurrence In occs
        result.Add(occ)
        If occ.DefinitionDocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
            Try
                If occ.SubOccurrences IsNot Nothing Then
                    CollectAllOccurrences(occ.SubOccurrences, result)
                End If
            Catch : End Try
        End If
    Next
End Sub
```

#### Detecting Derived Parts

**`DerivedPartComponents` collection**:

```vb
Dim dpcs As DerivedPartComponents = partDoc.ComponentDefinition.ReferenceComponents.DerivedPartComponents
If dpcs.Count > 0 Then
    ' This part has at least one derived-part reference
End If

' Get source document path
If partDoc.ReferencedDocuments.Count > 0 Then
    Dim masterPath As String = partDoc.ReferencedDocuments.Item(1).FullFileName
End If
```

**Determining which body a derived part came from** (from `Lib/MakeComponentsLib.vb:893-947`):

```vb
Dim dpDef As DerivedPartUniformScaleDef = dpcs.CreateUniformScaleDef(masterDoc.FullDocumentName)
For Each dpe As DerivedPartEntity In dpDef.Solids
    Dim refEntity As Object = dpe.ReferencedEntity
    If TypeOf refEntity Is SurfaceBody Then
        Dim bodyName As String = CType(refEntity, SurfaceBody).Name
        ' bodyName identifies which master body this entity references
    End If
Next
```

#### What's Missing for Step 2

- New `AssemblyTree`, `PartInfo`, `AsmInfo` data structures
- `ClassifyPart()` function to determine Master/Derived/Manual role
- `FindMasterParts()` to identify all masters in the tree
- Integration with drawing discovery

---

### 3. Variant Matrix (Step 3)

#### Parameter Manipulation API

**Get/Set parameters** (from `Lib/CenterPatternLib.vb:208-283`):

```vb
Public Function SetParameter(asmDoc As AssemblyDocument, paramName As String, value As Double, Optional units As String = "mm") As Parameter
    Dim params As Parameters = asmDoc.ComponentDefinition.Parameters
    Dim expression As String = value.ToString(System.Globalization.CultureInfo.InvariantCulture) & " " & units
    
    Try
        Dim param As Parameter = params.Item(paramName)
        param.Expression = expression
        Return param
    Catch
        Try
            Return params.UserParameters.AddByExpression(paramName, expression, UnitsTypeEnum.kDefaultDisplayLengthUnits)
        Catch
            Return Nothing
        End Try
    End Try
End Function

Public Function GetParameterValue(asmDoc As AssemblyDocument, paramName As String) As Double
    Try
        Return asmDoc.ComponentDefinition.Parameters.Item(paramName).Value
    Catch
        Return 0
    End Try
End Function
```

**Apply parameters from dictionary** (from `Lib/VariantReleaseLib.vb:555-577`):

```vb
Public Function ApplyParameters(doc As Document, parameters As Dictionary(Of String, String)) As Boolean
    Dim docParams As Parameters = doc.ComponentDefinition.Parameters
    For Each kvp As KeyValuePair(Of String, String) In parameters
        If kvp.Key.StartsWith("_") Then Continue For
        Try
            Dim param As Parameter = docParams.Item(kvp.Key)
            param.Expression = kvp.Value
        Catch
            success = False
        End Try
    Next
    Return success
End Function
```

**Force rebuild**: `doc.Update()` after parameter changes

#### Excel Reading

**`Lib/ExcelReaderLib.vb`** already exists with `ReleaseConfig`:

```vb
Public Class ReleaseConfig
    Public ConfigName As String
    Public PartNumber As String
    Public Parameters As Dictionary(Of String, String)
    
    Public Function GetParameter(key As String) As String
    Public Function GetParameterAsDouble(key As String) As Double
End Class

Public Function ReadVariantTable(excelPath As String, Optional sheetName As String = "") As List(Of ReleaseConfig)
    ' Excel COM: CreateObject("Excel.Application"), read-only workbook
    ' Row 1 headers, Col 1 name, Col 2 part number, Cols 3+ into Parameters
End Function
```

#### What's Missing for Step 3

- `SetParametersOnAllMasters()` - set params on multiple masters at once
- `FingerprintAllParts()` - fingerprint entire tree in current state
- `BuildVariantMatrix()` - orchestrate cycling through variants
- `IdentifyPartGroups()` - classify shared vs unique parts
- `VariantMatrix` data structure

---

### 4. Vault-Safe Parameter Cycling (Step 4)

#### Transaction API

**From `Taasta värvid.vb:26-52`**:

```vb
Dim trans As Transaction = app.TransactionManager.StartTransaction(doc, "Analysis")

Try
    ' ... change params, read fingerprints ...
    
    trans.End()  ' Commit as single undo step
Catch ex As Exception
    trans.Abort()  ' Roll back all changes
End Try
```

#### Document Dirty Check

```vb
If asmDoc.Dirty Then
    asmDoc.Save()
End If
```

#### Vault Checkout Detection

From `AGENTS.md`:
- Use `doc.ReservedForWriteByMe` (reliable)
- NOT `doc.IsModifiable` (unreliable for Vault)

#### What's Missing for Step 4

- `SnapshotAllParameters()` - capture expressions for all masters
- `RestoreAllParameters()` - restore from snapshot
- `AnalyzeVariantsVaultSafe()` - wrapper with full save/restore

---

### 5. Standalone Part Creation (Step 5)

#### SaveAs API

```vb
' Save as new primary file
doc.SaveAs(filePath, False)

' Save copy (document stays associated with original)
doc.SaveAs(copyPath, True)
```

#### Breaking Derivation Links

Two approaches exist:

1. **`DerivedPartComponent.BreakLinkToFile()`** - Keeps geometry, removes associative link (true "standalone")
2. **`DerivedPartComponent.Delete()`** - Removes feature entirely (geometry may disappear)

**Current codebase approach** (doesn't use `BreakLinkToFile`):
- Copy file with `File.Copy`
- Update references using `ReferencedFileDescriptors.PutLogicalFileName` or binary patching
- For refreshing derivation: delete + re-derive

**From `Lib/MakeComponentsLib.vb`** - delete existing derives:

```vb
' Delete all DerivedPartComponents
Dim dpcs As DerivedPartComponents = partDoc.ComponentDefinition.ReferenceComponents.DerivedPartComponents
Dim toDelete As New List(Of DerivedPartComponent)
For Each dpc As DerivedPartComponent In dpcs
    toDelete.Add(dpc)
Next
For Each dpc In toDelete
    dpc.Delete()
Next
```

#### What's Missing for Step 5

- `CreateStandaloneCopy()` using `BreakLinkToFile()` approach
- `VerifyStandalone()` to confirm no references remain
- Testing on sheet metal parts

---

### 6. Assembly Snapshot (Step 6)

#### ComponentOccurrence.Replace API

```vb
' Replace single occurrence
selectedOcc.Replace("C:\path\to\newpart.ipt", False)

' Replace all occurrences with same definition
selectedOcc.Replace("C:\path\to\newpart.ipt", True)
```

**From `Lib/VariantReleaseLib.vb:378-382`**:

```vb
If copyMap.ContainsKey(currentPath) Then
    Dim newPath As String = copyMap(currentPath)
    occ.Replace(newPath, False)
End If
```

#### What's Missing for Step 6

- `CreateAssemblySnapshot()` - copy assembly + replace all references per map
- `BuildReferenceMap()` - compute which files map to shared vs variant-specific
- Handle sub-assemblies bottom-up

---

### 7. Drawing Handling (Step 7)

#### Finding Drawings

**From `Lib/VariantReleaseLib.vb:56-127`** - `FindAllDrawings`:

```vb
Public Function FindAllDrawings(app As Inventor.Application, searchFolder As String, _
                                referencedFiles As List(Of String), _
                                Optional ByRef logMessages As List(Of String) = Nothing) As List(Of String)
    ' Scan *.idw and *.dwg, skip \OldVersions\
    For Each f In Directory.GetFiles(searchFolder, "*.idw", SearchOption.AllDirectories)
        If f.IndexOf("\OldVersions\", StringComparison.OrdinalIgnoreCase) < 0 Then
            drawingFiles.Add(f)
        End If
    Next
    ' Open each drawing, check ReferencedDocuments against referencedFiles set
    For Each refDoc As Document In drawDoc.ReferencedDocuments
        Dim refPath As String = Path.GetFullPath(refDoc.FullFileName)
        If refFileSet.Contains(refPath) Then
            referencesOurFiles = True
        End If
    Next
End Function
```

#### Updating Drawing References

**API approach** (from `Lib/VariantReleaseLib.vb:522-535`):

```vb
Dim refDescriptors As ReferencedFileDescriptors = drawDoc.ReferencedFileDescriptors
For i As Integer = 1 To refDescriptors.Count
    Dim rfd As ReferencedFileDescriptor = refDescriptors.Item(i)
    Dim oldPath As String = rfd.FullFileName
    If copyMap.ContainsKey(oldPath) Then
        rfd.PutLogicalFileName(copyMap(oldPath))
    End If
Next
drawDoc.Save()
```

**Binary approach** (used for iLogic compatibility):
- `Lib/BinaryReferenceUpdateLib.vb` - `UpdateFileReferencesBinary`
- Length-matched path substitution in closed files
- Requires variant folder names same length as originals

**Note from codebase**: `PutLogicalFileName` may not work in iLogic context; binary patching is the fallback.

#### What's Missing for Step 7

- `ClassifyDrawings()` - determine shared vs variant-specific drawings
- `CreateDrawingCopy()` with proper relinking
- Handle multiple drawings per part/assembly

---

### 8. Release Planning and Execution (Step 8)

#### Existing Release Pattern

**`Lib/VariantReleaseLib.vb`** has the full pipeline:

1. `GetAllReferencedFiles()` - collect dependency tree
2. `FindAllDrawings()` - discover drawings
3. `BuildCopyMap()` / `CreateTargetFolders()` / `CopyFiles()` - file operations
4. `UpdateAllReferences()` - open assemblies, `occ.Replace()` + `PutLogicalFileName`
5. `UpdateDrawingReferences()` or binary patching
6. `WriteLogFile()` - persist release log

#### What's Missing for Step 8

- `PlanRelease()` - smart planning based on fingerprint analysis
- `ExecuteRelease()` - orchestrate with proper ordering
- Shared vs unique part handling
- Parameter cycling during release (set params, create standalone, reset)

---

### 9. Manifest Tracking (Step 9)

#### JSON Handling

The codebase uses simple string-based property storage (iProperties). For JSON manifests, would need:

```vb
' Simple approach - serialize to string manually
' Or use System.Web.Script.Serialization.JavaScriptSerializer if available
```

#### What's Missing for Step 9

- `WriteManifest()` - JSON manifest creation
- `ReadManifest()` - JSON parsing
- `ComputeDelta()` - fingerprint comparison
- Manifest data structure

---

## Code References

### Key Library Files

| File | Purpose |
|------|---------|
| `Lib/MakeComponentsLib.vb` | Body fingerprinting, derived parts, body metadata |
| `Lib/VariantReleaseLib.vb` | Assembly tree, drawing discovery, file copy, reference updates |
| `Lib/ExcelReaderLib.vb` | Excel reading, `ReleaseConfig` structure |
| `Lib/BinaryReferenceUpdateLib.vb` | Binary path patching for drawings/derived refs |
| `Lib/CenterPatternLib.vb` | Parameter get/set patterns |
| `Lib/SupportPlacementLib.vb` | Parameter manipulation, part update |
| `Lib/VaultNumberingLib.vb` | Vault connection, folder creation, numbering |
| `Lib/UtilsLib.vb` | Logging, paths, geometry helpers |

### Test Files for Reference

| File | Demonstrates |
|------|-------------|
| `Katsetused/TestDerivedPart.vb` | `DerivedPartUniformScaleDef`, body selection |
| `Katsetused/TestDerivedPartRefs.vb` | `ReferencedFileDescriptors` diagnostics |
| `Katsetused/TestUpdateDerivedRef.vb` | Updating derived part references |
| `Katsetused/TestReplaceReference.vb` | `ComponentOccurrence.Replace` |
| `Katsetused/TestSaveCopyAs.vb` | `SaveAs(path, True)` behavior |
| `Taasta värvid.vb` | Transaction start/end/abort |

### Main Related Scripts

- `Loo komponendid.vb` - Creates derived parts from multibody masters (reference for derived part patterns)

---

## Architecture Recommendations

### Reusable from Existing Code

1. **`MakeComponentsLib.ComputeBodySignature`** - Extend for full part fingerprint
2. **`VariantReleaseLib.GetAllReferencedFiles`** - Use for tree discovery
3. **`VariantReleaseLib.FindAllDrawings`** - Use for drawing discovery
4. **`ExcelReaderLib.ReleaseConfig`** - Use directly for variant tables
5. **`VariantReleaseLib.CollectAllOccurrences`** - Recursive occurrence walk
6. **Binary reference update pattern** - For iLogic-compatible drawing updates

### New Structures Needed

1. **`AssemblyTree`** - Comprehensive tree representation
2. **`PartInfo`** with role classification (Master/Derived/Manual)
3. **`VariantMatrix`** - Part × Variant → Fingerprint
4. **`PartGroup`** - Shared/unique classification with fingerprint groupings
5. **`ReleasePlan`** - What to create and in what order
6. **`ReleaseManifest`** - JSON tracking for re-release

### API Gotchas to Remember

| Issue | Solution |
|-------|----------|
| `SurfaceBody.Area` doesn't exist | ~~Use `SurfaceBody.SurfaceArea(tolerance)`~~ **TESTED: Also doesn't exist!** Use `sum of Face.Evaluator.Area` |
| `SurfaceBody.SurfaceArea` doesn't exist | Iterate `body.Faces` and sum `face.Evaluator.Area` for per-body surface area |
| `PutLogicalFileName` unreliable in iLogic | Use binary patching fallback |
| Transaction nesting | Keep simple; avoid nested transactions |
| Parameter formula separators | Use `;` not `,` in formulas |
| Vault checkout | Check `doc.ReservedForWriteByMe` |
| Sheet metal flat pattern | Call `ExitEdit()` before save |

---

## Open Questions

1. **BreakLinkToFile reliability**: Not used in current codebase - needs testing for standalone part creation
2. **Transaction scope with multiple masters**: Does abort roll back changes to ALL documents touched?
3. **Binary patching length constraints**: Current implementation requires same-length paths - how to handle?
4. **Large assembly performance**: How does `AllReferencedDocuments` perform on 500+ file assemblies?
5. **Vault file locking during analysis**: Can we safely change parameters on read-only files?

---

## Web Research: Inventor 2026 API Documentation

### Official API Documentation Findings

#### SurfaceBody Volume and Area

**Official API** (from Autodesk Manufacturing DevBlog and API Help):

```vb
' SurfaceBody methods take a tolerance parameter
Dim vol As Double = body.Volume(0.001)       ' tolerance in cm
Dim area As Double = body.SurfaceArea(0.001) ' NOT body.Area - that doesn't exist

' Alternative: Sum face areas via SurfaceEvaluator
Private Function GetBodyArea(body As SurfaceBody) As Double
    Dim area As Double = 0
    For Each oFace As Face In body.Faces
        Dim oEval As SurfaceEvaluator = oFace.Evaluator
        area = area + oEval.Area
    Next
    Return area
End Function
```

**For overlapping bodies** (from DevBlog):
```vb
' Use TransientBRep to get combined volume of overlapping bodies
Dim oTransientBRep As TransientBRep = ThisApplication.TransientBRep
Dim oBody1 As SurfaceBody = oTransientBRep.Copy(oRealBody1)
Dim oBody2 As SurfaceBody = oTransientBRep.Copy(oRealBody2)
Call oTransientBRep.DoBoolean(oBody1, oBody2, BooleanTypeEnum.kBooleanTypeUnion)
Dim combinedVolume As Double = oBody1.Volume(0.001)
```

#### MassProperties (Whole-Part Properties)

**From official API samples** - for whole-part volume/area without per-body iteration:

```vb
' Get MassProperties object from part or assembly
Dim oMassProps As MassProperties = partDoc.ComponentDefinition.MassProperties

' Set accuracy level
oMassProps.Accuracy = kMediumAccuracy

' Read properties (whole document, all bodies combined)
Debug.Print "Area: " & oMassProps.Area
Debug.Print "Volume: " & oMassProps.Volume
Debug.Print "Mass: " & oMassProps.Mass
Debug.Print "Center of Mass: " & oMassProps.CenterOfMass.X & ", " & oMassProps.CenterOfMass.Y & ", " & oMassProps.CenterOfMass.Z

' CRITICAL: To avoid dirtying document during analysis:
oMassProps.CacheResultsOnCompute = False
```

#### DerivedPartComponent.BreakLinkToFile

**From Autodesk Forums and Help** - key findings:

1. **API exists**: `DerivedPartComponent.BreakLinkToFile()` is available
2. **CRITICAL**: Link must be in RESOLVED state to break it - cannot break if reference is missing/unresolved
3. **Permanent**: Once broken, cannot restore the link
4. **Alternative for corrupted files**: Use `Copy Object → Create New → Repair Geometry` workflow

```vb
' Break link on all derived components in assembly parts
For Each oDocument As Document In oAsmDoc.AllReferencedDocuments
    If oDocument.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
        For Each oDerivedComp As DerivedPartComponent In oDocument.ComponentDefinition.ReferenceComponents.DerivedPartComponents
            oDerivedComp.BreakLinkToFile()
        Next
    End If
Next
```

**Alternative approach** (from forums - for some derive types):
```vb
' Setting Adaptive = False can break link on NonParametricBaseFeature
Dim bf As NonParametricBaseFeature = doc.ComponentDefinition.Features(1)
bf.Adaptive = False
```

#### ComponentOccurrence.Replace

**Official documentation** (Inventor API Help):

```vb
' Replace signature
ComponentOccurrence.Replace(FileName As String, ReplaceAll As Boolean)

' Replace single occurrence
occ.Replace("C:\path\to\newpart.ipt", False)

' Replace ALL occurrences with same definition
occ.Replace("C:\path\to\newpart.ipt", True)

' Extended version with more options
occ.Replace2(FileName, ReplaceAll, SaveEdits, KeepAdaptivity)
```

**Key behaviors**:
- If replacement geometry is similar, constraints are maintained
- If geometry differs significantly, constraints may drop
- May fail if `DocumentDescriptor.OwnershipType = kExclusiveOwnership`

#### FileDescriptor.ReplaceReference (Post-Inventor 10)

**From Manufacturing DevBlog** - this is the NEWER recommended API:

```vb
' Access via File object, not ReferencedFileDescriptors directly
Dim oFD As FileDescriptor = doc.File.ReferencedFileDescriptors(1)
oFD.ReplaceReference("C:\NewPath\Part2.ipt")

' CRITICAL REQUIREMENT: Replacement file must have SAME HERITAGE
' Files must share same InternalName (copied via SaveCopyAs or File.Copy)
' Cannot replace with a brand new file!
```

**Works in both**:
- Inventor API (inside Inventor)
- Apprentice Server API

**Known issue**: After ReplaceReference on drawings, views may not update visually until file is re-opened.

#### TransactionManager

**Official API** (Inventor API Help):

```vb
' Start transaction
Dim trans As Transaction = app.TransactionManager.StartTransaction(doc, "My Operation")

' Commit (creates single undo point)
trans.End()

' Rollback (discard all changes)
trans.Abort()

' Additional methods on TransactionManager:
app.TransactionManager.UndoTransaction()      ' Undo current committed transaction
app.TransactionManager.RedoTransaction()      ' Redo undone transaction
app.TransactionManager.ClearAllTransactions() ' Clear undo/redo stack
app.TransactionManager.SetCheckPoint()        ' Bookmark within transaction
app.TransactionManager.GoToCheckPoint()       ' Abort back to checkpoint
```

**Best practice pattern** (from iLogic community):

```vb
Dim trans As Transaction = app.TransactionManager.StartTransaction(doc, "Analysis")
Try
    ' ... perform operations ...
    trans.End()
Catch ex As Exception
    trans.Abort()
    ' Handle error
End Try
```

---

### Inventor 2026 New API Features

**From official Autodesk Developer Blog** - relevant new capabilities:

| Feature | Description | Relevance to Moodulid |
|---------|-------------|----------------------|
| **Model States Edit Scope** | Write access to `ModelStatesInEdit` property, `MemberEditScope` control | Could help with variant management |
| **iProperties per Model State** | Access/modify iProperties per model state via `PropertySets` | Variant-specific metadata |
| **Simplify Feature API** | New `SimplifyFeatures` under `ComponentDefinition.Features` | Could create simplified geometry for fingerprinting |
| **Apprentice Server Standalone** | No longer auto-registered; requires separate installation | May affect batch processing setup |
| **IFC4x3 Export** | New `IFCExportOptions` for BIM workflows | Not directly relevant |
| **Sketch-Based Break Operations** | `BreakOperations.AddBySketch` for drawings | Not directly relevant |

**Apprentice Server 2026 Changes** (important for automation):
- No longer COM-registered by default with Inventor installation
- Requires separate standalone installation
- Must manually run `ApprenticeRegSrv.exe /install` for COM-based tools
- Update paths to reference standalone installation, not Inventor's `Bin` folder

---

### API Gotchas Summary (Updated with Web Research)

| Issue | Solution | Source |
|-------|----------|--------|
| `SurfaceBody.Area` doesn't exist | ~~Use `SurfaceBody.SurfaceArea(tolerance)`~~ **WRONG** - Use `sum of Face.Evaluator.Area` | **TESTED 2026-04-26** |
| `SurfaceBody.SurfaceArea` doesn't exist | Iterate faces: `For Each face In body.Faces : area += face.Evaluator.Area : Next` | **TESTED 2026-04-26** |
| `BreakLinkToFile` fails | Link must be RESOLVED first; cannot break missing refs | Forums |
| `ReplaceReference` requires same heritage | Files must have same InternalName (use File.Copy or SaveCopyAs) | DevBlog |
| Drawing views don't update after ReplaceReference | May need to close/reopen file | Forums |
| MassProperties dirties document | Set `CacheResultsOnCompute = False` before reading | API Help |
| Apprentice Server 2026 not registered | Install standalone and run `ApprenticeRegSrv.exe /install` | DevBlog |
| `Replace2` not always available | Use basic `Replace` if `Replace2` fails | Forums |

---

### Useful External Links

- [Inventor 2026 API Help](https://help.autodesk.com/view/INVNTOR/2026/ENU/)
- [Manufacturing DevBlog](https://adndevblog.typepad.com/manufacturing/)
- [What's New in Inventor 2026 API](https://blog.autodesk.io/whats-new-in-the-autodesk-inventor-2026-api-feature-highlights-and-enhancements/)
- [ComponentOccurrence.Replace](https://help.autodesk.com/cloudhelp/2021/ENU/Inventor-API/files/ComponentOccurrence_Replace.htm)
- [TransactionManager](https://help.autodesk.com/cloudhelp/2022/ENU/Inventor-API/files/TransactionManager.htm)
- [MassProperties Sample](https://help.autodesk.com/cloudhelp/2025/ENU/Inventor-API/files/MassProperties_Sample.htm)

---

## Related Research

- `docs/research/2026-04-25-loo-komponendid-failures.md` - Derived part creation failure analysis
- `docs/research/2026-04-25-dimensions-thickness-width-length.md` - Dimension property patterns

