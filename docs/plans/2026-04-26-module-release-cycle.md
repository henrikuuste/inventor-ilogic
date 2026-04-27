<!-- Copyright (c) 2026 Henri Kuuste -->
# Module Release Cycle Implementation Plan

## Overview

Create a smart module release system (`Moodulid/Loo moodulid.vb`) that releases parametric Inventor modules with optimal file sharing. The system analyzes variant parameters, computes geometry fingerprints, and creates standalone copies only where geometry differs. Shared parts are consolidated in a common folder (`Ühine`), reducing Vault file numbers and simplifying manufacturing.

**Key Innovation**: Uses Vault disconnect-save-reconnect workflow to bypass the uncontrollable "New File" dialog, allowing programmatic control over file locations.

**Development Mode**: Phases 1-5 are completely Vault-independent. Only Phase 6 (Vault Integration) requires Vault connection. This allows full testing and development without Vault.

## Current State Analysis

### Existing Infrastructure
- **`Lib/VariantReleaseLib.vb`**: Assembly tree discovery, file copying, reference updates
- **`Lib/MakeComponentsLib.vb`**: Body fingerprinting (`ComputeBodySignature`), derived part handling
- **`Lib/ExcelReaderLib.vb`**: Variant table reading (`ReleaseConfig`)
- **`Lib/VaultNumberingLib.vb`**: Vault connection, numbering schemes, file upload APIs
- **`Lib/BinaryReferenceUpdateLib.vb`**: Binary reference patching (fallback)

### Confirmed API Capabilities (from Tests)
| Feature | API | Status |
|---------|-----|--------|
| Part fingerprinting | `body.Volume(0.001)`, `face.Evaluator.Area`, `body.RangeBox` | ✅ Works |
| Break derivation | `DerivedPartComponent.BreakLinkToFile()` | ✅ Works |
| Transaction rollback | `Transaction.Abort()` | ✅ Works |
| Drawing reference update | `FileDescriptor.ReplaceReference()` | ✅ Works (preferred) |
| Vault logout/login | `CommandManager.ControlDefinitions.Item("LogoutCmdIntName").Execute()` | ✅ Works |
| Vault file upload | `DocumentService.AddUploadedFile()` | ✅ Works |
| Vault file sync | `AcquireFiles` with Download option | ✅ Works |

### Key Discoveries
1. **`PutLogicalFileName()` doesn't exist in iLogic** - Use `ReplaceReference()` instead
2. **`ReplaceReference()` requires heritage match** - Files must share InternalName (achieved via `File.Copy`)
3. **Vault "New File" dialog location cannot be controlled** - Must use disconnect-save-add workflow
4. **Programmatic login/logout WORKS** - Full automation possible

### Design Decisions (from review)

1. **Master Discovery**: Masters are identified by being the SOURCE of derivations (via `DerivedFromMaster` paths from derived parts), not by folder location. Masters can be multibody or sketch-only parts.

2. **Parameter Location**: All variant parameters MUST be in master parts. If a user needs a parameter elsewhere, they must define it in a master and derive/link from there.

3. **Masters in Assemblies**: Masters may appear directly in assemblies (for reference geometry, constraints). On release:
   - First attempt: Remove/suppress master occurrences
   - If removal fails (constraints depend on it): Create frozen copy and replace reference
   - Frozen copies go to Ühine, only for assembly references (derived parts break links entirely)

4. **Full Module Discovery**: Find all `Alusmoodulid/{ModuleName}/*.iam` (direct children). No automatic hierarchy filtering - all IAMs at that level are root assemblies.

5. **Module Naming**: 
   - Base module name: From folder (`Alusmoodulid/Iste/`) → `"Iste"`
   - Release module name: From Excel ConfigName → `"Iste 70"`, `"Iste 110"`
   - Release folder: `Moodulid/{ConfigName}/` → `Moodulid/Iste 70/`
   - Description properties: Replace base name with release name

6. **Ühine Structure**: Flat by submodule only (e.g., `Ühine/Karkass/Detailid/`), not organized by base module. Manifest tracks provenance.

7. **Sharing Classification**: A part goes to Ühine if:
   - Same SOURCE part (identified by Part Number iProperty) AND
   - Same geometry fingerprint AND  
   - Used by 2+ variants/modules
   
   Parts unique to ONE variant stay in that variant's folder. Two geometrically identical parts from DIFFERENT source parts are NOT shared (respects user intent).

8. **Fingerprint Formula**: `PartNumber + GeometryHash` - includes source part number, not path. If file is moved/renamed, it's still the same source part. Similar geometry from different source parts = different fingerprints.

9. **Re-release with Changed Geometry**: If a shared part's geometry changes, ALL modules/variants using that part must be re-released together. System blocks partial release and informs user: "Cannot release Iste 70 alone. The following must also be released: Iste 90, Selg 100. Proceed with full re-release?"

10. **Sharing Scope**: Part sharing is per-product only. The manifest and Ühine folder are product-scoped. Cross-product sharing is not supported.

11. **Assembly Handling**: Sub-assemblies are ALWAYS per-variant (no fingerprinting, no sharing). Even if all parts inside are shared, assemblies have variant-specific parameters, constraints, BOM metadata, and potentially suppressed components.

12. **Library/Standard Parts**: Parts outside the `Alusmoodulid/{ModuleName}/` tree are considered external/library parts. They are excluded from release processing - not copied, not modified, assemblies keep original references to them.

13. **Error Recovery**: Resume capability on failure. Save progress to `_release_progress.json` with list of completed files. On failure, show error and offer "Resume" on next run. Resume skips already-created files, continues from failure point. On successful completion, delete progress file.

14. **Drawing References**: Drawings reference parts and assemblies, never masters. On release, only replace references in the reference map (module parts); library/standard parts keep original paths unchanged.

15. **iProperty Updates on Release**:
    - **Part Number**: Set to the new Vault file number
    - **Description**: Replace base module name with release module name (if present)
    - **Project**: No change (same product)
    - **Other properties**: No explicit changes - existing "Uuenda" iLogic rules auto-update custom properties on save/parameter change
    - **Uuenda rule safe**: Confirmed - it only updates Thickness/Width/Length dimension properties, not Part Number or Description

16. **Suppressed Components**: Keep suppression state as-is in released assemblies. Still create released part files for suppressed components (simpler, and other variants may use them unsuppressed).

## Desired End State

### Folder Structure
```
$/Product/Moodulid/                     <-- release output root
  Ühine/                                <-- shared parts (same geometry across ALL variants)
    {Submodule}/                        <-- e.g., Karkass, Poroloon
      Detailid/
        00001.ipt                       <-- standalone shared part
      Joonised/
        00001.idw                       <-- shared drawing (if all refs are shared)
  
  {VariantName}/                        <-- variant folder (from Excel ConfigName)
    {Submodule}/
      Detailid/
        00002.ipt                       <-- variant-specific part
      Koostud/
        00003.iam                       <-- frozen assembly snapshot
      Joonised/
        00002.idw                       <-- variant-specific drawing
        00003.idw                       <-- assembly drawing
```

### Cross-Module Sharing

Parts may be shared not only across variants of the same module, but also across **different modules** of the same product. This is common in base modules (Alusmoodulid) where multiple modules share standard components (brackets, connectors, hardware).

**Sharing levels:**
1. **Variant-level sharing**: Same part used by multiple variants of ONE module → goes to `Ühine/`
2. **Module-level sharing**: Same part used by multiple MODULES of the product → also goes to `Ühine/`

**Detection approach:**
- When releasing a module, check if the part already exists in `Ühine/` (from a previous module release)
- Compare fingerprints: if identical, reuse the existing file (no new Vault number needed)
- If fingerprints differ, this indicates a design issue that should be flagged

**Manifest tracking:**
- The manifest (`Moodulid/_manifest.json`) tracks which modules reference each shared part
- When re-releasing, the system can identify orphaned shared parts (no longer referenced by any module)

```
$/Product/Moodulid/
  Ühine/
    Karkass/
      Detailid/
        00001.ipt      <-- used by Module A variant 1, Module A variant 2, Module B variant 1
    Poroloon/
      Detailid/
        00002.ipt      <-- used by all variants of all modules
  
  ModuleA-1800/
    ...               <-- references parts from Ühine/
  ModuleA-2100/
    ...
  ModuleB-Standard/
    ...               <-- also references same parts from Ühine/
```

### Success Criteria
1. All released parts are **standalone** (no derivation links)
2. Shared parts appear **once** in Ühine, referenced by all variant assemblies
3. Drawings update correctly to reference released parts
4. All files saved to correct Vault locations without manual folder selection
5. File numbers assigned via "Softcom numbriskeem"
6. Assemblies open and update without errors

## What We're NOT Doing

1. **Not copying xlsx, pdf files** - These are regenerated outside this system
2. **Not versioning within Moodulid** - Vault handles versioning; we overwrite
3. **Not preserving derivation links** - Released parts are frozen snapshots
4. **Not handling master parts** - Masters stay in Alusmoodulid, never copied
5. **Not creating lifecycle states** - Files added in default state
6. **Not handling standard library parts** - Referenced directly from original location

## Implementation Approach

### Overall Strategy
1. **Analysis phase** (read-only): Discover files, cycle parameters, compute fingerprints
2. **Planning phase** (read-only): Reserve Vault numbers, compute paths
3. **Execution phase** (write): Disconnect from Vault, save files, reconnect, upload

### File Flow
```
Alusmoodulid/XX/              -->  Moodulid/Ühine/XX/       (shared parts)
                              -->  Moodulid/{Variant}/XX/   (unique parts)
```

### Vertical Slices (Implementation Order)

Implementation follows vertical slices rather than horizontal phases. Each slice delivers working functionality that can be tested independently.

#### Slice 1: Single Part Release
**Goal**: Release one derived part as a standalone copy.
- Input: One `.ipt` file (derived from a master)
- Output: Standalone copy with broken derivation link
- Test: Open copy, verify no reference to master, geometry intact

**Key APIs**: `File.Copy()`, `DerivedPartComponent.BreakLinkToFile()`

#### Slice 2: Part + Drawing
**Goal**: Release a part with its drawing.
- Input: One `.ipt` + one `.idw`
- Output: Standalone part + drawing referencing the new part
- Test: Open drawing, verify it references the released part

**Key APIs**: Previous + `FileDescriptor.ReplaceReference()`

#### Slice 3: Assembly with Parts
**Goal**: Release an assembly with all its parts.
- Input: One `.iam` with multiple `.ipt` files
- Output: Standalone assembly + parts, all cross-referencing correctly
- Test: Open assembly, verify all references resolved

**Key APIs**: Previous + `ComponentOccurrence.Replace()`

#### Slice 4: Multi-Variant Analysis
**Goal**: Cycle through variants and compute fingerprints.
- Input: Excel variant table + assembly tree
- Output: Variant matrix with fingerprints per part per variant
- Test: Fingerprints change when parameters change, restore to original

**Key APIs**: Previous + parameter snapshot/restore, `ComputeGeometryFingerprint()`

#### Slice 5: Shared Part Detection
**Goal**: Detect shared parts within a module release.
- Input: Variant matrix from Slice 4
- Output: Parts classified as shared (→ Ühine) or unique (→ variant folder)
- Test: Same geometry across variants → one file in Ühine referenced by all

**Key APIs**: Previous + `PartGroup` classification, folder structure

#### Slice 6: Cross-Module Sharing
**Goal**: Detect and reuse shared parts across module releases.
- Input: Previous module manifest + new module release
- Output: Reuses existing shared parts, updates manifest
- Test: Release Module A, then Module B → no duplicate files for matching fingerprints

**Key APIs**: Previous + manifest read/write, `FindExistingSharedPart()`

#### Slice 7: Vault Integration
**Goal**: Full production workflow with Vault.
- Input: Complete release ready for upload
- Output: Files uploaded to Vault with correct numbers and locations
- Test: Files appear in Vault Explorer, can be checked out

**Key APIs**: Previous + `ReserveVaultNumbers()`, logout/login, `AddUploadedFile()`

---

Each slice builds on the previous. Start with Slice 1 and verify it works before proceeding. This allows early validation of core assumptions and surfaces integration issues early.

---

## Phase 1: UI and Discovery

### Overview
Create the user-facing dialog with mode selection and discover the Excel variant table.

### Changes Required:

#### 1. Main Script Shell
**File**: `Moodulid/Loo moodulid.vb`
**Changes**: Create initial orchestrator with UI

```vb
' Entry point - called from iLogic
Sub Main()
    Dim app As Inventor.Application = ThisApplication
    Dim logs As New List(Of String)
    
    ' Show mode selection dialog (non-blocking, not centered)
    Dim mode As ReleaseMode = ShowModeSelectionDialog(app)
    If mode = ReleaseMode.Cancelled Then Return
    
    ' Discover based on mode
    Dim context As ReleaseContext = DiscoverContext(app, mode, logs)
    If context Is Nothing Then Return
    
    ' Show analysis results and confirm
    If Not ShowAnalysisConfirmation(context, logs) Then Return
    
    ' Execute release
    ExecuteRelease(app, context, logs)
End Sub
```

#### 2. Mode Selection Dialog
**File**: `Lib/ModuleReleaseLib.vb` (new)
**Changes**: Non-blocking WinForms dialog

```vb
Public Enum ReleaseMode
    Cancelled = 0
    FullModule = 1       ' All IAMs from Excel discovery
    CurrentAssembly = 2  ' Only current open IAM tree
End Enum

Public Function ShowModeSelectionDialog(app As Inventor.Application) As ReleaseMode
    ' WinForms Form with:
    ' - "Full Module" button
    ' - "Current Assembly Only" button
    ' - Cancel button
    ' Position: NOT centered (top-right or offset from center)
End Function
```

#### 3. Excel Discovery
**File**: `Lib/ModuleReleaseLib.vb`
**Changes**: Find Excel file and parse variant table

```vb
Public Class ReleaseContext
    Public Mode As ReleaseMode
    Public ExcelPath As String                    ' Path to XX_moodulid.xlsx
    Public Variants As List(Of ExcelReaderLib.ReleaseConfig)
    Public SourceRoot As String                   ' Product/Alusmoodulid/XX
    Public TargetRoot As String                   ' Product/Moodulid
    Public AssemblyTree As AssemblyTree           ' From Phase 2
    Public VariantMatrix As VariantMatrix         ' From Phase 3
    Public ReleasePlan As ReleasePlan             ' From Phase 4
End Class

Public Function DiscoverExcel(sourceFolder As String) As String
    ' Find *_moodulid.xlsx in sourceFolder
    Dim files = Directory.GetFiles(sourceFolder, "*_moodulid.xlsx")
    If files.Length = 0 Then Return Nothing
    Return files(0)
End Function
```

### Success Criteria:

#### Automated Verification:
- [ ] Mode selection dialog displays correctly
- [ ] Excel file discovered from source folder
- [ ] Variant table parsed correctly

#### Manual Verification:
- [ ] Dialog is non-blocking (can interact with Inventor behind it)
- [ ] Dialog is not centered (positioned to allow Inventor interaction)
- [ ] Cancel works correctly

**Implementation Note**: After completing this phase and all automated verification passes, pause here for manual confirmation before proceeding to the next phase.

---

## Phase 2: File Discovery and Classification

### Overview
Walk the assembly dependency tree and classify all files by role.

### Changes Required:

#### 1. Assembly Tree Discovery
**File**: `Lib/ModuleReleaseLib.vb`
**Changes**: Discover all relevant files

```vb
Public Class AssemblyTree
    Public RootAssemblyPath As String
    Public SourceRoot As String                   ' Common ancestor folder
    Public Parts As Dictionary(Of String, PartInfo)
    Public Assemblies As Dictionary(Of String, AssemblyInfo)
    Public Drawings As List(Of DrawingInfo)
End Class

Public Class PartInfo
    Public FilePath As String
    Public RelativePath As String                 ' Relative to SourceRoot
    Public Role As PartRole                       ' Derived, Manual
    Public DerivedFromMaster As String            ' Path of master (if derived)
    Public BodyName As String                     ' Body name in master (if derived)
End Class

Public Enum PartRole
    Derived   ' Created via DeriveBodyAsNewPart from a master
    Manual    ' Standalone part, not derived
End Enum

Public Class DrawingInfo
    Public DrawingPath As String
    Public RelativePath As String
    Public ReferencedModelPaths As List(Of String)
End Class
```

#### 2. Tree Traversal
**File**: `Lib/ModuleReleaseLib.vb`

```vb
Public Function DiscoverAssemblyTree(app As Inventor.Application, _
                                      rootAsmPath As String, _
                                      ByRef logs As List(Of String)) As AssemblyTree
    Dim tree As New AssemblyTree()
    tree.RootAssemblyPath = rootAsmPath
    
    ' Open assembly (if not already open)
    Dim asmDoc As AssemblyDocument = OpenAssembly(app, rootAsmPath)
    
    ' Collect all referenced documents (transitive closure)
    For Each refDoc As Document In asmDoc.AllReferencedDocuments
        Dim ext As String = Path.GetExtension(refDoc.FullFileName).ToLower()
        
        If ext = ".ipt" Then
            Dim info As PartInfo = ClassifyPart(CType(refDoc, PartDocument))
            tree.Parts.Add(refDoc.FullFileName, info)
        ElseIf ext = ".iam" Then
            tree.Assemblies.Add(refDoc.FullFileName, New AssemblyInfo With {
                .FilePath = refDoc.FullFileName
            })
        End If
        ' Skip .idw - discovered separately
    Next
    
    ' Compute source root (common ancestor)
    tree.SourceRoot = ComputeCommonAncestor(tree)
    
    ' Compute relative paths
    For Each kvp In tree.Parts
        kvp.Value.RelativePath = GetRelativePath(tree.SourceRoot, kvp.Key)
    Next
    
    Return tree
End Function
```

#### 3. Part Classification
**File**: `Lib/ModuleReleaseLib.vb`

```vb
Public Function ClassifyPart(partDoc As PartDocument) As PartInfo
    Dim info As New PartInfo()
    info.FilePath = partDoc.FullFileName
    
    ' Check for derived part components
    Dim dpcs = partDoc.ComponentDefinition.ReferenceComponents.DerivedPartComponents
    If dpcs.Count > 0 AndAlso partDoc.ReferencedDocuments.Count > 0 Then
        info.Role = PartRole.Derived
        info.DerivedFromMaster = partDoc.ReferencedDocuments.Item(1).FullFileName
        ' Get body name from derived definition
        info.BodyName = GetDerivedBodyName(dpcs.Item(1))
    Else
        info.Role = PartRole.Manual
    End If
    
    Return info
End Function
```

#### 4. Drawing Discovery
**File**: `Lib/ModuleReleaseLib.vb`

```vb
Public Function DiscoverDrawings(app As Inventor.Application, _
                                  tree As AssemblyTree, _
                                  searchFolders As List(Of String), _
                                  ByRef logs As List(Of String)) As List(Of DrawingInfo)
    Dim drawings As New List(Of DrawingInfo)
    Dim treeFiles As New HashSet(Of String)(tree.Parts.Keys)
    treeFiles.UnionWith(tree.Assemblies.Keys)
    
    ' Scan for .idw files
    For Each folder In searchFolders
        For Each idwPath In Directory.GetFiles(folder, "*.idw", SearchOption.AllDirectories)
            If idwPath.Contains("\OldVersions\") Then Continue For
            
            ' Open and check references
            Dim drawDoc As DrawingDocument = CType(app.Documents.Open(idwPath, False), DrawingDocument)
            Dim refs As New List(Of String)
            
            For Each refDoc As Document In drawDoc.ReferencedDocuments
                If treeFiles.Contains(refDoc.FullFileName) Then
                    refs.Add(refDoc.FullFileName)
                End If
            Next
            
            If refs.Count > 0 Then
                drawings.Add(New DrawingInfo With {
                    .DrawingPath = idwPath,
                    .RelativePath = GetRelativePath(tree.SourceRoot, idwPath),
                    .ReferencedModelPaths = refs
                })
            End If
            
            drawDoc.Close(True)  ' Close without saving
        Next
    Next
    
    Return drawings
End Function
```

### Success Criteria:

#### Automated Verification:
- [ ] All parts in assembly tree discovered
- [ ] Derived parts correctly identified with master reference
- [ ] Manual parts correctly identified
- [ ] All drawings referencing tree files discovered

#### Manual Verification:
- [ ] Part count matches Inventor's `AllReferencedDocuments`
- [ ] Drawing discovery finds CAM, detail, and assembly drawings
- [ ] Source root correctly computed

---

## Phase 3: Variant Analysis

### Overview
Cycle through variants, compute fingerprints, and classify parts as shared vs unique.

### Changes Required:

#### 1. Fingerprint Computation
**File**: `Lib/ModuleReleaseLib.vb`

```vb
Public Function ComputeGeometryFingerprint(partDoc As PartDocument) As String
    ' Geometry-only fingerprint for intra-module variant comparison
    ' Source path is tracked separately for cross-module sharing
    
    Dim bodies = partDoc.ComponentDefinition.SurfaceBodies
    If bodies.Count = 0 Then Return "NO_BODIES"
    
    Dim fps As New List(Of String)
    For Each body As SurfaceBody In bodies
        If body.IsSolid Then
            fps.Add(ComputeBodyFingerprint(body))
        End If
    Next
    
    fps.Sort()  ' Deterministic ordering
    Return String.Join("|", fps)
End Function

' For cross-module sharing: source PART NUMBER + geometry must BOTH match
' Two geometrically identical parts from different source parts are NOT shared
' This respects user intent - separate parts = not shared
' Using Part Number (not path) so rename/move doesn't break sharing
Public Function ComputeFullFingerprint(partDoc As PartDocument) As String
    Dim geometryFp = ComputeGeometryFingerprint(partDoc)
    Dim partNumber = partDoc.PropertySets("Design Tracking Properties")("Part Number").Value.ToString()
    Return $"PN:{partNumber}|GEO:{geometryFp}"
End Function

Public Function GetPartNumber(partDoc As PartDocument) As String
    Return partDoc.PropertySets("Design Tracking Properties")("Part Number").Value.ToString()
End Function

Public Function ComputeBodyFingerprint(body As SurfaceBody) As String
    ' Volume
    Dim vol As Double = Math.Round(body.Volume(0.001), 4)
    
    ' Surface area (sum of face areas)
    Dim area As Double = 0
    For Each face As Face In body.Faces
        area += face.Evaluator.Area
    Next
    area = Math.Round(area, 4)
    
    ' Bounding box (sorted for orientation independence)
    Dim bb As Box = body.RangeBox
    Dim dims() As Double = {
        Math.Round(bb.MaxPoint.X - bb.MinPoint.X, 3),
        Math.Round(bb.MaxPoint.Y - bb.MinPoint.Y, 3),
        Math.Round(bb.MaxPoint.Z - bb.MinPoint.Z, 3)
    }
    Array.Sort(dims)
    
    Return $"V:{vol}|A:{area}|BB:{dims(0)}x{dims(1)}x{dims(2)}"
End Function
```

#### 2. Parameter Snapshot and Restore
**File**: `Lib/ModuleReleaseLib.vb`

```vb
Public Function SnapshotMasterParameters(app As Inventor.Application, _
                                          masterPaths As List(Of String)) As Dictionary(Of String, Dictionary(Of String, String))
    Dim snapshot As New Dictionary(Of String, Dictionary(Of String, String))
    
    For Each masterPath In masterPaths
        Dim doc As Document = app.Documents.ItemByName(masterPath)
        If doc Is Nothing Then Continue For
        
        Dim paramSnapshot As New Dictionary(Of String, String)
        Dim params As Parameters = doc.ComponentDefinition.Parameters
        
        For Each param As Parameter In params.ModelParameters
            paramSnapshot.Add(param.Name, param.Expression)
        Next
        
        snapshot.Add(masterPath, paramSnapshot)
    Next
    
    Return snapshot
End Function

Public Sub RestoreMasterParameters(app As Inventor.Application, _
                                    snapshot As Dictionary(Of String, Dictionary(Of String, String)))
    For Each kvp In snapshot
        Dim doc As Document = app.Documents.ItemByName(kvp.Key)
        If doc Is Nothing Then Continue For
        
        Dim params As Parameters = doc.ComponentDefinition.Parameters
        For Each paramKvp In kvp.Value
            Try
                params.Item(paramKvp.Key).Expression = paramKvp.Value
            Catch
            End Try
        Next
    Next
End Sub
```

#### 3. Variant Matrix Building
**File**: `Lib/ModuleReleaseLib.vb`

```vb
Public Class VariantMatrix
    Public PartPaths As List(Of String)
    Public VariantNames As List(Of String)
    Public Fingerprints As Dictionary(Of String, Dictionary(Of String, String))
    ' Fingerprints(variantName)(partPath) = fingerprintHash
End Class

Public Function BuildVariantMatrix(app As Inventor.Application, _
                                    tree As AssemblyTree, _
                                    variants As List(Of ExcelReaderLib.ReleaseConfig), _
                                    masterPaths As List(Of String), _
                                    ByRef logs As List(Of String)) As VariantMatrix
    Dim matrix As New VariantMatrix()
    matrix.PartPaths = tree.Parts.Keys.ToList()
    matrix.VariantNames = variants.Select(Function(v) v.ConfigName).ToList()
    matrix.Fingerprints = New Dictionary(Of String, Dictionary(Of String, String))
    
    ' Snapshot original parameters
    Dim snapshot = SnapshotMasterParameters(app, masterPaths)
    
    Try
        For Each variant In variants
            logs.Add($"Analyzing variant: {variant.ConfigName}")
            
            ' Set variant parameters on all masters
            For Each masterPath In masterPaths
                Dim doc As Document = app.Documents.ItemByName(masterPath)
                ApplyParameters(doc, variant.Parameters)
            Next
            
            ' Update all documents
            app.ActiveDocument.Update()
            
            ' Fingerprint all parts
            Dim variantFps As New Dictionary(Of String, String)
            For Each partPath In matrix.PartPaths
                Dim partDoc As PartDocument = CType(app.Documents.ItemByName(partPath), PartDocument)
                variantFps.Add(partPath, ComputePartFingerprint(partDoc))
            Next
            
            matrix.Fingerprints.Add(variant.ConfigName, variantFps)
        Next
    Finally
        ' ALWAYS restore original parameters
        RestoreMasterParameters(app, snapshot)
        app.ActiveDocument.Update()
    End Try
    
    Return matrix
End Function
```

#### 4. Part Group Classification
**File**: `Lib/ModuleReleaseLib.vb`

```vb
Public Class PartGroup
    Public PartPath As String
    Public RelativePath As String
    Public IsShared As Boolean
    Public UniqueFingerprints As Dictionary(Of String, List(Of String))
    ' fingerprint -> list of variant names that produce it
End Class

Public Function ClassifyPartGroups(matrix As VariantMatrix, _
                                    tree As AssemblyTree) As List(Of PartGroup)
    Dim groups As New List(Of PartGroup)
    
    For Each partPath In matrix.PartPaths
        Dim group As New PartGroup()
        group.PartPath = partPath
        group.RelativePath = tree.Parts(partPath).RelativePath
        group.UniqueFingerprints = New Dictionary(Of String, List(Of String))
        
        ' Group variants by geometry fingerprint
        ' (source path is already fixed - we're iterating one source part)
        For Each variantName In matrix.VariantNames
            Dim fp = matrix.Fingerprints(variantName)(partPath)
            If Not group.UniqueFingerprints.ContainsKey(fp) Then
                group.UniqueFingerprints.Add(fp, New List(Of String))
            End If
            group.UniqueFingerprints(fp).Add(variantName)
        Next
        
        ' Sharing classification:
        ' - Fingerprint shared by 2+ variants → goes to Ühine
        ' - Fingerprint unique to 1 variant → goes to that variant's folder
        ' A single source part can produce MULTIPLE released files (one per unique geometry)
        
        groups.Add(group)
    Next
    
    Return groups
End Function
```

### Success Criteria:

#### Automated Verification:
- [ ] Fingerprints are deterministic (same result on repeated calls)
- [ ] Parameters restore correctly after analysis
- [ ] Document dirty state restored to original
- [ ] Part groups correctly classified as shared/unique

#### Manual Verification:
- [ ] Changing master parameter changes derived part fingerprints
- [ ] Transaction approach doesn't cause Vault issues
- [ ] Analysis completes in reasonable time (<30 seconds per variant)

---

## Phase 4: Release Planning

### Overview
Reserve file numbers and compute target file paths.

**Vault-independent**: In development mode, generates local sequential numbers without Vault connection. In production mode, reserves numbers from Vault using the configured numbering scheme.

### Changes Required:

#### 1. Vault Number Reservation
**File**: `Lib/ModuleReleaseLib.vb`

```vb
Public Class ReleasePlan
    Public Files As List(Of PlannedFile)
    Public SharedFolder As String                 ' Product/Moodulid/Ühine
    Public VariantFolders As Dictionary(Of String, String)  ' ConfigName -> Product/Moodulid/{ConfigName}
End Class

Public Class PlannedFile
    Public SourcePath As String
    Public TargetVaultPath As String              ' $/Product/Moodulid/...
    Public TargetLocalPath As String              ' C:\_SoftcomVault/Product/Moodulid/...
    Public VaultNumber As String                  ' Reserved number or existing file number
    Public FileType As FileType                   ' Part, Assembly, Drawing
    Public IsShared As Boolean
    Public IsExisting As Boolean                  ' True if reusing existing file from Ühine (cross-module)
    Public ForVariants As List(Of String)         ' Which variants use this file
    Public ForModules As List(Of String)          ' Which modules use this file (for manifest)
End Class

Public Enum FileType
    Part
    Assembly
    Drawing
End Enum

' Configuration - change for production
Public Const NUMBERING_SCHEME As String = "Test numbriskeem"  ' Use "Softcom numbriskeem" for production
Public Const DEVELOPMENT_MODE As Boolean = True               ' Set False for production (enables Vault)

Public Function ReserveVaultNumbers(conn As Object, _
                                     count As Integer, _
                                     ByRef logs As List(Of String)) As List(Of String)
    logs.Add($"Reserving {count} Vault numbers (scheme: {NUMBERING_SCHEME})...")
    
    Dim scheme = VaultNumberingLib.FindSchemeByName(conn, NUMBERING_SCHEME)
    If scheme Is Nothing Then
        logs.Add($"ERROR: Numbering scheme '{NUMBERING_SCHEME}' not found!")
        Return Nothing
    End If
    
    Dim numbers As New List(Of String)
    For i As Integer = 1 To count
        Dim num = VaultNumberingLib.GenerateFileNumber(conn, scheme)
        numbers.Add(num)
    Next
    
    logs.Add($"Reserved numbers: {String.Join(", ", numbers)}")
    Return numbers
End Function

' Development mode: generate local sequential numbers without Vault
Public Function GenerateLocalNumbers(count As Integer, _
                                      outputRoot As String, _
                                      ByRef logs As List(Of String)) As List(Of String)
    logs.Add($"Generating {count} local numbers (development mode)...")
    
    ' Find highest existing number in output folder
    Dim startNum As Integer = 1
    Dim manifestPath = outputRoot & "\_manifest.json"
    If File.Exists(manifestPath) Then
        Dim manifest = ReadManifest(manifestPath)
        If manifest IsNot Nothing AndAlso manifest.SharedParts.Count > 0 Then
            startNum = manifest.SharedParts.Max(Function(p) Integer.Parse(p.VaultNumber)) + 1
        End If
    End If
    
    Dim numbers As New List(Of String)
    For i As Integer = 0 To count - 1
        numbers.Add((startNum + i).ToString("D5"))  ' 5-digit format: 00001, 00002, etc.
    Next
    
    logs.Add($"Generated numbers: {numbers.First()} to {numbers.Last()}")
    Return numbers
End Function
```

#### 2. Path Computation
**File**: `Lib/ModuleReleaseLib.vb`

```vb
Public Function GetFileNumbers(targetRoot As String, _
                                count As Integer, _
                                ByRef logs As List(Of String)) As List(Of String)
    If DEVELOPMENT_MODE Then
        Return GenerateLocalNumbers(count, targetRoot, logs)
    Else
        Dim conn = VaultNumberingLib.GetVaultConnection()
        If conn Is Nothing Then
            logs.Add("ERROR: No Vault connection available!")
            Return Nothing
        End If
        Return ReserveVaultNumbers(conn, count, logs)
    End If
End Function

Public Function ComputeReleasePlan(tree As AssemblyTree, _
                                    partGroups As List(Of PartGroup), _
                                    drawings As List(Of DrawingInfo), _
                                    variants As List(Of ExcelReaderLib.ReleaseConfig), _
                                    targetRoot As String, _
                                    fileNumbers As List(Of String), _
                                    ByRef logs As List(Of String)) As ReleasePlan
    Dim plan As New ReleasePlan()
    plan.SharedFolder = targetRoot & "/Ühine"
    plan.VariantFolders = New Dictionary(Of String, String)
    plan.Files = New List(Of PlannedFile)
    
    Dim numberIndex As Integer = 0
    
    ' Compute variant folders
    For Each variant In variants
        plan.VariantFolders.Add(variant.ConfigName, targetRoot & "/" & variant.ConfigName)
    Next
    
    ' Plan part files
    For Each group In partGroups
        If group.IsShared Then
            ' Check if this part already exists in Ühine (from previous module release)
            Dim existingPath = FindExistingSharedPart(plan.SharedFolder, group, logs)
            
            If existingPath IsNot Nothing Then
                ' Reuse existing shared part (no new Vault number needed)
                plan.Files.Add(New PlannedFile With {
                    .SourcePath = group.PartPath,
                    .TargetVaultPath = existingPath,
                    .TargetLocalPath = VaultNumberingLib.ConvertVaultPathToLocal(existingPath),
                    .VaultNumber = Path.GetFileNameWithoutExtension(existingPath),
                    .FileType = FileType.Part,
                    .IsShared = True,
                    .IsExisting = True,  ' Flag: don't create, just reference
                    .ForVariants = variants.Select(Function(v) v.ConfigName).ToList()
                })
                logs.Add($"  Reusing existing shared: {existingPath}")
            Else
                ' New shared file in Ühine
                plan.Files.Add(New PlannedFile With {
                    .SourcePath = group.PartPath,
                    .TargetVaultPath = plan.SharedFolder & "/" & group.RelativePath,
                    .VaultNumber = vaultNumbers(numberIndex),
                    .FileType = FileType.Part,
                    .IsShared = True,
                    .IsExisting = False,
                    .ForVariants = variants.Select(Function(v) v.ConfigName).ToList()
                })
                numberIndex += 1
            End If
        Else
            ' One file per unique fingerprint
            For Each fpKvp In group.UniqueFingerprints
                Dim firstVariant = fpKvp.Value(0)  ' Use first variant for folder
                plan.Files.Add(New PlannedFile With {
                    .SourcePath = group.PartPath,
                    .TargetVaultPath = plan.VariantFolders(firstVariant) & "/" & group.RelativePath,
                    .VaultNumber = vaultNumbers(numberIndex),
                    .FileType = FileType.Part,
                    .IsShared = False,
                    .ForVariants = fpKvp.Value
                })
                numberIndex += 1
            Next
        End If
    Next
    
    ' Plan assembly files (one per variant - assemblies always per-variant)
    For Each variant In variants
        For Each asmKvp In tree.Assemblies
            Dim relativePath = GetRelativePath(tree.SourceRoot, asmKvp.Key)
            plan.Files.Add(New PlannedFile With {
                .SourcePath = asmKvp.Key,
                .TargetVaultPath = plan.VariantFolders(variant.ConfigName) & "/" & relativePath,
                .VaultNumber = vaultNumbers(numberIndex),
                .FileType = FileType.Assembly,
                .IsShared = False,
                .ForVariants = New List(Of String) From {variant.ConfigName}
            })
            numberIndex += 1
        Next
    Next
    
    ' Plan drawing files
    For Each drawing In drawings
        Dim refersOnlyToShared = drawing.ReferencedModelPaths.All(
            Function(p) partGroups.FirstOrDefault(Function(g) g.PartPath = p)?.IsShared = True)
        
        If refersOnlyToShared Then
            ' Shared drawing
            plan.Files.Add(New PlannedFile With {
                .SourcePath = drawing.DrawingPath,
                .TargetVaultPath = plan.SharedFolder & "/" & drawing.RelativePath,
                .VaultNumber = vaultNumbers(numberIndex),
                .FileType = FileType.Drawing,
                .IsShared = True,
                .ForVariants = variants.Select(Function(v) v.ConfigName).ToList()
            })
            numberIndex += 1
        Else
            ' Per-variant drawing
            For Each variant In variants
                plan.Files.Add(New PlannedFile With {
                    .SourcePath = drawing.DrawingPath,
                    .TargetVaultPath = plan.VariantFolders(variant.ConfigName) & "/" & drawing.RelativePath,
                    .VaultNumber = vaultNumbers(numberIndex),
                    .FileType = FileType.Drawing,
                    .IsShared = False,
                    .ForVariants = New List(Of String) From {variant.ConfigName}
                })
                numberIndex += 1
            Next
        End If
    Next
    
    ' Compute local paths (for non-existing files)
    For Each file In plan.Files
        If Not file.IsExisting Then
            file.TargetLocalPath = VaultNumberingLib.ConvertVaultPathToLocal(file.TargetVaultPath)
        End If
    Next
    
    logs.Add($"Release plan: {plan.Files.Count} files total")
    logs.Add($"  - Reusing existing: {plan.Files.Count(Function(f) f.IsExisting)}")
    logs.Add($"  - Shared: {plan.Files.Count(Function(f) f.IsShared)}")
    logs.Add($"  - Variant-specific: {plan.Files.Count(Function(f) Not f.IsShared)}")
    
    Return plan
End Function
```

#### 3. Cross-Module Sharing Detection
**File**: `Lib/ModuleReleaseLib.vb`

```vb
Public Function FindExistingSharedPart(app As Inventor.Application, _
                                        sharedFolder As String, _
                                        group As PartGroup, _
                                        ByRef logs As List(Of String)) As String
    ' Look for existing file in Ühine with matching PART NUMBER + GEOMETRY
    ' Cross-module sharing requires BOTH:
    '   1. Same source part number (stable across renames/moves)
    '   2. Same geometry fingerprint
    ' This respects user intent - two similar parts from different sources are NOT shared
    
    ' Read manifest if exists
    Dim manifestPath = Path.GetDirectoryName(sharedFolder) & "\_manifest.json"
    Dim manifest = ReadManifest(manifestPath)
    If manifest Is Nothing Then Return Nothing
    
    ' Get target source part number and geometry fingerprint
    Dim partDoc As PartDocument = CType(app.Documents.ItemByName(group.PartPath), PartDocument)
    Dim targetPartNumber = GetPartNumber(partDoc)
    Dim targetGeometryFp = group.UniqueFingerprints.Keys.First()  ' For shared parts, only one geometry
    
    For Each entry In manifest.SharedParts
        ' BOTH source part number AND geometry must match
        If entry.SourcePartNumber = targetPartNumber AndAlso entry.GeometryFingerprint = targetGeometryFp Then
            ' Verify file still exists (local or Vault depending on mode)
            If FileExistsForMode(entry) Then
                logs.Add($"  Found cross-module match: {entry.VaultPath}")
                logs.Add($"    Source Part#: {targetPartNumber}")
                logs.Add($"    Geometry: {targetGeometryFp.Substring(0, Math.Min(30, targetGeometryFp.Length))}...")
                Return entry.VaultPath
            End If
        End If
    Next
    
    Return Nothing  ' No existing match found
End Function

Public Function FileExistsForMode(entry As SharedPartEntry) As Boolean
    If DEVELOPMENT_MODE Then
        ' Check local file exists
        Dim localPath = ConvertVaultPathToLocal(entry.VaultPath)
        Return File.Exists(localPath)
    Else
        ' Check Vault file exists
        Try
            Dim conn = VaultNumberingLib.GetVaultConnection()
            Dim file = conn.WebServiceManager.DocumentService.GetLatestFileByMasterPath(entry.VaultPath)
            Return file IsNot Nothing
        Catch
            Return False
        End Try
    End If
End Function

Public Function ConvertVaultPathToLocal(vaultPath As String) As String
    ' Convert $/Product/... to C:\_SoftcomVault/Product/...
    ' Uses project workspace mapping
    Return vaultPath.Replace("$/", "C:\_SoftcomVault\").Replace("/", "\")
End Function
```

#### 4. Confirmation Dialog
**File**: `Lib/ModuleReleaseLib.vb`

```vb
Public Function ShowPlanConfirmationDialog(plan As ReleasePlan, _
                                            logs As List(Of String)) As Boolean
    ' WinForms dialog showing:
    ' - Total files to create
    ' - Shared vs variant breakdown
    ' - Vault numbers that will be used
    ' - "Proceed" / "Cancel" buttons
    ' 
    ' Position: NOT centered
    
    Dim message As String = $"Release Plan Summary:
    
Total files to create: {plan.Files.Count}
Shared parts: {plan.Files.Count(Function(f) f.IsShared AndAlso f.FileType = FileType.Part)}
Variant parts: {plan.Files.Count(Function(f) Not f.IsShared AndAlso f.FileType = FileType.Part)}
Assemblies: {plan.Files.Count(Function(f) f.FileType = FileType.Assembly)}
Drawings: {plan.Files.Count(Function(f) f.FileType = FileType.Drawing)}

Vault numbers reserved: {plan.Files.First().VaultNumber} to {plan.Files.Last().VaultNumber}

Proceed with release?"
    
    Return MessageBox.Show(message, "Confirm Release", MessageBoxButtons.YesNo) = DialogResult.Yes
End Function
```

### Success Criteria:

#### Automated Verification:
- [ ] Vault numbers reserved successfully
- [ ] All planned files have valid target paths
- [ ] Shared files planned once, unique files planned per fingerprint
- [ ] Drawing classification (shared vs per-variant) is correct

#### Manual Verification:
- [ ] Plan summary is accurate
- [ ] Vault folder paths follow expected structure
- [ ] Number count matches expected file count

---

## Phase 5: File Release Execution

### Overview
Create all files locally. In production mode, disconnects from Vault first to avoid dialogs.

**Vault-independent**: In development mode, simply saves files to the local output folder without any Vault interaction. In production mode, logs out of Vault before saving to bypass the "New File" dialog.

### Changes Required:

#### 1. Vault Logout (Production Mode Only)
**File**: `Lib/ModuleReleaseLib.vb`

```vb
Public Sub DisconnectFromVaultIfNeeded(app As Inventor.Application, ByRef logs As List(Of String))
    If DEVELOPMENT_MODE Then
        logs.Add("Development mode: Skipping Vault logout")
        Return
    End If
    
    ' Check if connected to Vault
    Try
        Dim conn = VaultNumberingLib.GetVaultConnection()
        If conn Is Nothing Then
            logs.Add("Not connected to Vault, skipping logout")
            Return
        End If
    Catch
        logs.Add("Vault connection check failed, skipping logout")
        Return
    End Try
    
    ' Show informational dialog (non-blocking, non-centered)
    ' "The system will now disconnect from Vault to save files.
    '  This will take a few minutes. Please wait..."
    '
    ' NO cancel button - numbers are already reserved
    
    logs.Add("Disconnecting from Vault...")
    app.CommandManager.ControlDefinitions.Item("LogoutCmdIntName").Execute()
    logs.Add("Vault disconnected")
End Sub
```

#### 2. Standalone Part Creation
**File**: `Lib/ModuleReleaseLib.vb`

```vb
Public Function CreateStandalonePart(app As Inventor.Application, _
                                      sourcePartPath As String, _
                                      targetPath As String, _
                                      ByRef logs As List(Of String)) As Boolean
    ' Ensure target folder exists
    Directory.CreateDirectory(Path.GetDirectoryName(targetPath))
    
    ' Copy file (preserves InternalName - required for ReplaceReference)
    System.IO.File.Copy(sourcePartPath, targetPath, True)
    
    ' Open copy
    Dim partDoc As PartDocument = CType(app.Documents.Open(targetPath, True), PartDocument)
    
    ' Break derivation links
    Dim dpcs = partDoc.ComponentDefinition.ReferenceComponents.DerivedPartComponents
    For Each dpc As DerivedPartComponent In dpcs
        Try
            dpc.BreakLinkToFile()
        Catch ex As Exception
            logs.Add($"WARNING: Could not break link: {ex.Message}")
        End Try
    Next
    
    ' Save and close
    partDoc.Save()
    partDoc.Close()
    
    logs.Add($"Created standalone: {targetPath}")
    Return True
End Function
```

#### 3. Assembly Snapshot Creation
**File**: `Lib/ModuleReleaseLib.vb`

```vb
Public Function CreateAssemblySnapshot(app As Inventor.Application, _
                                         sourceAsmPath As String, _
                                         targetPath As String, _
                                         referenceMap As Dictionary(Of String, String), _
                                         variantParams As Dictionary(Of String, String), _
                                         ByRef logs As List(Of String)) As Boolean
    ' Ensure target folder exists
    Directory.CreateDirectory(Path.GetDirectoryName(targetPath))
    
    ' Copy file
    System.IO.File.Copy(sourceAsmPath, targetPath, True)
    
    ' Open copy
    Dim asmDoc As AssemblyDocument = CType(app.Documents.Open(targetPath, True), AssemblyDocument)
    
    ' Replace component references
    For Each occ As ComponentOccurrence In GetAllOccurrences(asmDoc)
        Dim currentPath As String = occ.Definition.Document.FullFileName
        If referenceMap.ContainsKey(currentPath) Then
            Try
                occ.Replace(referenceMap(currentPath), True)  ' replaceAll=True
            Catch ex As Exception
                logs.Add($"WARNING: Could not replace {currentPath}: {ex.Message}")
            End Try
        End If
    Next
    
    ' Set variant parameters
    ApplyParameters(asmDoc, variantParams)
    
    ' Update and save
    asmDoc.Update()
    asmDoc.Save()
    asmDoc.Close()
    
    logs.Add($"Created assembly snapshot: {targetPath}")
    Return True
End Function
```

#### 4. Drawing Copy with Reference Update
**File**: `Lib/ModuleReleaseLib.vb`

```vb
Public Function CreateDrawingCopy(app As Inventor.Application, _
                                   sourceDrawingPath As String, _
                                   targetPath As String, _
                                   referenceMap As Dictionary(Of String, String), _
                                   ByRef logs As List(Of String)) As Boolean
    ' Ensure target folder exists
    Directory.CreateDirectory(Path.GetDirectoryName(targetPath))
    
    ' Copy file (preserves heritage for ReplaceReference)
    System.IO.File.Copy(sourceDrawingPath, targetPath, True)
    
    ' Open copy
    Dim drawDoc As DrawingDocument = CType(app.Documents.Open(targetPath, True), DrawingDocument)
    
    ' Replace model references using FileDescriptor.ReplaceReference
    For i As Integer = 1 To drawDoc.File.ReferencedFileDescriptors.Count
        Dim fd As FileDescriptor = drawDoc.File.ReferencedFileDescriptors.Item(i)
        Dim oldPath As String = fd.FullFileName
        If referenceMap.ContainsKey(oldPath) Then
            Try
                fd.ReplaceReference(referenceMap(oldPath))
                logs.Add($"  Replaced: {Path.GetFileName(oldPath)} -> {Path.GetFileName(referenceMap(oldPath))}")
            Catch ex As Exception
                logs.Add($"WARNING: ReplaceReference failed for {oldPath}: {ex.Message}")
            End Try
        End If
    Next
    
    ' Update and save
    drawDoc.Update()
    drawDoc.Save()
    drawDoc.Close()
    
    logs.Add($"Created drawing: {targetPath}")
    Return True
End Function
```

#### 5. Release Execution Orchestrator
**File**: `Lib/ModuleReleaseLib.vb`

```vb
Public Function ExecuteRelease(app As Inventor.Application, _
                                context As ReleaseContext, _
                                ByRef logs As List(Of String)) As Boolean
    ' Build reference map for each variant
    Dim referenceMaps = BuildReferenceMaps(context)
    
    ' Snapshot master parameters
    Dim masterPaths = GetMasterPaths(context.AssemblyTree)
    Dim paramSnapshot = SnapshotMasterParameters(app, masterPaths)
    
    Try
        ' STEP 1: Disconnect from Vault if needed (production mode only)
        DisconnectFromVaultIfNeeded(app, logs)
        
        ' STEP 2: Create shared parts (no parameter changes needed)
        logs.Add("Creating shared parts...")
        For Each file In context.ReleasePlan.Files.Where(Function(f) f.IsShared AndAlso f.FileType = FileType.Part)
            CreateStandalonePart(app, file.SourcePath, file.TargetLocalPath, logs)
        Next
        
        ' STEP 3: Create variant-specific parts (grouped by fingerprint)
        logs.Add("Creating variant-specific parts...")
        Dim processedFingerprints As New HashSet(Of String)
        For Each group In context.PartGroups.Where(Function(g) Not g.IsShared)
            For Each fpKvp In group.UniqueFingerprints
                Dim fingerprint = fpKvp.Key
                If processedFingerprints.Contains(fingerprint) Then Continue For
                processedFingerprints.Add(fingerprint)
                
                ' Set parameters for first variant with this fingerprint
                Dim variantName = fpKvp.Value(0)
                Dim variant = context.Variants.First(Function(v) v.ConfigName = variantName)
                ApplyParametersToMasters(app, masterPaths, variant.Parameters)
                app.ActiveDocument.Update()
                
                ' Find the planned file for this fingerprint
                Dim plannedFile = context.ReleasePlan.Files.FirstOrDefault(
                    Function(f) f.SourcePath = group.PartPath AndAlso f.ForVariants.Contains(variantName))
                
                If plannedFile IsNot Nothing Then
                    CreateStandalonePart(app, plannedFile.SourcePath, plannedFile.TargetLocalPath, logs)
                End If
            Next
        Next
        
        ' STEP 4: Create assemblies (per variant)
        logs.Add("Creating assembly snapshots...")
        For Each variant In context.Variants
            Dim refMap = referenceMaps(variant.ConfigName)
            For Each file In context.ReleasePlan.Files.Where(
                Function(f) f.FileType = FileType.Assembly AndAlso f.ForVariants.Contains(variant.ConfigName))
                CreateAssemblySnapshot(app, file.SourcePath, file.TargetLocalPath, refMap, variant.Parameters, logs)
            Next
        Next
        
        ' STEP 5: Create drawings
        logs.Add("Creating drawings...")
        For Each file In context.ReleasePlan.Files.Where(Function(f) f.FileType = FileType.Drawing)
            Dim variantName = file.ForVariants(0)
            Dim refMap = referenceMaps(variantName)
            CreateDrawingCopy(app, file.SourcePath, file.TargetLocalPath, refMap, logs)
        Next
        
        logs.Add("All files created successfully!")
        Return True
        
    Finally
        ' ALWAYS restore parameters
        RestoreMasterParameters(app, paramSnapshot)
        app.ActiveDocument.Update()
    End Try
End Function
```

### Success Criteria:

#### Automated Verification:
- [ ] Vault logout executes successfully
- [ ] All target folders created
- [ ] All files saved to correct local paths
- [ ] Derivation links broken in parts
- [ ] Assembly references point to released files
- [ ] Drawing references point to released files
- [ ] Master parameters restored after execution

#### Manual Verification:
- [ ] Released assemblies open without errors
- [ ] Drawing views display correctly
- [ ] No unexpected Vault dialogs appear

---

## Phase 6: Vault Integration

### Overview
Log back into Vault, upload all files, and sync.

**Production mode only**: This entire phase is skipped in development mode. In development mode, files are saved locally and the manifest is updated, but no Vault upload occurs.

### Changes Required:

#### 1. Vault Login and Upload
**File**: `Lib/ModuleReleaseLib.vb`

```vb
Public Function UploadToVault(app As Inventor.Application, _
                               plan As ReleasePlan, _
                               ByRef logs As List(Of String)) As Boolean
    ' Skip in development mode
    If DEVELOPMENT_MODE Then
        logs.Add("Development mode: Skipping Vault upload")
        logs.Add($"Files saved locally to: {Path.GetDirectoryName(plan.Files.First().TargetLocalPath)}")
        Return True
    End If
    
    ' STEP 1: Login to Vault
    logs.Add("Logging into Vault...")
    app.CommandManager.ControlDefinitions.Item("LoginCmdIntName").Execute()
    
    ' Wait for connection (with timeout)
    Dim conn = WaitForVaultConnection(30000)  ' 30 second timeout
    If conn Is Nothing Then
        logs.Add("ERROR: Could not establish Vault connection!")
        Return False
    End If
    
    ' STEP 2: Ensure target folders exist in Vault
    logs.Add("Creating Vault folders...")
    EnsureVaultFoldersExist(conn, plan)
    
    ' STEP 3: Add all files to Vault
    logs.Add("Uploading files to Vault...")
    Dim localPaths = plan.Files.Select(Function(f) f.TargetLocalPath).ToList()
    Dim vaultPaths = plan.Files.Select(Function(f) f.TargetVaultPath).ToList()
    
    For i As Integer = 0 To plan.Files.Count - 1
        Dim file = plan.Files(i)
        Try
            Dim vaultFolder = Path.GetDirectoryName(file.TargetVaultPath).Replace("\", "/")
            VaultNumberingLib.AddFileToVault(conn, vaultFolder, file.TargetLocalPath, "Module release")
            logs.Add($"  Uploaded: {Path.GetFileName(file.TargetLocalPath)}")
        Catch ex As Exception
            logs.Add($"ERROR uploading {file.TargetLocalPath}: {ex.Message}")
            Return False
        End Try
    Next
    
    ' STEP 4: Sync files to establish tracking
    logs.Add("Syncing files from Vault...")
    For Each file In plan.Files
        Try
            VaultNumberingLib.SyncFileFromVault(file.TargetVaultPath)
        Catch ex As Exception
            logs.Add($"WARNING: Could not sync {file.TargetVaultPath}: {ex.Message}")
        End Try
    Next
    
    logs.Add("Vault upload complete!")
    Return True
End Function

Private Function WaitForVaultConnection(timeoutMs As Integer) As Object
    Dim startTime = DateTime.Now
    While (DateTime.Now - startTime).TotalMilliseconds < timeoutMs
        Try
            Dim conn = VaultNumberingLib.GetVaultConnection()
            If conn IsNot Nothing Then Return conn
        Catch
        End Try
        System.Threading.Thread.Sleep(500)
    End While
    Return Nothing
End Function

Private Sub EnsureVaultFoldersExist(conn As Object, plan As ReleasePlan)
    Dim folders As New HashSet(Of String)
    
    ' Collect all unique folder paths
    For Each file In plan.Files
        Dim folder = Path.GetDirectoryName(file.TargetVaultPath).Replace("\", "/")
        folders.Add(folder)
    End For
    
    ' Create folders that don't exist
    For Each folder In folders
        Try
            VaultNumberingLib.EnsureFolderExists(conn, folder)
        Catch
        End Try
    Next
End Sub
```

#### 2. Final Summary Dialog
**File**: `Lib/ModuleReleaseLib.vb`

```vb
Public Sub ShowCompletionSummary(plan As ReleasePlan, logs As List(Of String))
    Dim summary As String = $"Module Release Complete!
    
Files created: {plan.Files.Count}
  - Shared: {plan.Files.Count(Function(f) f.IsShared)}
  - Variant-specific: {plan.Files.Count(Function(f) Not f.IsShared)}

Vault numbers used: {plan.Files.First().VaultNumber} to {plan.Files.Last().VaultNumber}

All files have been uploaded to Vault and synced."

    MessageBox.Show(summary, "Release Complete", MessageBoxButtons.OK, MessageBoxIcon.Information)
End Sub
```

### Success Criteria:

#### Development Mode:
- [ ] Files saved to local output folder
- [ ] Manifest updated with file information
- [ ] All references work locally
- [ ] No Vault interaction occurs

#### Production Mode:
- [ ] Vault login succeeds
- [ ] All files added to Vault
- [ ] Files appear in correct Vault folders
- [ ] Files synced (visible in Inventor with Vault icons)
- [ ] Files visible in Vault Explorer
- [ ] Can check out released files
- [ ] File numbers match reserved numbers
- [ ] Assemblies reference correct Vault paths

---

## Manifest Structure

The manifest (`Moodulid/_manifest.json`) enables cross-module sharing detection and re-release optimization.

### Data Structure

```vb
Public Class ReleaseManifest
    Public LastUpdated As DateTime
    Public Modules As List(Of ModuleEntry)        ' All released modules
    Public SharedParts As List(Of SharedPartEntry) ' Cross-module shared parts
End Class

Public Class ModuleEntry
    Public ModuleName As String                   ' e.g., "Selg", "Iste"
    Public Variants As List(Of VariantEntry)
    Public ReleaseDate As DateTime
End Class

Public Class VariantEntry
    Public ConfigName As String                   ' e.g., "Selg-1800"
    Public VaultFolder As String                  ' $/Product/Moodulid/Selg-1800
    Public Parts As List(Of String)               ' Vault paths of parts used
    Public Assemblies As List(Of String)          ' Vault paths of assemblies
    Public Drawings As List(Of String)            ' Vault paths of drawings
End Class

Public Class SharedPartEntry
    Public VaultPath As String                    ' $/Product/Moodulid/Ühine/...
    Public VaultNumber As String                  ' e.g., "00001" (released file number)
    Public SourcePartNumber As String             ' Original Alusmoodulid part number (stable across renames)
    Public GeometryFingerprint As String          ' Geometry-only fingerprint
    Public UsedByModules As List(Of String)       ' ["Selg", "Iste", "Jalad"]
    Public UsedByVariants As List(Of String)      ' ["Iste 70", "Iste 90", ...]
    Public ReleaseDate As DateTime
End Class
```

### Manifest Operations

```vb
Public Function ReadManifest(manifestPath As String) As ReleaseManifest
    If Not File.Exists(manifestPath) Then Return Nothing
    Dim json As String = File.ReadAllText(manifestPath)
    Return DeserializeManifest(json)
End Function

Public Sub WriteManifest(manifestPath As String, manifest As ReleaseManifest)
    manifest.LastUpdated = DateTime.Now
    Dim json As String = SerializeManifest(manifest)
    File.WriteAllText(manifestPath, json)
End Sub

Public Sub UpdateManifestForRelease(manifest As ReleaseManifest, _
                                     moduleName As String, _
                                     plan As ReleasePlan, _
                                     partGroups As List(Of PartGroup))
    ' Add/update module entry
    Dim moduleEntry = manifest.Modules.FirstOrDefault(Function(m) m.ModuleName = moduleName)
    If moduleEntry Is Nothing Then
        moduleEntry = New ModuleEntry With {.ModuleName = moduleName}
        manifest.Modules.Add(moduleEntry)
    End If
    moduleEntry.ReleaseDate = DateTime.Now
    moduleEntry.Variants = BuildVariantEntries(plan)
    
    ' Update shared parts registry
    For Each file In plan.Files.Where(Function(f) f.IsShared AndAlso f.FileType = FileType.Part)
        Dim sharedEntry = manifest.SharedParts.FirstOrDefault(Function(s) s.VaultPath = file.TargetVaultPath)
        If sharedEntry Is Nothing Then
            Dim group = partGroups.First(Function(g) g.PartPath = file.SourcePath)
            sharedEntry = New SharedPartEntry With {
                .VaultPath = file.TargetVaultPath,
                .VaultNumber = file.VaultNumber,
                .Fingerprint = group.UniqueFingerprints.Keys.First(),
                .SourcePath = file.SourcePath,
                .UsedByModules = New List(Of String),
                .UsedByVariants = New List(Of String),
                .ReleaseDate = DateTime.Now
            }
            manifest.SharedParts.Add(sharedEntry)
        End If
        
        ' Update usage tracking
        If Not sharedEntry.UsedByModules.Contains(moduleName) Then
            sharedEntry.UsedByModules.Add(moduleName)
        End If
        For Each variantName In file.ForVariants
            If Not sharedEntry.UsedByVariants.Contains(variantName) Then
                sharedEntry.UsedByVariants.Add(variantName)
            End If
        Next
    Next
End Sub
```

### Cross-Module Workflow

1. **First module release** (e.g., Module A):
   - No manifest exists → create new
   - All shared parts added to `Ühine/` with new Vault numbers
   - Manifest records fingerprints and usage

2. **Second module release** (e.g., Module B):
   - Read existing manifest
   - For each shared part candidate, check manifest for matching fingerprint
   - If match found → reuse existing file (no new Vault number)
   - If no match → create new file in `Ühine/`
   - Update manifest with Module B's usage

3. **Re-release** (e.g., Module A updated):
   - Check fingerprints against manifest
   - If geometry unchanged → skip file creation
   - If geometry changed → create new version (Vault versioning)
   - Update manifest

---

## Testing Strategy

### Development Mode Testing

All phases 1-5 can be tested without Vault connection:
1. Set `DEVELOPMENT_MODE = True` in configuration
2. Set `NUMBERING_SCHEME = "Test numbriskeem"` (for when Vault IS connected)
3. Files are saved to local workspace folder
4. Manifest is written locally
5. Cross-module sharing works via local manifest

**Switch to production mode** by setting `DEVELOPMENT_MODE = False` and `NUMBERING_SCHEME = "Softcom numbriskeem"`.

### Unit Tests
- Fingerprint determinism: same file = same fingerprint
- Parameter snapshot/restore roundtrip
- Reference map computation
- Path conversion (Vault ↔ local)

### Integration Tests

#### Test Case 1: Simple Part Release
1. Single part with 2 variants
2. Verify standalone creation
3. Verify correct folder placement

#### Test Case 2: Assembly with Shared Parts
1. Assembly with 3 parts, 2 shared
2. Verify shared parts in Ühine
3. Verify assemblies reference shared parts

#### Test Case 3: Full Module Release
1. Multi-assembly module with drawings
2. 5+ variants
3. Verify complete folder structure
4. Verify all drawings reference correct files

#### Test Case 4: Cross-Module Sharing
1. Release Module A (creates shared parts in Ühine)
2. Release Module B (should reuse parts with matching fingerprints)
3. Verify manifest tracks both modules' usage
4. Verify no duplicate files created for identical geometry
5. Re-release Module A after geometry change
6. Verify Module B still references original shared parts

### Manual Testing Steps
1. Run full module release on test assembly
2. Open each variant assembly in fresh Inventor session
3. Verify no broken references
4. Verify drawing views display correctly
5. Check Vault file locations manually
6. Test re-release (should overwrite existing files)

---

## References

### Research Documents
- `docs/research/2026-04-26-moodulid-api-research.md` - API capabilities
- `docs/research/2026-04-26-drawing-reference-alternatives.md` - Drawing reference update approaches
- `docs/research/2026-04-26-vault-new-file-location.md` - Vault dialog workaround

### Test Results
- `Katsetused/Moodulid/README.md` - Concept test results
- Test1: Fingerprinting ✅
- Test2: BreakLinkToFile ✅
- Test3: Transaction rollback ✅
- Test8: ReplaceReference ✅
- Test10: Disconnect-Save-Add ✅
- Test12: Programmatic login/logout ✅

### Existing Code
- `Lib/VariantReleaseLib.vb` - Reference patterns
- `Lib/MakeComponentsLib.vb` - Fingerprinting
- `Lib/VaultNumberingLib.vb` - Vault APIs
- `Lib/ExcelReaderLib.vb` - Variant table reading

### Previous Plan
- `~/.cursor/plans/moodulid_step-by-step_implementation_7b11dbff.plan.md` - Earlier implementation plan (reference for data structures)
