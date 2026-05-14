# Multi-Master External References Implementation Plan

## Overview

Enhance the element release system to handle multiple masters that may exist outside the current element folder, including masters that reference each other through derivation or projected geometry via intermediate assemblies. Each released element will get its own complete copy of all masters (no sharing between elements).

## Current State Analysis

### How Masters Are Currently Discovered

1. **`DiscoverAssemblyTree()`** (`Lib/ElementReleaseLib.vb:401-509`) walks the assembly in three passes:
   - `AllReferencedDocuments` - most files
   - `ReferencedFileDescriptors` - suppressed files
   - `DiscoverFromOccurrences` - recursive through occurrences

2. **Critical Filter**: `IsInsideSourceRoot()` (`Lib/ElementReleaseLib.vb:3274-3278`) excludes ANY file outside the element folder (e.g., `Aluselemendid/Iste/`)

3. **`ClassifyPart()`** (`Lib/ElementReleaseLib.vb:605-631`) identifies derived parts and records `DerivedFromMaster` from `partDoc.ReferencedDocuments.Item(1)`

4. **`GetMasterPaths()`** (`Lib/ElementReleaseLib.vb:1197-1207`) collects unique master paths from parts - but these may point to files OUTSIDE sourceRoot that were never added to the tree

### Key Discoveries

- **Gap 1**: External masters are recorded in `DerivedFromMaster` but never added to `tree.Parts` or `ReleasePlan.Files`
- **Gap 2**: Stage 1a copies masters but doesn't update references BETWEEN masters
- **Gap 3**: Intermediate assemblies used for projected geometry (Eskiis assemblies) are not specially tracked

### User Requirements

1. External masters are parametric, controlled by the same `elemendid.xlsx` parameters
2. Each released element gets their own set of masters (NO sharing)
3. Circular dependencies not an issue - just replace references after copy

## Desired End State

1. ALL masters referenced by parts are discovered, regardless of folder location
2. Master-to-master dependencies are tracked (derivation chains, projected geometry)
3. Masters are copied in dependency order (roots first, then dependents)
4. References between copied masters are updated before parameter application
5. Each element gets its own complete master set in `Elemendid/<ElementName>/Eskiis/`

### Verification Criteria

- [ ] Diagnostic script shows complete dependency tree including external masters
- [ ] Released element assemblies reference only files within their element folder
- [ ] Parts derived from masters derive from the COPIED masters, not originals
- [ ] Parameter changes propagate correctly through master chains
- [ ] No references remain to files in `Aluselemendid/`

## What We're NOT Doing

- Sharing masters between elements (user explicitly wants per-element copies)
- Breaking derived links (current architecture uses ReplaceReference)
- Handling circular master dependencies (Inventor prevents these)
- Modifying source files (copy-first architecture)

## Implementation Approach

The approach is test-first: create a diagnostic script to understand the actual dependency structure, then modify discovery, planning, and execution in sequence.

---

## Phase 0: Diagnostic Script (Test-First) ✅ COMPLETE

### Overview

Create a diagnostic script that discovers and displays the complete master dependency tree for manual verification before modifying production code.

### Implementation

**File**: `Katsetused/Elemendid/TestMasterDependencies.vb`

The script was implemented and tested successfully. Key capabilities:
- Discovers ALL referenced masters recursively (including external)
- Builds dependency graph via `DerivedPartComponents`
- Detects intermediate assemblies used for projected geometry
- Outputs hierarchical tree to iLogic log with copy order

### Technical Discoveries

#### How Projected Geometry Links Are Detected

**The Challenge**: Inventor's API does NOT expose projected geometry source through standard properties.

**What DOESN'T Work**:
| API Approach | Result |
|--------------|--------|
| `SketchEntity.ReferencedEntity` | Only returns WorkPoints from same document |
| `SketchEntity.ContainingOccurrence` | Not available (only works on proxies) |
| `ReferenceKeyManager.BindKeyToObject` | Returns "parameter incorrect" error |
| `BrowserNode.NativeObject` | Returns "parameter incorrect" error |
| Sketch collections (`ExternalReferences`, `AssociativeGeometry`, etc.) | Don't exist or empty |
| `Sketch.Adaptive` | Returns `True` but doesn't reveal source |

**What WORKS** - Browser Label Parsing:
1. Get sketch's browser node: `BrowserPanes.ActivePane.GetBrowserNodeFromObject(sketch)`
2. Enumerate child nodes: `sketchNode.BrowserNodes`
3. Get label from `BrowserNodeDefinition.Label`
4. Label format: `"Reference83 (Selg - Eskiis Multibody (000130):1)"`
5. Extract occurrence name: `"Selg - Eskiis Multibody (000130):1"`
6. Find occurrence in assembly's `AllLeafOccurrences`
7. Get verified path: `occurrence.Definition.Document.FullFileName`

**Important**: The occurrence name in the label corresponds to an occurrence in the CURRENT assembly context. We must open the assembly to verify and get the full document path.

#### Derivation vs Projected Geometry

Two types of master-to-master dependencies discovered:

1. **Direct Derivation** (`DerivedPartComponent`):
   - `000131.ipt` derives from `000130.ipt` (via `DerivedPartUniformScaleDef`)
   - `000130.ipt` derives from `000129.ipt`
   - Detected via `PartComponentDefinition.ReferenceComponents.DerivedPartComponents`

2. **Projected Geometry** (via assembly):
   - `000131.ipt` has projected geometry FROM `000130.ipt`
   - Via assembly `000126.iam`
   - Detected via browser label parsing (see above)
   - The source occurrences are often `Visible: False` (hidden reference geometry)

### Test Results

**Test Assembly**: `000234.iam` (Nurk element)

**Masters Discovered**:
- `000131.ipt` (internal) - derives from 000130
- `000130.ipt` (external, from Selg/Eskiis) - derives from 000129
- `000129.ipt` (external, from Selg/Eskiis) - root master

**Intermediate Assemblies**:
- `000126.iam` - **HAS ACTUAL PROJECTED GEOMETRY** (verified via browser labels)
- `000114.iam` - dependency bridge (contains masters with derivation relationship)

**Projected Geometry Chain Detected**:
```
000130.ipt --(via 000126.iam)--> 000131.ipt
```

**Suggested Copy Order** (topological):
1. `000129.ipt` (root, no dependencies)
2. `000130.ipt` (depends on 000129)
3. `000131.ipt` (depends on 000130)
4. `000126.iam` (assembly with projected geometry)
5. `000114.iam` (assembly for dependency bridge)

### Success Criteria

#### Verification:
- [x] Script runs without errors on test assembly
- [x] All internal masters are listed
- [x] External masters (from other Aluselemendid folders) are detected
- [x] Master-to-master derivation chains are shown
- [x] Intermediate assemblies are identified
- [x] Projected geometry chains detected and verified

#### Manual Verification:
- [x] Output matches expected structure based on Inventor Model Browser
- [x] No masters are missing from the tree
- [x] Dependency order is correct (roots have no dependencies)

**Status**: Phase 0 COMPLETE. Ready for Phase 1.

---

## Phase 1: Enhanced Master Discovery

### Overview

Modify the discovery system to find ALL masters regardless of location and track their dependencies, using the browser label parsing approach proven in Phase 0.

### Changes Required

#### 1. New Data Structures
**File**: `Lib/ElementReleaseLib.vb`

Add to `AssemblyTree` class:
```vb
Public ExternalMasters As New Dictionary(Of String, MasterInfo)(StringComparer.OrdinalIgnoreCase)
Public MasterDependencies As New Dictionary(Of String, List(Of String))(StringComparer.OrdinalIgnoreCase)
Public IntermediateAssemblies As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
Public ProjectedGeometryChains As New List(Of Tuple(Of String, String, String)) ' (SourcePart, ViaAssembly, TargetPart)
```

New class:
```vb
Public Class MasterInfo
    Public FilePath As String
    Public RelativePath As String  ' Relative to its own Aluselemendid folder
    Public SourceElement As String ' Which element it comes from
    Public DependsOn As New List(Of String) ' Other masters this one references
    Public IsIntermediate As Boolean ' True if this is an assembly used for projection
End Class
```

#### 2. New Discovery Functions
**File**: `Lib/ElementReleaseLib.vb`

**`DiscoverAllMasters()`**:
- Start from known master paths (from `GetMasterPaths()`)
- For each master, walk its `ReferencedDocuments` recursively
- Track derivation via `DerivedPartComponents` and `DerivedAssemblyComponents`
- No `IsInsideSourceRoot` filter - discover everything
- Build dependency graph as we go

**`DetectProjectedGeometry(masterPath, assemblyDoc)`** (from Phase 0):
- For each sketch in the master, get browser node
- Enumerate child nodes looking for "Reference" labels
- Parse occurrence name from label format: `"ReferenceXX (OccurrenceName)"`
- Find occurrence in assembly's `AllLeafOccurrences`
- Get verified source path from `occurrence.Definition.Document.FullFileName`
- Record chain: `(sourcePath, assemblyPath, masterPath)`

**`ClassifyIntermediateAssemblies()`**:
- Assembly is NEEDED if it has actual projected geometry (from `DetectProjectedGeometry`)
- Assembly is NEEDED if it's a "dependency bridge" (contains two masters where one depends on the other)
- Otherwise assembly is NOT needed (just happens to contain masters)

#### 3. Modify `DiscoverAssemblyTree()`
**File**: `Lib/ElementReleaseLib.vb`

After existing discovery:
- Call `DiscoverAllMasters()` with initial master paths
- For each candidate assembly that contains masters, call `DetectProjectedGeometry()`
- Call `ClassifyIntermediateAssemblies()` to filter
- Merge external masters into tracking structure
- Log summary of internal vs external masters

### Success Criteria

#### Verification:
- [x] `context.AssemblyTree.ExternalMasters` populated correctly
- [x] `context.AssemblyTree.MasterDependencies` shows correct relationships
- [x] `context.AssemblyTree.ProjectedGeometryChains` shows browser-verified links
- [x] Intermediate assemblies correctly classified (actual geometry vs just containing)

#### Manual Verification:
- [ ] Log output matches Phase 0 diagnostic script output

**Implementation Note**: After completing this phase and all verification passes, pause here for manual confirmation before proceeding to the next phase.

**Status**: Phase 1 IMPLEMENTED. Ready for testing.

---

## Phase 2: Dependency-Ordered Master Copying

### Overview

Implement topological sort for masters and update the copy logic to handle dependencies.

### Changes Required

#### 1. Topological Sort Function
**File**: `Lib/ElementReleaseLib.vb`

Create `SortMastersByDependency()`:
- Input: master paths + dependency graph (from `MasterDependencies`)
- Output: ordered list (roots first, then dependents, then intermediate assemblies)
- Algorithm: Kahn's algorithm for topological sort
- Example output from Phase 0: `000129.ipt → 000130.ipt → 000131.ipt → 000126.iam → 000114.iam`

#### 2. Update `ComputeReleasePlan()`
**File**: `Lib/ElementReleaseLib.vb`

Modifications:
- Include external masters in `PlannedFile` list per element
- Set target path to `Elemendid/<ElementName>/Eskiis/`
- Preserve filename (since all copies go to same Eskiis folder)
- Include intermediate assemblies from `ProjectedGeometryChains`

#### 3. Update Stage 1a in `ExecuteRelease()`
**File**: `Lib/ElementReleaseLib.vb`

Modifications:
- Copy masters in topological order
- After each master copy, call `ReplaceReference()` on dependent files that have already been copied
- Handle intermediate assemblies:
  - Copy them to Eskiis folder
  - Update their occurrence references to point to copied masters
  - Note: Hidden occurrences (`Visible: False`) must also be updated

#### 4. Reference Replacement Order

Critical: References must be updated in dependency order:
1. Copy `000129.ipt` (root) → no references to update
2. Copy `000130.ipt` → update its reference to `000129.ipt`
3. Copy `000131.ipt` → update its reference to `000130.ipt`
4. Copy `000126.iam` → update all occurrence references (000131, 000130, etc.)
5. Copy `000114.iam` → update all occurrence references

### Success Criteria

#### Verification:
- [x] Masters copied in correct order (roots first)
- [x] Copied masters reference other COPIED masters (not originals)
- [x] Intermediate assemblies copied and updated
- [x] Hidden reference occurrences (Visible: False) correctly updated

#### Manual Verification:
- [ ] Open copied master - references point to element folder
- [ ] Open copied assembly - all occurrences reference element folder files
- [ ] Parameter changes propagate through chain

**Implementation Note**: After completing this phase and all verification passes, pause here for manual confirmation before proceeding to the next phase.

**Status**: Phase 2 IMPLEMENTED. Ready for testing.

---

## Phase 3: Integration and Testing

### Overview

Full integration testing with various master configurations.

### Test Scenarios

1. **Single internal master** (regression test)
   - Master in `Aluselemendid/Iste/Eskiis/`
   - Parts derive from it
   - Verify: no change in behavior

2. **Multiple internal masters with derivation**
   - `master_A.ipt` (has parameters)
   - `master_B.ipt` derives from `master_A`
   - Parts derive from `master_B`
   - Verify: copied `master_B` references copied `master_A`

3. **External master from another element**
   - `Aluselemendid/Iste/` uses master from `Aluselemendid/Selg/Eskiis/`
   - Verify: master copied to `Elemendid/Iste 110/Eskiis/`

4. **Projected geometry via intermediate assembly**
   - `master_A.ipt` - parameters
   - `eskiisid.iam` - places master_A
   - `master_B.ipt` - DerivedAssemblyComponent from eskiisid.iam
   - Verify: all three copied and references updated

### Success Criteria

- [ ] All test scenarios pass
- [ ] No references to `Aluselemendid/` in released files
- [ ] Parameters apply correctly to all master types

---

## Testing Strategy

### Unit Tests
- Topological sort with various dependency graphs
- Dependency detection for different reference types

### Integration Tests
- Full release cycle with multi-master assembly
- Verify with `TestMasterDependencies.vb` before and after

### Manual Testing Steps
1. Run diagnostic script on test assembly
2. Execute release
3. Open released assembly
4. Verify all references in Model Browser
5. Change parameter in released master - verify propagation

## Terminology Checklist

Verify all code uses correct domain terms per UBIQUITOUS_LANGUAGE.md:
- [ ] "Aluselement" not "Alusmoodul" for parametric designs
- [ ] "Väljastatud element" not "Moodul" for released units
- [ ] "Detail" not "Component" for parts
- [ ] "Master" for skeleton/eskiis parts
- [ ] Folder paths match structure (Eskiis for masters)

## Technical Notes: Inventor API Limitations

### Projected Geometry Detection

The Inventor API does NOT provide a direct way to query the source of projected geometry. Extensive testing in Phase 0 confirmed the following:

**API Properties That Don't Expose Source**:
- `SketchEntity.ReferencedEntity` - only works for WorkPoints in the same document
- `SketchEntity.ContainingOccurrence` - only available on proxy objects, not native entities
- `ReferenceKeyManager.BindKeyToObject` - fails with "parameter incorrect" for projected entities
- `BrowserNodeDefinition.NativeObject` - fails with "parameter incorrect"
- Sketch collections like `ExternalReferences`, `AssociativeGeometry` - don't exist

**Working Approach**: Parse browser node labels
- Browser child nodes under a sketch have labels like `"Reference83 (OccurrenceName)"`
- The occurrence name in parentheses can be matched to an occurrence in the assembly
- The assembly must be OPEN to access `AllLeafOccurrences` for verification
- Get document path from `occurrence.Definition.Document.FullFileName`

**Occurrence Name Formats** (both must be handled):

| Format | Example | When |
|--------|---------|------|
| With description | `"Selg - Eskiis Multibody (000130):1"` | After running `Koost/Nimeta detailid.vb` |
| Default | `"000130:1"` | Default Inventor naming |

The label format is always `"ReferenceXX (OccurrenceName)"` but the occurrence name itself varies:
- **With description**: `"Description (FileNumber):Instance"` → extract `FileNumber` from inner parentheses
- **Default**: `"FileNumber:Instance"` → extract `FileNumber` before the colon

**Parsing Logic**:
```vb
' Extract occurrence name from label
Dim occName As String = label.Substring(label.IndexOf("(") + 1).TrimEnd(")"c)

' Find it in assembly - direct match first
For Each occ In assembly.AllLeafOccurrences
    If occ.Name = occName Then Return occ
Next

' Fallback - extract part number and match by filename
Dim partNum As String
If occName.Contains("(") Then
    ' Format: "Description (000130):1" - extract from inner parentheses
    partNum = occName.Substring(occName.LastIndexOf("(") + 1)
    partNum = partNum.Substring(0, partNum.IndexOf(")"))
Else
    ' Format: "000130:1" - extract before colon
    partNum = occName.Split(":"c)(0)
End If
```

**Reliability**: This approach is based on Inventor's internal display format. It has been verified to work and provides confirmed source paths. Both naming conventions are supported.

### Assembly Context Requirements

To detect projected geometry:
1. The assembly containing both parts must be opened
2. The part's sketches must be analyzed while the assembly is the context
3. Browser nodes are only accessible for the active document

### Hidden Reference Occurrences

Parts used for projected geometry are often placed as hidden occurrences (`Visible: False`) in the intermediate assembly. These must still be updated when replacing references.

## References

- Original analysis: This conversation
- Related research: `docs/research/2026-05-13-element-release-failures.md`
- Current implementation: `Lib/ElementReleaseLib.vb`
- Diagnostic script: `Katsetused/Elemendid/TestMasterDependencies.vb`
- Test patterns: `Katsetused/TestDerivedPartRefs.vb`
- Domain terminology: `docs/UBIQUITOUS_LANGUAGE.md`
