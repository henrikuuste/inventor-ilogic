# Terminology Refactoring Plan

## Overview

Refactor the codebase to use the correct domain terminology established in `docs/UBIQUITOUS_LANGUAGE.md`. The key change is that what we previously called "Alusmoodul" (parametric design) is now "Aluselement", and "Alusmoodul" now refers to module assembly definitions.

## Current State Analysis

The codebase currently uses terminology that conflates two distinct concepts:
- **Alusmoodul** = parametric design with masters (should be **Aluselement**)
- **Moodul** = released manufactured unit (should be **Väljastatud element**)

This causes confusion because the true **Alusmoodul** (module assembly) and **Moodul** (final assembly product) concepts have no representation in the current code.

### Key Discoveries

1. **`Lib/ModuleReleaseLib.vb`** - Main library using old terminology throughout (2145 lines)
2. **`Lib/ExcelReaderLib.vb`** - Uses "MooduliNimi", "ReleaseConfig" for element releases
3. **`Moodulid/` folder** - Contains element release scripts, should be `Elemendid/`
4. **`Moodulid/README.md`** - Documentation using old terminology
5. **`AGENTS.md`** - Project conventions with old folder structure

### Files Requiring Changes

| File | Change Type | Priority |
|------|-------------|----------|
| `Lib/ModuleReleaseLib.vb` | Rename classes, variables, comments | HIGH |
| `Lib/ExcelReaderLib.vb` | Rename ReleaseConfig, column names | HIGH |
| `Moodulid/Loo moodulid.vb` | Rename to element release, update calls | HIGH |
| `Moodulid/Loo alusmoodul.vb` | Rename to base element creation | HIGH |
| `Moodulid/README.md` | Update all terminology | HIGH |
| `AGENTS.md` | Update folder structure documentation | HIGH |
| `Katsetused/Moodulid/*.vb` | Update test file terminology | MEDIUM |
| `docs/plans/2026-04-26-module-release-cycle.md` | Update plan terminology | LOW |
| `docs/research/*.md` | Update research docs | LOW |

## Desired End State

### Folder Structure Changes

```
CURRENT                          → TARGET
Moodulid/                        → Elemendid/  (scripts for element release)
  Loo moodulid.vb                →   Loo elemendid.vb
  Loo alusmoodul.vb              →   Loo aluselement.vb
  README.md                      →   README.md (updated)
Katsetused/Moodulid/             → Katsetused/Elemendid/
```

### Class/Type Renaming

| Current Name | New Name | Location |
|--------------|----------|----------|
| `ModuleReleaseLib` | `ElementReleaseLib` | `Lib/ModuleReleaseLib.vb` → `Lib/ElementReleaseLib.vb` |
| `ReleaseContext` | `ElementReleaseContext` | `Lib/ElementReleaseLib.vb` |
| `VariantMatrix` | `ElementMatrix` | `Lib/ElementReleaseLib.vb` |
| `ReleaseConfig` | `ElementConfig` | `Lib/ExcelReaderLib.vb` |
| `ReleaseConfig.ConfigName` | `ElementConfig.ElementName` | `Lib/ExcelReaderLib.vb` |

### Excel Column Renaming

| Current Column | New Column | Notes |
|----------------|------------|-------|
| `MooduliNimi` | `ElementiNimi` | Released element name |

### Variable Renaming Patterns

| Pattern | Replacement |
|---------|-------------|
| `moduleName` | `elementName` |
| `moduleFolder` | `elementFolder` |
| `ModuleName` | `ElementName` |
| `variantName` | `releasedElementName` |
| `variant` (when meaning released element) | `releasedElement` |

## What We're NOT Doing

- **Vault folder migration** - Existing Vault folders will be migrated manually, not by this refactoring
- **Module assembly features** - True Alusmoodul/Moodul (element combination) features are not yet implemented
- **Excel file migration** - Existing `moodulid.xlsx` files will continue to work during transition

## Implementation Approach

Refactor in phases, starting with documentation and libraries, then scripts. Each phase is independently testable.

---

## Phase 1: Update Documentation

### Overview
Update all documentation files to use correct terminology. This establishes the reference for code changes.

### Changes Required

#### 1. AGENTS.md
**File**: `AGENTS.md`
**Changes**:
- Update folder structure diagram (lines 18-32)
- Replace "Alusmoodulid" → "Aluselemendid" for parametric designs
- Add note about terminology transition
- Keep both old and new folder names documented during transition

#### 2. Moodulid/README.md
**File**: `Moodulid/README.md`
**Changes**:
- Rename to describe element release (will be moved with folder)
- Replace all "moodul" → "element" references
- Update Excel format documentation
- Update folder structure diagram

### Success Criteria

#### Verification:
- [ ] AGENTS.md terminology matches UBIQUITOUS_LANGUAGE.md
- [ ] README.md uses correct element terminology
- [ ] No references to "moodul" when meaning "element" in docs

---

## Phase 2: Rename Library Module

### Overview
Rename `ModuleReleaseLib` to `ElementReleaseLib` and update all internal terminology.

### Changes Required

#### 1. Rename file
**From**: `Lib/ModuleReleaseLib.vb`
**To**: `Lib/ElementReleaseLib.vb`

#### 2. Update module header
```vb
' CURRENT
Public Module ModuleReleaseLib

' NEW
Public Module ElementReleaseLib
```

#### 3. Update class names
```vb
' CURRENT
Public Class ReleaseContext
    Public ModuleName As String
    Public VariantMatrix As VariantMatrix

' NEW  
Public Class ElementReleaseContext
    Public ElementName As String
    Public ElementMatrix As ElementMatrix
```

#### 4. Update comments throughout
Replace all comments referencing "module" (when meaning element) with "element".

#### 5. Update ExcelReaderLib.vb
**File**: `Lib/ExcelReaderLib.vb`
**Changes**:
- Rename `ReleaseConfig` → `ElementConfig`
- Rename `ConfigName` → `ElementName`
- Update Excel format comments (MooduliNimi → ElementiNimi)
- Keep backward compatibility: accept both column names during transition

### Success Criteria

#### Verification:
- [ ] `Lib/ElementReleaseLib.vb` exists with renamed module
- [ ] All classes renamed (ReleaseContext → ElementReleaseContext, etc.)
- [ ] No compile errors in library
- [ ] ExcelReaderLib accepts both old and new column names

---

## Phase 3: Rename Script Folder and Files

### Overview
Rename the `Moodulid/` script folder to `Elemendid/` and update script names.

### Changes Required

#### 1. Rename folder
**From**: `Moodulid/`
**To**: `Elemendid/`

#### 2. Rename script files
| Current | New |
|---------|-----|
| `Loo moodulid.vb` | `Loo elemendid.vb` |
| `Loo alusmoodul.vb` | `Loo aluselement.vb` |

#### 3. Update script contents
Each script needs:
- Update `AddVbFile` paths (`Lib/ModuleReleaseLib.vb` → `Lib/ElementReleaseLib.vb`)
- Update class references
- Update log messages
- Update user-facing Estonian text

### Success Criteria

#### Verification:
- [ ] `Elemendid/` folder exists with renamed scripts
- [ ] Scripts run without compile errors
- [ ] Scripts can release elements with updated terminology

---

## Phase 4: Update Test Scripts

### Overview
Rename test folder and update test scripts to use new terminology.

### Changes Required

#### 1. Rename folder
**From**: `Katsetused/Moodulid/`
**To**: `Katsetused/Elemendid/`

#### 2. Update test scripts
Update all `AddVbFile` references and terminology in:
- `Test1_Fingerprint.vb`
- `Test2_BreakLink.vb`
- `Test3_TransactionRollback.vb`
- `Test4_DrawingRelink.vb`
- `Test5_ParameterCycle.vb`
- `Test6_StandaloneCopy.vb`
- `Test7_BinaryPatch.vb`
- `Test8_FileDescriptorReplaceReference.vb`
- `Test9_VaultSaveLocation.vb`
- `Test10_DisconnectSaveCheckin.vb`
- `Test11_*.vb` through `Test14_*.vb`

### Success Criteria

#### Verification:
- [ ] All test scripts compile
- [ ] Test scripts reference correct library paths

---

## Phase 5: Update Historical Documentation

### Overview
Update research documents and old plans for terminology consistency.

### Changes Required

#### 1. Research documents
**Files**: `docs/research/*.md`
- Add note at top indicating historical terminology
- Optionally update inline references

#### 2. Old plan
**File**: `docs/plans/2026-04-26-module-release-cycle.md`
- Add deprecation note at top
- Reference new terminology document

### Success Criteria

#### Verification:
- [ ] Historical docs marked with terminology note
- [ ] New plans reference UBIQUITOUS_LANGUAGE.md

---

## Testing Strategy

### Unit Tests
- Verify ElementReleaseLib compiles without errors
- Verify ExcelReaderLib reads both old and new column names

### Integration Tests
- Run element release on test assembly
- Verify output folder structure
- Verify fingerprint computation works

### Manual Testing Steps
1. Open existing base element assembly
2. Run "Loo elemendid" script
3. Verify dialog shows correct Estonian text
4. Verify released elements created in correct folders
5. Verify manifest JSON uses new terminology

## Terminology Checklist

Verify all code uses correct domain terms per UBIQUITOUS_LANGUAGE.md:
- [ ] "Aluselement" not "Alusmoodul" for parametric designs
- [ ] "Väljastatud element" not "Moodul" for released units  
- [ ] "Element" not "Module" when referring to manufactured units
- [ ] "Detail" not "Component" for parts
- [ ] "Elemendid" not "Moodulid" for script/output folders
- [ ] "ElementiNimi" not "MooduliNimi" in Excel columns
- [ ] Folder paths match new structure

## Backward Compatibility Notes

During transition, the code should:
1. Accept both `MooduliNimi` and `ElementiNimi` Excel columns
2. Log warnings when old terminology detected
3. Continue to work with existing Vault folder structures until manual migration

## Migration Sequence

```
1. Update UBIQUITOUS_LANGUAGE.md    ✓ (completed)
2. Update AGENTS.md                  
3. Rename Lib/ModuleReleaseLib.vb → Lib/ElementReleaseLib.vb
4. Update ExcelReaderLib.vb
5. Rename Moodulid/ → Elemendid/
6. Update all scripts with new AddVbFile paths
7. Rename Katsetused/Moodulid/ → Katsetused/Elemendid/
8. Update historical docs
```

## References

- Domain terminology: `docs/UBIQUITOUS_LANGUAGE.md`
- Project conventions: `AGENTS.md`
- Current element release plan: `docs/plans/2026-04-26-module-release-cycle.md`
