# Terminology Refactoring Plan

**Last updated**: 2026-05-12

## Overview

Refactor the codebase to use the correct domain terminology established in `docs/UBIQUITOUS_LANGUAGE.md`. The key change is that what we previously called "Alusmoodul" (parametric design) is now "Aluselement", and "Alusmoodul" now refers to module assembly definitions.

## Current State Analysis

The codebase currently uses terminology that conflates two distinct concepts:
- **Alusmoodul** = parametric design with masters (should be **Aluselement**)
- **Moodul** = released manufactured unit (should be **Väljastatud element**)

This causes confusion because the true **Alusmoodul** (module assembly) and **Moodul** (final assembly product) concepts have no representation in the current code.

### Good News: Centralized Constants

The codebase has been improved with centralized constants that make refactoring easier:

1. **`Lib/BaseModuleLayoutLib.vb`** - Folder path constants centralized here
   - Note in file header: "when terminology/folder naming is refactored (module -> element), update the constants in this file"
   - Changing `SEG_BASE_MODULES` propagates to all callers

2. **`Lib/StringsLib.vb`** - Estonian UI strings centralized here
   - All user-facing module/element text in one place

### Files Requiring Changes

| File | Change Type | Priority | Notes |
|------|-------------|----------|-------|
| **Lib/BaseModuleLayoutLib.vb** | Update constants, rename functions | HIGH | Central location - changes propagate |
| **Lib/StringsLib.vb** | Update UI string constants | HIGH | Central location - changes propagate |
| `Lib/ModuleReleaseLib.vb` | Rename module, classes, variables | HIGH | 2145 lines |
| `Lib/ExcelReaderLib.vb` | Rename ReleaseConfig, column names | HIGH | |
| `Moodulid/Loo moodulid.vb` | Rename to element release, update calls | HIGH | |
| `Moodulid/Loo alusmoodul.vb` | Rename to base element creation | HIGH | |
| `Loo detailid.vb` | Update AddVbFile paths | MEDIUM | Uses BaseElementLayoutLib |
| `Koost/Sorteeri detailid.vb` | Update AddVbFile paths | MEDIUM | Uses BaseModuleLayoutLib |
| `Lib/MaterialRoutingLib.vb` | Verify uses BaseModuleLayoutLib | LOW | Should auto-update via constants |
| `Moodulid/README.md` | Update all terminology | HIGH | |
| `AGENTS.md` | Update folder structure documentation | HIGH | |
| `Katsetused/Moodulid/*.vb` | Update 16 test files | MEDIUM | |
| `docs/plans/2026-04-26-module-release-cycle.md` | Add deprecation note | LOW | |
| `docs/research/*.md` | Add terminology note to 8 files | LOW | |

### New Files Since Original Plan

| File | Status | Notes |
|------|--------|-------|
| `Lib/BaseModuleLayoutLib.vb` | NEW | Centralizes folder constants - key refactoring target |
| `Lib/StringsLib.vb` | NEW | Centralizes UI strings - key refactoring target |
| `Lib/UILib.vb` | NEW | No terminology issues |
| `Lib/MaterialRoutingLib.vb` | NEW | Uses BaseModuleLayoutLib constants |
| `Detailid/Pinnalaotuse vaated.vb` | NEW | New script folder |
| `Test11_DrawingTitleBlockUpdate.vb` | NEW | Test file |
| `Test11_UserInfo.vb` | NEW | Test file |
| `Test12_VaultLoginLogout.vb` | NEW | Test file |
| `Test12_VerifyPartNumber.vb` | NEW | Test file |
| `Test13_DiagnoseTitleBlock.vb` | NEW | Test file |
| `Test14_CheckDrawingViewPN.vb` | NEW | Test file |

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

### Constant Changes in BaseModuleLayoutLib.vb

```vb
' CURRENT
Public Const SEG_BASE_MODULES As String = "Alusmoodulid"

' NEW
Public Const SEG_BASE_ELEMENTS As String = "Aluselemendid"
```

### String Changes in StringsLib.vb

| Current Constant | Current Value | New Value |
|------------------|---------------|-----------|
| `TITLE_MODULE_RELEASE` | "Moodulite väljastamine" | "Elementide väljastamine" |
| `BTN_ALL_MODULES` | "Kõik moodulid" | "Kõik elemendid" |
| `BTN_FIRST_MODULE` | "Esimene moodul" | "Esimene element" |
| `TITLE_CREATE_BASE_MODULE` | "Loo alusmoodul" | "Loo aluselement" |

### Class/Type Renaming

| Current Name | New Name | Location |
|--------------|----------|----------|
| `ModuleReleaseLib` | `ElementReleaseLib` | `Lib/ModuleReleaseLib.vb` → `Lib/ElementReleaseLib.vb` |
| `BaseModuleLayoutLib` | `BaseElementLayoutLib` | `Lib/BaseModuleLayoutLib.vb` → `Lib/BaseElementLayoutLib.vb` |
| `ReleaseContext` | `ElementReleaseContext` | `Lib/ElementReleaseLib.vb` |
| `VariantMatrix` | `ElementMatrix` | `Lib/ElementReleaseLib.vb` |
| `ReleaseConfig` | `ElementConfig` | `Lib/ExcelReaderLib.vb` |
| `ReleaseConfig.ConfigName` | `ElementConfig.ElementName` | `Lib/ExcelReaderLib.vb` |

### Excel Column Renaming

| Current Column | New Column | Notes |
|----------------|------------|-------|
| `MooduliNimi` | `Element` | Released element name |

### Variable Renaming Patterns

| Pattern | Replacement |
|---------|-------------|
| `moduleName` | `elementName` |
| `moduleFolder` | `elementFolder` |
| `ModuleName` | `ElementName` |
| `moduleRoot` | `elementRoot` |
| `variantName` | `releasedElementName` |
| `variant` (when meaning released element) | `releasedElement` |

## What We're NOT Doing

- **Vault folder migration** - Existing Vault folders will be migrated manually, not by this refactoring
- **Module assembly features** - True Alusmoodul/Moodul (element combination) features are not yet implemented
- **Excel file migration** - Existing `moodulid.xlsx` files will continue to work during transition

## Implementation Approach

Refactor in phases, starting with centralized constants, then libraries, then scripts. Each phase is independently testable.

---

## Phase 1: Update Centralized Constants

### Overview
Update the centralized constant libraries first. This propagates changes to all callers automatically.

### Changes Required

#### 1. BaseModuleLayoutLib.vb
**File**: `Lib/BaseModuleLayoutLib.vb`
**Changes**:
- Rename file to `Lib/BaseElementLayoutLib.vb`
- Rename module: `BaseModuleLayoutLib` → `BaseElementLayoutLib`
- Update constant: `SEG_BASE_MODULES` → `SEG_BASE_ELEMENTS` = "Aluselemendid"
- Update function names:
  - `DetectModuleRootFromMasterPath` → `DetectElementRootFromMasterPath`
  - `GetModuleName` → `GetElementName`
  - `EnumerateExpectedFolders` (keep name, update comments)
- Update variable names: `moduleRoot` → `elementRoot`, `moduleName` → `elementName`
- Update comments throughout

#### 2. StringsLib.vb
**File**: `Lib/StringsLib.vb`
**Changes**:
- Update section comment: "MODULE RELEASE" → "ELEMENT RELEASE"
- Update constants:
  ```vb
  ' CURRENT
  Public Const TITLE_MODULE_RELEASE As String = "Moodulite väljastamine"
  Public Const BTN_ALL_MODULES As String = "Kõik moodulid"
  Public Const BTN_FIRST_MODULE As String = "Esimene moodul"
  Public Const TITLE_CREATE_BASE_MODULE As String = "Loo alusmoodul"
  
  ' NEW
  Public Const TITLE_ELEMENT_RELEASE As String = "Elementide väljastamine"
  Public Const BTN_ALL_ELEMENTS As String = "Kõik elemendid"
  Public Const BTN_FIRST_ELEMENT As String = "Esimene element"
  Public Const TITLE_CREATE_BASE_ELEMENT As String = "Loo aluselement"
  ```

### Success Criteria

#### Verification:
- [ ] `Lib/BaseElementLayoutLib.vb` exists with renamed module and constants
- [ ] `Lib/StringsLib.vb` has updated element terminology
- [ ] No compile errors in libraries
- [ ] All callers of old constants identified and updated

---

## Phase 2: Update Main Release Library

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

#### 4. Update enum values
```vb
' CURRENT
Public Enum ReleaseMode
    FullModule = 1

' NEW
Public Enum ReleaseMode
    FullElement = 1
```

#### 5. Update comments throughout
Replace all comments referencing "module" (when meaning element) with "element".

#### 6. Update ExcelReaderLib.vb
**File**: `Lib/ExcelReaderLib.vb`
**Changes**:
- Rename `ReleaseConfig` → `ElementConfig`
- Rename `ConfigName` → `ElementName`
- Update Excel format comments (MooduliNimi → Element)
- Keep backward compatibility: accept both column names during transition

### Success Criteria

#### Verification:
- [ ] `Lib/ElementReleaseLib.vb` exists with renamed module
- [ ] All classes renamed (ReleaseContext → ElementReleaseContext, etc.)
- [ ] No compile errors in library
- [ ] ExcelReaderLib accepts both old and new column names

---

## Phase 3: Update Caller Scripts

### Overview
Update all scripts that use the renamed libraries.

### Changes Required

#### 1. Scripts using BaseModuleLayoutLib → BaseElementLayoutLib

| Script | AddVbFile Change |
|--------|------------------|
| `Moodulid/Loo alusmoodul.vb` | `Lib/BaseModuleLayoutLib.vb` → `Lib/BaseElementLayoutLib.vb` |
| `Loo detailid.vb` | Same |
| `Koost/Sorteeri detailid.vb` | Same |

#### 2. Scripts using ModuleReleaseLib → ElementReleaseLib

| Script | AddVbFile Change |
|--------|------------------|
| `Moodulid/Loo moodulid.vb` | `Lib/ModuleReleaseLib.vb` → `Lib/ElementReleaseLib.vb` |

#### 3. Update MaterialRoutingLib.vb
**File**: `Lib/MaterialRoutingLib.vb`
**Changes**:
- Update references: `BaseModuleLayoutLib.SEG_*` → `BaseElementLayoutLib.SEG_*`

### Success Criteria

#### Verification:
- [ ] All scripts compile without errors
- [ ] All AddVbFile paths updated

---

## Phase 4: Rename Script Folder and Files

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
- Update log messages
- Update user-facing Estonian text (now via StringsLib constants)

### Success Criteria

#### Verification:
- [ ] `Elemendid/` folder exists with renamed scripts
- [ ] Scripts run without compile errors
- [ ] Scripts can release elements with updated terminology

---

## Phase 5: Update Test Scripts

### Overview
Rename test folder and update test scripts to use new terminology.

### Changes Required

#### 1. Rename folder
**From**: `Katsetused/Moodulid/`
**To**: `Katsetused/Elemendid/`

#### 2. Update test scripts (16 files)
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
- `Test11_DrawingTitleBlockUpdate.vb`
- `Test11_UserInfo.vb`
- `Test12_VaultLoginLogout.vb`
- `Test12_VerifyPartNumber.vb`
- `Test13_DiagnoseTitleBlock.vb`
- `Test14_CheckDrawingViewPN.vb`

### Success Criteria

#### Verification:
- [ ] All 16 test scripts compile
- [ ] Test scripts reference correct library paths

---

## Phase 6: Update Documentation

### Overview
Update all documentation files to use correct terminology.

### Changes Required

#### 1. AGENTS.md
**File**: `AGENTS.md`
**Changes**:
- Update folder structure diagram (lines 18-32)
- Replace "Alusmoodulid" → "Aluselemendid" for parametric designs
- Add note about terminology transition
- Keep both old and new folder names documented during transition

#### 2. Moodulid/README.md → Elemendid/README.md
**File**: `Elemendid/README.md` (after folder rename)
**Changes**:
- Update title and description
- Replace all "moodul" → "element" references
- Update Excel format documentation
- Update folder structure diagram

#### 3. Historical Documentation
**Files**: `docs/research/*.md` (8 files), `docs/plans/2026-04-26-module-release-cycle.md`
- Add deprecation/terminology note at top
- Reference new terminology document

### Success Criteria

#### Verification:
- [ ] AGENTS.md terminology matches UBIQUITOUS_LANGUAGE.md
- [ ] README.md uses correct element terminology
- [ ] Historical docs marked with terminology note

---

## Testing Strategy

### Unit Tests
- Verify ElementReleaseLib compiles without errors
- Verify BaseElementLayoutLib compiles without errors
- Verify ExcelReaderLib reads both old and new column names

### Integration Tests
- Run element release on test assembly
- Verify output folder structure
- Verify fingerprint computation works

### Manual Testing Steps
1. Open existing base element assembly
2. Run "Loo elemendid" script
3. Verify dialog shows correct Estonian text ("Elementide väljastamine")
4. Verify released elements created in correct folders
5. Verify manifest JSON uses new terminology

## Terminology Checklist

Verify all code uses correct domain terms per UBIQUITOUS_LANGUAGE.md:
- [ ] "Aluselement" not "Alusmoodul" for parametric designs
- [ ] "Väljastatud element" not "Moodul" for released units  
- [ ] "Element" not "Module" when referring to manufactured units
- [ ] "Detail" not "Component" for parts
- [ ] "Elemendid" not "Moodulid" for script/output folders
- [ ] "Element" not "MooduliNimi" in Excel columns
- [ ] Folder paths match new structure

## Backward Compatibility Notes

During transition, the code should:
1. Accept both `MooduliNimi` and `Element` Excel columns
2. Log warnings when old terminology detected
3. Continue to work with existing Vault folder structures until manual migration

## Migration Sequence

```
1. Update UBIQUITOUS_LANGUAGE.md              ✓ (completed 2026-04-29)
2. Update BaseModuleLayoutLib.vb constants    ✓ (completed 2026-05-12)
3. Update StringsLib.vb UI strings            ✓ (completed 2026-05-12)
4. Rename Lib/ModuleReleaseLib.vb → Lib/ElementReleaseLib.vb  ✓ (completed 2026-05-12)
5. Update ExcelReaderLib.vb                   ✓ (completed 2026-05-12)
6. Update caller scripts (AddVbFile paths)    ✓ (completed 2026-05-12)
7. Rename Moodulid/ → Elemendid/              ✓ (completed 2026-05-12)
8. Rename Katsetused/Moodulid/ → Katsetused/Elemendid/  ✓ (completed 2026-05-12)
9. Update AGENTS.md                           ✓ (completed 2026-05-12)
10. Update historical docs                    ✓ (completed 2026-05-12)
```

## References

- Domain terminology: `docs/UBIQUITOUS_LANGUAGE.md`
- Project conventions: `AGENTS.md`
- Current element release plan: `docs/plans/2026-04-26-module-release-cycle.md`
- UI library plan: `docs/plans/2026-05-08-unified-ui-library.md`
