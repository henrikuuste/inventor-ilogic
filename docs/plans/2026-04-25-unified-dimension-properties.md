<!-- Copyright (c) 2026 Henri Kuuste -->
# Unified Dimension Properties Implementation Plan

## Overview

Unify the handling of Thickness/Width/Length dimension properties across all scripts using a self-contained update handler registered via DocumentUpdateLib. The generated code must work on any computer without external library dependencies.

## Current State Analysis

### Existing Approaches

| Script | Part Type | Approach | Issues |
|--------|-----------|----------|--------|
| `Mõõdud.vb` | Normal parts | Creates "Uuenda mõõdud" rule with `AddVbFile "Lib/CustomPropertiesLib.vb"` | Depends on external library |
| `Lehtmetall.vb` | Sheet metal | Calls `CustomPropertiesLib.ValidateAndFixDimensionProperties()` | No auto-update rule created |
| `Loo komponendid.vb` | Both | Calls `BoundingBoxStockLib.CreateOrUpdateRule()` | Same dependency issue |
| `BoundingBoxStockLib.BuildRuleText()` | Normal parts | Generates standalone rule text | Requires `AddVbFile` |

### Key Discoveries

1. **Dimension calculation paths**:
   - **Sheet Metal**: Thickness from `SheetMetalComponentDefinition.Thickness`, Width/Length from `=<Sheet Metal Width>` and `=<Sheet Metal Length>` formulas
   - **Normal Parts**: Bounding box calculation (axis-aligned or oriented via vertex projection)

2. **Property storage**:
   - Axis configuration stored in: `BB_ThicknessAxis`, `BB_WidthAxis` custom properties
   - Values currently stored as text iProperties with integer mm values

3. **DocumentUpdateLib pattern** (from `Lib/DocumentUpdateLib.vb:116-185`):
   - Registers UID-guarded sections in centralized "Uuenda" rule
   - Supports multiple triggers (parameter change, geometry change, etc.)
   - Code in sections must be self-contained

4. **Legacy rules to migrate**:
   - `"Uuenda mõõdud"` - old standalone rule from BoundingBoxStockLib

## Desired End State

1. **Single update handler** registered as "Dimensions" section in "Uuenda" rule via DocumentUpdateLib
2. **Self-contained code** - no `AddVbFile`, `AddReference`, or external dependencies
3. **Works for both part types**:
   - Sheet metal: reads flat pattern dimensions
   - Normal parts: calculates from bounding box/oriented box
4. **Number parameters** - `Thickness`, `Width`, `Length` as user parameters exposed as iProperties
5. **Consistent formatting** - integer mm on output, no unit suffix
6. **Migration support** - removes legacy "Uuenda mõõdud" rule when present

## What We're NOT Doing

- Not changing how axis detection works (keep existing geometry-based detection)
- Not modifying the UI in Mõõdud.vb (keep the DataGridView dialog)
- Not changing how sheet metal conversion works in SheetMetalLib
- Not modifying the multi-body detection in MakeComponentsLib
- Not adding new user-facing parameters beyond the existing T/W/L

## Implementation Approach

Create a new library function `DimensionUpdateLib.RegisterDimensionHandler()` that:
1. Generates self-contained VB code for dimension calculation
2. Registers it via DocumentUpdateLib with appropriate triggers
3. Removes legacy "Uuenda mõõdud" rule if present
4. Creates/updates Number parameters for T/W/L

---

## Phase 1: Create DimensionUpdateLib

### Overview
Create a new library module that generates self-contained dimension update code and registers it via DocumentUpdateLib.

### Changes Required

#### 1. Create new library file
**File**: `Lib/DimensionUpdateLib.vb`
**Purpose**: Generate self-contained dimension update code and register with DocumentUpdateLib

**Key Functions**:
- `RegisterDimensionHandler(doc, iLogicAuto, thicknessAxis, widthAxis, lengthAxis)` - Main entry point
- `BuildDimensionUpdateCode(thicknessAxis, widthAxis, lengthAxis)` - Generates self-contained VB code
- `RemoveLegacyRule(doc, iLogicAuto)` - Removes old "Uuenda mõõdud" rule
- `EnsureDimensionParameters(partDoc)` - Creates T/W/L Number parameters

**Code structure**:
```vb
' DimensionUpdateLib - Register self-contained dimension update handler
' Usage: AddVbFile "Lib/DimensionUpdateLib.vb"
'
' The generated update code is SELF-CONTAINED and does not depend on any
' external library files. It can run on any computer.

Imports Inventor

Public Module DimensionUpdateLib

    Public Const HANDLER_UID As String = "Dimensions"
    Private Const LEGACY_RULE_NAME As String = "Uuenda mõõdud"
    
    ' Register dimension handler and remove legacy rule
    Public Function RegisterDimensionHandler(...) As Boolean
    
    ' Build self-contained VB code for dimension updates
    Private Function BuildDimensionUpdateCode(...) As String()
    
    ' Ensure Number parameters exist for T/W/L
    Private Sub EnsureDimensionParameters(...)
    
    ' Remove legacy rule if present
    Private Sub RemoveLegacyRule(...)

End Module
```

### Success Criteria

#### Automated Verification:
- [x] `DimensionUpdateLib.vb` compiles without errors (test by including in a rule)
- [x] No `AddVbFile` or `AddReference` in generated code
- [x] Generated code handles both sheet metal and normal parts

#### Manual Verification:
- [ ] Open a normal part, run dimension setup, verify "Uuenda" rule contains "Dimensions" section
- [ ] Verify old "Uuenda mõõdud" rule is removed
- [ ] Verify dimension parameters exist and update on geometry change

**Implementation Note**: After completing this phase and all automated verification passes, pause here for manual confirmation before proceeding to the next phase.

---

## Phase 2: Update Mõõdud.vb to Use New Library

### Overview
Modify Mõõdud.vb to use DimensionUpdateLib instead of BoundingBoxStockLib.CreateOrUpdateRule.

### Changes Required

#### 1. Update imports and dependencies
**File**: `Mõõdud.vb`
**Changes**:
- Add `AddVbFile "Lib/DimensionUpdateLib.vb"`
- Add `AddVbFile "Lib/DocumentUpdateLib.vb"`
- Keep `AddVbFile "Lib/BoundingBoxStockLib.vb"` for axis detection/UI

#### 2. Replace rule creation call
**File**: `Mõõdud.vb` (line ~169)
**Current**:
```vb
BoundingBoxStockLib.CreateOrUpdateRule(partDoc, thicknessAxes(i), widthAxes(i), lengthAxes(i), iLogicVb.Automation)
```
**New**:
```vb
DimensionUpdateLib.RegisterDimensionHandler(partDoc, iLogicVb.Automation, thicknessAxes(i), widthAxes(i), lengthAxes(i))
```

### Success Criteria

#### Automated Verification:
- [x] Mõõdud.vb compiles without errors
- [x] No direct calls to `BoundingBoxStockLib.CreateOrUpdateRule` remain

#### Manual Verification:
- [ ] Run Mõõdud on a normal part, verify dimensions update correctly
- [ ] Change part geometry, verify T/W/L update automatically
- [ ] Verify no "Uuenda mõõdud" rule exists (only "Uuenda" with "Dimensions" section)

**Implementation Note**: After completing this phase and all automated verification passes, pause here for manual confirmation before proceeding to the next phase.

---

## Phase 3: Update Lehtmetall.vb for Unified Flow

### Overview
Modify Lehtmetall.vb to register the same dimension handler, ensuring sheet metal parts also get auto-updating dimensions.

### Changes Required

#### 1. Add library dependencies
**File**: `Lehtmetall.vb`
**Changes**:
- Add `AddVbFile "Lib/DocumentUpdateLib.vb"`
- Add `AddVbFile "Lib/DimensionUpdateLib.vb"`

#### 2. Register dimension handler after conversion
**File**: `Lehtmetall.vb` (after line ~96)
**Changes**: After setting sheet metal properties, register the handler:
```vb
' Register dimension update handler (self-contained, works on any computer)
DimensionUpdateLib.RegisterDimensionHandler(partDoc, iLogicVb.Automation, "", "", "")
```

Note: For sheet metal, axis parameters are empty since dimensions come from flat pattern formulas.

#### 3. Update ValidateAndRepairExistingSheetMetal
**File**: `Lehtmetall.vb` (in `ValidateAndRepairExistingSheetMetal` sub)
**Changes**: Also register handler when validating existing sheet metal

### Success Criteria

#### Automated Verification:
- [x] Lehtmetall.vb compiles without errors
- [x] DimensionUpdateLib.RegisterDimensionHandler is called after conversion

#### Manual Verification:
- [ ] Convert a normal part to sheet metal via Lehtmetall
- [ ] Verify "Uuenda" rule has "Dimensions" section
- [ ] Verify T/W/L properties update when flat pattern changes

**Implementation Note**: After completing this phase and all automated verification passes, pause here for manual confirmation before proceeding to the next phase.

---

## Phase 4: Update Loo komponendid.vb

### Overview
Modify Loo komponendid.vb to use the unified dimension handler for created parts.

### Changes Required

#### 1. Add library dependency
**File**: `Loo komponendid.vb`
**Changes**:
- Add `AddVbFile "Lib/DocumentUpdateLib.vb"`
- Add `AddVbFile "Lib/DimensionUpdateLib.vb"`

#### 2. Replace rule creation in part processing loop
**File**: `Loo komponendid.vb` (lines ~378-392)
**Current**:
```vb
BoundingBoxStockLib.CreateOrUpdateRule(newPart, bi.ThicknessVector, bi.WidthVector, bi.LengthVector, iLogicVb.Automation)
```
**New**:
```vb
DimensionUpdateLib.RegisterDimensionHandler(newPart, iLogicVb.Automation, bi.ThicknessVector, bi.WidthVector, bi.LengthVector)
```

### Success Criteria

#### Automated Verification:
- [x] Loo komponendid.vb compiles without errors

#### Manual Verification:
- [ ] Run Loo komponendid on a multi-body part
- [ ] Verify created parts have "Uuenda" rule with "Dimensions" section
- [ ] Verify both normal parts and sheet metal converted parts work

**Implementation Note**: After completing this phase and all automated verification passes, pause here for manual confirmation before proceeding to the next phase.

---

## Phase 5: Clean Up and Unify Code

### Overview
Remove duplicated code, deprecate old patterns, and ensure consistency across the codebase.

### Changes Required

#### 1. Deprecate BoundingBoxStockLib.CreateOrUpdateRule
**File**: `Lib/BoundingBoxStockLib.vb`
**Changes**:
- Add deprecation comment to `CreateOrUpdateRule`
- Add comment pointing to `DimensionUpdateLib.RegisterDimensionHandler`
- Keep function for backward compatibility but add warning log

#### 2. Deprecate BoundingBoxStockLib.BuildRuleText
**File**: `Lib/BoundingBoxStockLib.vb`
**Changes**:
- Add deprecation comment
- Function remains for reference but is no longer called

#### 3. Update CustomPropertiesLib
**File**: `Lib/CustomPropertiesLib.vb`
**Changes**:
- Ensure `ValidateAndFixDimensionProperties` works with Number parameters
- Add support for reading parameter values (not just text properties)

#### 4. Clean up MakeComponentsLib
**File**: `Lib/MakeComponentsLib.vb`
**Changes**:
- Update `SetDimensionProperties` to use the new parameter-based approach if needed
- Ensure consistency with DimensionUpdateLib

### Success Criteria

#### Automated Verification:
- [x] All modified files compile without errors
- [x] No warnings from deprecated function usage in active code paths

#### Manual Verification:
- [ ] Run all three scripts (Mõõdud, Lehtmetall, Loo komponendid)
- [ ] Verify consistent behavior across all scripts
- [ ] Verify migration works on parts with old "Uuenda mõõdud" rule

**Implementation Note**: After completing this phase and all automated verification passes, pause here for manual confirmation before proceeding to the next phase.

---

## Phase 6: Testing and Documentation

### Overview
Comprehensive testing across all scenarios and documentation updates.

### Test Cases

#### Normal Part Flow
1. New part → run Mõõdud → verify dimensions set and auto-update
2. Part with old "Uuenda mõõdud" rule → run Mõõdud → verify migration
3. Change geometry → verify T/W/L update automatically

#### Sheet Metal Flow
1. New solid → run Lehtmetall → verify conversion and dimensions
2. Existing sheet metal → run Lehtmetall → verify validation/repair
3. Modify flat pattern → verify dimensions update

#### Multi-Body Flow
1. Multi-body part → run Loo komponendid → verify all created parts have handler
2. Mix of sheet metal and normal bodies → verify correct handling for each
3. Re-run on same master → verify no duplicate handlers

#### Portability Test
1. Copy part file to computer without library files
2. Modify geometry
3. Verify "Uuenda" rule runs successfully (self-contained code)

### Documentation Updates

#### Update research document
**File**: `docs/research/2026-04-25-dimensions-thickness-width-length.md`
**Changes**: Add section documenting the new unified approach

### Success Criteria

#### Manual Verification:
- [ ] All test cases pass
- [x] Documentation updated with new patterns

---

## Testing Strategy

### Unit Tests (Manual)
- Verify generated code syntax by running on isolated test parts
- Verify parameter creation/update logic
- Verify legacy rule removal

### Integration Tests (Manual)
- Full workflow for each script
- Cross-script consistency (same part processed by different scripts)
- Migration from old format to new format

### Regression Tests (Manual)
- Existing parts with old rules continue to work
- No breaking changes to axis detection UI
- Material assignment still works in Loo komponendid

## References

- Research document: `docs/research/2026-04-25-dimensions-thickness-width-length.md`
- DocumentUpdateLib pattern: `Lib/DocumentUpdateLib.vb:116-185`
- Current rule generation: `Lib/BoundingBoxStockLib.vb:971-1148`
- Sheet metal handling: `Lib/CustomPropertiesLib.vb:54-68`
