# Element Release System - Failure Modes Analysis

## Date: 2026-05-13, Updated 2026-05-14

---

## BREAKTHROUGH FINDING (2026-05-14)

### What Actually Works
1. **Reference replacement works**: Parts correctly reference NEW masters (copies), not source masters
2. **Derived relationships work**: With links NOT broken, parts update when master params change
3. **Manual parameter change works**: Manually modifying master params → assemblies and parts update correctly

### Root Cause Identified
**We were modifying SOURCE masters, then trying to copy them.**
The correct approach: **Copy first, then modify COPIES.**

### Shared Parts Placement
Shared parts placed with "grounded at origin" may have wrong placement for some elements.
This is expected behavior - parts should use proper constraints that move with the rest.
**This is a design/user issue, not a script issue.**

---

## NEW ARCHITECTURE: Copy-First, Modify-Copies-Only

**IMPLEMENTED: 2026-05-14**

### Stage 1: Create File Copies (NO parameter changes)
```
For each variant:
  1a. Copy master(s) to target location (SaveAs with new GUID)
      - Set Part Number, Description properties
  
  1b. Copy all parts to target locations
      - SaveAs with new GUID
      - ReplaceReference: source master → copied master
      - Set properties
  
  1c. Copy assembly to target location
      - SaveAs with new GUID
      - Replace all component references: source parts → copied parts
      - Set properties
```

### Stage 2: Apply Parameters to COPIES Only
```
  2a. Open copied master(s)
      - Apply variant parameters
      - Update master
      - Save
  
  2b. Open copied assembly
      - Update (propagates through references to parts)
      - Save
```

### Stage 3: Create Drawings
```
  3. Copy drawings with reference replacement
```

**Key principles**:
- Source files are NEVER modified
- All parameter changes happen on COPIES
- Parts keep derived links (not broken) so they update from master
- Shared parts placement is a user responsibility (use proper constraints)

---

## The Two Failure Modes (Historical)

We keep oscillating between two distinct failure modes:

### Failure Mode A: Wrong Parameters/Geometry
- Parts are created with WRONG geometry (e.g., 700mm variant values when creating 1100mm variant)
- Even though params are applied and Update() is called, SaveAs produces parts with original/wrong geometry
- DimensionUpdate logs show wrong values (698mm, 976.4mm) during part creation

### Failure Mode B: Corrupt Assembly Placement  
- Parts have CORRECT geometry
- But assembly placement is wrong (e.g., 700mm element positions used for 1100mm element)
- Constraints don't resolve correctly after reference replacement

## What We Know For Certain (Verified Facts)

### 1. Fingerprinting Works Correctly
- The BuildElementMatrix/fingerprinting phase correctly identifies different geometries
- Parts ARE updated with correct params during fingerprinting
- Log shows distinct fingerprints for different laius values

### 2. Parameter Application Works
- `ApplyParameters()` correctly sets master params (log shows "Set laius: 700 mm -> 1100 mm")
- `Update()` on masters propagates changes
- DimensionUpdate logs during update show CORRECT values for some parts (1100mm range)

### 3. SaveAs Changes Document Identity
- After `partDoc.SaveAs(targetPath, False)`, the PartDocument object IS the target
- The original source file is no longer represented by that object
- The source file on disk is unchanged

### 4. Derived Parts Have Links
- Parts derived from masters have `DerivedPartComponent` links
- These links are broken AFTER SaveAs (log: "Broke DerivedPartComponent link: 00000.ipt")
- The geometry captured depends on master state at SaveAs time

### 5. Masters Can Be Closed Unexpectedly
- After creating parts, masters may not be open (log: "Master not open, skipping: 00000.ipt")
- RestoreMasterParameters only works on OPEN documents
- We added re-opening logic, which works (log: "Re-opened master: 00000.ipt")

### 6. Assembly Constraints Reference Geometry
- ComponentOccurrence.Transformation defines position
- After ReplaceReference, constraints try to resolve against new geometry
- Update2(True) recalculates positions

## Observed Timing/Sequence Issues

### Issue: DimensionUpdate Shows Wrong Values During Part Creation
From log (Processing Iste 110 / 1100mm variant):
```
Line 631: Updated source assembly
Line 633: Updated 15 parts  
Line 637: Creating variant-specific part: 01000.ipt
Line 644: Raw values (cm): T=0.3 W=30.2 L=69.8  <-- 698mm = WRONG!
Line 648: SaveAs with new GUID: 01000.ipt
```

The part being SaveAs'd has 700mm geometry even though we're in 1100mm variant processing.

### Issue: loadedParts Dictionary May Have Stale Data
- We load parts into `loadedParts` dictionary
- We update parts
- By the time we CreateStandalonePartFromDocument, the geometry is wrong
- Something happens between Update and CreateStandalone that loses the updated state

## Hypotheses (Unconfirmed)

### H1: SaveAs Triggers Re-Derivation
When we SaveAs a derived part, Inventor might re-read geometry from the master.
If master state has changed (e.g., through some operation), geometry would be wrong.

### H2: Document Objects Become Invalid
After SaveAs, the PartDocument object moves to target path.
Other documents in loadedParts might reference shared resources that get corrupted.

### H3: Transform Capture Affects Part State
Iterating through ComponentOccurrences to capture transforms might cause
Inventor to re-read parts from disk or refresh derived geometry.

### H4: Assembly Open Affects Derived Parts
When we have assembly + derived parts + master all open, operations on one
might affect the others through Inventor's internal reference tracking.

## What Has NOT Worked

### Approach 1: Apply Params → Update → SaveAs Parts
- Parts created with wrong geometry (Failure Mode A)

### Approach 2: Separate Pass 1 (Parts) and Pass 2 (Assembly)
- Pass 1: Apply params, update, SaveAs parts, restore params
- Pass 2: Open fresh, replace refs, update, SaveAs assembly
- Result: Either wrong geometry OR wrong placement depending on implementation

### Approach 3: Capture and Apply Occurrence Transforms
- Capture transforms when positions are correct (params applied)
- Apply transforms after reference replacement
- Result: Reverted to wrong geometry (Failure Mode A)

## Open Questions

1. Why does DimensionUpdate show wrong values AFTER we updated parts?
2. What operation between Update() and SaveAs() causes geometry to revert?
3. Is the master's in-memory state being used, or is Inventor reading from disk?
4. Does iterating occurrences for transform capture affect part documents?

## Current Test (2026-05-14)

**Test 2: Disable BreakAllExternalLinks**

Temporarily disabled breaking derived links in both:
- `CreateStandalonePart`
- `CreateStandalonePartFromDocument`

**Hypothesis**: Breaking links might be triggering geometry re-derivation or
causing parts to lose their updated state.

**What to observe**:
- Do parts have correct geometry now?
- If yes, the break operation was interfering with SaveAs
- Released parts will still reference source masters (not standalone)

---

## Previous Test (2026-05-13)

**Hypothesis**: Iterating through ComponentOccurrences to capture transforms causes 
Inventor to refresh/reload part documents, losing their updated geometry.

**Test**: Removed transform capture code. If parts now have correct geometry, 
this confirms the hypothesis.

**If parts are correct but assembly is wrong**: We need a different approach 
for capturing positions. Options:
- Apply params during assembly pass (but restore immediately after SaveAs)
- Store occurrence transforms in a file/external storage
- Different timing for when we capture transforms

## Next Steps to Investigate

1. ✓ Remove transform capture to test if that's the culprit
2. If geometry is correct, find alternative way to capture/apply positions
3. If still wrong, add logging IMMEDIATELY before SaveAs to verify part geometry
4. Consider completely different approach: File.Copy + Open + Update instead of in-memory SaveAs

## Key Insight

The problem seems to be that certain Inventor API operations (like iterating 
occurrences or accessing occurrence.Definition.Document) cause Inventor to 
refresh documents from disk or re-derive geometry, losing in-memory updates.

**Rule**: Minimize API calls between Update() and SaveAs(). Don't iterate 
through assembly structure after updating if you need to preserve part geometry.

## Drawing Numbering (2026-05-14)

**Decision**: Drawings do NOT get their own unique file numbers. Instead:

1. Drawing filename = [referenced model's number] + [original suffix] + .idw
2. Examples:
   - `00003.idw` → `01234.idw` (when model `00003.iam` becomes `01234.iam`)
   - `00005_sheet2.idw` → `01235_sheet2.idw` (preserves suffix)
3. Drawing's Part Number iProperty = the new filename (without .idw)

**Implementation**:
- In `ComputeReleasePlan`: Drawings look up their referenced model's PlannedFile and use its VaultNumber
- In `AllocateRealNumbers`: After parts/assemblies get real numbers, drawings are updated using a placeholder→real number map
- `CalculateRequiredNumbers` no longer counts drawings (they don't need reserved numbers)
- `IsPlaceholder = False` for drawings since they derive their number

**Rationale**: Drawings are tied 1:1 to their referenced models. Using the same number makes it easier to find related files and follows our existing naming conventions.
