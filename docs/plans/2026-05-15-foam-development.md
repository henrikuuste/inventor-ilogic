# Foam Part Development (Flat Blank from Curved Geometry) - Implementation Plan

## Overview

Foam parts are modeled as curved 3D shapes (sweep of a closed profile along a guide curve) but manufactured as straight extrusions with angled end cuts. This plan implements an iLogic-driven system to automatically compute the flat manufacturing blank geometry from the curved design model, keep both representations in the same part file (multi-body), and integrate with the existing `Mõõdud` dimension system.

## Current State Analysis

### Existing Pinnalaotus (Unwrap) Workflow

The codebase already has a complete workflow for parts where `Unwrap` → `Thicken` produces the manufactured shape:

- **`UnwrapLib.vb`**: Detects UnwrapFeature, finds associated ThickenFeature, resolves measurement body, manages DVRs (`Pinnalaotus` for flat/manufactured, `Komponent` for curved/assembly)
- **`Mõõdud.vb`**: Dimension calculator with dedicated "Pinnalaotus" mode that measures the thickened body using oriented extents
- **`DimensionUpdateLib.vb`**: Self-contained auto-update handler embedded in the part's "Uuenda" rule; stores `BB_DimensionSource`, `BB_ThicknessAxis`, etc. as custom properties
- **`DocumentUpdateLib.vb`**: Registers code sections in the "Uuenda" rule with configurable triggers

### Why Unwrap Doesn't Work for This Case

Unwrap computes isometric flattening of a single surface. For foam parts:
- Edges perpendicular to the unwrap direction can remain curved
- Multiple unwraps (inner/outer surfaces) don't align because different radii produce different developed lengths
- Non-planar cross-sections (D-profile) have no single "flat" reference surface
- The manufactured shape isn't an unwrapped surface — it's a straight extrusion with end cuts

### What's Needed

A **new dimension source** ("Arendus" = development) that:
1. Measures curved edge arc lengths on the original body
2. Computes end cut angles from sweep tangent directions
3. Creates a flat blank body (straight extrude + angled cuts) in the same part
4. Provides T/W/L dimensions of the flat blank to the Mõõdud system
5. Uses DVRs or Model States to control visibility per context

### Key Discoveries

- `WorkFeatureLib.vb:526-543`: Already uses `CurveEvaluator` for edge parameterization (`GetParamExtents`, `GetPointAtParam`)
- `CurveEvaluator.GetLengthAtParam(minParam, maxParam, length)`: API method for measuring arc length — not yet used in codebase
- `SweepFeature.Path`: Should give access to the sweep path curve directly
- `DimensionUpdateLib.vb`: The self-contained handler pattern can be extended for a new "Arendus" dimension source
- `UnwrapLib.vb:35-48`: DVR constants and property names — the Arendus system should follow the same pattern

## What We're NOT Doing

- **Not replacing the Unwrap workflow** — existing Pinnalaotus parts continue to use Unwrap+Thicken
- **Not handling complex 3D sweep paths** in the first iteration — focusing on planar sweep paths (curve in one plane)
- **Not auto-creating flat blank geometry** in the first iteration — POCs validate the measurement approach first; flat body creation comes in Phase 2
- **Not supporting variable-profile sweeps** — profile must be constant along the sweep

## Desired End State

1. User has a foam part with a sweep feature (closed profile + guide curve)
2. Running a rule measures key edge lengths and end angles automatically
3. A flat blank body exists in the same part file, driven by measured parameters
4. `Mõõdud.vb` detects the "Arendus" dimension source and measures the flat blank body
5. DVRs control which body is visible (curved for assembly, flat for manufacturing)
6. The "Uuenda" rule auto-updates measurements when geometry changes

## Implementation Approach

**Phase 0**: Proof-of-concept scripts to validate key API assumptions
**Phase 1**: Core measurement library (`FoamDevelopmentLib.vb`)
**Phase 2**: Flat blank body creation + Mõõdud integration
**Phase 3**: User-facing rule + auto-update handler

---

## Phase 0: Proof-of-Concept Scripts

### Overview
Three independent POC scripts to validate API capabilities and measurement approaches before committing to the full implementation. Run these on actual foam parts with sweep features to identify issues early.

### POC Scripts

Scripts are in `Katsetused/Poroloon/`.

#### POC-1: `TestEdgeArcLength.vb` — Edge Arc Length Measurement

**Goal**: Verify `CurveEvaluator.GetLengthAtParam` works correctly in iLogic for measuring curved edge lengths.

**What it does**:
- Iterates all edges on all bodies of the active part
- For each edge: measures arc length, identifies geometry type (Line, Arc, BSpline, etc.), logs start/end vertices
- Groups edges by face membership to understand topology
- Compares arc length vs. chord length (straight-line distance between endpoints)

**Success criteria**:
- [ ] Arc lengths are non-zero and reasonable for curved edges
- [ ] Straight edges show arc length = chord length (within tolerance)
- [ ] API doesn't throw on any standard geometry type (Line, Arc, BSpline, Circle)
- [ ] Works on both simple (circular arc sweep) and complex (spline sweep) foam parts

**Key questions answered**:
- Does `GetLengthAtParam` work reliably in iLogic?
- What edge geometry types appear on typical foam sweep parts?
- Are edge lengths in internal units (cm)?

#### POC-2: `TestSweepPathAndEdgeClassification.vb` — Sweep Feature Access + Edge Classification

**Goal**: Verify we can access the SweepFeature's path, and develop an algorithm to classify edges as longitudinal (along sweep) vs. transverse (cross-section).

**What it does**:
- Finds SweepFeature(s) in the part
- Accesses `SweepFeature.Path` and tries to measure path length via its evaluator
- Identifies the start/end faces of the sweep (the two planar end faces)
- Classifies edges:
  - **Transverse**: edges that lie entirely on a start or end face
  - **Longitudinal**: edges that connect start face to end face (run along the sweep)
  - **Side**: edges on side faces only
- Measures all longitudinal edges and logs min/max/average arc lengths
- Identifies the "neutral axis length" (average or centroid-traced length)

**Success criteria**:
- [ ] Can access `SweepFeature.Path` and measure path length
- [ ] Can identify start/end faces reliably (planar faces at sweep extremes)
- [ ] Longitudinal edges are correctly classified (verified visually)
- [ ] Arc lengths of longitudinal edges show expected variation (outer > inner)
- [ ] Path length correlates with average of longitudinal edge lengths

**Key questions answered**:
- Can we access the sweep path curve directly?
- Is face-adjacency-based edge classification robust?
- What's the relationship between sweep path length and edge arc lengths?

#### POC-3: `TestEndCutAngles.vb` — End Tangent and Cut Angle Computation

**Goal**: Compute the angles needed for end cuts on a flat blank.

**What it does**:
- Uses the edge classification from POC-2 to find longitudinal edges
- For each longitudinal edge: gets tangent vector at start and end using `GetFirstDerivatives`
- Computes chord direction (straight line from start to end vertex)
- Computes the angle between the tangent at each end and the chord direction
- For each end face: computes the normal vector
- Computes the angle between the end face normal and the chord/tangent
- Logs all computed angles in degrees

**Success criteria**:
- [ ] Tangent vectors are non-zero and have reasonable directions
- [ ] Angles at start/end are consistent across longitudinal edges on the same face
- [ ] For a circular arc sweep: angles match expected geometry (90° - half arc angle)
- [ ] End face normals are consistent with tangent directions

**Key questions answered**:
- Can we reliably get tangent vectors at edge endpoints?
- Are the computed angles sufficient to define planar end cuts?
- Do different longitudinal edges produce consistent cut angles?

### Phase 0 Decision Points

After running POCs, decide:
1. **Edge classification strategy**: Face-adjacency vs. sweep feature access vs. user selection?
2. **Length calculation**: Use max longitudinal edge (outer radius), sweep path (neutral axis), or user-selectable?
3. **End cut geometry**: Simple planar cuts (single angle per end) sufficient, or need compound/ruled cuts?
4. **Does the SweepFeature.Path API actually work?** If not, fall back to pure edge topology.

---

## Phase 1: Core Measurement Library (`FoamDevelopmentLib.vb`)

### Overview
Encapsulate the proven measurement approach into a reusable library module.

### Changes Required

#### 1. New Library: `Lib/FoamDevelopmentLib.vb`

**Functions** (based on POC results):
- `MeasureEdgeArcLength(edge As Edge) As Double` — arc length via CurveEvaluator
- `ClassifyBodyEdges(body As SurfaceBody) As Dictionary` — longitudinal/transverse/side classification
- `GetLongitudinalEdgeLengths(body As SurfaceBody) As List(Of Double)` — all longitudinal arc lengths
- `GetMaxDevelopedLength(body As SurfaceBody) As Double` — longest longitudinal edge
- `GetEndCutAngles(body As SurfaceBody, ByRef startAngle As Double, ByRef endAngle As Double)` — cut angles at each end
- `GetEdgeTangentAtEnd(edge As Edge, atStart As Boolean) As UnitVector` — tangent direction

### Success Criteria
- [ ] Library compiles without errors
- [ ] All measurements match POC script results
- [ ] Works on parts identified during POC testing

---

## Phase 2: Flat Blank Body Creation + Mõõdud Integration

### Overview
Create the physical flat blank geometry in the same part file and integrate with the dimension system.

### Changes Required

#### 1. `Lib/FoamDevelopmentLib.vb` — Flat Blank Creation
- `CreateFlatBlank(partDoc, sweepBody, profileSketch)` — creates extrude + angled cuts as a new body
- Stores body name in custom property (same pattern as `BB_PinnalaotusSolidBodyName`)

#### 2. `Lib/UnwrapLib.vb` — New Dimension Source
- Add `DIMENSION_SOURCE_ARENDUS As String = "Arendus"`
- Add to DVR management (flat blank visible in "Pinnalaotus" DVR, hidden in "Komponent" DVR)

#### 3. `Lib/DimensionUpdateLib.vb` — Arendus Handler
- Extend `BuildDimensionUpdateCode()` to handle `BB_DimensionSource = "Arendus"`
- Measure the flat blank body (same oriented-extent logic as Pinnalaotus)

#### 4. `Mõõdud.vb` — UI Integration
- Add "Arendus" option to the Telg dropdown alongside existing Pinnalaotus/Lehtmetall/Normal
- Detect foam development parts and default to Arendus mode
- Show flat blank T/W/L in the dialog

### Success Criteria
- [ ] Flat blank body created in same part file
- [ ] DVRs correctly show/hide curved vs. flat body
- [ ] `Mõõdud.vb` displays correct T/W/L for flat blank
- [ ] "Uuenda" rule auto-updates dimensions when sweep geometry changes

---

## Phase 3: User-Facing Rule + Auto-Update

### Overview
User workflow rule and automatic re-measurement on geometry changes.

### Changes Required

#### 1. New Rule: `Poroloon arendus.vb`
- User runs on a foam part with sweep feature
- Measures edges, computes blank dimensions
- Creates flat blank body (if not present) or updates parameters
- Registers dimension handler with "Arendus" source
- Sets up DVRs

#### 2. `Lib/DocumentUpdateLib.vb` Integration
- Register "FoamDevelopment" handler section in "Uuenda" rule
- Triggers: `PartGeometryChange`, `UserParameterChange`, `ModelParameterChange`
- Re-runs measurement and updates blank geometry on triggers

### Success Criteria
- [ ] End-to-end workflow: curved foam part → run rule → flat blank + dimensions
- [ ] Parametric: change sweep parameters → blank auto-updates
- [ ] Works in assembly context (assembly shows curved, drawings can show flat)

---

## Testing Strategy

### POC Testing (Phase 0)
Run each POC script on at least:
1. A simple foam part (circular arc sweep, rectangular or D-profile)
2. A complex foam part (spline guide curve, irregular profile)
3. A part with multiple bodies (to verify body isolation)

### Manual Testing (Phase 1-3)
1. Create test foam part with known geometry (e.g., 90° circular arc sweep, 50mm radius)
2. Manually calculate expected developed length and end angles
3. Compare script output to manual calculation
4. Change sweep parameters and verify auto-update

### Integration Testing
1. Place foam part in assembly → verify curved body visible, flat hidden
2. Create drawing from part → verify flat body visible in Pinnalaotus DVR
3. Run `Mõõdud.vb` on the part → verify correct T/W/L for flat blank
4. Release part via element release → verify dimensions carry through

## Terminology

Per `docs/UBIQUITOUS_LANGUAGE.md`:
- **Poroloon** = Foam (material classification)
- **Arendus** = Development (the process of computing flat blank from curved geometry)
- **Detail** = Part
- **Aluselement** = Base element (parametric design containing masters)

## References

- Existing Pinnalaotus workflow: `Lib/UnwrapLib.vb`, `Lib/DimensionUpdateLib.vb`
- Dimension calculator: `Mõõdud.vb`
- DVR management: `UnwrapLib.DVR_NAME_PINNALAOTUS`, `UnwrapLib.DVR_NAME_KOMPONENT`
- Edge evaluator usage: `Lib/WorkFeatureLib.vb:526-543`
- Inventor API: `CurveEvaluator.GetLengthAtParam`, `CurveEvaluator.GetFirstDerivatives`
- Auto-update pattern: `Lib/DocumentUpdateLib.vb`
