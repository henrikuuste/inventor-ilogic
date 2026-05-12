<!-- Copyright (c) 2026 Henri Kuuste -->
---
date: 2026-04-25T08:46:00+03:00
researcher: Claude Opus 4.5
git_commit: 60d1a9eb68274d18738495914b4154317a979fcc
branch: main
repository: Inventor-Rules
topic: "How Thickness/Width/Length dimensions are handled throughout the codebase"
tags: [research, codebase, dimensions, sheet-metal, bounding-box, iproperties]
status: complete
last_updated: 2026-04-25
---

# Research: Thickness/Width/Length Dimension Handling

**Date**: Saturday, April 25, 2026, 08:46 AM (UTC+3)
**Git Commit**: 60d1a9eb68274d18738495914b4154317a979fcc
**Branch**: main

## Research Question

How are dimensions and assigning properties for Thickness/Width/Length handled throughout the codebase? Both for sheet metal parts as well as normal parts. Where is the code? What is duplicated? What is different?

## Summary

The codebase handles Thickness/Width/Length dimensions through two main pathways:

1. **Sheet Metal Parts**: Use the Inventor sheet metal API (`SheetMetalComponentDefinition.Thickness`) and formula-linked iProperties for Width/Length
2. **Normal Parts**: Use bounding box calculations (axis-aligned or geometry-based oriented box) with text-based iProperties

**Central entry point**: `CustomPropertiesLib.ValidateAndFixDimensionProperties()` handles both paths based on part subtype detection.

**Property names**: English (`Thickness`, `Width`, `Length`) for iProperty keys; Estonian (`Paksus`, `Laius`, `Pikkus`) only in UI labels.

**Key libraries**:
- `CustomPropertiesLib.vb` - Central property assignment
- `BoundingBoxStockLib.vb` - Normal part dimension detection and rule generation
- `SheetMetalLib.vb` - Sheet metal conversion and thickness detection
- `MakeComponentsLib.vb` - Multi-body part dimension handling

## Detailed Findings

### CustomPropertiesLib - Central Property Management

**File**: `Lib/CustomPropertiesLib.vb`

This is the single source of truth for dimension property names and formatting.

**Constants (lines 6-11)**:
```vb
Public Const PROP_THICKNESS As String = "Thickness"
Public Const PROP_WIDTH As String = "Width"
Public Const PROP_LENGTH As String = "Length"
Public Const SHEET_METAL_GUID As String = "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}"
Public Const SHEET_METAL_WIDTH_FORMULA As String = "=<Sheet Metal Width>"
Public Const SHEET_METAL_LENGTH_FORMULA As String = "=<Sheet Metal Length>"
```

**Key Functions**:

| Function | Lines | Purpose |
|----------|-------|---------|
| `ValidateAndFixDimensionProperties` | 72-108 | Main entry point - branches on sheet metal vs normal |
| `SetNormalPartDimensionProperties` | 48-51 | Sets T/W/L as text properties (integer mm) |
| `SetSheetMetalDimensionProperties` | 54-57 | Sets W/L as formula strings |
| `EnsureSheetMetalThicknessExport` | 59-68 | Exposes thickness parameter as number property |
| `FormatDimensionText` | 27-30 | Converts cm to integer mm string |
| `IsSheetMetalPart` | 32-38 | Detects sheet metal via SubType GUID |

**Formatting Logic**:
- Input: values in **cm** (Inventor internal units)
- Output: integer **mm** string with no unit suffix
- Rounding: `Math.Round(..., 0, MidpointRounding.AwayFromZero)`

### Sheet Metal Parts

#### How Thickness is Obtained

**Method 1: Direct API read** (`SheetMetalComponentDefinition.Thickness`)

Used when reading existing sheet metal properties:
- `Lib/CAMDrawingLib.vb:247-248` - For flat pattern view layout
- `Lehtmetall.vb:240-241` - When comparing to measured geometry

**Method 2: Geometry measurement** (find smallest extent along face normals)

Used when detecting thickness direction:
- `Lib/SheetMetalLib.vb:30-78` - `DetectThicknessVector()` function
- `Lehtmetall.vb:120-165` - `MeasureThicknessAlongNormal()` function

**Algorithm** (smallest extent detection):
1. Iterate all faces, extract unique planar face normals
2. For each normal, project all body vertices to calculate extent
3. Normal with smallest extent = thickness direction

#### How Thickness is Written

**Setting the model parameter**:
- `Lib/SheetMetalLib.vb:193-201` - `SetThickness()`
- `Lehtmetall.vb:176-183` - `SetMeasuredThickness()`

Both use:
```vb
smCompDef.UseSheetMetalStyleThickness = False
smCompDef.Thickness.Value = thickness  ' in cm
```

**Exposing as iProperty** (`EnsureSheetMetalThicknessExport`):
```vb
thicknessParam.ExposedAsProperty = True
thicknessParam.CustomPropertyFormat.PropertyType = kNumberPropertyType
thicknessParam.CustomPropertyFormat.ShowUnitsString = False
thicknessParam.CustomPropertyFormat.Units = "mm"
```

#### How Width/Length are Set for Sheet Metal

Width and Length use formula strings that link to built-in sheet metal properties:
- `Width` = `"=<Sheet Metal Width>"`
- `Length` = `"=<Sheet Metal Length>"`

These are evaluated at runtime by Inventor.

### Normal Parts (Bounding Box)

#### How Dimensions are Calculated

**File**: `Lib/BoundingBoxStockLib.vb`

**Method 1: Axis-aligned bounding box**

```vb
Public Sub GetBoundingBoxSizes(partDoc As PartDocument, ByRef xSize, ySize, zSize)
    Dim rangebox As Box = partDoc.ComponentDefinition.RangeBox
    xSize = rangebox.MaxPoint.X - rangebox.MinPoint.X
    ySize = rangebox.MaxPoint.Y - rangebox.MinPoint.Y
    zSize = rangebox.MaxPoint.Z - rangebox.MinPoint.Z
End Sub
```

**Method 2: Oriented bounding box** (for rotated parts)

```vb
Public Function GetOrientedExtent(partDoc As PartDocument, dirX, dirY, dirZ) As Double
    ' Projects all vertices onto direction vector
    ' Returns maxProj - minProj
End Function
```

#### Dimension Sorting Convention

**Rule**: Smallest = Thickness, Middle = Width, Largest = Length

**Axis-aligned** (`AutoDetectAxes`, lines 219-248):
- Bubble sort X/Y/Z sizes ascending
- `axes(0)` = Thickness, `axes(1)` = Width, `axes(2)` = Length

**Geometry-based** (`AutoDetectAxesFromGeometry`, lines 255+):
1. Find face normal with smallest extent → Thickness direction
2. Compute two perpendicular directions
3. Compare their extents: larger = Length, smaller = Width

**Given a fixed thickness axis** (`AssignWidthLength`, lines 184-209):
```vb
If size1 >= size2 Then
    lengthAxis = axis1
    widthAxis = axis2
Else
    lengthAxis = axis2
    widthAxis = axis1
End If
```

#### How Values are Assigned to iProperties

The generated rule `"Uuenda mõõdud"` (built by `BuildRuleText()`) calls:
```vb
CustomPropertiesLib.ValidateAndFixDimensionProperties(partDoc, thicknessVal, widthVal, lengthVal)
```

**Axis configuration stored**:
- `BB_ThicknessAxis` - e.g., `"X"`, `"Y"`, `"Z"`, or `"V:0.707,0.707,0"`
- `BB_WidthAxis` - same format
- Length axis derived from cross product or remaining axis

### MakeComponentsLib - Multi-Body Parts

**File**: `Lib/MakeComponentsLib.vb`

Handles dimension detection per-body (for multi-body master documents).

**Key Functions**:

| Function | Lines | Purpose |
|----------|-------|---------|
| `DetectAxesForBody` | 693-770 | Same smallest-extent algorithm but per-body |
| `GetOrientedExtentForBody` | 1372-1388 | Extent calculation for single body |
| `SetDimensionProperties` | 1071-1087 | Delegates to `CustomPropertiesLib` |
| `ReadPropertiesFromPart` | 529-558 | Reads T/W/L from existing parts (mm → cm) |

**BodyInfo structure** stores per-body:
- `ThicknessVector`, `ThicknessValue`
- `WidthVector`, `WidthValue`
- `LengthVector`, `LengthValue`

### Support Parts (Different Semantics)

**File**: `Lib/SupportPlacementLib.vb`

Uses same property **names** but different **semantics**:
- `Thickness` = fixed beam thickness (constant)
- `Width` = beam width (from model parameter `Width`)
- `Height` = beam height (from model parameter `Length`, confusingly)

Formatting uses `UtilsLib.FormatDimensionMm()` - 3 decimal places with " mm" suffix.

## Code References

### Sheet Metal Path
- `Lib/CustomPropertiesLib.vb:59-68` - `EnsureSheetMetalThicknessExport`
- `Lib/CustomPropertiesLib.vb:54-57` - `SetSheetMetalDimensionProperties`
- `Lib/SheetMetalLib.vb:30-78` - `DetectThicknessVector`
- `Lib/SheetMetalLib.vb:193-201` - `SetThickness`
- `Lehtmetall.vb:120-165` - `MeasureThicknessAlongNormal`
- `Lehtmetall.vb:176-183` - `SetMeasuredThickness`

### Normal Parts Path
- `Lib/BoundingBoxStockLib.vb:57-62` - `GetBoundingBoxSizes`
- `Lib/BoundingBoxStockLib.vb:219-248` - `AutoDetectAxes`
- `Lib/BoundingBoxStockLib.vb:255-340` - `AutoDetectAxesFromGeometry`
- `Lib/BoundingBoxStockLib.vb:617-634` - `GetOrientedExtent`
- `Lib/BoundingBoxStockLib.vb:918-950` - `CreateOrUpdateRule`
- `Lib/BoundingBoxStockLib.vb:971-1045` - `BuildRuleText`

### Central Property Logic
- `Lib/CustomPropertiesLib.vb:72-108` - `ValidateAndFixDimensionProperties`
- `Lib/CustomPropertiesLib.vb:48-51` - `SetNormalPartDimensionProperties`
- `Lib/CustomPropertiesLib.vb:27-30` - `FormatDimensionText`

### Multi-Body Parts
- `Lib/MakeComponentsLib.vb:693-770` - `DetectAxesForBody`
- `Lib/MakeComponentsLib.vb:1372-1388` - `GetOrientedExtentForBody`
- `Lib/MakeComponentsLib.vb:1071-1087` - `SetDimensionProperties`

## Duplicated Code Patterns

### 1. Smallest Extent Along Face Normals (Thickness Detection)

**Same algorithm appears in**:

| File | Function |
|------|----------|
| `Lib/BoundingBoxStockLib.vb:255-340` | `AutoDetectAxesFromGeometry` |
| `Lib/SheetMetalLib.vb:30-78` | `DetectThicknessVector` |
| `Lib/MakeComponentsLib.vb:693-770` | `DetectAxesForBody` |
| `Katsetused/TestDerivedPart.vb` | `DetectThicknessVector` (local copy) |

**Pattern**:
1. Iterate faces, extract unique planar normals
2. Canonicalize normal direction (positive hemisphere)
3. Calculate extent along each normal
4. Smallest extent = thickness

### 2. Oriented Extent Calculation

**Duplicated in**:
- `Lib/BoundingBoxStockLib.vb:617-634` - `GetOrientedExtent` (whole part)
- `Lib/MakeComponentsLib.vb:1372-1388` - `GetOrientedExtentForBody` (single body)

**Same vertex projection logic**:
```vb
For Each vertex In body.Vertices
    proj = pt.X * dirX + pt.Y * dirY + pt.Z * dirZ
    If proj < minProj Then minProj = proj
    If proj > maxProj Then maxProj = proj
Next
Return maxProj - minProj
```

### 3. Perpendicular Vector Computation

**Duplicated in**:
- `Lib/BoundingBoxStockLib.vb` - `ComputePerpendicularVectors`
- `Lib/MakeComponentsLib.vb` - `ComputePerpendicularVectors` (private)

Both generate orthogonal basis from thickness direction.

### 4. Width/Length Swap Pattern

**Same comparison logic** (`lengthExtent >= widthExtent`):
- `Lib/BoundingBoxStockLib.vb:101-108` - In `RunConfigLoop`
- `Lib/BoundingBoxStockLib.vb:202-209` - In `AssignWidthLength`
- `Lib/MakeComponentsLib.vb:755-765` - In `DetectAxesForBody`
- `Mõõdud.vb` - In `CollectPartData`

### 5. Sheet Metal Conversion Flow

`Lehtmetall.vb` and `SheetMetalLib.vb` have parallel implementations:
- Comment in `SheetMetalLib.vb:10`: *"Based on Lehtmetall.vb but with automated A-side detection"*
- Both call `CustomPropertiesLib.ValidateAndFixDimensionProperties`
- Both set thickness via `smCompDef.Thickness.Value`

## Key Differences Between Approaches

### Sheet Metal vs Normal Parts

| Aspect | Sheet Metal | Normal Parts |
|--------|-------------|--------------|
| **Thickness source** | `SheetMetalComponentDefinition.Thickness` or geometry measurement | Bounding box smallest dimension |
| **Thickness iProperty** | Exposed parameter (number, mm) | Text property (integer mm string) |
| **Width/Length** | Formula strings linking to built-in | Text properties from bounding box |
| **Update mechanism** | Parameter change triggers formula | Generated rule `"Uuenda mõõdud"` |

### Axis-Aligned vs Geometry-Based Detection

| Aspect | Axis-Aligned | Geometry-Based |
|--------|--------------|----------------|
| **Use case** | Parts aligned to coordinate axes | Rotated/tilted parts |
| **Thickness** | Smallest of X/Y/Z sizes | Smallest extent along face normals |
| **Width/Length** | From remaining two axes | From perpendicular directions |
| **Stored as** | `"X"`, `"Y"`, `"Z"` | `"V:x,y,z"` vector format |

### Per-Part vs Per-Body

| Aspect | BoundingBoxStockLib | MakeComponentsLib |
|--------|---------------------|-------------------|
| **Scope** | Whole part document | Individual body |
| **Use case** | Standalone parts | Multi-body master documents |
| **Storage** | `BB_ThicknessAxis` on part | `LK_Body_*_TAxis` parameters |

## Architecture Documentation

### Data Flow

```
┌─────────────────────┐
│   User/Rule Call    │
└──────────┬──────────┘
           │
           ▼
┌─────────────────────────────────────────────────┐
│     CustomPropertiesLib.ValidateAndFix...       │
│  ┌───────────────┬───────────────────────────┐  │
│  │ IsSheetMetal? │                           │  │
│  └───────┬───────┘                           │  │
│          │                                   │  │
│    ┌─────┴─────┐                             │  │
│    ▼           ▼                             │  │
│  Sheet      Normal                           │  │
│  Metal      Part                             │  │
└────┬───────────┬────────────────────────────-┘  │
     │           │                                │
     ▼           ▼                                │
┌─────────┐  ┌──────────────────────┐             │
│Expose   │  │SetNormalPart...      │             │
│Thickness│  │(T/W/L text props)    │             │
│Param    │  └──────────────────────┘             │
│+ W/L    │                                       │
│formulas │                                       │
└─────────┘                                       │
```

### Generated Rule Architecture

`BoundingBoxStockLib.CreateOrUpdateRule()` generates an embedded iLogic rule that:

1. Reads axis configuration from `BB_ThicknessAxis`, `BB_WidthAxis`
2. Calculates sizes (axis-aligned or oriented)
3. Applies optional overrides (`ThicknessOverride`, `WidthOverride`, `LengthOverride` parameters)
4. Calls `CustomPropertiesLib.ValidateAndFixDimensionProperties()`

This rule auto-triggers on parameter changes.

## Related Research

- `docs/research/2026-04-25-loo-komponendid-failures.md` - Sheet metal conversion issues

## Open Questions

1. **Could thickness detection be consolidated?** The same algorithm exists in 4 places with minor variations.
2. **Should `GetOrientedExtent` vs `GetOrientedExtentForBody` be unified?** Only difference is scope (all bodies vs one body).
3. **Is the test file duplication intentional?** `TestDerivedPart.vb` has a local copy of `DetectThicknessVector`.

---

## Unified Approach (Implemented 2026-04-25)

### New Architecture: DimensionUpdateLib

The dimension handling has been unified using a new `DimensionUpdateLib` module that:

1. **Generates self-contained code** - No `AddVbFile` or `AddReference` in the generated rule code
2. **Uses DocumentUpdateLib** - Registers a "Dimensions" section in the centralized "Uuenda" rule
3. **Works on any computer** - Parts can be opened on machines without the library files

### Key Changes

| Script | Before | After |
|--------|--------|-------|
| `Mõõdud.vb` | `BoundingBoxStockLib.CreateOrUpdateRule()` | `DimensionUpdateLib.RegisterDimensionHandler()` |
| `Lehtmetall.vb` | No auto-update rule | `DimensionUpdateLib.RegisterDimensionHandler()` |
| `Loo detailid.vb` | `BoundingBoxStockLib.CreateOrUpdateRule()` | `DimensionUpdateLib.RegisterDimensionHandler()` |

### Generated Code Structure

The `DimensionUpdateLib.BuildDimensionUpdateCode()` generates inline VB code that:

1. Detects if part is sheet metal (via SubType GUID)
2. For sheet metal:
   - Exports Thickness parameter as iProperty (mm, no units string)
   - Sets Width/Length to `=<Sheet Metal Width>` / `=<Sheet Metal Length>` formulas
3. For normal parts:
   - Reads axis config from `BB_ThicknessAxis`, `BB_WidthAxis`
   - Supports both axis-aligned (X/Y/Z) and oriented (V:x,y,z) formats
   - Calculates sizes via bounding box or vertex projection
   - Applies optional overrides (ThicknessOverride, WidthOverride, LengthOverride)
   - Sets properties as integer mm text

### Migration

- Legacy "Uuenda mõõdud" rules are automatically removed when `RegisterDimensionHandler()` is called
- `BoundingBoxStockLib.CreateOrUpdateRule()` and `BuildRuleText()` are now deprecated
- Existing parts will be migrated on next dimension setup run

### Triggers

The dimension handler is registered with these triggers:
- `PartGeometryChange` - Updates when part geometry changes
- `UserParameterChange` - Updates when user parameters change
