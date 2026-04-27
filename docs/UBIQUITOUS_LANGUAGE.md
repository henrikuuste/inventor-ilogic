<!-- Copyright (c) 2026 Henri Kuuste -->
# Ubiquitous Language

Domain terminology for the Inventor Moodulid (Module Release) system.

## Module Lifecycle

| Term | Definition | Aliases to avoid |
|------|------------|------------------|
| **Alusmoodul** | A parametric base module containing masters and derived parts, stored in `Alusmoodulid/` | Base module, source module, design module |
| **Moodul** | A released, production-ready module with frozen geometry, stored in `Moodulid/` | Release, output, product |
| **Variant** | A specific parameter configuration defined in the Excel table | Configuration, version, instance |
| **Submodule** | A component group within a module (e.g., Karkass, Poroloon) | Component group, section, subsystem |

## Part Classification

| Term | Definition | Aliases to avoid |
|------|------------|------------------|
| **Master** | A multibody or sketch-only part used as the source for derivations; may appear in assemblies for reference (excluded from BOM) | Skeleton, parent, source part, Eskiis |
| **Derived Part** | A part created via `DeriveBodyAsNewPart` from a master body | Child part, generated part |
| **Manual Part** | A standalone part not derived from any master | Independent part, static part |
| **Standalone Copy** | A released part with all derivation links broken, preserving only geometry | Frozen part, disconnected part |

## Sharing Classification

| Term | Definition | Aliases to avoid |
|------|------------|------------------|
| **Shared Part** | A part with identical geometry across ALL variants of a module | Common part, universal part |
| **Unique Part** | A part with different geometry for different variants | Variant-specific part, custom part |
| **Ühine** | The shared folder (`Moodulid/Ühine/`) containing parts used across variants or modules | Common, shared folder |
| **Cross-Module Sharing** | Reuse of a shared part by multiple different modules of the same product | Inter-module sharing |

## Geometry Analysis

| Term | Definition | Aliases to avoid |
|------|------------|------------------|
| **Geometry Fingerprint** | A deterministic hash of part geometry (volume, surface area, bounding box) used to compare shapes | Signature, hash, geometry ID |
| **Full Fingerprint** | Source part number + geometry fingerprint; used for cross-module sharing to ensure same SOURCE part (stable across renames) | Combined fingerprint |
| **Variant Matrix** | A mapping of Part × Variant → Geometry Fingerprint used to classify parts as shared or unique | Analysis matrix, fingerprint table |

## File Operations

| Term | Definition | Aliases to avoid |
|------|------------|------------------|
| **Heritage** | The shared InternalName (GUID) between a source file and its copy, required for `ReplaceReference` | Ancestry, lineage |
| **Assembly Snapshot** | A frozen copy of an assembly with references replaced to point to released parts | Frozen assembly, variant assembly |
| **Reference Map** | A dictionary mapping source file paths to their released counterparts | Copy map, path map |

## Vault Integration

| Term | Definition | Aliases to avoid |
|------|------------|------------------|
| **Numbering Scheme** | A Vault configuration that generates sequential file numbers | Number generator, numbering rule |
| **Manifest** | A JSON file (`_manifest.json`) tracking released files, fingerprints, and cross-module usage | Release log, tracking file |
| **Disconnect-Save-Add Workflow** | The process of logging out of Vault, saving locally, then uploading via API to control file locations | Offline save, bypass workflow |

## Relationships

- A **Moodul** is created by releasing an **Alusmoodul**
- An **Alusmoodul** contains one or more **Masters**
- A **Master** generates multiple **Derived Parts**
- A **Variant** defines parameter values for all **Masters** in a module
- A **Shared Part** has one **Fingerprint** across all **Variants**
- A **Unique Part** has different **Fingerprints** for different **Variants**
- **Cross-Module Sharing** is detected by matching **Fingerprints** in the **Manifest**
- A **Standalone Copy** is created from a **Derived Part** by breaking **Heritage** links
- An **Assembly Snapshot** uses a **Reference Map** to point to released parts

## Example dialogue

> **Dev:** "When we release an **Alusmoodul**, do we copy the **Masters**?"

> **Domain expert:** "No — **Masters** stay in `Alusmoodulid/`. We only create **Standalone Copies** of **Derived Parts** and **Manual Parts**. The **Masters** are never released."

> **Dev:** "How do we know if a part should go to **Ühine** or a **Variant** folder?"

> **Domain expert:** "We compute the **Fingerprint** for each part across all **Variants**. If the **Fingerprint** is identical for ALL **Variants**, it's a **Shared Part** and goes to **Ühine**. Otherwise it's a **Unique Part** and each distinct geometry gets its own file in the **Variant** folder."

> **Dev:** "What about parts shared between different modules, like brackets used in both Selg and Iste?"

> **Domain expert:** "That's **Cross-Module Sharing**. When releasing the second module, we check the **Manifest** for matching **Fingerprints**. If we find a match in **Ühine**, we reuse that file instead of creating a duplicate."

> **Dev:** "So the **Manifest** is key for avoiding duplicate Vault numbers?"

> **Domain expert:** "Exactly. The **Manifest** tracks every **Shared Part** with its **Fingerprint**. It enables both **Cross-Module Sharing** detection and re-release optimization."

## Flagged ambiguities

- **"Module"** was used to mean both the parametric design (`Alusmoodul`) and the released output (`Moodul`). These are distinct: an **Alusmoodul** is editable with live parameters, while a **Moodul** is frozen for production.

- **"Shared"** can mean shared across variants (same module) OR shared across modules (different modules). Both go to **Ühine**, but the detection mechanism differs: variant sharing uses the **Variant Matrix**, while module sharing uses the **Manifest**.

- **"Copy"** is overloaded: `File.Copy` preserves **Heritage** (required for `ReplaceReference`), while a **Standalone Copy** explicitly breaks derivation links. Use "copy with heritage" vs "standalone copy" to distinguish.

- **"Eskiis"** (Sketch) is sometimes used as a synonym for **Master** in folder naming. Prefer **Master** in code and documentation; Eskiis is acceptable in folder names for user familiarity.

## Codebase inconsistencies (2026-04-27)

### HIGH: "Signature" vs "Fingerprint"
`Lib/MakeComponentsLib.vb` uses `Signature` and `ComputeBodySignature()` throughout, but the canonical term is **Fingerprint**. Newer test code (`Test1_Fingerprint.vb` etc.) correctly uses `ComputePartFingerprint()`. **Action**: Rename `Signature` → `Fingerprint` and `ComputeBodySignature()` → `ComputeBodyFingerprint()` in `MakeComponentsLib.vb`.

### MEDIUM: "skeleton" instead of "Master"
Several files use "skeleton" in comments to refer to **Master** parts:
- `Katsetused/UpdateSupportLengths.vb:7,18` - "skeleton or assembly geometry changes"
- `Katsetused/PlaceSupport.vb:16` - "assembly/skeleton"

### MEDIUM: "source part" instead of "Master"
- `Katsetused/RotateOriginAxes.vb:454,457` - "Source part may have no solid bodies"

### MEDIUM: "variant-specific" instead of "Unique Part"
`docs/plans/2026-04-26-module-release-cycle.md` extensively uses "variant-specific part" (lines 53, 57, 819, 1115-1116, 1297). The canonical term is **Unique Part**. Note: "variant-specific" as an adjective is acceptable; the issue is when it's used as a noun phrase replacing "Unique Part".

### LOW: "release log" instead of "Manifest"
- `docs/research/2026-04-26-moodulid-api-research.md:473` - "persist release log"

### LOW: ReleaseConfig class naming
`Lib/ExcelReaderLib.vb` defines `ReleaseConfig` class, mixing "Release" (alias for Moodul) with "Config" (alias for Variant). Consider renaming to `VariantDefinition` or `VariantSpec`.
