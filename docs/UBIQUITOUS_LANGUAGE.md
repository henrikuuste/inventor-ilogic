<!-- Copyright (c) 2026 Henri Kuuste -->
# Ubiquitous Language

Domain terminology for the Inventor Moodulid (Module Release) system.

## Module Lifecycle

| Term | Definition | Aliases to avoid |
|------|------------|------------------|
| **Alusmoodul** | A parametric base module containing masters and derived parts, stored in `Alusmoodulid/` | Base module, source module, design module |
| **Moodul** | A released, production-ready module with frozen geometry, stored in `Moodulid/{MooduliNimi}/` | Release, output, product, variant |
| **Väljastamine** | The process of creating a **Moodul** from an **Alusmoodul** | Release, vabastamine |
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
| **Shared Part** | A part with identical geometry across ALL moodulid of an alusmoodul | Common part, universal part |
| **Unique Part** | A part with different geometry for different moodulid | Variant-specific part, moodulispetsiifiline |
| **Ühine** | The shared folder (`Moodulid/Ühine/`) containing parts used across moodulid or alusmoodulid | Common, shared folder |
| **Cross-Module Sharing** | Reuse of a shared part by multiple different alusmoodulid of the same product | Inter-module sharing |

## Geometry Analysis

| Term | Definition | Aliases to avoid |
|------|------------|------------------|
| **Geometry Fingerprint** | A deterministic hash of part geometry (volume, surface area, bounding box) used to compare shapes | Signature, hash, geometry ID |
| **Full Fingerprint** | Source part number + geometry fingerprint; used for cross-module sharing to ensure same SOURCE part (stable across renames) | Combined fingerprint |
| **Mooduli Matrix** | A mapping of Part × Moodul → Geometry Fingerprint used to classify parts as shared or unique | Variant matrix, analysis matrix, fingerprint table |

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

- A **Moodul** is created by **Väljastamine** of an **Alusmoodul**
- An **Alusmoodul** contains one or more **Masters**
- A **Master** generates multiple **Derived Parts**
- Each row in `moodulid.xlsx` defines parameter values for one **Moodul**
- A **Shared Part** has one **Fingerprint** across all **Moodulid**
- A **Unique Part** has different **Fingerprints** for different **Moodulid**
- **Cross-Module Sharing** is detected by matching **Fingerprints** in the **Manifest**
- A **Standalone Copy** is created from a **Derived Part** by breaking **Heritage** links
- An **Assembly Snapshot** uses a **Reference Map** to point to released parts

## Example dialogue

> **Dev:** "When we do **Väljastamine** of an **Alusmoodul**, do we copy the **Masters**?"

> **Domain expert:** "No — **Masters** stay in `Alusmoodulid/`. We only create **Standalone Copies** of **Derived Parts** and **Manual Parts**. The **Masters** are never released."

> **Dev:** "How do we know if a part should go to **Ühine** or a **Moodul** folder?"

> **Domain expert:** "We compute the **Fingerprint** for each part across all **Moodulid**. If the **Fingerprint** is identical for ALL **Moodulid**, it's a **Shared Part** and goes to **Ühine**. Otherwise it's a **Unique Part** and each distinct geometry gets its own file in the **Moodul** folder."

> **Dev:** "What about parts shared between different alusmoodulid, like brackets used in both Selg and Iste?"

> **Domain expert:** "That's **Cross-Module Sharing**. During **Väljastamine** of the second alusmoodul, we check the **Manifest** for matching **Fingerprints**. If we find a match in **Ühine**, we reuse that file instead of creating a duplicate."

> **Dev:** "So the **Manifest** is key for avoiding duplicate Vault numbers?"

> **Domain expert:** "Exactly. The **Manifest** tracks every **Shared Part** with its **Fingerprint**. It enables both **Cross-Module Sharing** detection and re-release optimization."

## Flagged ambiguities

- **"Module"** was used to mean both the parametric design (`Alusmoodul`) and the released output (`Moodul`). These are distinct: an **Alusmoodul** is editable with live parameters, while a **Moodul** is frozen for production.

- **"Variant"** should NOT be used in Estonian. In English code/docs, "variant" may appear, but in Estonian UI and user-facing text, always use **moodul/moodulid**.

- **"vabastamine"** is incorrect for "release" in this context. Use **väljastamine** instead.

- **"Shared"** can mean shared across moodulid (same alusmoodul) OR shared across alusmoodulid (different source modules). Both go to **Ühine**, but the detection mechanism differs: moodul sharing uses the **Mooduli Matrix**, while cross-module sharing uses the **Manifest**.

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
`docs/plans/2026-04-26-module-release-cycle.md` extensively uses "variant-specific part" (lines 53, 57, 819, 1115-1116, 1297). The canonical term is **Unique Part**. Note: In Estonian, use "moodulispetsiifiline" not "variandispetsiifiline".

### LOW: "release log" instead of "Manifest"
- `docs/research/2026-04-26-moodulid-api-research.md:473` - "persist release log"

### LOW: ReleaseConfig class naming
`Lib/ExcelReaderLib.vb` defines `ReleaseConfig` class, mixing "Release" (alias for Moodul) with "Config" (alias for Variant). Consider renaming to `MoodulDefinition` or `MoodulSpec`.
