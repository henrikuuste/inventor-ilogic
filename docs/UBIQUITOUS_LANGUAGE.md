<!-- Copyright (c) 2026 Henri Kuuste -->
# Ubiquitous Language

Domain terminology for the Inventor furniture engineering and manufacturing system (mööbli konstruktsioon ja tootmine).

## Product Hierarchy

| Term | Estonian | Definition | Aliases to avoid |
|------|----------|------------|------------------|
| **Product Family** | Toote perekond | A collection of related furniture products, stored at `$/Tooted/<name>/` in Vault. All files reference this in their "Project" property. | Product line, collection |
| **Module** | Moodul | A finished unit created during final assembly by combining elements on the assembly line. | Product, variant |
| **Element** | Element | An entity that goes through manufacturing separately, made up of parts. Usually an assembly created in manufacturing, not final assembly. Example: an armrest is an element, the full sofa is a module. | Component, subassembly |

## Element Lifecycle

| Term | Estonian | Definition | Aliases to avoid |
|------|----------|------------|------------------|
| **Base Element** | Aluselement | A parametric design containing masters and derived parts, describing how an entity is manufactured (parts and operations). Stored in `Aluselemendid/<name>/`. | Source element, design element, ~~alusmoodul~~ (old term) |
| **Released Element** | Väljastatud element | A production-ready element with frozen geometry, created from a base element by applying parameters or mirroring. Stored in `Elemendid/<name>/`. | Variant, ~~moodul~~ (old term) |
| **Release** | Väljastamine | The process of creating a released element from a base element. | Vabastamine (incorrect) |

## Module Lifecycle

| Term | Estonian | Definition | Aliases to avoid |
|------|----------|------------|------------------|
| **Base Module** | Alusmoodul | An assembly file (`.iam`) defining how elements are arranged to form a module. Contains placed base elements with instance names. Stored flat in `Alusmoodulid/`. | Module template, assembly definition |
| **Released Module** | Väljastatud moodul | A production-ready module assembly referencing released elements. Stored flat in `Moodulid/`. | Variant module |
| **Module Matrix** | Koosluste tabel | A generated output showing all released modules × all released elements with quantities. Based on released data only, no base references. | Assembly matrix, BOM matrix |

## Part Classification

| Term | Estonian | Definition | Aliases to avoid |
|------|----------|------------|------------------|
| **Part** | Detail | A manufactured piece within an element. | Component, piece |
| **Assembly** | Koost | A group of parts assembled together. Elements usually contain assemblies. | - |
| **Subassembly** | Alamkoost | A nested assembly within an element's assembly. | Sub-component |
| **Master** | Master | A multibody or sketch-only part used as the source for derivations; excluded from BOM. | Skeleton, Eskiis (acceptable in folder names only) |
| **Derived Part** | Tuletatud detail | A part created via `DeriveBodyAsNewPart` from a master body. | Child part, generated part |
| **Standalone Copy** | Iseseisev koopia | A released part with all derivation links broken, preserving only geometry. | Frozen part, disconnected part |

## Material Classification

| Term | Estonian | Definition | Notes |
|------|----------|------------|-------|
| **Wood** | Puit | Wood-based materials: vineer, PLP, kask, etc. | Complexity determines if separate drawing needed |
| **Cardboard** | Papp | Sheet materials, usually HDF. | Grouped with frame for BOM |
| **Foam** | Poroloon | Cushioning material. | Complex shapes need drawings |
| **Metal** | Metall | Metal parts and hardware. | May get separate BOM in future |
| **Frame** | Karkass | The structural wood/cardboard assembly. | - |

## BOM Types

| Term | Estonian | Definition | Aliases to avoid |
|------|----------|------------|------------------|
| **Frame BOM** | Puiduspets | BOM for wood frame parts, sent to frame contractor. Includes wood and cardboard. | Wood spec |
| **Foam BOM** | Poroloonispets | BOM for foam parts, sent to padding contractor. | Foam spec |

## Sharing Classification

| Term | Estonian | Definition | Aliases to avoid |
|------|----------|------------|------------------|
| **Shared Part** | Ühine detail | A part with identical geometry across multiple released elements. Stored in `Elemendid/Ühine/`. | Common part, universal part |
| **Unique Part** | Unikaalne detail | A part with different geometry for different released elements. Stored in the element's folder. | Element-specific part |
| **Shared Folder** | Ühine | The folder `Elemendid/Ühine/` containing parts shared across elements within a product family. | Common folder |

## Geometry Analysis

| Term | Estonian | Definition | Aliases to avoid |
|------|----------|------------|------------------|
| **Geometry Fingerprint** | Geomeetria sõrmejälg | A deterministic hash of part geometry (volume, surface area, bounding box) used to compare shapes. | Signature, hash |
| **Element Matrix** | Elemendi maatriks | A mapping of Part × Released Element → Fingerprint used to classify parts as shared or unique. | Variant matrix, fingerprint table |

## File Operations

| Term | Estonian | Definition | Aliases to avoid |
|------|----------|------------|------------------|
| **Heritage** | Pärand | The shared InternalName (GUID) between a source file and its copy, required for `ReplaceReference`. | Ancestry, lineage |
| **Reference Map** | Viidete kaart | A dictionary mapping source file paths to their released counterparts. | Copy map, path map |
| **Manifest** | Manifest | A JSON file (`_manifest.json`) tracking released files, fingerprints, and cross-element usage. | Release log |

## Mirroring

| Term | Estonian | Definition | Aliases to avoid |
|------|----------|------------|------------------|
| **Mirror Plane** | Peegli tasand | Origin plane or work plane used to create a mirrored released element from a base element. Empty = no mirroring. | - |
| **Mirrored Part** | Peegeldatud detail | A part that cannot be rotated to match its mirror image; requires separate manufacturing and part number. | - |
| **Symmetric Part** | Sümmeetriline detail | A part unchanged by mirroring; can be shared between left/right elements. | - |

## Vault Integration

| Term | Estonian | Definition | Aliases to avoid |
|------|----------|------------|------------------|
| **Numbering Scheme** | Numeratsiooniskeem | A Vault configuration that generates sequential file numbers. | Number generator |
| **Disconnect-Save-Add Workflow** | Lahti-Salvesta-Lisa töövoog | The process of logging out of Vault, saving locally, then uploading via API to control file locations. | Offline save |

## Excel Input Files

| File | Location | Purpose |
|------|----------|---------|
| **elemendid.xlsx** | `Aluselemendid/<element>/` | Defines released element permutations: name, parameters, mirror plane |
| **moodulid.xlsx** | `Alusmoodulid/` | Defines released module permutations: base module, released name, instance→element mappings |

## Vault Folder Structure

```
$/Tooted/<Product Family>/
  ├── Aluselemendid/                    (base elements with parametric masters)
  │   ├── Käetugi/                      (armrest base element)
  │   │   ├── elemendid.xlsx            (release definitions)
  │   │   ├── *.ipt, *.iam              (masters, parts, assemblies)
  │   │   └── *.idw                     (drawings)
  │   ├── Selg/                         (backrest)
  │   └── Iste/                         (seat)
  │
  ├── Elemendid/                        (released elements)
  │   ├── Ühine/                        (shared parts across elements)
  │   ├── Käetugi_V/                    (left armrest - released)
  │   │   ├── *.ipt, *.iam              (frozen parts, assemblies)
  │   │   ├── *.idw                     (drawings)
  │   │   ├── puiduspets.xlsx           (frame BOM)
  │   │   └── poroloonispets.xlsx       (foam BOM)
  │   ├── Käetugi_P/                    (right armrest - mirrored)
  │   └── Iste_90/, Iste_110/           (seat variants by parameter)
  │
  ├── Alusmoodulid/                     (flat - assembly files only)
  │   ├── moodulid.xlsx                 (module definitions)
  │   ├── Diivan_2K.iam                 (2-seat sofa base module)
  │   └── Diivan_3K.iam                 (3-seat sofa base module)
  │
  └── Moodulid/                         (flat - released module assemblies)
      ├── Diivan_2K_V.iam               (2-seat with left armrest)
      └── Diivan_2K_P.iam               (2-seat with right armrest)
```

## Relationships

- A **Product Family** contains **Base Elements**, **Released Elements**, **Base Modules**, and **Released Modules**
- A **Base Element** is released into one or more **Released Elements** via **Väljastamine**
- A **Released Element** is created by applying parameters and/or mirroring to a **Base Element**
- A **Base Module** defines how **Base Elements** are arranged (with instance names)
- A **Released Module** is created by mapping instance names to specific **Released Elements**
- An **Element** contains **Parts**, **Assemblies**, and **Subassemblies**
- A **Part** with identical **Fingerprint** across elements is a **Shared Part** (goes to **Ühine**)
- A **Mirrored Part** that cannot be rotated requires a separate part number
- Each **Released Element** has a **Puiduspets** (frame BOM) and **Poroloonispets** (foam BOM)
- The **Koosluste tabel** shows all **Released Modules** × **Released Elements** with quantities

## Example Dialogue

> **Dev:** "When we do **Väljastamine** of a **Aluselement**, what gets copied?"

> **Domain expert:** "We create **Standalone Copies** of all **Derived Parts** and assemblies. The **Masters** stay in the **Aluselement** folder - they're never released. Each copy goes either to the **Released Element** folder or to **Ühine** if it's shared."

> **Dev:** "How do we know if a part goes to **Ühine**?"

> **Domain expert:** "We compute the **Fingerprint** for each part across all **Released Elements** of that **Base Element**. If the **Fingerprint** is identical for ALL releases, it's a **Shared Part** and goes to **Ühine**."

> **Dev:** "What about the mirrored armrest? Left and right are different **Elements**?"

> **Domain expert:** "Yes. **Käetugi_V** and **Käetugi_P** are separate **Released Elements** from the same **Base Element**. Parts that can't be rotated to match get new part numbers. **Symmetric Parts** are shared in **Ühine**."

> **Dev:** "And for **Modules** - how does **moodulid.xlsx** work?"

> **Domain expert:** "The **Base Module** `.iam` defines which **Base Elements** go where, with instance names like 'Käetugi:1'. The Excel maps each instance to a specific **Released Element**. So one row might say: Base Module 'Diivan_2K', Released Module 'Diivan_2K_V', and map 'Käetugi:1' → 'Käetugi_V'."

## Flagged Ambiguities

- **"Alusmoodul"** previously meant what we now call **Aluselement** (the parametric design). Going forward, **Alusmoodul** means the assembly file that defines module composition. All existing code using "Alusmoodul" for the parametric design should be refactored to use **Aluselement**.

- **"Moodul"** previously meant what we now call **Released Element** (a single manufactured unit). Going forward, **Moodul** means the final assembly unit created by combining elements. Existing code needs refactoring.

- **"Variant"** should NOT be used. Use **Released Element** or **Released Module** instead. The variation comes from parameters and mirroring, not "variants."

- **"Kaetoes"** is incorrect Estonian. Use **Käetugi** for armrest.

## Codebase Migration Required

The following terminology changes need to be applied throughout the codebase:

| Old Term | New Term | Notes |
|----------|----------|-------|
| Alusmoodul (parametric design) | Aluselement | Major change - affects folder names, class names, variables |
| Moodul (released unit) | Väljastatud element | When referring to released manufactured units |
| Alusmoodulid/ folder | Aluselemendid/ | Vault and local folder structure |
| Moodulid/ folder (releases) | Elemendid/ | For released elements |
| moodulid.xlsx (element definitions) | elemendid.xlsx | Per-element release definitions |
| Variant | Released Element/Module | Avoid "variant" terminology |
| Signature | Fingerprint | Already partially migrated |
