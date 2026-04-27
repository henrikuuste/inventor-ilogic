<!-- Copyright (c) 2026 Henri Kuuste -->
---
date: 2026-04-25T18:45:00+03:00
researcher: claude-opus-4
git_commit: bf7f102d8d0f7a4c05bf77cc0d28052379596082
branch: main
repository: Inventor Rules
topic: "Moving assembly component patterns to browser folders via API"
tags: [research, codebase, browser-folders, patterns, OccurrencePatterns]
status: complete
last_updated: 2026-04-25
---

# Research: Moving Assembly Component Patterns to Browser Folders via API

**Date**: 2026-04-25 18:45 EEST  
**Git Commit**: bf7f102d8d0f7a4c05bf77cc0d28052379596082  
**Branch**: main

## Research Question

Is there a way to move component patterns (Rectangular, Circular, Mirror) to browser folders via the Inventor API? The current implementation in `SortingLib.vb` skips patterns and logs them for manual movement.

## Summary

**There IS a way to move patterns to browser folders via API**, but it requires a different approach than adding to existing folders:

| Method | Works for Patterns? | Notes |
|--------|---------------------|-------|
| `BrowserFolder.Add(node)` | **NO** | Fails with E_FAIL for all pattern types |
| `oPane.AddBrowserFolder(name, objectCollection)` | **YES** | Patterns must be included when creating the folder |

The solution is to use `ObjectCollection` with the `AddBrowserFolder` overload that accepts a collection of nodes to include at folder creation time.

## Detailed Findings

### Current Implementation (SortingLib.vb)

The current code at `Lib/SortingLib.vb:143-180` detects pattern elements and skips them:

```vb
If isPatternElement Then
    ' Component patterns cannot be moved to folders via API.
    ' The pattern itself must be moved manually...
    patternCount += 1
    Continue For
Else
    ' Standalone occurrence - move it directly
    folder.Add(movableNode)  ' This works for non-patterns
End If
```

### The Working Solution

From Autodesk forum (April 2025): https://forums.autodesk.com/t5/inventor-programming-forum/add-pattern-to-browser-folder/td-p/13426389

```vb
' 1. Create ObjectCollection
Dim oOccurrenceNodes As ObjectCollection
oOccurrenceNodes = ThisApplication.TransientObjects.CreateObjectCollection

' 2. Add regular occurrences (skip pattern elements to avoid duplicates)
For Each occ As ComponentOccurrence In oOccs
    If occ.Name.Contains(PartName) And occ.IsPatternElement = False Then
        oNode = oPane.GetBrowserNodeFromObject(occ)
        oOccurrenceNodes.Add(oNode)
    End If
Next

' 3. Add patterns via OccurrencePatterns collection
For Each pattern As OccurrencePattern In oPatterns
    If pattern.OccurrencePatternElements.Item(1).Occurrences.Item(1).Name.Contains(PartName) Then
        oNode = oPane.GetBrowserNodeFromObject(pattern)
        oOccurrenceNodes.Add(oNode)
    End If
Next

' 4. Create folder WITH the collection
oPFolder = oPane.AddBrowserFolder(FolderName, oOccurrenceNodes)
```

### Key Differences from Current Approach

| Current Approach | Working Approach |
|------------------|------------------|
| Create folder first, then add items | Collect all items, create folder with collection |
| Uses `BrowserFolder.Add(node)` | Uses `oPane.AddBrowserFolder(name, collection)` |
| Iterates `ComponentOccurrences` only | Also iterates `OccurrencePatterns` |
| Pattern elements detected via `occ.IsPatternElement` | Patterns accessed via `asmDef.OccurrencePatterns` |

### API Objects Involved

- **`OccurrencePatterns`**: Collection of all patterns in assembly (`asmDef.OccurrencePatterns`)
- **`OccurrencePattern`**: Individual pattern (Rectangular, Circular, Mirror)
- **`OccurrencePatternElements`**: Elements within a pattern
- **`ObjectCollection`**: Transient collection for grouping nodes

### Accessing Pattern Information

To get material of a pattern (for folder matching):
```vb
Dim pattern As OccurrencePattern
Dim firstOcc As ComponentOccurrence = pattern.OccurrencePatternElements.Item(1).Occurrences.Item(1)
Dim material As String = UtilsLib.GetOccurrenceMaterial(firstOcc)
```

### Historical Context

- **Pre-2023**: Patterns could not be moved via API at all. Workarounds included suppress-move-unsuppress or creating parent folders.
- **Inventor 2023+**: Direct drag-drop of patterns into folders was fixed in the UI. API support via `AddBrowserFolder` collection parameter works.

## Code References

- `Lib/SortingLib.vb:83-191` - Current `ApplyFolders` function that skips patterns
- `Lib/SortingLib.vb:143-161` - Pattern detection and skip logic
- `Lib/SortingLib.vb:196-214` - `GetPatternName` helper function
- `AGENTS.md:908` - Current constraint documentation

## Implementation Considerations

### For New Folders
When creating a new folder, collect all nodes (occurrences + patterns) first, then create:
```vb
folder = oPane.AddBrowserFolder(folderName, nodeCollection)
```

### For Existing Folders
Two options:
1. **Delete and recreate**: Delete existing folder, collect all items, create new folder with collection
2. **Use UI command**: Similar to current suppression approach, select folder and execute move command

### Pattern Types
All pattern types from `OccurrencePatterns` should work:
- Rectangular Component Pattern
- Circular Component Pattern  
- Mirror Component Pattern

### Filtering Pattern Elements
When iterating occurrences, use `occ.IsPatternElement = False` to avoid adding pattern child occurrences that would duplicate what's under the pattern parent.

## Final Implementation (Verified Working)

The implemented solution in `SortingLib.ApplyFolders`:

1. **Detect patterns via browser tree**: When an occurrence has `IsPatternElement=True`, walk up the browser tree using `FindParentPatternNode()` to find the pattern's browser node
2. **Mirror pattern detection**: Mirror patterns throw E_NOTIMPL on `NativeObject` - detect by checking if parent is assembly root (label contains `.iam`)
3. **Folder handling strategy**:
   - **New folders**: Create with `AddBrowserFolder(name, collection)` - supports patterns
   - **Existing folders, regular items**: Add with `BrowserFolder.Add(node)`
   - **Existing folders, patterns**: Recreate folder with fresh node references (delete folder, re-query occurrences by name, create new folder)

Key insight: After deleting a folder, browser node references become stale. Must re-query occurrence nodes by name to get fresh references.

## Related Research

- `docs/research/2026-04-25-loo-komponendid-failures.md` - Related component/pattern issues

## References

- [Forum: Add pattern to browser folder (April 2025)](https://forums.autodesk.com/t5/inventor-programming-forum/add-pattern-to-browser-folder/td-p/13426389)
- [Forum: Inventor 2016 Moving patterns into folders](https://forums.autodesk.com/t5/inventor-forum/inventor-2016-moving-assembly-patterns-into-folders-in-the/td-p/6590946)
- [Inventor API Help: BrowserFolder](https://help.autodesk.com/cloudhelp/2022/ENU/Inventor-API/files/BrowserFolder.htm)
