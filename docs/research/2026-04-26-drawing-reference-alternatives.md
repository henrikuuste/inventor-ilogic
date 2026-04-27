<!-- Copyright (c) 2026 Henri Kuuste -->
---
date: 2026-04-26T09:45:00+03:00
researcher: Claude
git_commit: 455e30425b6ad57a2e7fe980feebe7fa0d5faa38
branch: main
repository: Inventor-Rules
topic: "Alternatives to Binary Patching for Inventor Drawing References"
tags: [research, drawing-references, binary-patching, api, inventor-2026, moodulid]
status: complete
last_updated: 2026-04-26
---

# Research: Alternatives to Binary Patching for Inventor Drawing References

**Date**: 2026-04-26T09:45:00+03:00  
**Git Commit**: 455e30425b6ad57a2e7fe980feebe7fa0d5faa38  
**Branch**: main

## Research Question

Given the findings from Test4 (`PutLogicalFileName` doesn't exist in iLogic) and Test7 (binary patching works but requires exact path length matching), are there viable alternatives for updating Inventor drawing references in Inventor 2026 that don't have the path length constraint?

## Summary

**The path length constraint is fundamental, not a limitation of our implementation.** After extensive research, no alternative approaches completely bypass the path length issue. However, several strategies can work around it:

| Approach | Path Length Constraint | Heritage Constraint | Complexity | Recommended |
|----------|----------------------|---------------------|------------|-------------|
| **FileDescriptor.ReplaceReference** | **No** | Yes (same InternalName) | Low | ✅ **PREFERRED** |
| Binary Patching (fallback) | Yes (new ≤ old) | No | Low | ⚠️ Fallback only |
| Apprentice + PutLogicalFileNameUsingFull | Unknown | Unknown | High | ❌ Not needed |
| OLE Structured Storage manipulation | Yes (inherent to format) | No | Very High | ❌ No benefit |
| Path Length Planning | N/A (avoidance) | N/A | Medium | ⚠️ For fallback |
| Recreate Drawing Views | No | No | Very High | ❌ Last resort |

**Recommended approach**: Use `FileDescriptor.ReplaceReference` as primary method (works for all Moodulid releases since we copy masters). Keep binary patching only as fallback for edge cases where files don't share heritage.

**✅ CONFIRMED by Test8 (2026-04-26)**: `FileDescriptor.ReplaceReference` works in iLogic with no path length constraint!

---

## Detailed Findings

### 1. Why Path Length Matters

Inventor files (`.ipt`, `.iam`, `.idw`) use **OLE2 Compound Document Format** (Microsoft Structured Storage). File references are stored as **fixed-length Unicode strings** within the binary structure.

When we patch a reference:
- The new path **overwrites** the old path bytes
- If new path is shorter: can pad with null bytes (`\0`)
- If new path is longer: **cannot expand** without restructuring the entire file

This is not a limitation of our binary patching code—it's inherent to how Inventor stores references in the file format.

---

### 2. API Approaches Investigated

#### 2.1 FileDescriptor.ReplaceReference (Recommended for Copied Files)

**Source**: [Autodesk Manufacturing DevBlog](https://adndevblog.typepad.com/manufacturing/2012/08/replace-the-file-reference-by-inventor-api.html)

```vb
' Access via doc.File (not doc.ReferencedFileDescriptors)
Dim oFD As FileDescriptor = doc.File.ReferencedFileDescriptors(1)
oFD.ReplaceReference("C:\NewPath\Part2.ipt")
```

**Key findings**:
- Available in both **Inventor API** and **Apprentice Server**
- **Heritage constraint**: Replacement file must have same `InternalName` (GUID)
- Files share InternalName if created via `File.Copy` or `SaveCopyAs`
- **No path length constraint** - can replace with any valid path
- Works well for our Moodulid use case because released files ARE copies of masters

**Test4 issue**: The test tried `ReferencedFileDescriptor.PutLogicalFileName()` which doesn't exist. The correct API is `doc.File.ReferencedFileDescriptors.Item(n).ReplaceReference()`.

**Usage in iLogic** (corrected from Test4):
```vb
Dim drawDoc As DrawingDocument = CType(ThisApplication.ActiveDocument, DrawingDocument)
For i As Integer = 1 To drawDoc.File.ReferencedFileDescriptors.Count
    Dim fd As FileDescriptor = drawDoc.File.ReferencedFileDescriptors.Item(i)
    If fd.FullFileName = oldPath Then
        fd.ReplaceReference(newPath)
    End If
Next
drawDoc.Update()
drawDoc.Save()
```

#### 2.2 ReferencedFileDescriptor.PutLogicalFileNameUsingFull (Legacy/Apprentice)

**Source**: [Autodesk Forums](https://forums.autodesk.com/t5/inventor-programming-forum/change-inventor-references-using-powershell/td-p/8185466)

```vb
' Apprentice Server only - legacy API
Dim oApprentice As New ApprenticeServerComponent
Dim oDoc As ApprenticeServerDocument = oApprentice.Open("C:\drawing.idw")

For Each oRefFileDesc As ReferencedFileDescriptor In oDoc.ReferencedFileDescriptors
    If oRefFileDesc.FullFileName = "C:\OldPart.ipt" Then
        oRefFileDesc.PutLogicalFileNameUsingFull("C:\NewPart.ipt")
    End If
Next

oApprentice.FileSaveAs.AddFileToSave(oDoc, oDoc.FullFileName)
oApprentice.FileSaveAs.ExecuteSave
```

**Key findings**:
- **Pre-Inventor 10 API**, now deprecated
- Only available in **Apprentice Server** (not iLogic)
- May not have heritage constraint (unconfirmed for Inventor 2026)
- Requires external tool/script, not usable from iLogic rules
- **Apprentice Server 2026** requires separate installation and manual COM registration

**Status**: Needs testing. Could be viable for batch processing tools.

#### 2.3 OLE Structured Storage Direct Manipulation

**Source**: [Stack Overflow](https://stackoverflow.com/questions/55008271/python-writing-a-bytestream-to-overwrite-an-existing-microsoft-structured-stora)

```python
# Python example using pythoncom
import pythoncom
from win32com.storagecon import *

mode = STGM_READWRITE | STGM_SHARE_EXCLUSIVE
istorage = pythoncom.StgOpenStorageEx(filename, mode, STGFMT_STORAGE, 0, pythoncom.IID_IStorage)
# Navigate to streams, modify content
```

**Key findings**:
- Inventor files are OLE2/CFB (Compound Binary File) format
- Can use `pythoncom.StgOpenStorageEx` to open and navigate storage
- **Still has size constraints**: OLE streams have defined sizes
- Writing larger data requires resizing streams, which is complex
- No advantage over binary patching for this use case

**Verdict**: More complex than binary patching with no benefit for path replacement.

---

### 3. The InternalName/Heritage Issue

**Source**: [Autodesk Forums](https://forums.autodesk.com/t5/inventor-programming-ilogic/internal-name-is-it-possible-to-chance/td-p/8721976)

Every Inventor document has an `InternalName` (GUID) that:
- Is assigned when document is **created**
- **Cannot be changed** after first save
- Can ONLY be set via `Document.PutInternalNameAndRevisionId()` on **unsaved** documents
- Is preserved when file is copied (`File.Copy` or `SaveCopyAs`)

```vb
' This only works on NEW, UNSAVED documents
oDoc.PutInternalNameAndRevisionId(newInternalName, newRevisionId, oldInternalName, oldRevisionId)
```

**Why this matters**:
- `FileDescriptor.ReplaceReference` requires matching InternalName
- Files copied from masters **share** the master's InternalName
- Files created fresh have **different** InternalNames
- Cannot "forge" a matching InternalName after file creation

**For Moodulid**: Our released files ARE copies of masters, so they share InternalName. This means `ReplaceReference` should work!

---

### 4. Recommended Strategies

#### Strategy A: Use ReplaceReference for Copied Files (Best)

Since Moodulid releases files by copying masters:
1. **File.Copy** the master part/assembly to release location
2. Use `FileDescriptor.ReplaceReference` on drawings (works because heritage matches)
3. No path length constraint

**Test code to validate** (add to `Katsetused/Moodulid/`):
```vb
' Test8_FileDescriptorReplaceReference.vb
' Test if ReplaceReference works on drawing when target is a COPY of the original

Dim drawDoc As DrawingDocument = CType(ThisApplication.ActiveDocument, DrawingDocument)
Dim originalRef As String = drawDoc.File.ReferencedFileDescriptors.Item(1).FullFileName

' Create a COPY (preserves InternalName)
Dim copyPath As String = originalRef.Replace(".ipt", "_Copy.ipt")
System.IO.File.Copy(originalRef, copyPath, True)

' Now replace reference
drawDoc.File.ReferencedFileDescriptors.Item(1).ReplaceReference(copyPath)
drawDoc.Update()
drawDoc.Save()
```

#### Strategy B: Path Length Planning (For Binary Patching Fallback)

When `ReplaceReference` isn't suitable, use path length matching:

```vb
' Current implementation in BinaryReferenceUpdateLib.vb
Public Function CalculateMatchingVariantPath(masterFolder As String, _
                                              releaseRoot As String, _
                                              variantName As String) As String
    ' Returns a path where: releaseRoot + variantFolder = masterFolder length
    ' Uses padding with underscores if needed
End Function
```

**Folder structure example**:
```
Master:     C:\Project\CAD\MasterAsm\       (26 chars)
Release:    C:\Project\CAD\V001_____\       (26 chars) ← padded
```

#### Strategy C: Hybrid Approach (Recommended)

1. **For files created by copying masters**: Use `ReplaceReference` API
2. **For binary patching fallback**: Use path length planning
3. **Detect at runtime** which approach works:

```vb
Function UpdateDrawingReference(drawingPath As String, oldRef As String, newRef As String) As Boolean
    ' Try API approach first
    Try
        Dim drawDoc As DrawingDocument = CType(ThisApplication.Documents.Open(drawingPath), DrawingDocument)
        For i As Integer = 1 To drawDoc.File.ReferencedFileDescriptors.Count
            Dim fd As FileDescriptor = drawDoc.File.ReferencedFileDescriptors.Item(i)
            If fd.FullFileName.Equals(oldRef, StringComparison.OrdinalIgnoreCase) Then
                fd.ReplaceReference(newRef)
            End If
        Next
        drawDoc.Save()
        drawDoc.Close()
        Return True
    Catch ex As Exception
        ' API failed (probably heritage mismatch)
        ' Fall back to binary patching
    End Try
    
    ' Binary patching (requires matching path lengths)
    If oldRef.Length <> newRef.Length Then
        Logger.Error("Path length mismatch - cannot binary patch")
        Return False
    End If
    
    Return BinaryReferenceUpdateLib.UpdateFileReferencesBinary(drawingPath, pathMap, logs)
End Function
```

---

### 5. What Test4 Should Have Done

Test4 failed because it tested the wrong API. Here's the corrected approach:

**Wrong (what Test4 tried)**:
```vb
' This doesn't exist!
rfd.PutLogicalFileName(newPath)  ' ❌ FAILS
```

**Correct (via doc.File)**:
```vb
' Access FileDescriptor via doc.File
Dim fd As FileDescriptor = drawDoc.File.ReferencedFileDescriptors.Item(i)
fd.ReplaceReference(newPath)  ' ✅ WORKS (if heritage matches)
```

**Recommendation**: Create Test8 to validate `ReplaceReference` with copied files.

---

### 6. Inventor 2026 Apprentice Server Changes

**Important for automation tools**:
- Apprentice Server 2026 is **no longer auto-registered** with Inventor installation
- Requires separate standalone installation
- Must manually run `ApprenticeRegSrv.exe /install` for COM registration
- Update paths to reference standalone installation, not Inventor's `Bin` folder

---

## Code References

| File | Relevance |
|------|-----------|
| `Lib/BinaryReferenceUpdateLib.vb` | Current binary patching implementation |
| `Katsetused/Moodulid/Test4_DrawingRelink.vb` | Failed API test (wrong method) |
| `Katsetused/Moodulid/Test7_BinaryPatch.vb` | Successful binary patching test |
| `Lib/VariantReleaseLib.vb:522-535` | Original `PutLogicalFileName` attempt |

---

## Conclusions

1. **Binary patching path length constraint is fundamental** - no workaround exists that avoids it
2. **`FileDescriptor.ReplaceReference` is the better API** - no path length constraint, but requires heritage match
3. **Heritage match is automatic for copied files** - which is our Moodulid use case
4. **Test4 tested the wrong API** - should use `doc.File.ReferencedFileDescriptors.ReplaceReference`
5. **Hybrid approach recommended**: API first, binary patching as fallback with length planning

---

## Action Items

1. [x] Create Test8 to validate `FileDescriptor.ReplaceReference` on drawings with copied models - **PASSED 2026-04-26**
2. [ ] Update `VariantReleaseLib` to use `ReplaceReference` instead of `PutLogicalFileName`
3. [ ] Keep binary patching as fallback for edge cases (non-heritage files)
4. [ ] Document folder naming conventions for path length matching (fallback only)

## Test8 Results (2026-04-26)

```
FileDescriptor access via doc.File: WORKS
ReplaceReference with copied file: SUCCESS

Heritage verified:
- Original InternalName: {4B39C1CB-4F8A-D0B8-D803-A895A9E0ADC1}
- Copy InternalName: {4B39C1CB-4F8A-D0B8-D803-A895A9E0ADC1}
- Heritage: MATCH

Drawing views: 3 total, 0 errors
All views updated correctly to new reference
```

**CONFIRMED**: `FileDescriptor.ReplaceReference` is the preferred approach for Moodulid!

---

## External Links

- [Replace File Reference by Inventor API](https://adndevblog.typepad.com/manufacturing/2012/08/replace-the-file-reference-by-inventor-api.html) - Manufacturing DevBlog
- [Change Inventor References using PowerShell](https://forums.autodesk.com/t5/inventor-programming-forum/change-inventor-references-using-powershell/td-p/8185466) - Apprentice Server examples
- [InternalName Discussion](https://forums.autodesk.com/t5/inventor-programming-ilogic/internal-name-is-it-possible-to-chance/td-p/8721976) - Heritage/GUID explanation
- [FileDescriptor.ReplaceReference Must Share Ancestry](https://forums.autodesk.com/t5/inventor-programming-forum/filedescriptor-replacereference-method-quot-must-share-ancestry/td-p/13723228) - Heritage constraint discussion
- [OLE Structured Storage Manipulation](https://stackoverflow.com/questions/55008271/python-writing-a-bytestream-to-overwrite-an-existing-microsoft-structured-stora) - Python/pythoncom approach

---

## Related Research

- `docs/research/2026-04-26-moodulid-api-research.md` - Overall Moodulid API findings
- `Katsetused/Moodulid/README.md` - Test results tracking
