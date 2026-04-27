<!-- Copyright (c) 2026 Henri Kuuste -->
# Moodulid Concept Tests

This folder contains focused concept tests to validate API features needed for the Moodulid smart variant release system.

**IMPORTANT**: Run these tests before implementing the release system to verify assumptions!

## Test Matrix

| Test | API Being Validated | Risk Level | Prerequisites |
|------|---------------------|------------|---------------|
| Test1_Fingerprint | `SurfaceBody.Volume()`, `SurfaceArea()`, `RangeBox` | Low | Any part file |
| Test2_BreakLink | `DerivedPartComponent.BreakLinkToFile()` | **HIGH** | Derived part (from Loo komponendid) |
| Test3_TransactionRollback | `TransactionManager.StartTransaction/Abort` | Medium | Part with user parameters |
| Test4_DrawingRelink | `ReferencedFileDescriptor.PutLogicalFileName()` | **HIGH** | Drawing file with model references |
| Test5_ParameterCycle | Parameter save/restore across masters | Medium | Assembly with parametric masters |
| Test6_StandaloneCopy | SaveAs + `DerivedPartComponent.Delete()` | **HIGH** | Derived part |
| Test7_BinaryPatch | Binary patching for drawing refs | **HIGH** | Drawing file (will be modified!) |
| Test8_FileDescriptorReplaceReference | FileDescriptor.ReplaceReference API | Medium | Drawing file with model references |
| Test9_VaultSaveLocation | Vault save location APIs, path conversion | Low | Any document, logged into Vault |
| Test10_DisconnectSaveCheckin | Get number → save locally → check in | Medium | NEW document, logged into Vault |
| Test11_UserInfo | User data available through Inventor API | Low | Any document (or none) |
| Test12_VaultLoginLogout | Programmatic Vault login/logout approaches | Medium | Logged into Vault |

## Test Results Tracking

Run each test and record results here:

### Test1_Fingerprint
- **Status**: ✅ PASS
- **Date**: 2026-04-26
- **Result**: All fingerprinting APIs work correctly
- **Notes**: 
  - ✅ `body.Volume(0.001)` - WORKS (1977.74 cm³)
  - ❌ `body.SurfaceArea(0.001)` - **DOES NOT EXIST** (research doc was wrong!)
  - ✅ `body.RangeBox` - WORKS (sorted for orientation independence)
  - ✅ `body.Faces.Count` - WORKS (9 faces)
  - ✅ `face.Evaluator.Area` - WORKS (sum = 3014.29 cm², matches MassProperties.Area exactly!)
  - ✅ `MassProperties.Area` - WORKS (3014.29 cm², doesn't dirty document)
  - ✅ Fingerprints are **deterministic** (3 calls = identical results)
  - Fingerprint format: `V:volume|A:area|F:facecount|BB:d1xd2xd3` 

### Test2_BreakLink
- **Status**: ✅ PASS
- **Date**: 2026-04-26
- **Result**: `BreakLinkToFile()` works? **YES!**
- **Alternative needed?**: NO - this is the preferred approach
- **Notes**:
  - ✅ `DerivedPartComponent.BreakLinkToFile()` completes without error
  - ✅ `ReferencedDocuments.Count` goes from 1 → 0 (link broken!)
  - ✅ Geometry preserved (fingerprint matches exactly before/after)
  - ⚠️ `DerivedPartComponents.Count` still 1 (feature object exists but disconnected)
  - ⚠️ Document becomes dirty (must save to persist)
  - This is the **preferred approach** for creating standalone parts! 

### Test3_TransactionRollback
- **Status**: ✅ FULL PASS
- **Date**: 2026-04-26
- **Result**: Transaction rollback works perfectly for variant analysis!
- **Notes**:
  - ✅ Found 66 model parameters (laius, sügavus, selja_kõrgus, etc.)
  - ✅ Changed `laius` from 1100mm → 1210mm (10% increase)
  - ✅ Geometry CHANGED during transaction (fingerprints different)
  - ✅ `Transaction.Abort()` restored parameter to original
  - ✅ Fingerprint RESTORED exactly to original
  - ✅ `Document.Dirty` restored to False
  - **CONFIRMED**: Transactions are SAFE for variant analysis! 

### Test4_DrawingRelink
- **Status**: ❌ API FAILS - Binary patching required
- **Date**: 2026-04-26
- **Result**: `PutLogicalFileName()` works in iLogic? **NO - DOES NOT EXIST**
- **Alternative needed?**: **YES - Must use binary patching (VariantReleaseLib approach)**
- **Notes**:
  - ✅ `ReferencedFileDescriptors` collection - WORKS
  - ✅ `rfd.FullFileName` - WORKS (can read reference path)
  - ❌ `rfd.LogicalFileName` - DOES NOT EXIST
  - ❌ `rfd.ReferenceMissing` - DOES NOT EXIST
  - ❌ `rfd.PutLogicalFileName()` - **DOES NOT EXIST in iLogic!**
  - ❌ `File.ReferencedFileDescriptors` - Different interface, cast fails
  - **SOLUTION**: Use `VariantReleaseLib.UpdateSingleFileBinary` for drawing ref updates 

### Test5_ParameterCycle
- **Status**: ✅ FULL PASS
- **Date**: 2026-04-26
- **Result**: Parameters restore correctly? **YES!**
- **Derived parts update?**: **YES - 10 parts changed!**
- **Notes**:
  - ✅ Found master `00000.ipt` with 66 model parameters
  - ✅ Changed `laius` from 1100mm → 1210mm (10% increase)
  - ✅ **10 derived parts automatically updated** with new geometry
  - ✅ Restored `laius` back to 1100mm
  - ✅ All fingerprints restored exactly to original
  - **CONFIRMED**: Parameter cycling works for variant analysis! 

### Test6_StandaloneCopy
- **Status**: ✅ TESTED - Confirms Delete approach is WRONG
- **Date**: 2026-04-26
- **Result**: Geometry preserved after delete? **NO - ALL GEOMETRY REMOVED!**
- **Notes**:
  - ✅ SaveAs copy works
  - ✅ `DerivedPartComponent.Delete()` executes without error
  - ❌ **ALL SOLID BODIES REMOVED** after delete!
  - Fingerprint: `V:1040.92...` → `NO_BODIES`
  - **CONCLUSION**: Must use `BreakLinkToFile()` (Test2), NOT Delete!

### Test7_BinaryPatch
- **Status**: ✅ FULL PASS
- **Date**: 2026-04-26
- **Result**: Binary patching works for drawing refs? **YES!**
- **Notes**:
  - ✅ Path lengths must match (or new shorter, padded with nulls)
  - ✅ Binary patch replaced 2 reference occurrences in file
  - ✅ Drawing reopened successfully after patch
  - ✅ All 3 views intact, 0 errors
  - ✅ `BinaryReferenceUpdateLib.UpdateFileReferencesBinary` works!
  - **CONFIRMED**: Binary patching is the solution for drawing references 

### Test8_FileDescriptorReplaceReference
- **Status**: ✅ FULL PASS
- **Date**: 2026-04-26
- **Result**: FileDescriptor.ReplaceReference works in iLogic? **YES!**
- **Alternative needed?**: NO - This is the PREFERRED approach for Moodulid!
- **Notes**:
  - ✅ `doc.File.ReferencedFileDescriptors` - ACCESSIBLE in iLogic
  - ✅ `FileDescriptor.ReplaceReference()` - WORKS!
  - ✅ Heritage verified: File.Copy preserves InternalName (GUID matches exactly)
  - ✅ Drawing views update correctly (3 views, 0 errors)
  - ✅ **NO PATH LENGTH CONSTRAINT** - can use any valid path!
  - **BETTER than binary patching** for our use case:
    - Works through official API
    - No path length matching required
    - Views update automatically
  - **REQUIREMENT**: Target file must share heritage (same InternalName)
  - **For Moodulid**: File.Copy preserves heritage - perfect for release workflow!

### Test9_VaultSaveLocation
- **Status**: ❌ NO API AVAILABLE
- **Date**: 2026-04-26
- **Result**: Can set Vault save location programmatically? **NO**
- **Notes**:
  - ❌ **No API exists** to set "New File" dialog default folder from iLogic
  - ❌ Saving to local folder DOES NOT affect Vault dialog location
  - ❌ `WorkingFoldersManager.SetWorkingFolder()` - Inventor ignores it
  - ❌ `FileUIEvents.OnFileSaveAsDialog` - Vault Add-in intercepts first
  - ❌ Setting `$Prop["_FilePath"]` - Read-only in dialog context
  - The dialog uses "last used folder" stored internally by Vault Add-in
  - **ALTERNATIVES**:
    - Data Standard `GetParentFolderName` function (requires installation)
    - coolOrange powerEvents (third-party)
    - Manual folder selection by user
    - **Disconnect-Save-Add workflow** (see Test10)

### Test10_DisconnectSaveCheckin
- **Status**: ✅ PASS - WORKING ALTERNATIVE WORKFLOW
- **Date**: 2026-04-26
- **Result**: Successfully bypasses "New File" dialog by disconnecting from Vault
- **Notes**:
  - ✅ **Working workflow discovered** that bypasses Vault dialog entirely
  - ✅ Get file number from Vault while connected
  - ✅ Log out from Vault (manual step - non-modal dialog allows interaction)
  - ✅ Save document locally (no Vault dialog when disconnected!)
  - ✅ Log back in to Vault
  - ✅ Add file via `DocumentService.AddUploadedFile()` API
  - ✅ Sync file via `AcquireFiles` with Download option
  - **Key APIs added to VaultNumberingLib.vb**:
    - `UploadFileToVault()` - Upload file bytes to filestore
    - `AddFileToVault()` - Add local file to Vault folder
    - `AddFilesToVault()` - Batch add multiple files
    - `SyncFileFromVault()` - Download to establish tracking
    - `SyncFilesFromVault()` - Batch sync multiple files
  - **Limitations**:
    - Requires manual logout/login (no reliable programmatic toggle)
    - For assemblies, component references need separate handling
    - File added in default lifecycle state

### Test11_UserInfo
- **Status**: 🔲 NOT RUN
- **Date**: 
- **Result**: 
- **Notes**:
  - Tests the following data sources:
    1. `app.UserName` - Inventor username setting
    2. `app.GeneralOptions.UserName` - User settings
    3. `app.WebServicesManager` - Autodesk cloud account
    4. Vault Add-in automation - Vault logged-in user
    5. `System.Environment.UserName` - Windows username
    6. Document PropertySets - Author, Creator, Manager
    7. Environment variables - Autodesk-related vars
  - NOTE: WindowsIdentity and Registry APIs not available in iLogic runtime

### Test12_VaultLoginLogout
- **Status**: ✅ PASS
- **Date**: 2026-04-26
- **Result**: Programmatic logout/login works? **YES!**
- **Notes**:
  - ✅ **`LogoutCmdIntName` WORKS** - Successfully disconnects from Vault
  - ✅ **`LoginCmdIntName` WORKS** - Re-connects to Vault (showed dialog, but completed)
  - ✅ All three connection checks confirm state change:
    - `VaultNumberingLib.GetVaultConnection()` - reflects connection state
    - `VB.ConnectionManager.Instance.Connection` - reflects connection state
    - `EdmSecurity.Instance.IsSignedIn()` - reflects connection state
  - **Working code:**
    ```vb
    ' Logout
    app.CommandManager.ControlDefinitions.Item("LogoutCmdIntName").Execute()
    ' ... do operations without Vault dialogs ...
    ' Login (may show brief dialog even with auto-login)
    app.CommandManager.ControlDefinitions.Item("LoginCmdIntName").Execute()
    ```
  - **CONCLUSION**: The Disconnect-Save-Add workflow from Test10 can be FULLY AUTOMATED!
  - **Alternative commands found**: `VaultLogout`, `VaultLogin` (aliases)

## Critical API Findings

### Working APIs (Confirmed by Testing)
- ✅ `SurfaceBody.Volume(tolerance)` - Returns volume in cm³, deterministic
- ✅ `SurfaceBody.RangeBox` - Returns bounding box (sort dims for orientation independence)
- ✅ `SurfaceBody.Faces.Count` - Returns face count
- ✅ `Face.Evaluator.Area` - Returns individual face area (sum for per-body area)
- ✅ `MassProperties.Area` - Returns whole-document surface area (doesn't dirty doc)
- ✅ Fingerprinting is deterministic across multiple calls
- ✅ `DerivedPartComponent.BreakLinkToFile()` - Breaks derivation link, preserves geometry!
- ✅ `TransactionManager.StartTransaction()` - Creates transaction context
- ✅ `Transaction.Abort()` - Rolls back changes, restores fingerprints
- ✅ `BinaryReferenceUpdateLib.UpdateFileReferencesBinary()` - Patches drawing refs (closed file)
- ✅ `doc.File.ReferencedFileDescriptors.Item(n).ReplaceReference()` - **PREFERRED!** No path length constraint!

### Problematic APIs
- ❌ `SurfaceBody.SurfaceArea(tolerance)` - DOES NOT EXIST (use Face.Evaluator.Area sum)
- ❌ `ReferencedFileDescriptor.PutLogicalFileName()` - DOES NOT EXIST in iLogic!
- ❌ `ReferencedFileDescriptor.LogicalFileName` - DOES NOT EXIST in iLogic
- ❌ `ReferencedFileDescriptor.ReferenceMissing` - DOES NOT EXIST in iLogic
- ❌ `File.ReferencedFileDescriptors` cast - E_NOINTERFACE error
- ❌ `DerivedPartComponent.Delete()` - **REMOVES ALL GEOMETRY!** (use BreakLinkToFile instead)

### Alternative Approaches Needed
- For body surface area: iterate `body.Faces` and sum `face.Evaluator.Area`
- For drawing reference updates: **Two approaches**:
  1. **`FileDescriptor.ReplaceReference`** (RECOMMENDED for Moodulid):
     - Access via `doc.File.ReferencedFileDescriptors.Item(n).ReplaceReference(newPath)`
     - Requires files to share "heritage" (same InternalName/GUID)
     - File.Copy and SaveCopyAs preserve heritage - perfect for release workflow!
     - **No path length constraint**
  2. **Binary patching** (fallback for non-heritage cases):
     - Close drawing, patch Unicode paths directly in file bytes
     - Requires length-matched paths (new path ≤ old path length)
     - Use `BinaryReferenceUpdateLib.UpdateFileReferencesBinary`

### Vault Login/Logout APIs (from Test12 - CONFIRMED WORKING)
- ✅ `app.CommandManager.ControlDefinitions.Item("LogoutCmdIntName").Execute()` - **WORKS!** Fully disconnects
- ✅ `app.CommandManager.ControlDefinitions.Item("LoginCmdIntName").Execute()` - **WORKS!** Re-connects (may show dialog)
- ✅ `EdmSecurity.Instance.IsSignedIn()` - Check if signed in (reflects actual state)
- ✅ `EdmSecurity.Instance.OnLoginButtonExecute(true)` - Simulate login button
- ⚠️ `VDF.Vault.Library.ConnectionManager.LogIn/LogOut` - Creates **separate** connection, not Inventor's
- ⚠️ `app.ApplicationAddIns.ItemById("{48B682BC...}").Deactivate/Activate` - Disables UI only, not tested

**RECOMMENDED APPROACH for bypassing Vault "New File" dialog:**
```vb
' 1. Get file number while connected
' 2. Logout
app.CommandManager.ControlDefinitions.Item("LogoutCmdIntName").Execute()
' 3. Save files locally (no Vault dialog!)
' 4. Login
app.CommandManager.ControlDefinitions.Item("LoginCmdIntName").Execute()
' 5. Add files via API + sync
```

## How to Run Tests

1. **Test1_Fingerprint**: Open any `.ipt` file → Run rule
2. **Test2_BreakLink**: Open a derived part (created by Loo komponendid) → Run rule
3. **Test3_TransactionRollback**: Open a part with user parameters → Run rule
4. **Test4_DrawingRelink**: Open a `.idw` file that references a model → Run rule
5. **Test5_ParameterCycle**: Open an assembly with master/derived parts → Run rule
6. **Test6_StandaloneCopy**: Open a derived part → Run rule
7. **Test7_BinaryPatch**: Open a `.idw` file → Run rule (will close, patch, reopen)
8. **Test8_FileDescriptorReplaceReference**: Open a `.idw` file → Run rule (tests correct API)
9. **Test9_VaultSaveLocation**: Open any document, log into Vault → Run rule
10. **Test10_DisconnectSaveCheckin**: Open NEW document, log into Vault → Run rule (tests alternative workflow)
11. **Test11_UserInfo**: Open any document (or none) → Run rule (explores all user data sources)
12. **Test12_VaultLoginLogout**: Log into Vault → Run rule → Select test approach

## Dependencies

All tests use `Lib/UtilsLib.vb` for logging utilities.

## Next Steps

After all tests pass (or alternatives are identified):
1. Update `docs/research/2026-04-26-moodulid-api-research.md` with findings
2. Create detailed implementation plan in `docs/plans/`
3. Implement `VariantAnalysisLib.vb` and `SmartReleaseLib.vb`
