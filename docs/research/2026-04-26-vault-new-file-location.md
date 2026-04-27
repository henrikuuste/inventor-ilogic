<!-- Copyright (c) 2026 Henri Kuuste -->
---
date: 2026-04-26T10:00:00+03:00
researcher: Claude (Cursor Agent)
git_commit: 455e304
branch: main
repository: Inventor-Rules
topic: "Setting Vault Location Programmatically for New File Dialog"
tags: [research, vault, api, ilogic, save-location]
status: complete
last_updated: 2026-04-26
---

# Research: Setting Vault Location Programmatically for New File Dialog

**Date**: 2026-04-26  
**Git Commit**: 455e304  
**Branch**: main

## Research Question

When saving a new file in Inventor 2026 connected to Vault, a "New File" dialog appears with a "Vault Location" field that shows the target folder path (e.g., `$/Tooted/Andorra Dining/Alusmoodulid/Selg/Eskiis`). This location defaults to the "last used" folder.

**Question**: Is there a way to programmatically set this location so users don't have to manually browse every time?

## Summary

**CONCLUSION: There is NO API to set the Vault Location in the "New File" dialog from iLogic.**

**HOWEVER: A working alternative workflow was discovered that BYPASSES the dialog entirely.**

The Vault "New File" dialog maintains its own "last used folder" state that is:
- Stored internally by the Vault Add-in
- NOT connected to the document's local save path
- NOT exposed via any public API

### Attempted Approaches That Do NOT Work:

1. **Saving document to local folder first** - ❌ Vault dialog ignores document's current path; uses "last used" folder
2. **`WorkingFoldersManager.SetWorkingFolder()`** - ❌ Inventor ignores Vault client folder mappings
3. **Setting `$Prop["_FilePath"]`** - ❌ Read-only in the dialog context (confirmed by forum posts)
4. **`FileUIEvents.OnFileSaveAsDialog`** - ❌ Vault Add-in intercepts this event first
5. **Inventor command `VaultCheckin`** - ❌ Only works for files already in Vault

### WORKING SOLUTION: Disconnect-Save-Add Workflow ✅

**Tested and confirmed working in Test10_DisconnectSaveCheckin.vb**

1. **Data Standard customization** (requires installation) - Can use `GetParentFolderName` function
2. **Disconnect-Save-Add workflow** - Bypass dialog by adding file via Vault API (see below)
3. **Manual user selection** - User must browse to the correct folder each time

## WORKING SOLUTION: Disconnect-Save-Add Workflow

**Status**: ✅ Tested and confirmed working (Test10_DisconnectSaveCheckin.vb)

This workflow bypasses the "New File" dialog entirely by:
1. Getting file numbers while connected to Vault
2. Logging out from Vault
3. Saving files locally (no Vault dialog appears when disconnected)
4. Logging back in to Vault
5. Adding files directly to Vault via API
6. Syncing files to establish proper Vault tracking

### Workflow Steps (Single File)

```vb
' 1. Get file number while connected
Dim conn = VaultNumberingLib.GetVaultConnection()
Dim scheme = VaultNumberingLib.FindSchemeByName(conn, "My Numbering Scheme")
Dim fileNumber As String = VaultNumberingLib.GenerateFileNumber(conn, scheme)

' 2. Calculate target paths
Dim targetVaultPath As String = "$/Tooted/MyProject"
Dim targetLocalPath As String = "C:\_SoftcomVault\Tooted\MyProject"
Dim fullFilePath As String = targetLocalPath & "\" & fileNumber & ".ipt"

' 3. LOG OUT from Vault (programmatic - confirmed working!)
app.CommandManager.ControlDefinitions.Item("LogoutCmdIntName").Execute()

' 4. Save document locally (no Vault dialog when disconnected!)
doc.SaveAs(fullFilePath, False)

' 5. LOG BACK IN to Vault (programmatic - confirmed working!)
app.CommandManager.ControlDefinitions.Item("LoginCmdIntName").Execute()

' 6. Add file to Vault via API
Dim addedFile = VaultNumberingLib.AddFileToVault(conn, targetVaultPath, fullFilePath, "Added via iLogic")

' 7. Sync file from Vault to establish tracking
Dim vaultFilePath As String = targetVaultPath & "/" & fileNumber & ".ipt"
VaultNumberingLib.SyncFileFromVault(vaultFilePath)
```

### Workflow Steps (Batch - Multiple Files)

```vb
' 1. Get all file numbers while connected
Dim numbers As List(Of String) = VaultNumberingLib.GenerateFileNumbers(conn, scheme, count)

' 2. LOG OUT from Vault (programmatic)
app.CommandManager.ControlDefinitions.Item("LogoutCmdIntName").Execute()

' 3. Save all documents locally
For i As Integer = 0 To documents.Count - 1
    documents(i).SaveAs(localPath & "\" & numbers(i) & ".ipt", False)
Next

' 4. LOG BACK IN to Vault (programmatic)
app.CommandManager.ControlDefinitions.Item("LoginCmdIntName").Execute()

' 5. Add all files to Vault
Dim localPaths As New List(Of String)
For Each num In numbers
    localPaths.Add(localPath & "\" & num & ".ipt")
Next
Dim results = VaultNumberingLib.AddFilesToVault(conn, vaultPath, localPaths, "Batch add")

' 6. Sync all files
Dim vaultPaths As New List(Of String)
For Each num In numbers
    vaultPaths.Add(vaultPath & "/" & num & ".ipt")
Next
VaultNumberingLib.SyncFilesFromVault(vaultPaths)
```

### Key Functions Added to VaultNumberingLib.vb

| Function | Purpose |
|----------|---------|
| `UploadFileToVault(wsm, filename, fileContents)` | Upload file bytes to Vault filestore |
| `AddUploadedFileToVault(wsm, folderId, filename, comment, ticket)` | Add uploaded file to Vault folder |
| `AddFileToVault(conn, vaultFolderPath, localFilePath, comment)` | Complete add operation (upload + add) |
| `AddFilesToVault(conn, vaultFolderPath, localFilePaths, comment)` | Batch add multiple files |
| `SyncFileFromVault(vaultFilePath)` | Download file to sync local with Vault |
| `SyncFilesFromVault(vaultFilePaths)` | Batch sync multiple files |
| `AddAndSyncFileToVault(conn, vaultFolderPath, localFilePath, comment)` | Complete workflow in one call |

### Why This Works

1. **When disconnected from Vault**, Inventor's Vault Add-in doesn't intercept save operations
2. **Saving locally** creates the file in the correct workspace location
3. **Adding via API** (`DocumentService.AddUploadedFile`) puts the file directly in the target Vault folder
4. **Syncing via AcquireFiles** downloads the Vault version, establishing proper version tracking

### Limitations

1. ~~**Requires manual logout/login**~~ - **SOLVED!** See Test12_VaultLoginLogout results below
2. **File associations** - For assemblies, component references need separate handling
3. **No automatic lifecycle state** - File is added in default state; may need manual state change

### UPDATE: Programmatic Login/Logout WORKS! (Test12)

**Date**: 2026-04-26

Testing confirmed that Vault login/logout can be automated via Inventor command execution:

```vb
' Logout from Vault
app.CommandManager.ControlDefinitions.Item("LogoutCmdIntName").Execute()

' ... perform operations without Vault dialogs ...

' Login to Vault (may show brief progress dialog with auto-login)
app.CommandManager.ControlDefinitions.Item("LoginCmdIntName").Execute()
```

**This means the Disconnect-Save-Add workflow can be FULLY AUTOMATED!**

## Detailed Findings

### How Vault Location is Determined

The "Vault Location" field in the New File dialog is populated based on:

1. **Local path to Vault path mapping** - Defined in the Inventor Project (`.ipj`) file via `VaultVirtualFolder` XML element
2. **Last used folder** - Vault remembers the last folder used for save operations
3. **Document's current save path** - If the document has a local path, it's converted to Vault path

**Key insight**: The Vault add-in maps local file system paths to Vault paths. If you save a document to `C:\_SoftcomVault\Tooted\Project\Subfolder\Part.ipt`, and your IPJ maps `C:\_SoftcomVault` to `$/`, then the Vault Location will show `$/Tooted/Project/Subfolder`.

### Approach 1: Save to Correct Local Folder (RECOMMENDED)

**How it works**: Before triggering the Vault save dialog, save the document to the local folder that corresponds to your desired Vault location.

```vb
' Convert desired Vault path to local path
Dim desiredVaultPath As String = "$/Tooted/MyProject/NewFolder"
Dim localPath As String = VaultNumberingLib.ConvertVaultPathToLocal(conn, desiredVaultPath)

' Save document to that local location first
doc.SaveAs(localPath & "\Part1.ipt", False)

' Now when user triggers Vault check-in, the location will be pre-filled
```

**Pros**:
- Works with standard Vault workflow
- No special APIs needed
- User sees correct location in dialog

**Cons**:
- Requires knowing the Vault-to-local path mapping
- File must be saved locally first

### Approach 2: Bypass Dialog with FileManager.AddFile()

**How it works**: Use Vault API to add files directly to a specific folder, bypassing the interactive dialog entirely.

```vb
' Get folder entity
Dim folder As Object = conn.WebServiceManager.DocumentService.GetFolderByPath("$/Tooted/MyProject/NewFolder")
Dim folderEntity As Object = New Autodesk.DataManagement.Client.Framework.Vault.Currency.Entities.Folder(conn, folder)

' Add file directly
Dim vdfPath As Object = New Autodesk.DataManagement.Client.Framework.Currency.FilePathAbsolute(localFilePath)
Dim addedFile As Object = conn.FileManager.AddFile(folderEntity, "Added by iLogic", Nothing, Nothing, 0, False, vdfPath)
```

**Pros**:
- Complete control over target folder
- No user interaction required
- Can be fully automated

**Cons**:
- Only works for standalone files (not assemblies with references)
- Must handle file associations manually for CAD files
- Bypasses user confirmation

### Approach 3: WorkingFoldersManager.SetWorkingFolder()

**How it works**: Vault stores working folder mappings that determine local-to-Vault path relationships.

```vb
' Set working folder mapping
conn.WorkingFoldersManager.SetWorkingFolder("$/Tooted/MyProject", "C:\MyLocalFolder")
```

**Important Limitations**:
- Inventor ignores Vault client folder mappings; it uses `.ipj` file instead
- Setting persists until user logs in via Vault Explorer (which overwrites it)
- Primarily affects Vault Explorer, not Inventor's save dialog

### Approach 4: DataStandard Customization

If **Autodesk Vault Data Standard** is installed, you can customize the save dialog behavior via PowerShell scripts.

**Location**: `C:\ProgramData\Autodesk\Vault <version>\Extensions\DataStandard\CAD\addins\Default.ps1`

**Key function**:
```powershell
function GetParentFolderName
{
    # Return desired folder path to override selection
    $folderName = "$/Tooted/MyProject/TargetFolder"
    return $folderName
}
```

**Key variables**:
- `$Prop["_FilePath"]` - Current file's Vault path
- `$Prop["_SuggestedVaultPath"]` - Path from related file (e.g., assembly for component)
- `$Prop["_VaultVirtualPath"]` - Mapped virtual folder from IPJ

**Pros**:
- Rich customization options
- Can use folder properties, categories, etc.
- Works within standard Vault workflow

**Cons**:
- Requires DataStandard (separate installation)
- PowerShell-based customization
- Changes affect all users

### Approach 5: Modify Settings XML (NOT RECOMMENDED)

Vault stores settings in XML files:
- `C:\Users\<user>\AppData\Roaming\Autodesk\VaultCommon\Servers\<version>\<server>\Vaults\<vault>\Objects\WorkingFolders.xml`
- `C:\Users\<user>\AppData\Roaming\Autodesk\Autodesk Vault Professional <version>\ApplicationPreferences.xml`

**Risk**: Direct modification of these files is unsupported and may be overwritten.

## API References

### Vault Connection (from Inventor)

```vb
' Get existing Vault connection
Dim conn As Object = Connectivity.InventorAddin.EdmAddin.EdmSecurity.Instance.VaultConnection()
```

### Path Conversion

```vb
' Get local folder for Vault path
Dim localPath As String = conn.WorkingFoldersManager.GetWorkingFolder("$/MyFolder").FullPath

' Get Vault folder by path
Dim folder As Object = conn.WebServiceManager.DocumentService.GetFolderByPath("$/Tooted/Project")
```

### IPJ File Mapping

From Inventor API:
```vb
Dim project As DesignProject = app.DesignProjectManager.ActiveDesignProject
Dim vaultPath As String = project.VirtualVaultPath  ' e.g., "$/Tooted"
Dim localPath As String = project.WorkspacePath     ' e.g., "C:\_SoftcomVault\Tooted"
```

## Test Scripts

Created test scripts in `Katsetused/Moodulid/` to validate these approaches:

| Test | Description |
|------|-------------|
| Test9_VaultSaveLocation.vb | Tests path conversion and AddFile API |

## Recommendations

1. **For automated workflows**: Use `FileManager.AddFile()` to add files directly to the desired folder. Works well for non-CAD files or standalone parts.

2. **For interactive workflows**: Save documents to the correct local folder before check-in. The dialog will show the corresponding Vault path.

3. **For enterprise-wide control**: Consider DataStandard customization if consistent folder behavior is needed across all users.

4. **Key helper function needed**: Add `ConvertVaultPathToLocal()` to `VaultNumberingLib.vb` to complement the existing `ConvertLocalPathToVaultPath()`.

## Related Research

- `docs/research/2026-04-26-moodulid-api-research.md` - Related Vault API research

## Open Questions

1. Can we intercept the FileUIEvents to modify the dialog's initial folder selection?
2. Does the Vault add-in expose any additional APIs not documented publicly?
3. Would setting `doc.FullFileName` before save affect the dialog's location?
