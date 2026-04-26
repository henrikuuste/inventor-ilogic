' Copyright (c) 2026 Henri Kuuste
' ============================================================================
' VaultNumberingLib - Vault WebServices API wrapper for numbering, folders, and file operations
' 
' Provides functions to:
' - Check Vault connection status
' - Enumerate available numbering schemes
' - Generate file numbers from a specific scheme
' - Create and manage folders in Vault
' - Upload and add new files to Vault (bypassing "New File" dialog)
' - Sync files between local workspace and Vault
'
' Dependencies:
'   UtilsLib - logging via UtilsLib.LogInfo / UtilsLib.LogWarn (set logger in caller:
'              UtilsLib.SetLogger(Logger) from Sub Main)
'
' Usage: 
'   In calling script (BEFORE AddVbFile):
'     AddReference "Autodesk.Connectivity.WebServices"
'     AddReference "Autodesk.DataManagement.Client.Framework.Vault"
'     AddReference "Connectivity.InventorAddin.EdmAddin"
'     AddVbFile "Lib/UtilsLib.vb"
'     AddVbFile "Lib/VaultNumberingLib.vb"
'
' ============================================================================
' WORKFLOW: Adding New Files to Vault (bypassing "New File" dialog)
' ============================================================================
' This workflow allows saving files to Vault without the interactive "New File"
' dialog, which is useful for batch operations.
'
' STEPS:
' 1. Get file numbers from Vault while connected (GenerateFileNumber)
' 2. Log out from Vault (user must do this manually or via Command)
' 3. Save files locally while disconnected
' 4. Log back in to Vault
' 5. Add files to Vault via API (AddFileToVault)
' 6. Sync files from Vault to establish tracking (SyncFileFromVault)
'
' BATCH EXAMPLE:
'   ' 1. Get all numbers while connected
'   Dim numbers As List(Of String) = GenerateFileNumbers(conn, scheme, count)
'   
'   ' 2. Log out (execute "LoginCmdIntName" command to toggle)
'   
'   ' 3. Save each document locally
'   For Each num In numbers
'       doc.SaveAs(localPath & "\" & num & ".ipt", False)
'   Next
'   
'   ' 4. Log back in (execute "LoginCmdIntName" command)
'   
'   ' 5-6. Add and sync all files
'   Dim results = AddFilesToVault(wsm, vaultFolder, localPaths, comment)
'   SyncFilesFromVault(vdfConn, vaultPaths)
' ============================================================================

Public Module VaultNumberingLib

    ' Get the current Vault connection, or Nothing if not logged in
    Public Function GetVaultConnection() As Object
        Try
            Return Connectivity.InventorAddin.EdmAddin.EdmSecurity.Instance.VaultConnection()
        Catch
            Return Nothing
        End Try
    End Function
    
    ' Check if user is logged into Vault
    Public Function IsVaultConnected() As Boolean
        Return GetVaultConnection() IsNot Nothing
    End Function
    
    ' Get connection info for logging
    Public Function GetConnectionInfo(conn As Object) As String
        If conn Is Nothing Then Return "Not connected"
        Try
            Return "Server: " & conn.Server & ", Vault: " & conn.Vault & ", User: " & conn.UserName
        Catch
            Return "Connected (details unavailable)"
        End Try
    End Function
    
    ' Get available numbering schemes for files
    Public Function GetNumberingSchemes(conn As Object) As Object()
        If conn Is Nothing Then
            UtilsLib.LogWarn("VaultNumberingLib: No Vault connection")
            Return Nothing
        End If
        
        Try
            Dim schemes As Object() = conn.WebServiceManager.NumberingService.GetNumberingSchemes("FILE", Nothing)
            UtilsLib.LogInfo("VaultNumberingLib: Found " & schemes.Length & " numbering scheme(s)")
            Return schemes
        Catch ex As Exception
            UtilsLib.LogWarn("VaultNumberingLib: Error getting schemes: " & ex.Message)
            Return Nothing
        End Try
    End Function
    
    ' Get scheme names as a list (for dropdown)
    Public Function GetSchemeNames(conn As Object) As System.Collections.Generic.List(Of String)
        Dim names As New System.Collections.Generic.List(Of String)
        Dim schemes As Object() = GetNumberingSchemes(conn)
        
        If schemes IsNot Nothing Then
            For Each scheme As Object In schemes
                names.Add(scheme.Name)
            Next
        End If
        
        Return names
    End Function
    
    ' Find a scheme by name
    Public Function FindSchemeByName(conn As Object, _
                                     schemeName As String) As Object
        Dim schemes As Object() = GetNumberingSchemes(conn)
        
        If schemes Is Nothing Then Return Nothing
        
        Dim searchName As String = schemeName.Trim()
        UtilsLib.LogInfo("VaultNumberingLib: Looking for scheme '" & searchName & "' (len=" & searchName.Length & ")")
        
        For Each scheme As Object In schemes
            Dim schName As String = CStr(scheme.Name).Trim()
            UtilsLib.LogInfo("VaultNumberingLib:   Comparing with '" & schName & "' (len=" & schName.Length & ")")
            If schName.Equals(searchName, StringComparison.OrdinalIgnoreCase) Then
                UtilsLib.LogInfo("VaultNumberingLib: Found matching scheme")
                Return scheme
            End If
        Next
        
        UtilsLib.LogWarn("VaultNumberingLib: Scheme '" & searchName & "' not found")
        Return Nothing
    End Function
    
    ' Generate a file number from a specific scheme
    Public Function GenerateFileNumber(conn As Object, _
                                       scheme As Object) As String
        If conn Is Nothing Then
            UtilsLib.LogWarn("VaultNumberingLib: No Vault connection")
            Return ""
        End If
        
        If scheme Is Nothing Then
            UtilsLib.LogWarn("VaultNumberingLib: No scheme specified")
            Return ""
        End If
        
        Try
            Dim numGenArgs() As String = {""}
            Dim number As String = conn.WebServiceManager.DocumentService.GenerateFileNumber(scheme.SchmID, numGenArgs)
            UtilsLib.LogInfo("VaultNumberingLib: Generated number: " & number)
            Return number
        Catch ex As Exception
            UtilsLib.LogWarn("VaultNumberingLib: Error generating number: " & ex.Message)
            Return ""
        End Try
    End Function
    
    ' Generate a file number by scheme name (convenience function)
    Public Function GenerateFileNumberByName(conn As Object, _
                                             schemeName As String) As String
        Dim scheme As Object = FindSchemeByName(conn, schemeName)
        If scheme Is Nothing Then Return ""
        Return GenerateFileNumber(conn, scheme)
    End Function
    
    ' Generate multiple file numbers at once
    Public Function GenerateFileNumbers(conn As Object, _
                                        scheme As Object, _
                                        count As Integer) As System.Collections.Generic.List(Of String)
        Dim numbers As New System.Collections.Generic.List(Of String)
        
        For i As Integer = 1 To count
            Dim num As String = GenerateFileNumber(conn, scheme)
            If String.IsNullOrEmpty(num) Then
                UtilsLib.LogWarn("VaultNumberingLib: Failed to generate number " & i & " of " & count)
                Exit For
            End If
            numbers.Add(num)
        Next
        
        Return numbers
    End Function
    
    ' ============================================================================
    ' Vault Folder Operations
    ' ============================================================================
    
    ' Convert local file system path to Vault path format
    ' Example: "C:\_SoftcomVault\Tooted\Test" -> "$/Tooted/Test"
    Public Function ConvertLocalPathToVaultPath(localPath As String, workspaceRoot As String) As String
        ' Ensure paths are normalized
        localPath = localPath.TrimEnd("\"c)
        workspaceRoot = workspaceRoot.TrimEnd("\"c)
        
        ' Check if local path starts with workspace root
        If Not localPath.StartsWith(workspaceRoot, StringComparison.OrdinalIgnoreCase) Then
            ' Paths don't match - cannot convert
            Return ""
        End If
        
        ' Get relative path (portion after workspace root)
        Dim relativePath As String = localPath.Substring(workspaceRoot.Length)
        relativePath = relativePath.TrimStart("\"c)
        
        ' Convert to Vault format: $/ prefix with forward slashes
        If String.IsNullOrEmpty(relativePath) Then
            Return "$"
        End If
        
        Return "$/" & relativePath.Replace("\", "/")
    End Function
    
    ' Get workspace root path - the local folder that maps to $/ in Vault
    ' This detects the root by testing path prefixes against Vault
    Public Function GetWorkspaceRoot(app As Object) As String
        ' Try to get from Inventor project first (just for logging)
        Try
            Dim project As Object = app.DesignProjectManager.ActiveDesignProject
            Dim projectWorkspace As String = project.WorkspacePath
            If Not String.IsNullOrEmpty(projectWorkspace) Then
                UtilsLib.LogInfo("VaultNumberingLib: Project workspace: " & projectWorkspace)
            End If
        Catch ex As Exception
            UtilsLib.LogWarn("VaultNumberingLib: Could not get workspace from project: " & ex.Message)
        End Try
        
        Return ""
    End Function
    
    ' Detect the Vault workspace root by testing path prefixes against Vault
    ' Returns the local path that corresponds to $/ in Vault
    Public Function DetectWorkspaceRoot(conn As Object, _
                                        localPath As String) As String
        If conn Is Nothing Then
            UtilsLib.LogWarn("VaultNumberingLib: No Vault connection for DetectWorkspaceRoot")
            Return ""
        End If
        
        If String.IsNullOrEmpty(localPath) Then
            UtilsLib.LogWarn("VaultNumberingLib: No local path for DetectWorkspaceRoot")
            Return ""
        End If
        
        ' Normalize path
        localPath = localPath.TrimEnd("\"c)
        
        ' Split path into parts
        Dim parts() As String = localPath.Split("\"c)
        
        ' Try progressively longer prefixes until we find one that maps to a valid Vault folder
        ' Start with drive letter (e.g., "C:\_SoftcomVault")
        For prefixLen As Integer = 2 To parts.Length - 1
            ' Build prefix path
            Dim prefix As String = String.Join("\", parts, 0, prefixLen)
            
            ' Build what the Vault path would be if this prefix is the root
            Dim remainingParts() As String = New String(parts.Length - prefixLen - 1) {}
            Array.Copy(parts, prefixLen, remainingParts, 0, parts.Length - prefixLen)
            Dim vaultPath As String = "$/" & String.Join("/", remainingParts)
            
            ' Test if this Vault path exists
            Try
                Dim folder As Object = conn.WebServiceManager.DocumentService.GetFolderByPath(vaultPath)
                If folder IsNot Nothing Then
                    UtilsLib.LogInfo("VaultNumberingLib: Detected workspace root: " & prefix)
                    UtilsLib.LogInfo("VaultNumberingLib: Vault path test succeeded: " & vaultPath)
                    Return prefix
                End If
            Catch
                ' Folder doesn't exist with this prefix, try next
            End Try
        Next
        
        ' Fallback: try common patterns
        Dim possibleRoots() As String = {"C:\_SoftcomVault", "C:\VaultWS", "C:\Vault"}
        For Each root As String In possibleRoots
            If localPath.StartsWith(root, StringComparison.OrdinalIgnoreCase) Then
                UtilsLib.LogInfo("VaultNumberingLib: Using fallback workspace root: " & root)
                Return root
            End If
        Next
        
        UtilsLib.LogWarn("VaultNumberingLib: Could not detect workspace root")
        Return ""
    End Function
    
    ' Get folder by Vault path, returns folder object or Nothing if not found
    Public Function GetVaultFolder(conn As Object, _
                                   vaultPath As String) As Object
        If conn Is Nothing Then
            UtilsLib.LogWarn("VaultNumberingLib: No Vault connection for GetVaultFolder")
            Return Nothing
        End If
        
        If String.IsNullOrEmpty(vaultPath) Then
            UtilsLib.LogWarn("VaultNumberingLib: Empty vault path for GetVaultFolder")
            Return Nothing
        End If
        
        Try
            Dim folder As Object = conn.WebServiceManager.DocumentService.GetFolderByPath(vaultPath)
            Return folder
        Catch ex As Exception
            ' Folder not found is expected in some cases, don't log as error
            UtilsLib.LogInfo("VaultNumberingLib: Folder not found: " & vaultPath)
            Return Nothing
        End Try
    End Function
    
    ' Ensure a folder exists in Vault, creating it if necessary
    ' Returns the folder object, or Nothing if creation failed
    Public Function EnsureVaultFolder(conn As Object, _
                                      vaultPath As String) As Object
        If conn Is Nothing Then
            UtilsLib.LogWarn("VaultNumberingLib: No Vault connection for EnsureVaultFolder")
            Return Nothing
        End If
        
        If String.IsNullOrEmpty(vaultPath) Then
            UtilsLib.LogWarn("VaultNumberingLib: Empty vault path for EnsureVaultFolder")
            Return Nothing
        End If
        
        ' First check if folder already exists
        Dim existingFolder As Object = GetVaultFolder(conn, vaultPath)
        If existingFolder IsNot Nothing Then
            UtilsLib.LogInfo("VaultNumberingLib: Vault folder already exists: " & vaultPath)
            Return existingFolder
        End If
        
        ' Parse path to get parent and folder name
        Dim lastSlash As Integer = vaultPath.LastIndexOf("/")
        If lastSlash <= 0 Then
            UtilsLib.LogWarn("VaultNumberingLib: Cannot parse parent path from: " & vaultPath)
            Return Nothing
        End If
        
        Dim parentPath As String = vaultPath.Substring(0, lastSlash)
        Dim folderName As String = vaultPath.Substring(lastSlash + 1)
        
        If String.IsNullOrEmpty(folderName) Then
            UtilsLib.LogWarn("VaultNumberingLib: Empty folder name in path: " & vaultPath)
            Return Nothing
        End If
        
        ' Get parent folder (it must exist)
        Dim parentFolder As Object = GetVaultFolder(conn, parentPath)
        If parentFolder Is Nothing Then
            UtilsLib.LogWarn("VaultNumberingLib: Parent folder not found in Vault: " & parentPath)
            Return Nothing
        End If
        
        ' Create the folder
        Try
            ' AddFolder(name, parentId, isLibrary)
            Dim newFolder As Object = conn.WebServiceManager.DocumentService.AddFolder(folderName, parentFolder.Id, False)
            UtilsLib.LogInfo("VaultNumberingLib: Created Vault folder: " & vaultPath)
            Return newFolder
        Catch ex As Exception
            ' Check for "folder exists" error (error code 1011)
            If ex.Message.Contains("1011") OrElse ex.Message.ToLower().Contains("exists") Then
                UtilsLib.LogInfo("VaultNumberingLib: Folder already exists (concurrent creation): " & vaultPath)
                ' Try to get the folder again
                Return GetVaultFolder(conn, vaultPath)
            End If
            
            UtilsLib.LogWarn("VaultNumberingLib: Failed to create folder: " & ex.Message)
            Return Nothing
        End Try
    End Function
    
    ' Ensure a subfolder exists in both local file system and Vault
    ' Returns the local path of the folder
    Public Function EnsureLocalAndVaultFolder(localPath As String, _
                                              conn As Object, _
                                              workspaceRoot As String) As Boolean
        Dim success As Boolean = True
        
        ' Create local folder if it doesn't exist
        Try
            If Not System.IO.Directory.Exists(localPath) Then
                System.IO.Directory.CreateDirectory(localPath)
                UtilsLib.LogInfo("VaultNumberingLib: Created local folder: " & localPath)
            End If
        Catch ex As Exception
            UtilsLib.LogWarn("VaultNumberingLib: Failed to create local folder: " & ex.Message)
            success = False
        End Try
        
        ' Create Vault folder if connected and workspace is known
        If conn IsNot Nothing AndAlso Not String.IsNullOrEmpty(workspaceRoot) Then
            Dim vaultPath As String = ConvertLocalPathToVaultPath(localPath, workspaceRoot)
            If Not String.IsNullOrEmpty(vaultPath) Then
                Dim vaultFolder As Object = EnsureVaultFolder(conn, vaultPath)
                If vaultFolder Is Nothing Then
                    UtilsLib.LogWarn("VaultNumberingLib: Could not ensure Vault folder: " & vaultPath)
                    ' Don't fail completely - local folder may still work
                End If
            Else
                UtilsLib.LogWarn("VaultNumberingLib: Could not convert path to Vault format: " & localPath)
            End If
        End If
        
        Return success
    End Function
    
    ' Ensure a local folder exists in Vault (folder must already exist on disk)
    ' This is useful when the local folder was created by the user manually
    ' Parameters:
    '   localPath - The full local folder path (must exist on disk)
    '   conn - Vault connection object (from GetVaultConnection)
    '   workspaceRoot - Local workspace root path (maps to $/ in Vault)
    ' Returns: True if folder is ready (exists in Vault or was created)
    Public Function EnsureFolderInVault(localPath As String, _
                                        conn As Object, _
                                        workspaceRoot As String) As Boolean
        ' Skip if folder doesn't exist on disk
        If Not System.IO.Directory.Exists(localPath) Then
            UtilsLib.LogWarn("VaultNumberingLib: Folder does not exist on disk: " & localPath)
            Return False
        End If
        
        ' Skip if no Vault connection
        If conn Is Nothing Then
            UtilsLib.LogInfo("VaultNumberingLib: No Vault connection, skipping Vault folder creation")
            Return True
        End If
        
        ' Skip if no workspace root
        If String.IsNullOrEmpty(workspaceRoot) Then
            UtilsLib.LogInfo("VaultNumberingLib: No workspace root, skipping Vault folder creation")
            Return True
        End If
        
        ' Convert local path to Vault path
        Dim vaultPath As String = ConvertLocalPathToVaultPath(localPath, workspaceRoot)
        
        If String.IsNullOrEmpty(vaultPath) Then
            UtilsLib.LogWarn("VaultNumberingLib: Path not in workspace, cannot create Vault folder: " & localPath)
            Return True
        End If
        
        ' Ensure folder exists in Vault
        Dim vaultFolder As Object = EnsureVaultFolder(conn, vaultPath)
        If vaultFolder IsNot Nothing Then
            UtilsLib.LogInfo("VaultNumberingLib: Vault folder ready: " & vaultPath)
            Return True
        Else
            UtilsLib.LogWarn("VaultNumberingLib: Could not create Vault folder: " & vaultPath)
            Return False
        End If
    End Function
    
    ' Ensure a Vault folder exists, creating all parent folders as needed (recursive)
    ' Unlike EnsureVaultFolder which only creates the last segment, this function
    ' creates the entire folder hierarchy from root to leaf.
    ' Example: EnsureVaultFolderRecursive(conn, "$/Tooted/Lume/Alusmoodulid/Iste")
    '   Creates: $/Tooted (if needed), then $/Tooted/Lume, then $/Tooted/Lume/Alusmoodulid, etc.
    ' Parameters:
    '   conn - Vault connection object (from GetVaultConnection)
    '   vaultPath - Full Vault path starting with $/ (e.g., "$/Tooted/Lume/Subfolder")
    ' Returns: The folder object, or Nothing if creation failed
    Public Function EnsureVaultFolderRecursive(conn As Object, vaultPath As String) As Object
        If conn Is Nothing Then
            UtilsLib.LogWarn("VaultNumberingLib: No Vault connection for EnsureVaultFolderRecursive")
            Return Nothing
        End If
        
        If String.IsNullOrEmpty(vaultPath) OrElse Not vaultPath.StartsWith("$") Then
            UtilsLib.LogWarn("VaultNumberingLib: Invalid vault path for recursive creation: " & vaultPath)
            Return Nothing
        End If
        
        ' First check if folder already exists (fast path)
        Dim existingFolder As Object = GetVaultFolder(conn, vaultPath)
        If existingFolder IsNot Nothing Then
            Return existingFolder
        End If
        
        ' Split path into segments: "$/Tooted/Lume/Sub" -> ["$", "Tooted", "Lume", "Sub"]
        Dim segments() As String = vaultPath.Split("/"c)
        If segments.Length < 2 Then
            UtilsLib.LogWarn("VaultNumberingLib: Path too short for recursive creation: " & vaultPath)
            Return Nothing
        End If
        
        ' Build path incrementally and create each level
        Dim currentPath As String = segments(0) ' Start with "$"
        Dim lastFolder As Object = Nothing
        
        For i As Integer = 1 To segments.Length - 1
            currentPath = currentPath & "/" & segments(i)
            
            ' Try to get existing folder first
            Dim folder As Object = GetVaultFolder(conn, currentPath)
            If folder Is Nothing Then
                ' Need to create this level
                folder = EnsureVaultFolder(conn, currentPath)
                If folder Is Nothing Then
                    UtilsLib.LogWarn("VaultNumberingLib: Failed to create folder level: " & currentPath)
                    Return Nothing
                End If
            End If
            
            lastFolder = folder
        Next
        
        Return lastFolder
    End Function
    
    ' Ensure a folder exists in both local file system and Vault, creating parent folders as needed
    ' Combines Directory.CreateDirectory (which handles nested paths) with recursive Vault creation
    ' Parameters:
    '   localPath - Full local path to create
    '   conn - Vault connection object (from GetVaultConnection)
    '   workspaceRoot - Local workspace root path (maps to $/ in Vault)
    ' Returns: True if folder is ready (both local and Vault)
    Public Function EnsureLocalAndVaultFolderRecursive(localPath As String, _
                                                       conn As Object, _
                                                       workspaceRoot As String) As Boolean
        Dim success As Boolean = True
        
        ' Create local folder (CreateDirectory handles nested paths automatically)
        Try
            If Not System.IO.Directory.Exists(localPath) Then
                System.IO.Directory.CreateDirectory(localPath)
                UtilsLib.LogInfo("VaultNumberingLib: Created local folder: " & localPath)
            End If
        Catch ex As Exception
            UtilsLib.LogWarn("VaultNumberingLib: Failed to create local folder: " & ex.Message)
            Return False
        End Try
        
        ' Create Vault folder recursively if connected
        If conn IsNot Nothing AndAlso Not String.IsNullOrEmpty(workspaceRoot) Then
            Dim vaultPath As String = ConvertLocalPathToVaultPath(localPath, workspaceRoot)
            If Not String.IsNullOrEmpty(vaultPath) Then
                Dim vaultFolder As Object = EnsureVaultFolderRecursive(conn, vaultPath)
                If vaultFolder Is Nothing Then
                    UtilsLib.LogWarn("VaultNumberingLib: Could not ensure Vault folder (recursive): " & vaultPath)
                    success = False
                End If
            Else
                UtilsLib.LogWarn("VaultNumberingLib: Could not convert path to Vault format: " & localPath)
            End If
        End If
        
        Return success
    End Function
    
    ' ============================================================================
    ' Vault Path to Local Path Conversion
    ' ============================================================================
    
    ' Convert Vault path to local file system path using workspace root
    ' This is the reverse of ConvertLocalPathToVaultPath
    ' Example: "$/Tooted/Project/Subfolder" -> "C:\_SoftcomVault\Tooted\Project\Subfolder"
    ' Parameters:
    '   vaultPath - Vault path starting with $/ or $ (e.g., "$/Tooted/Project")
    '   workspaceRoot - Local workspace root that maps to $ in Vault (e.g., "C:\_SoftcomVault")
    ' Returns: Local path, or empty string if conversion fails
    Public Function ConvertVaultPathToLocalPath(vaultPath As String, workspaceRoot As String) As String
        If String.IsNullOrEmpty(vaultPath) Then Return ""
        If String.IsNullOrEmpty(workspaceRoot) Then Return ""
        
        ' Normalize paths
        vaultPath = vaultPath.TrimEnd("/"c)
        workspaceRoot = workspaceRoot.TrimEnd("\"c)
        
        ' Handle different Vault path formats
        Dim relativePath As String = ""
        
        If vaultPath = "$" Then
            ' Root folder
            Return workspaceRoot
        ElseIf vaultPath.StartsWith("$/") Then
            ' Standard format: $/Folder/Subfolder
            relativePath = vaultPath.Substring(2)
        ElseIf vaultPath.StartsWith("$\") Then
            ' Alternative format with backslash
            relativePath = vaultPath.Substring(2)
        Else
            ' Invalid format - must start with $
            UtilsLib.LogWarn("VaultNumberingLib: Invalid Vault path (must start with $): " & vaultPath)
            Return ""
        End If
        
        ' Convert forward slashes to backslashes and combine
        relativePath = relativePath.Replace("/", "\")
        
        If String.IsNullOrEmpty(relativePath) Then
            Return workspaceRoot
        End If
        
        Return workspaceRoot & "\" & relativePath
    End Function
    
    ' Get local path for a Vault path using detected workspace root
    ' Convenience function that combines DetectWorkspaceRoot and ConvertVaultPathToLocalPath
    ' Parameters:
    '   conn - Vault connection object
    '   vaultPath - Vault path (e.g., "$/Tooted/Project")
    '   anyLocalPath - Any known local path in the Vault workspace (used to detect root)
    ' Returns: Local path, or empty string if conversion fails
    Public Function GetLocalPathForVaultPath(conn As Object, _
                                             vaultPath As String, _
                                             anyLocalPath As String) As String
        If conn Is Nothing Then
            UtilsLib.LogWarn("VaultNumberingLib: No Vault connection for GetLocalPathForVaultPath")
            Return ""
        End If
        
        ' Detect workspace root from the provided local path
        Dim workspaceRoot As String = DetectWorkspaceRoot(conn, anyLocalPath)
        If String.IsNullOrEmpty(workspaceRoot) Then
            UtilsLib.LogWarn("VaultNumberingLib: Could not detect workspace root")
            Return ""
        End If
        
        Return ConvertVaultPathToLocalPath(vaultPath, workspaceRoot)
    End Function
    
    ' ============================================================================
    ' Vault File Upload Operations
    ' ============================================================================
    
    ' Upload file bytes to Vault filestore and return upload ticket
    ' This is the first step in adding a new file to Vault.
    ' Parameters:
    '   wsm - WebServiceManager from conn.WebServiceManager
    '   filename - Name of the file (e.g., "Part1.ipt")
    '   fileContents - Byte array of file contents
    ' Returns: ByteArray upload ticket, or Nothing if upload fails
    Public Function UploadFileToVault(wsm As Object, _
                                      filename As String, _
                                      fileContents As Byte()) As Autodesk.Connectivity.WebServices.ByteArray
        Try
            Dim filestoreService = wsm.FilestoreService
            
            ' Set up file transfer header
            filestoreService.FileTransferHeaderValue = New Autodesk.Connectivity.WebServices.FileTransferHeader()
            filestoreService.FileTransferHeaderValue.Identity = Guid.NewGuid()
            filestoreService.FileTransferHeaderValue.Extension = System.IO.Path.GetExtension(filename)
            filestoreService.FileTransferHeaderValue.Vault = wsm.WebServiceCredentials.VaultName
            
            Dim uploadTicket As New Autodesk.Connectivity.WebServices.ByteArray()
            Dim bytesTotal As Integer = If(fileContents IsNot Nothing, fileContents.Length, 0)
            Dim bytesTransferred As Integer = 0
            Dim MAX_FILE_TRANSFER_SIZE As Integer = 49 * 1024 * 1024  ' 49 MB chunks
            
            UtilsLib.LogInfo("VaultNumberingLib: Uploading " & bytesTotal & " bytes for " & filename)
            
            Do
                ' Calculate buffer size for this chunk
                Dim remaining As Integer = bytesTotal - bytesTransferred
                Dim bufferSize As Integer = Math.Min(remaining, MAX_FILE_TRANSFER_SIZE)
                
                Dim buffer As Byte()
                If bufferSize = bytesTotal AndAlso bytesTransferred = 0 Then
                    ' Single chunk - use original array
                    buffer = fileContents
                Else
                    ' Multi-chunk - copy portion
                    buffer = New Byte(bufferSize - 1) {}
                    Array.Copy(fileContents, bytesTransferred, buffer, 0, bufferSize)
                End If
                
                ' Set transfer header properties
                filestoreService.FileTransferHeaderValue.Compression = Autodesk.Connectivity.WebServices.Compression.None
                filestoreService.FileTransferHeaderValue.IsComplete = ((bytesTransferred + bufferSize) = bytesTotal)
                filestoreService.FileTransferHeaderValue.UncompressedSize = bufferSize
                
                ' Upload this chunk
                Using fileContentsStream As New System.IO.MemoryStream(buffer)
                    uploadTicket.Bytes = filestoreService.UploadFilePart(fileContentsStream)
                End Using
                
                bytesTransferred += bufferSize
                
            Loop While bytesTransferred < bytesTotal
            
            UtilsLib.LogInfo("VaultNumberingLib: Upload complete, ticket size: " & uploadTicket.Bytes.Length & " bytes")
            Return uploadTicket
            
        Catch ex As Exception
            UtilsLib.LogWarn("VaultNumberingLib: Upload failed: " & ex.Message)
            Return Nothing
        End Try
    End Function
    
    ' Add a new file to Vault using an upload ticket
    ' This is the second step in adding a new file to Vault (after UploadFileToVault).
    ' Parameters:
    '   wsm - WebServiceManager from conn.WebServiceManager
    '   folderId - Target Vault folder ID
    '   filename - Name of the file (e.g., "Part1.ipt")
    '   comment - Check-in comment
    '   uploadTicket - Upload ticket from UploadFileToVault
    ' Returns: The added File object, or Nothing if add fails
    Public Function AddUploadedFileToVault(wsm As Object, _
                                           folderId As Long, _
                                           filename As String, _
                                           comment As String, _
                                           uploadTicket As Autodesk.Connectivity.WebServices.ByteArray) As Object
        Try
            Dim docService = wsm.DocumentService
            
            Dim addedFile = docService.AddUploadedFile( _
                folderId, _
                filename, _
                comment, _
                DateTime.Now, _
                Nothing, _
                Nothing, _
                Autodesk.Connectivity.WebServices.FileClassification.None, _
                False, _
                uploadTicket)
            
            If addedFile IsNot Nothing Then
                UtilsLib.LogInfo("VaultNumberingLib: File added - ID: " & addedFile.Id & ", Name: " & addedFile.Name)
            End If
            
            Return addedFile
            
        Catch ex As Exception
            UtilsLib.LogWarn("VaultNumberingLib: AddUploadedFile failed: " & ex.Message)
            Return Nothing
        End Try
    End Function
    
    ' Add a new file to Vault (combines upload and add in one call)
    ' This is the main function for adding a new local file to Vault.
    ' Parameters:
    '   conn - Vault connection object (from GetVaultConnection)
    '   vaultFolderPath - Target Vault folder path (e.g., "$/Tooted/Test")
    '   localFilePath - Full local file path
    '   comment - Check-in comment (optional, defaults to "Added via API")
    ' Returns: The added File object, or Nothing if operation fails
    Public Function AddFileToVault(conn As Object, _
                                   vaultFolderPath As String, _
                                   localFilePath As String, _
                                   Optional comment As String = "Added via API") As Object
        If conn Is Nothing Then
            UtilsLib.LogWarn("VaultNumberingLib: No Vault connection for AddFileToVault")
            Return Nothing
        End If
        
        If Not System.IO.File.Exists(localFilePath) Then
            UtilsLib.LogWarn("VaultNumberingLib: Local file not found: " & localFilePath)
            Return Nothing
        End If
        
        Try
            Dim wsm = conn.WebServiceManager
            Dim docService = wsm.DocumentService
            
            ' Get target folder
            Dim targetFolder = docService.GetFolderByPath(vaultFolderPath)
            If targetFolder Is Nothing Then
                UtilsLib.LogWarn("VaultNumberingLib: Vault folder not found: " & vaultFolderPath)
                Return Nothing
            End If
            
            ' Read file bytes
            Dim fileBytes() As Byte = System.IO.File.ReadAllBytes(localFilePath)
            Dim filename As String = System.IO.Path.GetFileName(localFilePath)
            
            UtilsLib.LogInfo("VaultNumberingLib: Adding " & filename & " to " & vaultFolderPath)
            
            ' Upload to filestore
            Dim uploadTicket = UploadFileToVault(wsm, filename, fileBytes)
            If uploadTicket Is Nothing OrElse uploadTicket.Bytes Is Nothing Then
                UtilsLib.LogWarn("VaultNumberingLib: Upload failed for " & filename)
                Return Nothing
            End If
            
            ' Add file to Vault
            Return AddUploadedFileToVault(wsm, targetFolder.Id, filename, comment, uploadTicket)
            
        Catch ex As Exception
            UtilsLib.LogWarn("VaultNumberingLib: AddFileToVault failed: " & ex.Message)
            Return Nothing
        End Try
    End Function
    
    ' Add multiple files to Vault in a batch
    ' Parameters:
    '   conn - Vault connection object
    '   vaultFolderPath - Target Vault folder path
    '   localFilePaths - List of local file paths
    '   comment - Check-in comment (optional)
    ' Returns: List of results (file objects or Nothing for failures)
    Public Function AddFilesToVault(conn As Object, _
                                    vaultFolderPath As String, _
                                    localFilePaths As System.Collections.Generic.List(Of String), _
                                    Optional comment As String = "Added via API") As System.Collections.Generic.List(Of Object)
        Dim results As New System.Collections.Generic.List(Of Object)
        
        If conn Is Nothing Then
            UtilsLib.LogWarn("VaultNumberingLib: No Vault connection for AddFilesToVault")
            Return results
        End If
        
        UtilsLib.LogInfo("VaultNumberingLib: Adding " & localFilePaths.Count & " files to " & vaultFolderPath)
        
        Dim successCount As Integer = 0
        For Each localPath As String In localFilePaths
            Dim result = AddFileToVault(conn, vaultFolderPath, localPath, comment)
            results.Add(result)
            If result IsNot Nothing Then
                successCount += 1
            End If
        Next
        
        UtilsLib.LogInfo("VaultNumberingLib: Batch add complete - " & successCount & "/" & localFilePaths.Count & " succeeded")
        Return results
    End Function
    
    ' ============================================================================
    ' Vault File Sync Operations
    ' ============================================================================
    
    ' Sync a file from Vault to establish proper local tracking
    ' After adding a file via API, this function downloads it to ensure
    ' the local file is properly linked to Vault's version tracking.
    ' Parameters:
    '   vaultFilePath - Full Vault file path (e.g., "$/Tooted/Test/Part1.ipt")
    ' Returns: True if sync succeeded
    ' Note: Requires Connectivity.Application.VaultBase reference
    Public Function SyncFileFromVault(vaultFilePath As String) As Boolean
        Try
            ' Get VDF connection via EdmAddin (same as GetVaultConnection)
            Dim vdfConn As Autodesk.DataManagement.Client.Framework.Vault.Currency.Connections.Connection
            vdfConn = Connectivity.InventorAddin.EdmAddin.EdmSecurity.Instance.VaultConnection()
            
            If vdfConn Is Nothing Then
                UtilsLib.LogWarn("VaultNumberingLib: No VDF connection for SyncFileFromVault")
                Return False
            End If
            
            ' Find the file
            Dim files = vdfConn.WebServiceManager.DocumentService.FindLatestFilesByPaths(New String() {vaultFilePath})
            If files Is Nothing OrElse files.Length = 0 OrElse files(0) Is Nothing Then
                UtilsLib.LogWarn("VaultNumberingLib: File not found in Vault: " & vaultFilePath)
                Return False
            End If
            
            Dim vaultFile = files(0)
            UtilsLib.LogInfo("VaultNumberingLib: Syncing file - ID: " & vaultFile.Id & ", Name: " & vaultFile.Name)
            
            ' Create FileIteration for AcquireFiles
            Dim fileIteration As New Autodesk.DataManagement.Client.Framework.Vault.Currency.Entities.FileIteration(vdfConn, vaultFile)
            
            ' Set up acquire settings - Download to sync local with Vault
            Dim aqSettings As New Autodesk.DataManagement.Client.Framework.Vault.Settings.AcquireFilesSettings(vdfConn, False)
            aqSettings.AddFileToAcquire(fileIteration, Autodesk.DataManagement.Client.Framework.Vault.Settings.AcquireFilesSettings.AcquisitionOption.Download)
            
            ' Execute the download/sync
            Dim results = vdfConn.FileManager.AcquireFiles(aqSettings)
            
            If results IsNot Nothing AndAlso results.FileResults IsNot Nothing Then
                For Each fileResult In results.FileResults
                    UtilsLib.LogInfo("VaultNumberingLib: Sync result: " & fileResult.File.EntityName & " - " & fileResult.Status.ToString())
                    If fileResult.Status.ToString() = "Success" Then
                        Return True
                    End If
                Next
            End If
            
            Return False
            
        Catch ex As Exception
            UtilsLib.LogWarn("VaultNumberingLib: SyncFileFromVault failed: " & ex.Message)
            Return False
        End Try
    End Function
    
    ' Sync multiple files from Vault in a batch
    ' Parameters:
    '   vaultFilePaths - List of full Vault file paths
    ' Returns: Number of successfully synced files
    Public Function SyncFilesFromVault(vaultFilePaths As System.Collections.Generic.List(Of String)) As Integer
        Dim successCount As Integer = 0
        
        UtilsLib.LogInfo("VaultNumberingLib: Syncing " & vaultFilePaths.Count & " files from Vault")
        
        For Each vaultPath As String In vaultFilePaths
            If SyncFileFromVault(vaultPath) Then
                successCount += 1
            End If
        Next
        
        UtilsLib.LogInfo("VaultNumberingLib: Batch sync complete - " & successCount & "/" & vaultFilePaths.Count & " succeeded")
        Return successCount
    End Function
    
    ' Complete workflow: Add file to Vault and sync it
    ' Combines AddFileToVault and SyncFileFromVault for convenience.
    ' Parameters:
    '   conn - Vault connection object
    '   vaultFolderPath - Target Vault folder path
    '   localFilePath - Full local file path
    '   comment - Check-in comment (optional)
    ' Returns: True if both add and sync succeeded
    Public Function AddAndSyncFileToVault(conn As Object, _
                                          vaultFolderPath As String, _
                                          localFilePath As String, _
                                          Optional comment As String = "Added via API") As Boolean
        ' Add file to Vault
        Dim addedFile = AddFileToVault(conn, vaultFolderPath, localFilePath, comment)
        If addedFile Is Nothing Then
            Return False
        End If
        
        ' Build Vault file path for sync
        Dim filename As String = System.IO.Path.GetFileName(localFilePath)
        Dim vaultFilePath As String = vaultFolderPath & "/" & filename
        
        ' Sync from Vault
        Return SyncFileFromVault(vaultFilePath)
    End Function

End Module
