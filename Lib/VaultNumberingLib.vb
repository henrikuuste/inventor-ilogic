' Copyright (c) 2026 Henri Kuuste
' ============================================================================
' VaultNumberingLib - Vault WebServices API wrapper for numbering and folders
' 
' Provides functions to:
' - Check Vault connection status
' - Enumerate available numbering schemes
' - Generate file numbers from a specific scheme
' - Create and manage folders in Vault
'
' Dependencies:
'   UtilsLib - logging via UtilsLib.LogInfo / UtilsLib.LogWarn (set logger in caller:
'              UtilsLib.SetLogger(Logger) from Sub Main)
'
' Usage: 
'   In calling script (BEFORE AddVbFile):
'     AddReference "Autodesk.Connectivity.WebServices"
'     AddReference "Autodesk.DataManagement.Client.Framework.Vault"
'     AddReference "Autodesk.DataManagement.Client.Framework.Vault.Forms"
'     AddReference "Connectivity.InventorAddin.EdmAddin"
'     AddVbFile "Lib/UtilsLib.vb"
'     AddVbFile "Lib/VaultNumberingLib.vb"
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

End Module
