' ============================================================================
' VaultNumberingLib - Vault WebServices API wrapper for numbering and folders
' 
' Provides functions to:
' - Check Vault connection status
' - Enumerate available numbering schemes
' - Generate file numbers from a specific scheme
' - Create and manage folders in Vault
'
' Usage: 
'   In calling script (BEFORE AddVbFile):
'     AddReference "Autodesk.Connectivity.WebServices"
'     AddReference "Autodesk.DataManagement.Client.Framework.Vault"
'     AddReference "Autodesk.DataManagement.Client.Framework.Vault.Forms"
'     AddReference "Connectivity.InventorAddin.EdmAddin"
'     AddVbFile "Lib/VaultNumberingLib.vb"
'
' Note: Logger is not available in library modules.
'       Pass a List(Of String) to collect log messages.
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
    Public Function GetNumberingSchemes(conn As Object, _
                                        logs As System.Collections.Generic.List(Of String)) As Object()
        If conn Is Nothing Then
            logs.Add("VaultNumberingLib: No Vault connection")
            Return Nothing
        End If
        
        Try
            Dim schemes As Object() = conn.WebServiceManager.NumberingService.GetNumberingSchemes("FILE", Nothing)
            logs.Add("VaultNumberingLib: Found " & schemes.Length & " numbering scheme(s)")
            Return schemes
        Catch ex As Exception
            logs.Add("VaultNumberingLib: Error getting schemes: " & ex.Message)
            Return Nothing
        End Try
    End Function
    
    ' Get scheme names as a list (for dropdown)
    Public Function GetSchemeNames(conn As Object, _
                                   logs As System.Collections.Generic.List(Of String)) As System.Collections.Generic.List(Of String)
        Dim names As New System.Collections.Generic.List(Of String)
        Dim schemes As Object() = GetNumberingSchemes(conn, logs)
        
        If schemes IsNot Nothing Then
            For Each scheme As Object In schemes
                names.Add(scheme.Name)
            Next
        End If
        
        Return names
    End Function
    
    ' Find a scheme by name
    Public Function FindSchemeByName(conn As Object, _
                                     schemeName As String, _
                                     logs As System.Collections.Generic.List(Of String)) As Object
        Dim schemes As Object() = GetNumberingSchemes(conn, logs)
        
        If schemes Is Nothing Then Return Nothing
        
        Dim searchName As String = schemeName.Trim()
        logs.Add("VaultNumberingLib: Looking for scheme '" & searchName & "' (len=" & searchName.Length & ")")
        
        For Each scheme As Object In schemes
            Dim schName As String = CStr(scheme.Name).Trim()
            logs.Add("VaultNumberingLib:   Comparing with '" & schName & "' (len=" & schName.Length & ")")
            If schName.Equals(searchName, StringComparison.OrdinalIgnoreCase) Then
                logs.Add("VaultNumberingLib: Found matching scheme")
                Return scheme
            End If
        Next
        
        logs.Add("VaultNumberingLib: Scheme '" & searchName & "' not found")
        Return Nothing
    End Function
    
    ' Generate a file number from a specific scheme
    Public Function GenerateFileNumber(conn As Object, _
                                       scheme As Object, _
                                       logs As System.Collections.Generic.List(Of String)) As String
        If conn Is Nothing Then
            logs.Add("VaultNumberingLib: No Vault connection")
            Return ""
        End If
        
        If scheme Is Nothing Then
            logs.Add("VaultNumberingLib: No scheme specified")
            Return ""
        End If
        
        Try
            Dim numGenArgs() As String = {""}
            Dim number As String = conn.WebServiceManager.DocumentService.GenerateFileNumber(scheme.SchmID, numGenArgs)
            logs.Add("VaultNumberingLib: Generated number: " & number)
            Return number
        Catch ex As Exception
            logs.Add("VaultNumberingLib: Error generating number: " & ex.Message)
            Return ""
        End Try
    End Function
    
    ' Generate a file number by scheme name (convenience function)
    Public Function GenerateFileNumberByName(conn As Object, _
                                             schemeName As String, _
                                             logs As System.Collections.Generic.List(Of String)) As String
        Dim scheme As Object = FindSchemeByName(conn, schemeName, logs)
        If scheme Is Nothing Then Return ""
        Return GenerateFileNumber(conn, scheme, logs)
    End Function
    
    ' Generate multiple file numbers at once
    Public Function GenerateFileNumbers(conn As Object, _
                                        scheme As Object, _
                                        count As Integer, _
                                        logs As System.Collections.Generic.List(Of String)) As System.Collections.Generic.List(Of String)
        Dim numbers As New System.Collections.Generic.List(Of String)
        
        For i As Integer = 1 To count
            Dim num As String = GenerateFileNumber(conn, scheme, logs)
            If String.IsNullOrEmpty(num) Then
                logs.Add("VaultNumberingLib: Failed to generate number " & i & " of " & count)
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
    Public Function GetWorkspaceRoot(app As Object, _
                                     logs As System.Collections.Generic.List(Of String)) As String
        ' Try to get from Inventor project first (just for logging)
        Dim projectWorkspace As String = ""
        Try
            Dim project As Object = app.DesignProjectManager.ActiveDesignProject
            projectWorkspace = project.WorkspacePath
            If Not String.IsNullOrEmpty(projectWorkspace) Then
                logs.Add("VaultNumberingLib: Project workspace: " & projectWorkspace)
            End If
        Catch ex As Exception
            logs.Add("VaultNumberingLib: Could not get workspace from project: " & ex.Message)
        End Try
        
        Return ""
    End Function
    
    ' Detect the Vault workspace root by testing path prefixes against Vault
    ' Returns the local path that corresponds to $/ in Vault
    Public Function DetectWorkspaceRoot(conn As Object, _
                                        localPath As String, _
                                        logs As System.Collections.Generic.List(Of String)) As String
        If conn Is Nothing Then
            logs.Add("VaultNumberingLib: No Vault connection for DetectWorkspaceRoot")
            Return ""
        End If
        
        If String.IsNullOrEmpty(localPath) Then
            logs.Add("VaultNumberingLib: No local path for DetectWorkspaceRoot")
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
                    logs.Add("VaultNumberingLib: Detected workspace root: " & prefix)
                    logs.Add("VaultNumberingLib: Vault path test succeeded: " & vaultPath)
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
                logs.Add("VaultNumberingLib: Using fallback workspace root: " & root)
                Return root
            End If
        Next
        
        logs.Add("VaultNumberingLib: Could not detect workspace root")
        Return ""
    End Function
    
    ' Get folder by Vault path, returns folder object or Nothing if not found
    Public Function GetVaultFolder(conn As Object, _
                                   vaultPath As String, _
                                   logs As System.Collections.Generic.List(Of String)) As Object
        If conn Is Nothing Then
            logs.Add("VaultNumberingLib: No Vault connection for GetVaultFolder")
            Return Nothing
        End If
        
        If String.IsNullOrEmpty(vaultPath) Then
            logs.Add("VaultNumberingLib: Empty vault path for GetVaultFolder")
            Return Nothing
        End If
        
        Try
            Dim folder As Object = conn.WebServiceManager.DocumentService.GetFolderByPath(vaultPath)
            Return folder
        Catch ex As Exception
            ' Folder not found is expected in some cases, don't log as error
            logs.Add("VaultNumberingLib: Folder not found: " & vaultPath)
            Return Nothing
        End Try
    End Function
    
    ' Ensure a folder exists in Vault, creating it if necessary
    ' Returns the folder object, or Nothing if creation failed
    Public Function EnsureVaultFolder(conn As Object, _
                                      vaultPath As String, _
                                      logs As System.Collections.Generic.List(Of String)) As Object
        If conn Is Nothing Then
            logs.Add("VaultNumberingLib: No Vault connection for EnsureVaultFolder")
            Return Nothing
        End If
        
        If String.IsNullOrEmpty(vaultPath) Then
            logs.Add("VaultNumberingLib: Empty vault path for EnsureVaultFolder")
            Return Nothing
        End If
        
        ' First check if folder already exists
        Dim existingFolder As Object = GetVaultFolder(conn, vaultPath, logs)
        If existingFolder IsNot Nothing Then
            logs.Add("VaultNumberingLib: Vault folder already exists: " & vaultPath)
            Return existingFolder
        End If
        
        ' Parse path to get parent and folder name
        Dim lastSlash As Integer = vaultPath.LastIndexOf("/")
        If lastSlash <= 0 Then
            logs.Add("VaultNumberingLib: Cannot parse parent path from: " & vaultPath)
            Return Nothing
        End If
        
        Dim parentPath As String = vaultPath.Substring(0, lastSlash)
        Dim folderName As String = vaultPath.Substring(lastSlash + 1)
        
        If String.IsNullOrEmpty(folderName) Then
            logs.Add("VaultNumberingLib: Empty folder name in path: " & vaultPath)
            Return Nothing
        End If
        
        ' Get parent folder (it must exist)
        Dim parentFolder As Object = GetVaultFolder(conn, parentPath, logs)
        If parentFolder Is Nothing Then
            logs.Add("VaultNumberingLib: Parent folder not found in Vault: " & parentPath)
            Return Nothing
        End If
        
        ' Create the folder
        Try
            ' AddFolder(name, parentId, isLibrary)
            Dim newFolder As Object = conn.WebServiceManager.DocumentService.AddFolder(folderName, parentFolder.Id, False)
            logs.Add("VaultNumberingLib: Created Vault folder: " & vaultPath)
            Return newFolder
        Catch ex As Exception
            ' Check for "folder exists" error (error code 1011)
            If ex.Message.Contains("1011") OrElse ex.Message.ToLower().Contains("exists") Then
                logs.Add("VaultNumberingLib: Folder already exists (concurrent creation): " & vaultPath)
                ' Try to get the folder again
                Return GetVaultFolder(conn, vaultPath, logs)
            End If
            
            logs.Add("VaultNumberingLib: Failed to create folder: " & ex.Message)
            Return Nothing
        End Try
    End Function
    
    ' Ensure a subfolder exists in both local file system and Vault
    ' Returns the local path of the folder
    Public Function EnsureLocalAndVaultFolder(localPath As String, _
                                              conn As Object, _
                                              workspaceRoot As String, _
                                              logs As System.Collections.Generic.List(Of String)) As Boolean
        Dim success As Boolean = True
        
        ' Create local folder if it doesn't exist
        Try
            If Not System.IO.Directory.Exists(localPath) Then
                System.IO.Directory.CreateDirectory(localPath)
                logs.Add("VaultNumberingLib: Created local folder: " & localPath)
            End If
        Catch ex As Exception
            logs.Add("VaultNumberingLib: Failed to create local folder: " & ex.Message)
            success = False
        End Try
        
        ' Create Vault folder if connected and workspace is known
        If conn IsNot Nothing AndAlso Not String.IsNullOrEmpty(workspaceRoot) Then
            Dim vaultPath As String = ConvertLocalPathToVaultPath(localPath, workspaceRoot)
            If Not String.IsNullOrEmpty(vaultPath) Then
                Dim vaultFolder As Object = EnsureVaultFolder(conn, vaultPath, logs)
                If vaultFolder Is Nothing Then
                    logs.Add("VaultNumberingLib: Could not ensure Vault folder: " & vaultPath)
                    ' Don't fail completely - local folder may still work
                End If
            Else
                logs.Add("VaultNumberingLib: Could not convert path to Vault format: " & localPath)
            End If
        End If
        
        Return success
    End Function

End Module
