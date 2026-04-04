' ============================================================================
' TestVaultFolderCreation - Test Vault WebServices API for folder operations
' 
' Tests:
' - Can we get a folder by Vault path (GetFolderByPath)?
' - Can we create a new folder (AddFolder)?
' - Path conversion from local to Vault format
'
' Usage: Run while logged into Vault with write permissions.
'        Creates a test folder under the current document's Vault folder.
' ============================================================================

AddReference "Autodesk.Connectivity.WebServices"
AddReference "Autodesk.DataManagement.Client.Framework.Vault"
AddReference "Autodesk.DataManagement.Client.Framework.Vault.Forms"
AddReference "Connectivity.InventorAddin.EdmAddin"

Imports ACW = Autodesk.Connectivity.WebServices
Imports VDF = Autodesk.DataManagement.Client.Framework
Imports edm = Connectivity.InventorAddin.EdmAddin

Sub Main()
    Logger.Info("TestVaultFolderCreation: Starting Vault folder API tests...")
    
    Dim app As Inventor.Application = ThisApplication
    Dim doc As Document = app.ActiveDocument
    
    If doc Is Nothing Then
        Logger.Error("TestVaultFolderCreation: No active document")
        MessageBox.Show("Ava dokument enne testi käivitamist.", "TestVaultFolderCreation")
        Exit Sub
    End If
    
    ' Get local path info
    Dim localPath As String = System.IO.Path.GetDirectoryName(doc.FullDocumentName)
    Logger.Info("TestVaultFolderCreation: Local document path: " & localPath)
    
    ' Test 1: Get Vault connection
    Logger.Info("TestVaultFolderCreation: Test 1 - Getting Vault connection...")
    Dim conn As VDF.Vault.Currency.Connections.Connection = Nothing
    
    Try
        conn = edm.EdmSecurity.Instance.VaultConnection()
    Catch ex As Exception
        Logger.Error("TestVaultFolderCreation: Exception getting Vault connection: " & ex.Message)
        MessageBox.Show("Vault ühenduse viga: " & ex.Message, "TestVaultFolderCreation")
        Exit Sub
    End Try
    
    If conn Is Nothing Then
        Logger.Warn("TestVaultFolderCreation: No Vault connection available.")
        MessageBox.Show("Vault ühendus puudub. Palun logi Vault'i sisse.", "TestVaultFolderCreation")
        Exit Sub
    End If
    
    Logger.Info("TestVaultFolderCreation: Vault connection OK - " & conn.Server & "/" & conn.Vault)
    
    ' Test 2: Detect workspace root by testing against Vault
    Logger.Info("TestVaultFolderCreation: Test 2 - Detecting workspace root...")
    Dim workspaceRoot As String = ""
    
    ' First show what the project reports (for info only)
    Try
        Dim project As Inventor.DesignProject = app.DesignProjectManager.ActiveDesignProject
        Logger.Info("TestVaultFolderCreation: Project workspace (for info): " & project.WorkspacePath)
    Catch ex As Exception
        Logger.Warn("TestVaultFolderCreation: Could not get project workspace: " & ex.Message)
    End Try
    
    ' Detect the actual Vault root by testing path prefixes
    workspaceRoot = DetectWorkspaceRootLocal(conn, localPath)
    
    If String.IsNullOrEmpty(workspaceRoot) Then
        Logger.Error("TestVaultFolderCreation: Could not determine workspace root")
        MessageBox.Show("Workspace'i juurkausta ei leitud." & vbCrLf & _
                        "Veendu, et kaust eksisteerib Vault'is.", "TestVaultFolderCreation")
        Exit Sub
    End If
    
    Logger.Info("TestVaultFolderCreation: Detected workspace root: " & workspaceRoot)
    
    ' Test 3: Convert local path to Vault path
    Logger.Info("TestVaultFolderCreation: Test 3 - Converting local path to Vault path...")
    Dim vaultPath As String = ConvertLocalPathToVaultPath(localPath, workspaceRoot)
    Logger.Info("TestVaultFolderCreation: Local path: " & localPath)
    Logger.Info("TestVaultFolderCreation: Vault path: " & vaultPath)
    
    ' Test 4: Get folder by path
    Logger.Info("TestVaultFolderCreation: Test 4 - Getting folder by path...")
    Dim folder As ACW.Folder = Nothing
    
    Try
        folder = conn.WebServiceManager.DocumentService.GetFolderByPath(vaultPath)
    Catch ex As Exception
        Logger.Error("TestVaultFolderCreation: Exception getting folder: " & ex.Message)
    End Try
    
    If folder Is Nothing Then
        Logger.Warn("TestVaultFolderCreation: Folder not found in Vault: " & vaultPath)
        Logger.Info("TestVaultFolderCreation: This might mean the folder only exists locally, not in Vault")
    Else
        Logger.Info("TestVaultFolderCreation: Folder found! ID: " & folder.Id & ", Name: " & folder.Name)
    End If
    
    ' Test 5: Create a test subfolder
    Logger.Info("TestVaultFolderCreation: Test 5 - Creating test subfolder...")
    Dim testFolderName As String = "TestFolder_" & DateTime.Now.ToString("yyyyMMdd_HHmmss")
    Dim testVaultPath As String = vaultPath & "/" & testFolderName
    
    If folder Is Nothing Then
        Logger.Warn("TestVaultFolderCreation: Cannot create subfolder - parent folder not found in Vault")
        Logger.Info("TestVaultFolderCreation: ========================================")
        Logger.Info("TestVaultFolderCreation: TEST SUMMARY")
        Logger.Info("TestVaultFolderCreation: ========================================")
        Logger.Info("TestVaultFolderCreation: Vault connection: OK")
        Logger.Info("TestVaultFolderCreation: Workspace root: " & workspaceRoot)
        Logger.Info("TestVaultFolderCreation: Path conversion: " & vaultPath)
        Logger.Info("TestVaultFolderCreation: Folder lookup: FAILED (folder not in Vault)")
        Logger.Info("TestVaultFolderCreation: Folder creation: SKIPPED")
        Logger.Info("TestVaultFolderCreation: ========================================")
        
        MessageBox.Show("Vault'i kaust ei leitud: " & vaultPath & vbCrLf & vbCrLf & _
                        "Veendu, et fail on Vault'is registreeritud.", "TestVaultFolderCreation")
        Exit Sub
    End If
    
    Dim newFolder As ACW.Folder = Nothing
    Try
        ' AddFolder(name, parentId, isLibrary)
        newFolder = conn.WebServiceManager.DocumentService.AddFolder(testFolderName, folder.Id, False)
        Logger.Info("TestVaultFolderCreation: Created folder successfully!")
        Logger.Info("TestVaultFolderCreation:   Name: " & newFolder.Name)
        Logger.Info("TestVaultFolderCreation:   ID: " & newFolder.Id)
        Logger.Info("TestVaultFolderCreation:   Path: " & testVaultPath)
    Catch ex As Exception
        Logger.Error("TestVaultFolderCreation: Exception creating folder: " & ex.Message)
        
        ' Check if folder already exists (error 1011)
        If ex.Message.Contains("1011") OrElse ex.Message.ToLower().Contains("exists") Then
            Logger.Info("TestVaultFolderCreation: Folder already exists - this is OK for idempotent operations")
        End If
    End Try
    
    ' Test 6: Verify the folder was created
    Logger.Info("TestVaultFolderCreation: Test 6 - Verifying created folder...")
    Dim verifyFolder As ACW.Folder = Nothing
    
    Try
        verifyFolder = conn.WebServiceManager.DocumentService.GetFolderByPath(testVaultPath)
        If verifyFolder IsNot Nothing Then
            Logger.Info("TestVaultFolderCreation: Verification successful - folder exists in Vault")
        Else
            Logger.Warn("TestVaultFolderCreation: Verification failed - folder not found")
        End If
    Catch ex As Exception
        Logger.Error("TestVaultFolderCreation: Exception verifying folder: " & ex.Message)
    End Try
    
    ' Also create the local folder
    Logger.Info("TestVaultFolderCreation: Creating local folder...")
    Dim localTestPath As String = System.IO.Path.Combine(localPath, testFolderName)
    Try
        If Not System.IO.Directory.Exists(localTestPath) Then
            System.IO.Directory.CreateDirectory(localTestPath)
            Logger.Info("TestVaultFolderCreation: Created local folder: " & localTestPath)
        Else
            Logger.Info("TestVaultFolderCreation: Local folder already exists: " & localTestPath)
        End If
    Catch ex As Exception
        Logger.Error("TestVaultFolderCreation: Exception creating local folder: " & ex.Message)
    End Try
    
    ' Summary
    Logger.Info("TestVaultFolderCreation: ========================================")
    Logger.Info("TestVaultFolderCreation: TEST SUMMARY")
    Logger.Info("TestVaultFolderCreation: ========================================")
    Logger.Info("TestVaultFolderCreation: Vault connection: OK")
    Logger.Info("TestVaultFolderCreation: Workspace root: " & workspaceRoot)
    Logger.Info("TestVaultFolderCreation: Path conversion: " & vaultPath)
    Logger.Info("TestVaultFolderCreation: Folder lookup: " & If(folder IsNot Nothing, "OK", "FAILED"))
    Logger.Info("TestVaultFolderCreation: Folder creation: " & If(newFolder IsNot Nothing OrElse verifyFolder IsNot Nothing, "OK", "FAILED"))
    Logger.Info("TestVaultFolderCreation: Local folder: " & localTestPath)
    Logger.Info("TestVaultFolderCreation: ========================================")
    Logger.Info("TestVaultFolderCreation: All tests completed!")
    
    If newFolder IsNot Nothing OrElse verifyFolder IsNot Nothing Then
        MessageBox.Show("Vault kausta loomine töötab!" & vbCrLf & vbCrLf & _
                        "Loodud kaust: " & testVaultPath & vbCrLf & _
                        "Kohalik kaust: " & localTestPath, "TestVaultFolderCreation")
    Else
        MessageBox.Show("Vault kausta loomine ebaõnnestus." & vbCrLf & _
                        "Vaata iLogic logi detailide jaoks.", "TestVaultFolderCreation")
    End If
End Sub

' Convert local path to Vault path
' Example: "C:\_SoftcomVault\Tooted\Test" -> "$/Tooted/Test"
Function ConvertLocalPathToVaultPath(localPath As String, workspaceRoot As String) As String
    ' Ensure paths are normalized
    localPath = localPath.TrimEnd("\"c)
    workspaceRoot = workspaceRoot.TrimEnd("\"c)
    
    ' Get relative path
    If Not localPath.StartsWith(workspaceRoot, StringComparison.OrdinalIgnoreCase) Then
        ' Paths don't match, return empty
        Return ""
    End If
    
    Dim relativePath As String = localPath.Substring(workspaceRoot.Length)
    relativePath = relativePath.TrimStart("\"c)
    
    ' Convert to Vault format
    If String.IsNullOrEmpty(relativePath) Then
        Return "$"
    End If
    
    Return "$/" & relativePath.Replace("\", "/")
End Function

' Detect the Vault workspace root by testing path prefixes against Vault
' Returns the local path that corresponds to $/ in Vault
Function DetectWorkspaceRootLocal(conn As VDF.Vault.Currency.Connections.Connection, localPath As String) As String
    If conn Is Nothing OrElse String.IsNullOrEmpty(localPath) Then
        Return ""
    End If
    
    ' Normalize path
    localPath = localPath.TrimEnd("\"c)
    
    ' Split path into parts
    Dim parts() As String = localPath.Split("\"c)
    
    ' Try progressively longer prefixes until we find one that maps to a valid Vault folder
    ' Start with drive + first folder (e.g., "C:\_SoftcomVault")
    For prefixLen As Integer = 2 To parts.Length - 1
        ' Build prefix path
        Dim prefix As String = String.Join("\", parts, 0, prefixLen)
        
        ' Build what the Vault path would be if this prefix is the root
        Dim remainingParts() As String = New String(parts.Length - prefixLen - 1) {}
        Array.Copy(parts, prefixLen, remainingParts, 0, parts.Length - prefixLen)
        Dim vaultPath As String = "$/" & String.Join("/", remainingParts)
        
        Logger.Info("TestVaultFolderCreation: Testing prefix '" & prefix & "' -> '" & vaultPath & "'")
        
        ' Test if this Vault path exists
        Try
            Dim folder As ACW.Folder = conn.WebServiceManager.DocumentService.GetFolderByPath(vaultPath)
            If folder IsNot Nothing Then
                Logger.Info("TestVaultFolderCreation: Found valid mapping!")
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
            Logger.Warn("TestVaultFolderCreation: Using fallback root: " & root)
            Return root
        End If
    Next
    
    Return ""
End Function
