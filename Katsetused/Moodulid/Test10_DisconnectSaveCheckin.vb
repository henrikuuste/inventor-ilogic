' Copyright (c) 2026 Henri Kuuste
' Test10_DisconnectSaveCheckin.vb
' PURPOSE: Test alternative workflow to bypass "New File" dialog
' 
' WORKFLOW:
' 1. Get a new file number from Vault (using "Test numbriskeem") - MUST BE CONNECTED
' 2. Ensure folders exist in Vault
' 3. LOG OUT from Vault
' 4. Save document locally with the number as filename
' 5. LOG BACK IN to Vault
' 6. Check in the file (should use "Add File" instead of "New File" dialog)
'
' TARGET FOLDER: $/Tooted/Test -> C:\_SoftcomVault\Tooted\Test
'
' RUN: Open a NEW (unsaved) document while connected to Vault

AddReference "Autodesk.Connectivity.WebServices"
AddReference "Autodesk.DataManagement.Client.Framework.Vault"
AddReference "Connectivity.Application.VaultBase"
AddReference "Connectivity.InventorAddin.EdmAddin"
AddVbFile "Lib/StringsLib.vb"
AddVbFile "Lib/UtilsLib.vb"
AddVbFile "Lib/VaultNumberingLib.vb"

Sub Main()
    Dim app As Inventor.Application = ThisApplication
    Dim doc As Document = app.ActiveDocument
    
    ' Configuration
    Const NUMBERING_SCHEME As String = "Test numbriskeem"
    Const TARGET_VAULT_PATH As String = "$/Tooted/Test"
    
    Logger.Info("=== Test10_DisconnectSaveCheckin: Starting ===")
    Logger.Info("")
    Logger.Info("WORKFLOW:")
    Logger.Info("1. Get number from Vault (must be connected)")
    Logger.Info("2. Ensure folders exist")
    Logger.Info("3. LOG OUT from Vault")
    Logger.Info("4. Save locally while disconnected")
    Logger.Info("5. LOG BACK IN to Vault")
    Logger.Info("6. Check in file")
    Logger.Info("")
    
    ' === STEP 1: Check Prerequisites ===
    Logger.Info("--- STEP 1: Check Prerequisites ---")
    
    If doc Is Nothing Then
        Logger.Error("No document open!")
        MessageBox.Show("Ava esmalt UUS dokument (ipt, iam, idw).", "Test10")
        Return
    End If
    
    Logger.Info("Document: " & doc.DisplayName)
    
    ' Get extension based on document type
    Dim extension As String = GetExtension(doc)
    Logger.Info("Extension: " & extension)
    
    ' Check if already saved
    If Not String.IsNullOrEmpty(doc.FullFileName) Then
        Logger.Warn("Document is already saved: " & doc.FullFileName)
        Dim cont As DialogResult = MessageBox.Show( _
            "Dokument on juba salvestatud:" & vbCrLf & doc.FullFileName & vbCrLf & vbCrLf & _
            "Kas jätkata siiski? (fail kirjutatakse üle)", _
            "Test10", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
        If cont = DialogResult.No Then Return
    End If
    
    Dim vaultConn As Object = VaultNumberingLib.GetVaultConnection()
    If vaultConn Is Nothing Then
        Logger.Error("NOT connected to Vault!")
        MessageBox.Show("Logi esmalt Vault'i sisse.", "Test10")
        Return
    End If
    
    Logger.Info("Vault: " & vaultConn.Vault & " (User: " & vaultConn.UserName & ")")
    Logger.Info("")
    
    ' === STEP 2: Calculate Local Path ===
    Logger.Info("--- STEP 2: Calculate Local Path ---")
    
    Dim project As Object = app.DesignProjectManager.ActiveDesignProject
    Dim workspacePath As String = project.WorkspacePath.TrimEnd("\"c)
    Logger.Info("Workspace: " & workspacePath)
    
    ' Convert Vault path to local path
    Dim targetLocalPath As String = ConvertVaultPathToLocal(TARGET_VAULT_PATH, workspacePath)
    
    If String.IsNullOrEmpty(targetLocalPath) Then
        Logger.Error("Could not convert Vault path to local!")
        Return
    End If
    
    Logger.Info("Target Vault Path: " & TARGET_VAULT_PATH)
    Logger.Info("Target Local Path: " & targetLocalPath)
    Logger.Info("")
    
    ' === STEP 3: Ensure Folders Exist ===
    Logger.Info("--- STEP 3: Ensure Folders Exist ---")
    
    ' Check/create Vault folder
    Dim vaultFolder As Object = VaultNumberingLib.GetVaultFolder(vaultConn, TARGET_VAULT_PATH)
    If vaultFolder Is Nothing Then
        Logger.Info("Creating Vault folder: " & TARGET_VAULT_PATH)
        vaultFolder = VaultNumberingLib.EnsureVaultFolderRecursive(vaultConn, TARGET_VAULT_PATH)
        If vaultFolder Is Nothing Then
            Logger.Error("Failed to create Vault folder!")
            Return
        End If
    End If
    Logger.Info("Vault folder ID: " & vaultFolder.Id)
    
    ' Create local folder
    If Not System.IO.Directory.Exists(targetLocalPath) Then
        Logger.Info("Creating local folder: " & targetLocalPath)
        System.IO.Directory.CreateDirectory(targetLocalPath)
    End If
    Logger.Info("Local folder exists: True")
    Logger.Info("")
    
    ' === STEP 4: Get File Number from Vault ===
    Logger.Info("--- STEP 4: Get File Number from Vault ---")
    
    Dim scheme As Object = VaultNumberingLib.FindSchemeByName(vaultConn, NUMBERING_SCHEME)
    If scheme Is Nothing Then
        Logger.Error("Numbering scheme not found: " & NUMBERING_SCHEME)
        Logger.Info("Available schemes:")
        Dim names As List(Of String) = VaultNumberingLib.GetSchemeNames(vaultConn)
        For Each name As String In names
            Logger.Info("  - " & name)
        Next
        Return
    End If
    
    Logger.Info("Found scheme: " & scheme.Name)
    
    Dim fileNumber As String = VaultNumberingLib.GenerateFileNumber(vaultConn, scheme)
    If String.IsNullOrEmpty(fileNumber) Then
        Logger.Error("Failed to generate file number!")
        Return
    End If
    
    Logger.Info("Generated number: " & fileNumber)
    
    Dim fileName As String = fileNumber & extension
    Dim fullFilePath As String = targetLocalPath & "\" & fileName
    Logger.Info("Full file path: " & fullFilePath)
    Logger.Info("")
    
    ' === STEP 5: Wait for user to log out (NON-MODAL) ===
    Logger.Info("--- STEP 5: Log Out from Vault (manual) ---")
    Logger.Info("Showing non-modal dialog - user can interact with Inventor")
    
    ' Create a non-modal form for logout instructions
    Dim logoutResult As Boolean = ShowNonModalWaitDialog( _
        "Test10 - Logi Vaultist välja", _
        "Saadud number: " & fileNumber & vbCrLf & _
        "Fail salvestatakse: " & fullFilePath & vbCrLf & vbCrLf & _
        "SAMM 1: Logi Vaultist välja" & vbCrLf & _
        "   File > Vault > Log Out" & vbCrLf & vbCrLf & _
        "See aken EI BLOKEERI Inventorit." & vbCrLf & _
        "Saad menüüsid kasutada.", _
        "Jätka (olen välja loginud)", _
        "Tühista")
    
    If Not logoutResult Then
        Logger.Info("Cancelled by user")
        Return
    End If
    
    ' Verify disconnected
    Dim connAfterLogout As Object = VaultNumberingLib.GetVaultConnection()
    If connAfterLogout Is Nothing Then
        Logger.Info("PASS: Disconnected from Vault")
    Else
        Logger.Warn("Still shows connected - but proceeding anyway")
    End If
    Logger.Info("")
    
    ' === STEP 7: Save Document Locally (should be disconnected now) ===
    Logger.Info("--- STEP 7: Save Document Locally ---")
    
    Try
        ' Set Part Number iProperty to match filename
        Try
            Dim propSets As PropertySets = doc.PropertySets
            Dim designProps As PropertySet = propSets.Item("Design Tracking Properties")
            designProps.Item("Part Number").Value = fileNumber
            Logger.Info("Set Part Number iProperty: " & fileNumber)
        Catch ex As Exception
            Logger.Warn("Could not set Part Number: " & ex.Message)
        End Try
        
        ' Save the document
        app.SilentOperation = True
        doc.SaveAs(fullFilePath, False)
        app.SilentOperation = False
        
        Logger.Info("PASS: Document saved to: " & fullFilePath)
        Logger.Info("File exists: " & System.IO.File.Exists(fullFilePath))
        
    Catch ex As Exception
        app.SilentOperation = False
        Logger.Error("Failed to save: " & ex.Message)
        MessageBox.Show("Salvestamine ebaõnnestus: " & ex.Message, "Test10")
        Return
    End Try
    Logger.Info("")
    
    ' === STEP 8: Log Back In to Vault (NON-MODAL) ===
    Logger.Info("--- STEP 8: Log Back In to Vault ---")
    Logger.Info("Showing non-modal dialog - user can interact with Inventor")
    
    Dim loginResult As Boolean = ShowNonModalWaitDialog( _
        "Test10 - Logi Vaulti sisse", _
        "Fail salvestatud: " & fullFilePath & vbCrLf & vbCrLf & _
        "SAMM 2: Logi Vaulti sisse" & vbCrLf & _
        "   File > Vault > Log In" & vbCrLf & vbCrLf & _
        "See aken EI BLOKEERI Inventorit.", _
        "Jätka (olen sisse loginud)", _
        "Tühista")
    
    If Not loginResult Then
        Logger.Info("Cancelled by user")
        Return
    End If
    
    ' Verify connected
    Dim connAfterLogin As Object = VaultNumberingLib.GetVaultConnection()
    If connAfterLogin IsNot Nothing Then
        Logger.Info("PASS: Connected to Vault: " & connAfterLogin.Vault)
    Else
        Logger.Warn("Not connected to Vault - but proceeding to try check-in anyway")
    End If
    Logger.Info("")
    
    ' === STEP 9: Add File to Vault via API ===
    Logger.Info("--- STEP 9: Add File to Vault via API ---")
    
    Logger.Info("File path: " & fullFilePath)
    Logger.Info("Target Vault folder: " & TARGET_VAULT_PATH)
    Logger.Info("")
    
    Dim addSuccess As Boolean = False
    
    Try
        ' Get fresh connection
        Dim conn As Object = VaultNumberingLib.GetVaultConnection()
        If conn Is Nothing Then
            Throw New Exception("No Vault connection")
        End If
        
        Dim wsm = conn.WebServiceManager
        Dim docService = wsm.DocumentService
        Dim filestoreService = wsm.FilestoreService
        
        ' Get target folder
        Dim targetFolder = docService.GetFolderByPath(TARGET_VAULT_PATH)
        If targetFolder Is Nothing Then
            Throw New Exception("Target folder not found: " & TARGET_VAULT_PATH)
        End If
        Logger.Info("Target folder ID: " & targetFolder.Id)
        
        ' Read file bytes
        Dim fileBytes() As Byte = System.IO.File.ReadAllBytes(fullFilePath)
        Logger.Info("File size: " & fileBytes.Length & " bytes")
        
        ' Upload file to Vault filestore using proper API
        Logger.Info("Uploading to Vault filestore...")
        Dim uploadTicket = UploadFileToVault(wsm, fileName, fileBytes)
        
        If uploadTicket Is Nothing OrElse uploadTicket.Bytes Is Nothing Then
            Throw New Exception("Upload failed - no ticket returned")
        End If
        Logger.Info("Upload ticket received: " & uploadTicket.Bytes.Length & " bytes")
        
        ' Add file to Vault using AddUploadedFile
        Logger.Info("Adding file to Vault folder...")
        Dim addedFile = docService.AddUploadedFile( _
            targetFolder.Id, _
            fileName, _
            "Added via iLogic Test10", _
            DateTime.Now, _
            Nothing, _
            Nothing, _
            Autodesk.Connectivity.WebServices.FileClassification.None, _
            False, _
            uploadTicket)
        
        If addedFile IsNot Nothing Then
            Logger.Info("PASS: File added to Vault!")
            Logger.Info("  File ID: " & addedFile.Id)
            Logger.Info("  File Name: " & addedFile.Name)
            addSuccess = True
        End If
        
    Catch ex As Exception
        Logger.Error("Vault API add failed: " & ex.Message)
        Logger.Info("")
        Logger.Info("Exception type: " & ex.GetType().Name)
        If ex.InnerException IsNot Nothing Then
            Logger.Info("Inner: " & ex.InnerException.Message)
        End If
    End Try
    
    If Not addSuccess Then
        Logger.Warn("Automatic add to Vault failed")
        Logger.Info("")
        Logger.Info("MANUAL CHECK-IN REQUIRED:")
        Logger.Info("  Right-click file in Vault browser > Check In")
        Logger.Info("  Or: File > Vault > Check In")
        
        MessageBox.Show( _
            "Automaatne Vaulti lisamine ei õnnestunud." & vbCrLf & vbCrLf & _
            "Palun tee käsitsi:" & vbCrLf & _
            "File > Vault > Check In" & vbCrLf & vbCrLf & _
            "Fail: " & fullFilePath, _
            "Test10 - Käsitsi check-in", _
            MessageBoxButtons.OK, _
            MessageBoxIcon.Information)
        
        Logger.Info("")
        Logger.Info("=== Test10_DisconnectSaveCheckin: Complete (with manual step required) ===")
        Return
    End If
    
    ' === STEP 10: Get file from Vault to sync ===
    Logger.Info("")
    Logger.Info("--- STEP 10: Get file from Vault to sync ---")
    Logger.Info("Getting file from Vault to establish proper link...")
    
    Try
        ' Close the current document first
        doc.Close(True)  ' True = skip save
        Logger.Info("Document closed")
        
        ' Use Vault Framework to get the file properly
        Dim vdfConn As Autodesk.DataManagement.Client.Framework.Vault.Currency.Connections.Connection
        vdfConn = Connectivity.Application.VaultBase.ConnectionManager.Instance.Connection
        
        If vdfConn Is Nothing Then
            Throw New Exception("No VDF connection")
        End If
        
        ' Find the file we just added
        Dim vaultFilePath As String = TARGET_VAULT_PATH & "/" & fileName
        Logger.Info("Looking for: " & vaultFilePath)
        
        Dim files = vdfConn.WebServiceManager.DocumentService.FindLatestFilesByPaths(New String() {vaultFilePath})
        If files Is Nothing OrElse files.Length = 0 OrElse files(0) Is Nothing Then
            Throw New Exception("Could not find file in Vault: " & vaultFilePath)
        End If
        
        Dim vaultFile = files(0)
        Logger.Info("Found file - ID: " & vaultFile.Id & ", Name: " & vaultFile.Name)
        
        ' Create FileIteration for AcquireFiles
        Dim fileIteration As New Autodesk.DataManagement.Client.Framework.Vault.Currency.Entities.FileIteration(vdfConn, vaultFile)
        
        ' Set up acquire settings - Download to sync local with Vault
        Dim aqSettings As New Autodesk.DataManagement.Client.Framework.Vault.Settings.AcquireFilesSettings(vdfConn, False)
        aqSettings.AddFileToAcquire(fileIteration, Autodesk.DataManagement.Client.Framework.Vault.Settings.AcquireFilesSettings.AcquisitionOption.Download)
        
        ' Execute the download/sync
        Logger.Info("Downloading from Vault to sync...")
        Dim results = vdfConn.FileManager.AcquireFiles(aqSettings)
        
        If results IsNot Nothing AndAlso results.FileResults IsNot Nothing Then
            For Each fileResult In results.FileResults
                Logger.Info("  Result: " & fileResult.File.EntityName & " - " & fileResult.Status.ToString())
            Next
        End If
        
        Logger.Info("PASS: File synced from Vault")
        
        ' Now open the synced file
        System.Threading.Thread.Sleep(500)
        Dim reopenedDoc As Document = app.Documents.Open(fullFilePath, False)
        Logger.Info("Document re-opened: " & reopenedDoc.FullFileName)
        
        MessageBox.Show( _
            "Fail edukalt Vaulti lisatud ja sünkroniseeritud!" & vbCrLf & vbCrLf & _
            "Fail: " & fileName & vbCrLf & _
            "Asukoht: " & TARGET_VAULT_PATH & vbCrLf & vbCrLf & _
            "Staatus peaks nüüd olema 'Latest' või 'Up to date'", _
            "Test10 - Õnnestus", _
            MessageBoxButtons.OK, _
            MessageBoxIcon.Information)
            
    Catch ex As Exception
        Logger.Error("Sync failed: " & ex.Message)
        If ex.InnerException IsNot Nothing Then
            Logger.Info("Inner: " & ex.InnerException.Message)
        End If
        Logger.Info("")
        Logger.Info("You may need to manually: File > Vault > Get")
        
        MessageBox.Show( _
            "Fail on Vaultis, aga sünkroniseerimine ebaõnnestus." & vbCrLf & vbCrLf & _
            "Palun tee käsitsi: File > Vault > Get" & vbCrLf & vbCrLf & _
            "Viga: " & ex.Message, _
            "Test10 - Vajab käsitsi sammu", _
            MessageBoxButtons.OK, _
            MessageBoxIcon.Warning)
    End Try
    
    Logger.Info("")
    Logger.Info("=== Test10_DisconnectSaveCheckin: Complete ===")
    Logger.Info("")
    Logger.Info("If you cancelled, you can delete the test file:")
    Logger.Info("  " & fullFilePath)
End Sub

' Get file extension based on document type
Function GetExtension(doc As Document) As String
    Select Case doc.DocumentType
        Case DocumentTypeEnum.kPartDocumentObject
            Return ".ipt"
        Case DocumentTypeEnum.kAssemblyDocumentObject
            Return ".iam"
        Case DocumentTypeEnum.kDrawingDocumentObject
            Return ".idw"
        Case Else
            Return ".ipt"
    End Select
End Function

' Convert Vault path to local path
Function ConvertVaultPathToLocal(vaultPath As String, workspacePath As String) As String
    If String.IsNullOrEmpty(vaultPath) Then Return ""
    
    ' Normalize
    vaultPath = vaultPath.TrimEnd("/"c)
    workspacePath = workspacePath.TrimEnd("\"c)
    
    ' Detect actual mapping from workspace path
    ' Workspace: C:\_SoftcomVault\Tooted -> maps to $/Tooted
    Dim workspaceParts() As String = workspacePath.Split("\"c)
    Dim actualVaultMapping As String = "$"
    
    ' Find Vault root indicator in path
    For i As Integer = 0 To workspaceParts.Length - 2
        Dim part As String = workspaceParts(i).ToLower()
        If part.Contains("vault") OrElse part.Contains("softcom") Then
            If i < workspaceParts.Length - 1 Then
                Dim subfolders As String = String.Join("/", workspaceParts, i + 1, workspaceParts.Length - i - 1)
                actualVaultMapping = "$/" & subfolders
                Exit For
            End If
        End If
    Next
    
    Logger.Info("  Detected mapping: " & actualVaultMapping & " -> " & workspacePath)
    
    ' Convert
    If vaultPath.StartsWith(actualVaultMapping, StringComparison.OrdinalIgnoreCase) Then
        Dim relativePath As String = vaultPath.Substring(actualVaultMapping.Length).TrimStart("/"c)
        If String.IsNullOrEmpty(relativePath) Then
            Return workspacePath
        End If
        Return workspacePath & "\" & relativePath.Replace("/", "\")
    ElseIf actualVaultMapping.StartsWith(vaultPath, StringComparison.OrdinalIgnoreCase) Then
        ' Target is parent of workspace
        Dim remainder As String = actualVaultMapping.Substring(vaultPath.Length).TrimStart("/"c)
        Dim levelsUp As Integer = remainder.Split("/"c).Length
        Dim localDir As String = workspacePath
        For i As Integer = 1 To levelsUp
            localDir = System.IO.Path.GetDirectoryName(localDir)
        Next
        Return localDir
    End If
    
    Return ""
End Function

' Show a non-modal dialog that allows Inventor interaction
' Returns True if user clicked Continue, False if Cancel
Function ShowNonModalWaitDialog(title As String, message As String, continueText As String, cancelText As String) As Boolean
    Dim result As Boolean = False
    Dim finished As Boolean = False
    
    ' Create form
    Dim frm As New Form()
    frm.Text = title
    frm.Width = 450
    frm.Height = 280
    frm.StartPosition = FormStartPosition.Manual
    ' Position in top-right area to not block menus (use fixed offset from right)
    frm.Left = 50
    frm.Top = 150
    frm.FormBorderStyle = FormBorderStyle.FixedDialog
    frm.MaximizeBox = False
    frm.MinimizeBox = False
    frm.TopMost = True  ' Keep on top but don't block
    
    ' Message label
    Dim lbl As New Label()
    lbl.Text = message
    lbl.SetBounds(15, 15, 410, 150)
    frm.Controls.Add(lbl)
    
    ' Continue button
    Dim btnContinue As New Button()
    btnContinue.Text = continueText
    btnContinue.SetBounds(120, 180, 200, 35)
    AddHandler btnContinue.Click, Sub(s, e)
        result = True
        finished = True
        frm.Close()
    End Sub
    frm.Controls.Add(btnContinue)
    
    ' Cancel button
    Dim btnCancel As New Button()
    btnCancel.Text = cancelText
    btnCancel.SetBounds(330, 180, 90, 35)
    AddHandler btnCancel.Click, Sub(s, e)
        result = False
        finished = True
        frm.Close()
    End Sub
    frm.Controls.Add(btnCancel)
    
    ' Handle form close via X button
    AddHandler frm.FormClosing, Sub(s, e)
        If Not finished Then
            result = False
            finished = True
        End If
    End Sub
    
    ' Show non-modal
    frm.Show()
    
    ' Process events while waiting for user to click
    Do While Not finished
        System.Windows.Forms.Application.DoEvents()
        System.Threading.Thread.Sleep(50)
    Loop
    
    Return result
End Function

' Upload file to Vault filestore and return upload ticket
' Based on Autodesk forum solution for proper file upload
Function UploadFileToVault(wsm As Object, filename As String, fileContents As Byte()) As Autodesk.Connectivity.WebServices.ByteArray
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
    
    Logger.Info("  Uploading " & bytesTotal & " bytes...")
    
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
        Logger.Info("  Transferred: " & bytesTransferred & " / " & bytesTotal)
        
    Loop While bytesTransferred < bytesTotal
    
    Return uploadTicket
End Function
