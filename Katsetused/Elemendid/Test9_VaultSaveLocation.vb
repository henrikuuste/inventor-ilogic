' Copyright (c) 2026 Henri Kuuste
' Test9_VaultSaveLocation.vb
' PURPOSE: Research test to understand Vault "New File" dialog location behavior
' 
' CONCLUSION: There is NO API to set the default folder in Vault's "New File" dialog
' 
' FINDINGS:
' 1. The Vault "New File" dialog uses a "last used folder" stored internally
' 2. This folder is NOT based on document's local save path
' 3. No public API exists to modify this folder from iLogic
' 4. FileUIEvents.OnFileSaveAsDialog is intercepted by Vault Add-in
' 5. Data Standard (if installed) can customize via GetParentFolderName function
'
' RUN: Open any NEW document while connected to Vault

AddReference "Autodesk.Connectivity.WebServices"
AddReference "Autodesk.DataManagement.Client.Framework.Vault"
AddReference "Connectivity.InventorAddin.EdmAddin"
AddVbFile "Lib/StringsLib.vb"
AddVbFile "Lib/UtilsLib.vb"
AddVbFile "Lib/VaultNumberingLib.vb"

Sub Main()
    Dim app As Inventor.Application = ThisApplication
    Dim doc As Document = app.ActiveDocument
    
    Logger.Info("=== Test9_VaultSaveLocation: Research Results ===")
    Logger.Info("")
    
    ' === Check Environment ===
    Logger.Info("--- Environment Check ---")
    
    If doc Is Nothing Then
        Logger.Error("No document open!")
        MessageBox.Show("Ava esmalt mingi dokument.", "Test9")
        Return
    End If
    
    Logger.Info("Document: " & doc.DisplayName)
    Logger.Info("Saved: " & If(String.IsNullOrEmpty(doc.FullFileName), "NO (new document)", "YES"))
    
    Dim vaultConn As Object = VaultNumberingLib.GetVaultConnection()
    If vaultConn Is Nothing Then
        Logger.Error("NOT connected to Vault!")
        MessageBox.Show("Logi esmalt Vault'i sisse.", "Test9")
        Return
    End If
    
    Logger.Info("Vault: " & vaultConn.Vault)
    Logger.Info("")
    
    ' === IPJ Mapping Info ===
    Logger.Info("--- IPJ Mapping (informational) ---")
    
    Dim project As Object = app.DesignProjectManager.ActiveDesignProject
    Logger.Info("Workspace: " & project.WorkspacePath)
    
    Try
        Logger.Info("VirtualVaultPath: " & project.VirtualVaultPath)
    Catch
        Logger.Info("VirtualVaultPath: $ (default)")
    End Try
    Logger.Info("")
    
    ' === Research Conclusions ===
    Logger.Info("=== RESEARCH CONCLUSIONS ===")
    Logger.Info("")
    Logger.Info("The Vault 'New File' dialog folder CANNOT be set via API.")
    Logger.Info("")
    Logger.Info("Tested approaches that DO NOT WORK:")
    Logger.Info("  1. Saving document to local folder - Vault ignores document path")
    Logger.Info("  2. WorkingFoldersManager.SetWorkingFolder() - Inventor ignores it")
    Logger.Info("  3. FileUIEvents.OnFileSaveAsDialog - Vault Add-in intercepts first")
    Logger.Info("  4. Setting $Prop['_FilePath'] - Read-only in dialog context")
    Logger.Info("")
    Logger.Info("The dialog uses 'last used folder' stored internally by Vault Add-in.")
    Logger.Info("")
    Logger.Info("ALTERNATIVES:")
    Logger.Info("  1. Data Standard - GetParentFolderName function in Default.ps1")
    Logger.Info("     Location: C:\ProgramData\Autodesk\Vault <ver>\Extensions\DataStandard\CAD\addins\")
    Logger.Info("  2. Manual selection - User browses to folder each time")
    Logger.Info("  3. coolOrange powerEvents - May offer customization options")
    Logger.Info("")
    Logger.Info("See: docs/research/2026-04-26-vault-new-file-location.md")
    Logger.Info("")
    
    ' === Demonstrate current behavior ===
    Dim result As DialogResult = MessageBox.Show( _
        "JÄRELDUS: Vault dialoogi kausta EI SAA API kaudu seada." & vbCrLf & vbCrLf & _
        "Vault kasutab 'viimati kasutatud kausta' mis on salvestatud" & vbCrLf & _
        "Vault Add-in'i sisemiselt ja pole API kaudu kättesaadav." & vbCrLf & vbCrLf & _
        "ALTERNATIIVID:" & vbCrLf & _
        "1. Data Standard - kui on installitud, saab kasutada" & vbCrLf & _
        "   GetParentFolderName funktsiooni Default.ps1 failis" & vbCrLf & _
        "2. Kasutaja valib kausta käsitsi" & vbCrLf & vbCrLf & _
        "Kas avada Vault salvestamise dialoog demonstratsiooniks?", _
        "Test9 - Tulemused", _
        MessageBoxButtons.YesNo, _
        MessageBoxIcon.Information)
    
    If result = DialogResult.Yes Then
        Logger.Info("Opening Vault check-in dialog for demonstration...")
        Logger.Info("NOTE: The folder shown is the 'last used' folder, not controllable via API.")
        
        Try
            Dim cmdMgr As CommandManager = app.CommandManager
            Dim ctrlDef As ControlDefinition = cmdMgr.ControlDefinitions.Item("VaultCheckin")
            ctrlDef.Execute()
        Catch ex As Exception
            Logger.Warn("Could not trigger Vault dialog: " & ex.Message)
        End Try
    End If
    
    Logger.Info("")
    Logger.Info("=== Test9_VaultSaveLocation: Complete ===")
End Sub
