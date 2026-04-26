' Copyright (c) 2026 Henri Kuuste
' ============================================================================
' Loo alusmoodul - Create base module folder structure
' 
' Creates the standard folder structure for a parametric base module:
'   Alusmoodulid/<ModuleName>/
'     Eskiis/
'     Karkass/Detailid/, Karkass/Joonised/
'     Poroloon/Detailid/, Poroloon/Joonised/
'
' Folders are created both on disk and in Vault (if connected).
'
' Usage: Run from any open document in the target project
' ============================================================================

' References must come FIRST, before any AddVbFile
AddReference "Autodesk.Connectivity.WebServices"
AddReference "Autodesk.DataManagement.Client.Framework.Vault"
AddReference "Connectivity.InventorAddin.EdmAddin"

' Libraries
AddVbFile "Lib/UtilsLib.vb"
AddVbFile "Lib/VaultNumberingLib.vb"

Imports System.Windows.Forms
Imports Inventor

Sub Main()
    Dim app As Inventor.Application = ThisApplication
    UtilsLib.SetLogger(Logger)
    
    UtilsLib.LogInfo("Loo alusmoodul: Starting...")
    
    ' Get default project path from active document
    Dim defaultProjectPath As String = ""
    If app.ActiveDocument IsNot Nothing Then
        defaultProjectPath = UtilsLib.GetProjectPath(app.ActiveDocument.FullDocumentName)
    End If
    
    If String.IsNullOrEmpty(defaultProjectPath) Then
        UtilsLib.LogWarn("Loo alusmoodul: Could not detect project path from active document")
    Else
        UtilsLib.LogInfo("Loo alusmoodul: Detected project path: " & defaultProjectPath)
    End If
    
    ' Get Vault connection
    Dim vaultConn As Object = VaultNumberingLib.GetVaultConnection()
    Dim vaultConnected As Boolean = (vaultConn IsNot Nothing)
    
    If vaultConnected Then
        UtilsLib.LogInfo("Loo alusmoodul: Vault connected - " & VaultNumberingLib.GetConnectionInfo(vaultConn))
    Else
        UtilsLib.LogWarn("Loo alusmoodul: Vault not connected - folders will be created locally only")
    End If
    
    ' Get workspace root for Vault path conversion
    Dim workspaceRoot As String = ""
    If vaultConnected AndAlso Not String.IsNullOrEmpty(defaultProjectPath) Then
        workspaceRoot = VaultNumberingLib.DetectWorkspaceRoot(vaultConn, defaultProjectPath)
    End If
    
    ' Show dialog and get user input
    Dim projectPath As String = defaultProjectPath
    Dim moduleName As String = ""
    
    Dim result As DialogResult = ShowInputDialog(projectPath, moduleName)
    
    If result <> DialogResult.OK Then
        UtilsLib.LogInfo("Loo alusmoodul: Cancelled by user")
        Exit Sub
    End If
    
    ' Validate input
    If String.IsNullOrEmpty(projectPath) Then
        UtilsLib.LogError("Loo alusmoodul: Project path is required")
        MessageBox.Show("Projekti kaust on kohustuslik.", "Loo alusmoodul")
        Exit Sub
    End If
    
    If String.IsNullOrEmpty(moduleName) Then
        UtilsLib.LogError("Loo alusmoodul: Module name is required")
        MessageBox.Show("Mooduli nimi on kohustuslik.", "Loo alusmoodul")
        Exit Sub
    End If
    
    ' Validate module name (no invalid characters)
    Dim invalidChars() As Char = System.IO.Path.GetInvalidFileNameChars()
    For Each c As Char In invalidChars
        If moduleName.Contains(c) Then
            UtilsLib.LogError("Loo alusmoodul: Invalid character in module name: " & c)
            MessageBox.Show("Mooduli nimes on keelatud sümbol: " & c, "Loo alusmoodul")
            Exit Sub
        End If
    Next
    
    UtilsLib.LogInfo("Loo alusmoodul: Creating module '" & moduleName & "' in " & projectPath)
    
    ' Update workspace root if project path changed
    If vaultConnected AndAlso projectPath <> defaultProjectPath Then
        workspaceRoot = VaultNumberingLib.DetectWorkspaceRoot(vaultConn, projectPath)
    End If
    
    ' Create the folder structure
    Dim foldersToCreate As New System.Collections.Generic.List(Of String)
    
    Dim alusmoodulid As String = System.IO.Path.Combine(projectPath, "Alusmoodulid")
    Dim moduleRoot As String = System.IO.Path.Combine(alusmoodulid, moduleName)
    
    foldersToCreate.Add(System.IO.Path.Combine(moduleRoot, "Eskiis"))
    foldersToCreate.Add(System.IO.Path.Combine(moduleRoot, "Karkass", "Detailid"))
    foldersToCreate.Add(System.IO.Path.Combine(moduleRoot, "Karkass", "Joonised"))
    foldersToCreate.Add(System.IO.Path.Combine(moduleRoot, "Poroloon", "Detailid"))
    foldersToCreate.Add(System.IO.Path.Combine(moduleRoot, "Poroloon", "Joonised"))
    
    Dim successCount As Integer = 0
    Dim failCount As Integer = 0
    
    For Each folderPath As String In foldersToCreate
        Try
            ' Create local folder (CreateDirectory handles nested paths)
            If Not System.IO.Directory.Exists(folderPath) Then
                System.IO.Directory.CreateDirectory(folderPath)
                UtilsLib.LogInfo("Loo alusmoodul: Created local folder: " & folderPath)
            Else
                UtilsLib.LogInfo("Loo alusmoodul: Folder already exists: " & folderPath)
            End If
            
            ' Create in Vault if connected
            If vaultConnected AndAlso Not String.IsNullOrEmpty(workspaceRoot) Then
                Dim vaultPath As String = VaultNumberingLib.ConvertLocalPathToVaultPath(folderPath, workspaceRoot)
                If Not String.IsNullOrEmpty(vaultPath) Then
                    Dim vaultFolder As Object = VaultNumberingLib.EnsureVaultFolderRecursive(vaultConn, vaultPath)
                    If vaultFolder IsNot Nothing Then
                        UtilsLib.LogInfo("Loo alusmoodul: Vault folder ready: " & vaultPath)
                    Else
                        UtilsLib.LogWarn("Loo alusmoodul: Could not create Vault folder: " & vaultPath)
                    End If
                End If
            End If
            
            successCount += 1
        Catch ex As Exception
            UtilsLib.LogError("Loo alusmoodul: Failed to create folder: " & folderPath & " - " & ex.Message)
            failCount += 1
        End Try
    Next
    
    ' Summary
    If failCount = 0 Then
        UtilsLib.LogInfo("Loo alusmoodul: Successfully created " & successCount & " folders for module '" & moduleName & "'")
        MessageBox.Show("Moodul '" & moduleName & "' loodud edukalt!" & vbCrLf & _
                       "Kaustu loodud: " & successCount, "Loo alusmoodul")
    Else
        UtilsLib.LogWarn("Loo alusmoodul: Created " & successCount & " folders, " & failCount & " failed")
        MessageBox.Show("Moodul '" & moduleName & "' loomine osaliselt ebaõnnestus." & vbCrLf & _
                       "Õnnestus: " & successCount & ", Ebaõnnestus: " & failCount, "Loo alusmoodul")
    End If
End Sub

Function ShowInputDialog(ByRef projectPath As String, ByRef moduleName As String) As DialogResult
    Dim frm As New System.Windows.Forms.Form()
    frm.Text = "Loo alusmoodul"
    frm.Width = 500
    frm.Height = 200
    frm.StartPosition = FormStartPosition.CenterScreen
    frm.FormBorderStyle = FormBorderStyle.FixedDialog
    frm.MaximizeBox = False
    frm.MinimizeBox = False
    
    Dim yPos As Integer = 15
    
    ' Project path label
    Dim lblProject As New System.Windows.Forms.Label()
    lblProject.Text = "Projekti kaust:"
    lblProject.Left = 15
    lblProject.Top = yPos
    lblProject.Width = 100
    frm.Controls.Add(lblProject)
    
    ' Project path textbox
    Dim txtProject As New System.Windows.Forms.TextBox()
    txtProject.Name = "txtProject"
    txtProject.Text = projectPath
    txtProject.Left = 120
    txtProject.Top = yPos
    txtProject.Width = 270
    frm.Controls.Add(txtProject)
    
    ' Browse button
    Dim btnBrowse As New System.Windows.Forms.Button()
    btnBrowse.Text = "..."
    btnBrowse.Left = 395
    btnBrowse.Top = yPos - 2
    btnBrowse.Width = 40
    btnBrowse.Height = 23
    AddHandler btnBrowse.Click, Sub(s, e)
        Dim fbd As New FolderBrowserDialog()
        fbd.Description = "Vali projekti kaust"
        fbd.ShowNewFolderButton = True
        If Not String.IsNullOrEmpty(txtProject.Text) AndAlso System.IO.Directory.Exists(txtProject.Text) Then
            fbd.SelectedPath = txtProject.Text
        End If
        If fbd.ShowDialog() = DialogResult.OK Then
            txtProject.Text = fbd.SelectedPath
        End If
    End Sub
    frm.Controls.Add(btnBrowse)
    
    yPos += 35
    
    ' Module name label
    Dim lblModule As New System.Windows.Forms.Label()
    lblModule.Text = "Mooduli nimi:"
    lblModule.Left = 15
    lblModule.Top = yPos
    lblModule.Width = 100
    frm.Controls.Add(lblModule)
    
    ' Module name textbox
    Dim txtModule As New System.Windows.Forms.TextBox()
    txtModule.Name = "txtModule"
    txtModule.Text = ""
    txtModule.Left = 120
    txtModule.Top = yPos
    txtModule.Width = 315
    frm.Controls.Add(txtModule)
    
    yPos += 45
    
    ' OK button
    Dim btnOK As New System.Windows.Forms.Button()
    btnOK.Text = "Loo"
    btnOK.Left = 250
    btnOK.Top = yPos
    btnOK.Width = 80
    btnOK.Height = 28
    btnOK.DialogResult = DialogResult.OK
    frm.Controls.Add(btnOK)
    frm.AcceptButton = btnOK
    
    ' Cancel button
    Dim btnCancel As New System.Windows.Forms.Button()
    btnCancel.Text = "Tühista"
    btnCancel.Left = 340
    btnCancel.Top = yPos
    btnCancel.Width = 80
    btnCancel.Height = 28
    btnCancel.DialogResult = DialogResult.Cancel
    frm.Controls.Add(btnCancel)
    frm.CancelButton = btnCancel
    
    ' Show dialog
    Dim result As DialogResult = frm.ShowDialog()
    
    ' Read values after dialog closes (avoiding ByRef in lambda issues)
    If result = DialogResult.OK Then
        projectPath = txtProject.Text.Trim()
        moduleName = txtModule.Text.Trim()
    End If
    
    Return result
End Function
