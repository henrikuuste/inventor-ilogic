' Copyright (c) 2026 Henri Kuuste
' ============================================================================
' Loo aluselement - Create base element folder structure
' 
' Creates the standard folder structure for a parametric base element:
'   Aluselemendid/<ElementName>/
'     Eskiis/
'     Karkass/Detailid/, Karkass/Joonised/
'     Poroloon/Detailid/, Poroloon/Joonised/
'
' Folders are created both on disk and in Vault (if connected).
'
' Terminology updated 2026-05-12 per docs/UBIQUITOUS_LANGUAGE.md:
'   - "Alusmoodul" (old) → "Aluselement" (base element)
'
' Usage: Run from any open document in the target project
' ============================================================================

' References must come FIRST, before any AddVbFile
AddReference "Autodesk.Connectivity.WebServices"
AddReference "Autodesk.DataManagement.Client.Framework.Vault"
AddReference "Connectivity.InventorAddin.EdmAddin"

' Libraries
AddVbFile "Lib/RuntimeLib.vb"
AddVbFile "Lib/StringsLib.vb"
AddVbFile "Lib/UtilsLib.vb"
AddVbFile "Lib/VaultNumberingLib.vb"
AddVbFile "Lib/BaseElementLayoutLib.vb"

Imports System.Windows.Forms
Imports Inventor

Sub Main()
    If Not AppRuntime.Initialize(ThisApplication) Then Return
    
    Dim app As Inventor.Application = ThisApplication
    UtilsLib.SetLogger(Logger)
    
    UtilsLib.LogInfo("Loo aluselement: Starting...")
    
    ' Get default project path from active document
    Dim defaultProjectPath As String = ""
    If app.ActiveDocument IsNot Nothing Then
        defaultProjectPath = UtilsLib.GetProjectPath(app.ActiveDocument.FullDocumentName)
    End If
    
    If String.IsNullOrEmpty(defaultProjectPath) Then
        UtilsLib.LogWarn("Loo aluselement: Could not detect project path from active document")
    Else
        UtilsLib.LogInfo("Loo aluselement: Detected project path: " & defaultProjectPath)
    End If
    
    ' Get Vault connection
    Dim vaultConn As Object = VaultNumberingLib.GetVaultConnection()
    Dim vaultConnected As Boolean = (vaultConn IsNot Nothing)
    
    If vaultConnected Then
        UtilsLib.LogInfo("Loo aluselement: Vault connected - " & VaultNumberingLib.GetConnectionInfo(vaultConn))
    Else
        UtilsLib.LogWarn("Loo aluselement: Vault not connected - folders will be created locally only")
    End If
    
    ' Get workspace root for Vault path conversion
    Dim workspaceRoot As String = ""
    If vaultConnected AndAlso Not String.IsNullOrEmpty(defaultProjectPath) Then
        workspaceRoot = VaultNumberingLib.DetectWorkspaceRoot(vaultConn, defaultProjectPath)
    End If
    
    ' Show dialog and get user input
    Dim projectPath As String = defaultProjectPath
    Dim elementName As String = ""
    
    Dim result As DialogResult = ShowInputDialog(projectPath, elementName)
    
    If result <> DialogResult.OK Then
        UtilsLib.LogInfo("Loo aluselement: Cancelled by user")
        Exit Sub
    End If
    
    ' Validate input
    If String.IsNullOrEmpty(projectPath) Then
        UtilsLib.LogError("Loo aluselement: Project path is required")
        MessageBox.Show("Projekti kaust on kohustuslik.", StringsLib.TITLE_CREATE_BASE_ELEMENT)
        Exit Sub
    End If
    
    If String.IsNullOrEmpty(elementName) Then
        UtilsLib.LogError("Loo aluselement: Element name is required")
        MessageBox.Show("Elemendi nimi on kohustuslik.", StringsLib.TITLE_CREATE_BASE_ELEMENT)
        Exit Sub
    End If
    
    ' Validate element name (no invalid characters)
    Dim invalidChars() As Char = System.IO.Path.GetInvalidFileNameChars()
    For Each c As Char In invalidChars
        If elementName.Contains(c) Then
            UtilsLib.LogError("Loo aluselement: Invalid character in element name: " & c)
            MessageBox.Show("Elemendi nimes on keelatud sümbol: " & c, StringsLib.TITLE_CREATE_BASE_ELEMENT)
            Exit Sub
        End If
    Next
    
    UtilsLib.LogInfo("Loo aluselement: Creating element '" & elementName & "' in " & projectPath)
    
    ' Update workspace root if project path changed
    If vaultConnected AndAlso projectPath <> defaultProjectPath Then
        workspaceRoot = VaultNumberingLib.DetectWorkspaceRoot(vaultConn, projectPath)
    End If
    
    ' Create the folder structure via shared library
    Dim elementRoot As String = BaseElementLayoutLib.EnsureBaseElementLayout(projectPath, elementName, vaultConn, workspaceRoot)
    Dim expectedFolders As System.Collections.Generic.List(Of String) = BaseElementLayoutLib.EnumerateExpectedFolders(projectPath, elementName)
    Dim successCount As Integer = 0
    Dim failCount As Integer = 0
    For Each folderPath As String In expectedFolders
        If System.IO.Directory.Exists(folderPath) Then
            successCount += 1
        Else
            failCount += 1
        End If
    Next
    
    ' Summary
    If failCount = 0 Then
        UtilsLib.LogInfo("Loo aluselement: Successfully created " & successCount & " folders for element '" & elementName & "'")
        MessageBox.Show("Element '" & elementName & "' loodud edukalt!" & vbCrLf & _
                       "Kaustu loodud: " & successCount, StringsLib.TITLE_CREATE_BASE_ELEMENT)
    Else
        UtilsLib.LogWarn("Loo aluselement: Created " & successCount & " folders, " & failCount & " failed")
        MessageBox.Show("Element '" & elementName & "' loomine osaliselt ebaõnnestus." & vbCrLf & _
                       "Õnnestus: " & successCount & ", Ebaõnnestus: " & failCount, StringsLib.TITLE_CREATE_BASE_ELEMENT)
    End If
End Sub

Function ShowInputDialog(ByRef projectPath As String, ByRef elementName As String) As DialogResult
    Dim frm As New System.Windows.Forms.Form()
    frm.Text = StringsLib.TITLE_CREATE_BASE_ELEMENT
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
    
    ' Element name label
    Dim lblModule As New System.Windows.Forms.Label()
    lblModule.Text = "Elemendi nimi:"
    lblModule.Left = 15
    lblModule.Top = yPos
    lblModule.Width = 100
    frm.Controls.Add(lblModule)
    
    ' Element name textbox
    Dim txtElement As New System.Windows.Forms.TextBox()
    txtElement.Name = "txtElement"
    txtElement.Text = ""
    txtElement.Left = 120
    txtElement.Top = yPos
    txtElement.Width = 315
    frm.Controls.Add(txtElement)
    
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
        elementName = txtElement.Text.Trim()
    End If
    
    Return result
End Function
