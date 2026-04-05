' ============================================================================
' Loo 1:1 joonised - Create 1:1 scale drawings for CAM applications
' 
' Features:
' - Create drawings with all 6 orthographic views at 1:1 scale
' - Auto-size sheet to fit part extents with dimension space (50% padding)
' - Add extent dimensions to all views
' - Shows existing drawing status for each part (by Part Number)
' - Copies Description and Project properties from part to drawing
' - Stores BB_SourcePartNumber for part-drawing association
'
' Usage: 
' - From part document: Creates drawing for active part
' - From assembly: Select components or use all parts (with checkboxes)
'
' Template: Uses Drawing.1.1.idw (must exist in templates folder)
' ============================================================================

' References for Vault integration
AddReference "Autodesk.Connectivity.WebServices"
AddReference "Autodesk.DataManagement.Client.Framework.Vault"
AddReference "Autodesk.DataManagement.Client.Framework.Vault.Forms"
AddReference "Connectivity.InventorAddin.EdmAddin"

' Libraries
AddVbFile "Lib/VaultNumberingLib.vb"
AddVbFile "Lib/CAMDrawingLib.vb"

Imports System.Collections.Generic
Imports System.Windows.Forms
Imports Inventor

Sub Main()
    Dim app As Inventor.Application = ThisApplication
    Dim logs As New List(Of String)
    
    ' Validate document
    If app.ActiveDocument Is Nothing Then
        Logger.Error("Loo 1:1 joonised: No active document")
        MessageBox.Show("Ava esmalt detail või koost.", "Loo 1:1 joonised")
        Exit Sub
    End If
    
    Dim doc As Document = app.ActiveDocument
    Dim docType As DocumentTypeEnum = doc.DocumentType
    
    If docType <> DocumentTypeEnum.kPartDocumentObject AndAlso _
       docType <> DocumentTypeEnum.kAssemblyDocumentObject Then
        Logger.Error("Loo 1:1 joonised: Invalid document type")
        MessageBox.Show("See reegel töötab ainult detaili või koostuga.", "Loo 1:1 joonised")
        Exit Sub
    End If
    
    Logger.Info("Loo 1:1 joonised: Starting for " & doc.DisplayName)
    
    ' Get Vault connection
    Dim vaultConn As Object = VaultNumberingLib.GetVaultConnection()
    Dim vaultConnected As Boolean = (vaultConn IsNot Nothing)
    
    If vaultConnected Then
        Logger.Info("Loo 1:1 joonised: Vault connected - " & VaultNumberingLib.GetConnectionInfo(vaultConn))
    Else
        Logger.Warn("Loo 1:1 joonised: Vault not connected")
    End If
    
    ' Get workspace root for Vault path conversion
    Dim workspaceRoot As String = ""
    If vaultConnected Then
        Dim docFolder As String = System.IO.Path.GetDirectoryName(doc.FullDocumentName)
        workspaceRoot = VaultNumberingLib.DetectWorkspaceRoot(vaultConn, docFolder, logs)
        For Each log As String In logs : Logger.Info(log) : Next
        logs.Clear()
    End If
    
    ' Collect part data: List of (PartDocument, PartNumber, DisplayName, HasDrawing, Selected)
    ' Using parallel lists instead of custom class to avoid iLogic type exposure issues
    Dim partDocs As New List(Of PartDocument)
    Dim partNumbers As New List(Of String)
    Dim displayNames As New List(Of String)
    Dim hasDrawings As New List(Of Boolean)
    Dim selectedFlags As New List(Of Boolean)
    
    If docType = DocumentTypeEnum.kPartDocumentObject Then
        ' Single part
        Dim partDoc As PartDocument = CType(doc, PartDocument)
        partDocs.Add(partDoc)
        partNumbers.Add(CAMDrawingLib.GetPartNumber(partDoc))
        displayNames.Add(partDoc.DisplayName)
        hasDrawings.Add(False)
        selectedFlags.Add(True)
    Else
        ' Assembly - get unique parts from occurrences
        Dim asmDoc As AssemblyDocument = CType(doc, AssemblyDocument)
        Dim partPaths As New HashSet(Of String)
        
        ' Check if user has selected specific occurrences
        Dim selectedOccurrences As New List(Of ComponentOccurrence)
        For Each obj As Object In asmDoc.SelectSet
            If TypeOf obj Is ComponentOccurrence Then
                selectedOccurrences.Add(CType(obj, ComponentOccurrence))
            End If
        Next
        
        ' Get parts from selection or all occurrences
        Dim occurrencesToProcess As IEnumerable(Of ComponentOccurrence)
        If selectedOccurrences.Count > 0 Then
            occurrencesToProcess = selectedOccurrences
            Logger.Info("Loo 1:1 joonised: Using " & selectedOccurrences.Count & " selected occurrence(s)")
        Else
            occurrencesToProcess = CType(asmDoc.ComponentDefinition.Occurrences, IEnumerable(Of ComponentOccurrence))
            Logger.Info("Loo 1:1 joonised: No selection, using all occurrences")
        End If
        
        For Each occ As ComponentOccurrence In occurrencesToProcess
            Try
                If occ.DefinitionDocumentType = DocumentTypeEnum.kPartDocumentObject Then
                    Dim partPath As String = occ.ReferencedFileDescriptor.FullFileName
                    If Not partPaths.Contains(partPath) Then
                        partPaths.Add(partPath)
                        Dim partDoc As PartDocument = CType(occ.Definition.Document, PartDocument)
                        partDocs.Add(partDoc)
                        partNumbers.Add(CAMDrawingLib.GetPartNumber(partDoc))
                        displayNames.Add(partDoc.DisplayName)
                        hasDrawings.Add(False)
                        selectedFlags.Add(True)
                    End If
                End If
            Catch
            End Try
        Next
    End If
    
    If partDocs.Count = 0 Then
        Logger.Error("Loo 1:1 joonised: No parts found")
        MessageBox.Show("Detaile ei leitud.", "Loo 1:1 joonised")
        Exit Sub
    End If
    
    Logger.Info("Loo 1:1 joonised: Found " & partDocs.Count & " part(s)")
    
    ' Default output folder (same as part/assembly folder)
    Dim outputFolder As String = System.IO.Path.GetDirectoryName(doc.FullDocumentName)
    
    ' Track existing drawing paths
    Dim existingDrawingPaths As New List(Of String)
    For i As Integer = 0 To partDocs.Count - 1
        existingDrawingPaths.Add("")
    Next
    
    ' Search root for existing drawings - use workspace root if available, otherwise doc folder
    Dim searchRoot As String = If(Not String.IsNullOrEmpty(workspaceRoot), workspaceRoot, outputFolder)
    Logger.Info("Loo 1:1 joonised: Drawing search root: " & searchRoot)
    
    ' Check for existing 1:1 drawings (in open documents and on disk)
    For i As Integer = 0 To partDocs.Count - 1
        If Not String.IsNullOrEmpty(partNumbers(i)) Then
            Dim foundPath As String = CAMDrawingLib.FindDrawingForPart( _
                partNumbers(i), searchRoot, app, logs, CAMDrawingLib.DRAWING_TYPE_1TO1, True)
            
            If Not String.IsNullOrEmpty(foundPath) Then
                hasDrawings(i) = True
                existingDrawingPaths(i) = foundPath
            End If
        End If
    Next
    For Each log As String In logs : Logger.Info(log) : Next
    logs.Clear()
    
    ' ========================================================================
    ' Show Dialog
    ' ========================================================================
    
    Dim frm As New System.Windows.Forms.Form()
    frm.Text = "Loo 1:1 joonised"
    frm.Width = 750
    frm.Height = 500
    frm.StartPosition = FormStartPosition.CenterScreen
    frm.FormBorderStyle = FormBorderStyle.Sizable
    frm.MinimizeBox = True
    frm.MaximizeBox = True
    
    Dim currentY As Integer = 10
    
    ' Parts list header
    Dim lblParts As New System.Windows.Forms.Label()
    lblParts.Text = "Detailid (" & partDocs.Count & "):"
    lblParts.Left = 10
    lblParts.Top = currentY
    lblParts.Width = 200
    frm.Controls.Add(lblParts)
    
    currentY += 20
    
    ' DataGridView for parts
    Dim dgv As New System.Windows.Forms.DataGridView()
    dgv.Name = "dgvParts"
    dgv.Left = 10
    dgv.Top = currentY
    dgv.Width = 710
    dgv.Height = 280
    dgv.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Bottom
    dgv.AllowUserToAddRows = False
    dgv.AllowUserToDeleteRows = False
    dgv.RowHeadersVisible = False
    dgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect
    dgv.MultiSelect = False
    dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
    
    ' Column: Selected (checkbox)
    Dim colSelected As New DataGridViewCheckBoxColumn()
    colSelected.Name = "colSelected"
    colSelected.HeaderText = "Vali"
    colSelected.Width = 40
    dgv.Columns.Add(colSelected)
    
    ' Column: Part Number
    Dim colPartNum As New DataGridViewTextBoxColumn()
    colPartNum.Name = "colPartNum"
    colPartNum.HeaderText = "Artikkel"
    colPartNum.Width = 120
    colPartNum.ReadOnly = True
    dgv.Columns.Add(colPartNum)
    
    ' Column: Name
    Dim colName As New DataGridViewTextBoxColumn()
    colName.Name = "colName"
    colName.HeaderText = "Nimi"
    colName.Width = 300
    colName.ReadOnly = True
    dgv.Columns.Add(colName)
    
    ' Column: Status
    Dim colStatus As New DataGridViewTextBoxColumn()
    colStatus.Name = "colStatus"
    colStatus.HeaderText = "Olek"
    colStatus.Width = 220
    colStatus.ReadOnly = True
    dgv.Columns.Add(colStatus)
    
    ' Populate rows
    For i As Integer = 0 To partDocs.Count - 1
        Dim rowIndex As Integer = dgv.Rows.Add()
        dgv.Rows(rowIndex).Tag = i
        dgv.Rows(rowIndex).Cells("colSelected").Value = selectedFlags(i) AndAlso Not hasDrawings(i)
        dgv.Rows(rowIndex).Cells("colPartNum").Value = If(String.IsNullOrEmpty(partNumbers(i)), "(puudub)", partNumbers(i))
        dgv.Rows(rowIndex).Cells("colName").Value = displayNames(i)
        
        If hasDrawings(i) Then
            dgv.Rows(rowIndex).Cells("colStatus").Value = "* 1:1 joonis olemas"
        Else
            dgv.Rows(rowIndex).Cells("colStatus").Value = "(uus)"
        End If
    Next
    
    frm.Controls.Add(dgv)
    
    currentY += 290
    
    ' Select all / none buttons
    Dim btnSelectAll As New System.Windows.Forms.Button()
    btnSelectAll.Text = "Vali kõik"
    btnSelectAll.Left = 10
    btnSelectAll.Top = currentY
    btnSelectAll.Width = 80
    btnSelectAll.Height = 25
    btnSelectAll.Anchor = AnchorStyles.Bottom Or AnchorStyles.Left
    frm.Controls.Add(btnSelectAll)
    
    AddHandler btnSelectAll.Click, Sub(s, e)
        For Each row As DataGridViewRow In dgv.Rows
            Dim idx As Integer = CInt(row.Tag)
            If Not hasDrawings(idx) Then
                row.Cells("colSelected").Value = True
            End If
        Next
    End Sub
    
    Dim btnSelectNone As New System.Windows.Forms.Button()
    btnSelectNone.Text = "Tühista"
    btnSelectNone.Left = 95
    btnSelectNone.Top = currentY
    btnSelectNone.Width = 80
    btnSelectNone.Height = 25
    btnSelectNone.Anchor = AnchorStyles.Bottom Or AnchorStyles.Left
    frm.Controls.Add(btnSelectNone)
    
    AddHandler btnSelectNone.Click, Sub(s, e)
        For Each row As DataGridViewRow In dgv.Rows
            row.Cells("colSelected").Value = False
        Next
    End Sub
    
    Dim btnSelectNew As New System.Windows.Forms.Button()
    btnSelectNew.Text = "Ainult uued"
    btnSelectNew.Left = 180
    btnSelectNew.Top = currentY
    btnSelectNew.Width = 90
    btnSelectNew.Height = 25
    btnSelectNew.Anchor = AnchorStyles.Bottom Or AnchorStyles.Left
    frm.Controls.Add(btnSelectNew)
    
    AddHandler btnSelectNew.Click, Sub(s, e)
        For Each row As DataGridViewRow In dgv.Rows
            Dim idx As Integer = CInt(row.Tag)
            row.Cells("colSelected").Value = Not hasDrawings(idx)
        Next
    End Sub
    
    currentY += 35
    
    ' Output folder
    Dim lblOutput As New System.Windows.Forms.Label()
    lblOutput.Text = "Väljundkaust:"
    lblOutput.Left = 10
    lblOutput.Top = currentY + 3
    lblOutput.Width = 80
    lblOutput.Anchor = AnchorStyles.Bottom Or AnchorStyles.Left
    frm.Controls.Add(lblOutput)
    
    Dim txtOutput As New System.Windows.Forms.TextBox()
    txtOutput.Name = "txtOutput"
    txtOutput.Text = outputFolder
    txtOutput.Left = 95
    txtOutput.Top = currentY
    txtOutput.Width = 570
    txtOutput.ReadOnly = True
    txtOutput.Anchor = AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right
    frm.Controls.Add(txtOutput)
    
    Dim btnBrowseOutput As New System.Windows.Forms.Button()
    btnBrowseOutput.Text = "..."
    btnBrowseOutput.Left = 670
    btnBrowseOutput.Top = currentY
    btnBrowseOutput.Width = 40
    btnBrowseOutput.Height = 23
    btnBrowseOutput.Anchor = AnchorStyles.Bottom Or AnchorStyles.Right
    frm.Controls.Add(btnBrowseOutput)
    
    AddHandler btnBrowseOutput.Click, Sub(s, e)
        Dim fbd As New FolderBrowserDialog()
        fbd.Description = "Vali väljundkaust joonistele"
        fbd.ShowNewFolderButton = True
        fbd.SelectedPath = txtOutput.Text
        If fbd.ShowDialog() = DialogResult.OK Then
            txtOutput.Text = fbd.SelectedPath
        End If
    End Sub
    
    currentY += 40
    
    ' OK/Cancel buttons
    Dim btnOK As New System.Windows.Forms.Button()
    btnOK.Text = "Loo joonised"
    btnOK.Left = 540
    btnOK.Top = currentY
    btnOK.Width = 100
    btnOK.Height = 28
    btnOK.Anchor = AnchorStyles.Bottom Or AnchorStyles.Right
    btnOK.DialogResult = DialogResult.OK
    frm.Controls.Add(btnOK)
    frm.AcceptButton = btnOK
    
    Dim btnCancel As New System.Windows.Forms.Button()
    btnCancel.Text = "Tühista"
    btnCancel.Left = 645
    btnCancel.Top = currentY
    btnCancel.Width = 70
    btnCancel.Height = 28
    btnCancel.Anchor = AnchorStyles.Bottom Or AnchorStyles.Right
    btnCancel.DialogResult = DialogResult.Cancel
    frm.Controls.Add(btnCancel)
    frm.CancelButton = btnCancel
    
    ' Show dialog
    Dim dlgResult As DialogResult = frm.ShowDialog()
    
    ' Extract values before disposing
    outputFolder = txtOutput.Text
    
    ' Update selection flags from grid
    For Each row As DataGridViewRow In dgv.Rows
        Dim idx As Integer = CInt(row.Tag)
        selectedFlags(idx) = CBool(row.Cells("colSelected").Value)
    Next
    
    frm.Dispose()
    
    If dlgResult <> DialogResult.OK Then
        Logger.Info("Loo 1:1 joonised: Cancelled by user")
        Exit Sub
    End If
    
    ' Get selected parts
    Dim selectedIndices As New List(Of Integer)
    For i As Integer = 0 To partDocs.Count - 1
        If selectedFlags(i) Then selectedIndices.Add(i)
    Next
    
    If selectedIndices.Count = 0 Then
        Logger.Warn("Loo 1:1 joonised: No parts selected")
        Exit Sub
    End If
    
    Logger.Info("Loo 1:1 joonised: Processing " & selectedIndices.Count & " part(s)")
    
    ' Ensure output folder exists
    If Not System.IO.Directory.Exists(outputFolder) Then
        Logger.Error("Loo 1:1 joonised: Output folder does not exist: " & outputFolder)
        MessageBox.Show("Väljundkausta ei leitud: " & vbCrLf & outputFolder, "Loo 1:1 joonised")
        Exit Sub
    End If
    
    ' Ensure folder exists in Vault
    If vaultConnected Then
        VaultNumberingLib.EnsureFolderInVault(outputFolder, vaultConn, workspaceRoot, logs)
        For Each log As String In logs : Logger.Info(log) : Next
        logs.Clear()
    End If
    
    ' Find drawing template
    Dim templatePath As String = CAMDrawingLib.FindDrawingTemplate(app, "Drawing.1.1.idw", logs)
    For Each log As String In logs : Logger.Info(log) : Next
    logs.Clear()
    
    If String.IsNullOrEmpty(templatePath) Then
        Logger.Error("Loo 1:1 joonised: Drawing template not found")
        MessageBox.Show("Joonise šablooni 'Drawing.1.1.idw' ei leitud.", "Loo 1:1 joonised")
        Exit Sub
    End If
    
    ' Process each part
    Dim createdDrawings As New List(Of String)
    Dim skippedCount As Integer = 0
    
    For Each idx As Integer In selectedIndices
        Dim partDoc As PartDocument = partDocs(idx)
        Logger.Info("Loo 1:1 joonised: Processing " & partDoc.DisplayName)
        
        ' Skip if drawing already exists and user selected it anyway
        If hasDrawings(idx) Then
            Logger.Info("Loo 1:1 joonised: Skipping - drawing already exists")
            skippedCount += 1
            Continue For
        End If
        
        ' Create drawing from template
        Dim drawDoc As DrawingDocument = CAMDrawingLib.CreateDrawingFromTemplate(app, templatePath, logs)
        For Each log As String In logs : Logger.Info(log) : Next
        logs.Clear()
        
        If drawDoc Is Nothing Then
            Logger.Error("Loo 1:1 joonised: Failed to create drawing for " & partDoc.DisplayName)
            Continue For
        End If
        
        Dim sheet As Sheet = drawDoc.ActiveSheet
        
        ' Set drawing association (copies properties + sets BB_SourcePartNumber)
        CAMDrawingLib.SetDrawingAssociation(drawDoc, partDoc, logs)
        For Each log As String In logs : Logger.Info(log) : Next
        logs.Clear()
        
        ' Add all 6 views at 1:1 scale
        Dim views As List(Of DrawingView) = CAMDrawingLib.AddAllViews(sheet, partDoc, app, logs)
        For Each log As String In logs : Logger.Info(log) : Next
        logs.Clear()
        
        ' Tag views as auto-generated (for future smart updates)
        For Each view As DrawingView In views
            CAMDrawingLib.TagAutoGeneratedView(view)
        Next
        
        ' Add extent dimensions to all views (dimensions are auto-tagged)
        CAMDrawingLib.AddExtentDimensionsToViews(sheet, views, app, logs)
        For Each log As String In logs : Logger.Info(log) : Next
        logs.Clear()
        
        ' Fit sheet to content with 50% padding
        CAMDrawingLib.FitSheetToContent(sheet, app, logs, 0.5)
        For Each log As String In logs : Logger.Info(log) : Next
        logs.Clear()
        
        ' Generate filename based on part name (Vault will assign number on save)
        Dim partName As String = System.IO.Path.GetFileNameWithoutExtension(partDoc.FullDocumentName)
        Dim drawingFileName As String = partName & ".idw"
        Dim drawingPath As String = System.IO.Path.Combine(outputFolder, drawingFileName)
        
        ' Save drawing (Vault dialog will appear for numbering)
        Try
            drawDoc.SaveAs(drawingPath, False)
            
            ' Get actual path after save (Vault may have renamed)
            Dim actualPath As String = drawDoc.FullDocumentName
            Logger.Info("Loo 1:1 joonised: Saved " & actualPath)
            createdDrawings.Add(actualPath)
            
        Catch ex As Exception
            Logger.Error("Loo 1:1 joonised: Failed to save: " & ex.Message)
        End Try
    Next
    
    ' Summary
    Logger.Info("Loo 1:1 joonised: ========================================")
    Logger.Info("Loo 1:1 joonised: SUMMARY")
    Logger.Info("Loo 1:1 joonised: ========================================")
    Logger.Info("Loo 1:1 joonised: Parts selected: " & selectedIndices.Count)
    Logger.Info("Loo 1:1 joonised: Drawings created: " & createdDrawings.Count)
    Logger.Info("Loo 1:1 joonised: Skipped (existing): " & skippedCount)
    Logger.Info("Loo 1:1 joonised: ========================================")
End Sub
