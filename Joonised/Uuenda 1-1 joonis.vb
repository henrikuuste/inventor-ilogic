' Copyright (c) 2026 Henri Kuuste
' ============================================================================
' Uuenda 1:1 joonis - Update existing 1:1 scale drawing
' 
' Features:
' - Works from drawing, part, or assembly context
' - Syncs Description and Project properties from part
' - Refreshes extent dimensions (removes old, adds new)
' - Fits sheet to content with configurable padding
'
' Usage: 
' - From drawing: Updates active drawing
' - From part: Finds and updates associated drawing (by Part Number)
'              Searches both open documents and disk
' - From assembly: Shows list of parts with associated drawings
'
' Note: This script does NOT add or remove views.
'       Use "Lisa vaated" to add views or "Loo 1:1 joonised" for new drawings.
' ============================================================================

' References for Vault integration (workspace root detection)
AddReference "Autodesk.Connectivity.WebServices"
AddReference "Autodesk.DataManagement.Client.Framework.Vault"
AddReference "Autodesk.DataManagement.Client.Framework.Vault.Forms"
AddReference "Connectivity.InventorAddin.EdmAddin"

' Libraries (UtilsLib before VaultNumberingLib for Vault logging)
AddVbFile "Lib/RuntimeLib.vb"
AddVbFile "Lib/UtilsLib.vb"
AddVbFile "Lib/VaultNumberingLib.vb"
AddVbFile "Lib/FileSearchLib.vb"
AddVbFile "Lib/CAMDrawingLib.vb"

Imports System.Collections.Generic
Imports System.Windows.Forms
Imports Inventor

Sub Main()
    If Not AppRuntime.Initialize(ThisApplication) Then Return
    
    Dim app As Inventor.Application = ThisApplication
    UtilsLib.SetLogger(Logger)
    UtilsLib.LogInfo("Uuenda 1:1 joonis: Starting...")
    
    ' Determine context and get the drawing to update
    Dim drawDoc As DrawingDocument = Nothing
    Dim partDoc As PartDocument = Nothing
    
    If app.ActiveDocument Is Nothing Then
        UtilsLib.LogError("Uuenda 1:1 joonis: No active document")
        MessageBox.Show("Aktiivne dokument puudub.", "Uuenda 1:1 joonis")
        Exit Sub
    End If
    
    Dim activeDoc As Document = app.ActiveDocument
    
    Select Case activeDoc.DocumentType
        Case DocumentTypeEnum.kDrawingDocumentObject
            ' Already in a drawing - use it directly
            drawDoc = CType(activeDoc, DrawingDocument)
            UtilsLib.LogInfo("Uuenda 1:1 joonis: Using active drawing: " & drawDoc.DisplayName)
            
        Case DocumentTypeEnum.kPartDocumentObject
            ' In a part - find associated drawing (open docs + disk)
            partDoc = CType(activeDoc, PartDocument)
            Dim partNumber As String = CAMDrawingLib.GetPartNumber(partDoc)
            
            If String.IsNullOrEmpty(partNumber) Then
                UtilsLib.LogWarn("Uuenda 1:1 joonis: Part has no Part Number - cannot find drawing")
                MessageBox.Show("Detailil puudub artikli number. Joonist ei saa tuvastada.", "Uuenda 1:1 joonis")
                Exit Sub
            End If
            
            UtilsLib.LogInfo("Uuenda 1:1 joonis: Part: " & partDoc.DisplayName & " (" & partNumber & ")")
            
            ' Get workspace root for disk search (depth-first search boundary)
            Dim partFolder As String = System.IO.Path.GetDirectoryName(partDoc.FullDocumentName)
            Dim vaultRoot As String = partFolder
            Dim vaultConn As Object = VaultNumberingLib.GetVaultConnection()
            If vaultConn IsNot Nothing Then
                Dim workspaceRoot As String = VaultNumberingLib.DetectWorkspaceRoot(vaultConn, partFolder)
                If Not String.IsNullOrEmpty(workspaceRoot) Then
                    vaultRoot = workspaceRoot
                End If
            End If
            
            UtilsLib.LogInfo("Uuenda 1:1 joonis: Searching for drawing - start: " & partFolder & ", limit: " & vaultRoot)
            
            ' Search in open documents AND on disk for 1:1 drawing (depth-first from part folder)
            Dim drawingPath As String = CAMDrawingLib.FindDrawingForPart( _
                partNumber, vaultRoot, app, CAMDrawingLib.DRAWING_TYPE_1TO1, True, partFolder)
            If Not String.IsNullOrEmpty(drawingPath) Then
                ' Found - check if already open
                drawDoc = CAMDrawingLib.FindDrawingForPartInOpenDocs(partNumber, app, CAMDrawingLib.DRAWING_TYPE_1TO1)
                If drawDoc Is Nothing Then
                    ' Open from disk
                    drawDoc = CAMDrawingLib.OpenExistingDrawing(app, drawingPath)
                End If
            End If
            
            If drawDoc Is Nothing Then
                ' Ask user to select a drawing file
                Dim ofd As New OpenFileDialog()
                ofd.Title = "Vali detaili '" & partDoc.DisplayName & "' joonis"
                ofd.Filter = "Inventor Drawing|*.idw;*.dwg"
                ofd.InitialDirectory = System.IO.Path.GetDirectoryName(partDoc.FullDocumentName)
                
                If ofd.ShowDialog() <> DialogResult.OK Then
                    UtilsLib.LogInfo("Uuenda 1:1 joonis: Cancelled by user")
                    Exit Sub
                End If
                
                drawDoc = CAMDrawingLib.OpenExistingDrawing(app, ofd.FileName)
                If drawDoc Is Nothing Then
                    UtilsLib.LogError("Uuenda 1:1 joonis: Failed to open drawing")
                    MessageBox.Show("Joonise avamine ebaõnnestus.", "Uuenda 1:1 joonis")
                    Exit Sub
                End If
            End If
            
        Case DocumentTypeEnum.kAssemblyDocumentObject
            ' In an assembly - show list of parts with drawings (search open docs + disk)
            UtilsLib.LogInfo("Uuenda 1:1 joonis: Assembly context - searching for parts with drawings")
            
            Dim asmDoc As AssemblyDocument = CType(activeDoc, AssemblyDocument)
            Dim partsWithDrawings As New List(Of Tuple(Of PartDocument, DrawingDocument))
            Dim partsWithDrawingPaths As New List(Of Tuple(Of PartDocument, String)) ' Parts with drawings on disk (not open)
            Dim partPaths As New HashSet(Of String)
            
            ' Get workspace root for disk search (depth-first search boundary)
            Dim asmFolder As String = System.IO.Path.GetDirectoryName(asmDoc.FullDocumentName)
            Dim vaultRoot As String = asmFolder
            Dim vaultConn As Object = VaultNumberingLib.GetVaultConnection()
            If vaultConn IsNot Nothing Then
                Dim workspaceRoot As String = VaultNumberingLib.DetectWorkspaceRoot(vaultConn, asmFolder)
                If Not String.IsNullOrEmpty(workspaceRoot) Then
                    vaultRoot = workspaceRoot
                End If
            End If
            
            UtilsLib.LogInfo("Uuenda 1:1 joonis: Searching for drawings - limit: " & vaultRoot)
            
            ' Find all unique parts and their drawings
            For Each occ As ComponentOccurrence In asmDoc.ComponentDefinition.Occurrences
                Try
                    If occ.DefinitionDocumentType = DocumentTypeEnum.kPartDocumentObject Then
                        Dim partPath As String = occ.ReferencedFileDescriptor.FullFileName
                        If Not partPaths.Contains(partPath) Then
                            partPaths.Add(partPath)
                            Dim pd As PartDocument = CType(occ.Definition.Document, PartDocument)
                            Dim pn As String = CAMDrawingLib.GetPartNumber(pd)
                            
                            If Not String.IsNullOrEmpty(pn) Then
                                ' First check open documents
                                Dim dd As DrawingDocument = CAMDrawingLib.FindDrawingForPartInOpenDocs( _
                                    pn, app, CAMDrawingLib.DRAWING_TYPE_1TO1)
                                If dd IsNot Nothing Then
                                    partsWithDrawings.Add(New Tuple(Of PartDocument, DrawingDocument)(pd, dd))
                                Else
                                    ' Search on disk (depth-first from part's folder)
                                    Dim partFolder As String = System.IO.Path.GetDirectoryName(pd.FullDocumentName)
                                    Dim drawingPath As String = CAMDrawingLib.FindDrawingForPart( _
                                        pn, vaultRoot, app, CAMDrawingLib.DRAWING_TYPE_1TO1, True, partFolder)
                                    If Not String.IsNullOrEmpty(drawingPath) Then
                                        partsWithDrawingPaths.Add(New Tuple(Of PartDocument, String)(pd, drawingPath))
                                    End If
                                End If
                            End If
                        End If
                    End If
                Catch
                End Try
            Next
            ' Open drawings found on disk
            For Each pdPair As Tuple(Of PartDocument, String) In partsWithDrawingPaths
                Dim dd As DrawingDocument = CAMDrawingLib.OpenExistingDrawing(app, pdPair.Item2)
                If dd IsNot Nothing Then
                    partsWithDrawings.Add(New Tuple(Of PartDocument, DrawingDocument)(pdPair.Item1, dd))
                End If
            Next
            If partsWithDrawings.Count = 0 Then
                MessageBox.Show("1:1 jooniseid ei leitud (avatud ega kettal)." & vbCrLf & _
                               "Käivita 'Loo 1:1 joonised' uute jooniste loomiseks.", "Uuenda 1:1 joonis")
                Exit Sub
            End If
            
            ' Show selection dialog if multiple
            If partsWithDrawings.Count = 1 Then
                partDoc = partsWithDrawings(0).Item1
                drawDoc = partsWithDrawings(0).Item2
            Else
                Dim result As Tuple(Of PartDocument, DrawingDocument) = ShowPartDrawingSelectionDialog(partsWithDrawings)
                If result Is Nothing Then
                    UtilsLib.LogInfo("Uuenda 1:1 joonis: Cancelled by user")
                    Exit Sub
                End If
                partDoc = result.Item1
                drawDoc = result.Item2
            End If
            
        Case Else
            UtilsLib.LogError("Uuenda 1:1 joonis: Invalid document type")
            MessageBox.Show("See reegel töötab ainult joonise, detaili või koostuga.", "Uuenda 1:1 joonis")
            Exit Sub
    End Select
    
    If drawDoc Is Nothing Then
        UtilsLib.LogError("Uuenda 1:1 joonis: No drawing to update")
        MessageBox.Show("Joonist ei leitud.", "Uuenda 1:1 joonis")
        Exit Sub
    End If
    
    ' Get the referenced part document from drawing if not already known
    If partDoc Is Nothing Then
        partDoc = CAMDrawingLib.GetReferencedPartDocument(drawDoc)
    End If
    
    If partDoc Is Nothing Then
        UtilsLib.LogWarn("Uuenda 1:1 joonis: No referenced part found in drawing")
    Else
        UtilsLib.LogInfo("Uuenda 1:1 joonis: Referenced part: " & partDoc.DisplayName)
    End If
    
    ' Get the active sheet
    Dim sheet As Sheet = drawDoc.ActiveSheet
    
    ' Show info about current state
    Dim viewCount As Integer = sheet.DrawingViews.Count
    Dim currentWidth As Double = sheet.Width * 10
    Dim currentHeight As Double = sheet.Height * 10
    Dim dimCount As Integer = sheet.DrawingDimensions.Count
    
    UtilsLib.LogInfo("Uuenda 1:1 joonis: Current sheet: " & FormatNumber(currentWidth, 1) & " x " & _
                FormatNumber(currentHeight, 1) & " mm")
    UtilsLib.LogInfo("Uuenda 1:1 joonis: Views: " & viewCount & ", Dimensions: " & dimCount)
    
    ' Show confirmation dialog
    Dim dlgResult As DialogResult = ShowUpdateDialog(drawDoc, partDoc, sheet, viewCount, dimCount)
    
    If dlgResult <> DialogResult.OK Then
        UtilsLib.LogInfo("Uuenda 1:1 joonis: Cancelled by user")
        Exit Sub
    End If
    
    ' ========================================================================
    ' Update Process
    ' ========================================================================
    
    UtilsLib.LogInfo("Uuenda 1:1 joonis: Starting update process...")
    
    ' Step 1: Sync properties from part to drawing (if part is available)
    If partDoc IsNot Nothing Then
        UtilsLib.LogInfo("Uuenda 1:1 joonis: Syncing properties from part...")
        CAMDrawingLib.CopyPropertiesToDrawing(partDoc, drawDoc)
    End If
    
    ' Step 2: Reposition views (geometry may have changed)
    UtilsLib.LogInfo("Uuenda 1:1 joonis: Repositioning views...")
    CAMDrawingLib.RepositionViews(sheet, app)
    ' Step 3: Update tagged extent dimensions (smart update)
    ' - Only recreates dimensions that were auto-generated and still exist
    ' - Preserves user-added dimensions (no tag)
    ' - Doesn't recreate dimensions user deleted (not in tagged list)
    UtilsLib.LogInfo("Uuenda 1:1 joonis: Updating tagged extent dimensions...")
    CAMDrawingLib.UpdateTaggedExtentDimensions(drawDoc, sheet, app)
    ' Step 4: Fit sheet to content
    UtilsLib.LogInfo("Uuenda 1:1 joonis: Fitting sheet to content...")
    CAMDrawingLib.FitSheetToContent(sheet, app)
    ' Step 5: Save the drawing
    Try
        drawDoc.Save()
        UtilsLib.LogInfo("Uuenda 1:1 joonis: Drawing saved")
    Catch ex As Exception
        UtilsLib.LogError("Uuenda 1:1 joonis: Failed to save: " & ex.Message)
    End Try
    
    ' Final sizes
    Dim newWidth As Double = sheet.Width * 10
    Dim newHeight As Double = sheet.Height * 10
    Dim newDimCount As Integer = sheet.DrawingDimensions.Count
    
    ' Summary
    UtilsLib.LogInfo("Uuenda 1:1 joonis: ========================================")
    UtilsLib.LogInfo("Uuenda 1:1 joonis: UPDATE COMPLETE")
    UtilsLib.LogInfo("Uuenda 1:1 joonis: ========================================")
    UtilsLib.LogInfo("Uuenda 1:1 joonis: Sheet size: " & FormatNumber(currentWidth, 0) & "x" & FormatNumber(currentHeight, 0) & _
                " -> " & FormatNumber(newWidth, 0) & "x" & FormatNumber(newHeight, 0) & " mm")
    UtilsLib.LogInfo("Uuenda 1:1 joonis: Dimensions: " & dimCount & " -> " & newDimCount)
    UtilsLib.LogInfo("Uuenda 1:1 joonis: Properties synced: " & If(partDoc IsNot Nothing, "Yes", "No"))
    UtilsLib.LogInfo("Uuenda 1:1 joonis: ========================================")
End Sub

' ============================================================================
' Update Confirmation Dialog
' ============================================================================

Function ShowUpdateDialog(drawDoc As DrawingDocument, _
                          partDoc As PartDocument, _
                          sheet As Sheet, _
                          viewCount As Integer, _
                          dimCount As Integer) As DialogResult
    
    Dim frm As New System.Windows.Forms.Form()
    frm.Text = "Uuenda 1:1 joonis"
    frm.Width = 500
    frm.Height = 360
    frm.StartPosition = FormStartPosition.CenterScreen
    frm.FormBorderStyle = FormBorderStyle.FixedDialog
    frm.MinimizeBox = False
    frm.MaximizeBox = False
    
    Dim currentY As Integer = 15
    
    ' Drawing info
    Dim lblDrawing As New System.Windows.Forms.Label()
    lblDrawing.Text = "Joonis:"
    lblDrawing.Left = 15
    lblDrawing.Top = currentY
    lblDrawing.Width = 80
    frm.Controls.Add(lblDrawing)
    
    Dim lblDrawingValue As New System.Windows.Forms.Label()
    lblDrawingValue.Text = drawDoc.DisplayName
    lblDrawingValue.Left = 100
    lblDrawingValue.Top = currentY
    lblDrawingValue.Width = 370
    frm.Controls.Add(lblDrawingValue)
    
    currentY += 25
    
    ' Part info
    Dim lblPart As New System.Windows.Forms.Label()
    lblPart.Text = "Detail:"
    lblPart.Left = 15
    lblPart.Top = currentY
    lblPart.Width = 80
    frm.Controls.Add(lblPart)
    
    Dim lblPartValue As New System.Windows.Forms.Label()
    lblPartValue.Text = If(partDoc IsNot Nothing, partDoc.DisplayName, "(tundmatu)")
    lblPartValue.Left = 100
    lblPartValue.Top = currentY
    lblPartValue.Width = 370
    frm.Controls.Add(lblPartValue)
    
    currentY += 25
    
    ' Current sheet size
    Dim lblSheet As New System.Windows.Forms.Label()
    lblSheet.Text = "Lehe suurus:"
    lblSheet.Left = 15
    lblSheet.Top = currentY
    lblSheet.Width = 80
    frm.Controls.Add(lblSheet)
    
    Dim lblSheetValue As New System.Windows.Forms.Label()
    lblSheetValue.Text = FormatNumber(sheet.Width * 10, 1) & " x " & FormatNumber(sheet.Height * 10, 1) & " mm"
    lblSheetValue.Left = 100
    lblSheetValue.Top = currentY
    lblSheetValue.Width = 200
    frm.Controls.Add(lblSheetValue)
    
    currentY += 25
    
    ' View count
    Dim lblViews As New System.Windows.Forms.Label()
    lblViews.Text = "Vaated:"
    lblViews.Left = 15
    lblViews.Top = currentY
    lblViews.Width = 80
    frm.Controls.Add(lblViews)
    
    Dim lblViewsValue As New System.Windows.Forms.Label()
    lblViewsValue.Text = viewCount.ToString()
    lblViewsValue.Left = 100
    lblViewsValue.Top = currentY
    lblViewsValue.Width = 50
    frm.Controls.Add(lblViewsValue)
    
    ' Dimension count
    Dim lblDims As New System.Windows.Forms.Label()
    lblDims.Text = "Mõõtmed:"
    lblDims.Left = 160
    lblDims.Top = currentY
    lblDims.Width = 60
    frm.Controls.Add(lblDims)
    
    Dim lblDimsValue As New System.Windows.Forms.Label()
    lblDimsValue.Text = dimCount.ToString()
    lblDimsValue.Left = 225
    lblDimsValue.Top = currentY
    lblDimsValue.Width = 50
    frm.Controls.Add(lblDimsValue)
    
    currentY += 35
    
    ' Separator
    Dim separator As New System.Windows.Forms.Label()
    separator.Text = "─────────────────────────────────────────────────"
    separator.Left = 15
    separator.Top = currentY
    separator.Width = 460
    frm.Controls.Add(separator)
    
    currentY += 25
    
    ' What will be updated
    Dim lblInfo As New System.Windows.Forms.Label()
    lblInfo.Text = "Uuendamisel tehakse:" & vbCrLf & vbCrLf & _
                   If(partDoc IsNot Nothing, "✓ Atribuudid (Project, Description) sünkroniseeritakse" & vbCrLf, "") & _
                   "✓ Automaatsed gabariidimõõtmed uuendatakse" & vbCrLf & _
                   "✓ Käsitsi lisatud mõõtmed säilivad" & vbCrLf & _
                   "✓ Kustutatud mõõtmeid ei taastata" & vbCrLf & _
                   "✓ Lehe suurus kohandatakse sisule" & vbCrLf & vbCrLf & _
                   "NB: Vaated jäävad muutmata."
    lblInfo.Left = 15
    lblInfo.Top = currentY
    lblInfo.Width = 460
    lblInfo.Height = 140
    frm.Controls.Add(lblInfo)
    
    currentY += 140
    
    ' Buttons
    Dim btnOK As New System.Windows.Forms.Button()
    btnOK.Text = "Uuenda"
    btnOK.Left = 300
    btnOK.Top = currentY
    btnOK.Width = 85
    btnOK.Height = 28
    btnOK.DialogResult = DialogResult.OK
    frm.Controls.Add(btnOK)
    frm.AcceptButton = btnOK
    
    Dim btnCancel As New System.Windows.Forms.Button()
    btnCancel.Text = "Tühista"
    btnCancel.Left = 390
    btnCancel.Top = currentY
    btnCancel.Width = 85
    btnCancel.Height = 28
    btnCancel.DialogResult = DialogResult.Cancel
    frm.Controls.Add(btnCancel)
    frm.CancelButton = btnCancel
    
    ' Show dialog
    Dim result As DialogResult = frm.ShowDialog()
    frm.Dispose()
    
    Return result
End Function

' ============================================================================
' Part-Drawing Selection Dialog (for assembly context)
' ============================================================================

Function ShowPartDrawingSelectionDialog(partsWithDrawings As List(Of Tuple(Of PartDocument, DrawingDocument))) As Tuple(Of PartDocument, DrawingDocument)
    
    Dim frm As New System.Windows.Forms.Form()
    frm.Text = "Vali joonis uuendamiseks"
    frm.Width = 600
    frm.Height = 400
    frm.StartPosition = FormStartPosition.CenterScreen
    frm.FormBorderStyle = FormBorderStyle.Sizable
    
    Dim currentY As Integer = 10
    
    Dim lblInfo As New System.Windows.Forms.Label()
    lblInfo.Text = "Leitud avatud joonised (" & partsWithDrawings.Count & "):"
    lblInfo.Left = 10
    lblInfo.Top = currentY
    lblInfo.Width = 400
    frm.Controls.Add(lblInfo)
    
    currentY += 25
    
    Dim lstItems As New System.Windows.Forms.ListBox()
    lstItems.Name = "lstItems"
    lstItems.Left = 10
    lstItems.Top = currentY
    lstItems.Width = 560
    lstItems.Height = 280
    lstItems.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Bottom
    
    For i As Integer = 0 To partsWithDrawings.Count - 1
        Dim pd As PartDocument = partsWithDrawings(i).Item1
        Dim dd As DrawingDocument = partsWithDrawings(i).Item2
        lstItems.Items.Add(pd.DisplayName & " -> " & dd.DisplayName)
    Next
    
    If lstItems.Items.Count > 0 Then lstItems.SelectedIndex = 0
    frm.Controls.Add(lstItems)
    
    currentY += 290
    
    ' Buttons
    Dim btnOK As New System.Windows.Forms.Button()
    btnOK.Text = "Uuenda"
    btnOK.Left = 400
    btnOK.Top = currentY
    btnOK.Width = 85
    btnOK.Height = 28
    btnOK.Anchor = AnchorStyles.Bottom Or AnchorStyles.Right
    btnOK.DialogResult = DialogResult.OK
    frm.Controls.Add(btnOK)
    frm.AcceptButton = btnOK
    
    Dim btnCancel As New System.Windows.Forms.Button()
    btnCancel.Text = "Tühista"
    btnCancel.Left = 490
    btnCancel.Top = currentY
    btnCancel.Width = 85
    btnCancel.Height = 28
    btnCancel.Anchor = AnchorStyles.Bottom Or AnchorStyles.Right
    btnCancel.DialogResult = DialogResult.Cancel
    frm.Controls.Add(btnCancel)
    frm.CancelButton = btnCancel
    
    ' Store selected index in Tag
    frm.Tag = -1
    
    Dim result As DialogResult = frm.ShowDialog()
    
    If result = DialogResult.OK AndAlso lstItems.SelectedIndex >= 0 Then
        frm.Dispose()
        Return partsWithDrawings(lstItems.SelectedIndex)
    End If
    
    frm.Dispose()
    Return Nothing
End Function
