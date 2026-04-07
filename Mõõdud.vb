' ============================================================================
' Mõõdud - Materjali gabariitmõõtude kalkulaator
' 
' Töötab nii detaili kui koostu dokumentidega:
' - Detailis: töötleb aktiivset detaili
' - Koostus: töötleb valitud detailid
'
' Shows all parts in a single DataGridView dialog where user can:
' - See T/W/L measurements for each part
' - Change thickness axis (X/Y/Z/Custom)
' - Flip width/length
' - Pick a face for custom axis orientation
'
' Loob igasse detaili lokaalse reegli "Uuenda mõõdud", mis uuendab
' iProperties väärtusi (Paksus, Laius, Pikkus) gabariitmõõtude alusel.
' ============================================================================

AddVbFile "Lib/UtilsLib.vb"
AddVbFile "Lib/BoundingBoxStockLib.vb"

Imports System.Collections.Generic
Imports System.Windows.Forms
Imports Inventor

Sub Main()
    UtilsLib.SetLogger(Logger)
    
    Dim app As Inventor.Application = ThisApplication
    Dim doc As Document = app.ActiveDocument

    If doc Is Nothing Then
        Logger.Error("Mõõdud: No active document")
        MessageBox.Show("Aktiivne dokument puudub.", "Mõõdud")
        Exit Sub
    End If

    ' Collect parts to process
    Dim partDocs As New List(Of PartDocument)
    Dim partNames As New List(Of String)
    Dim thicknessAxes As New List(Of String)
    Dim widthAxes As New List(Of String)
    Dim lengthAxes As New List(Of String)
    Dim customAxisDescs As New List(Of String)
    Dim selectedFlags As New List(Of Boolean)

    If doc.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
        ' Single part document
        Dim partDoc As PartDocument = CType(doc, PartDocument)
        
        If IsSheetMetalPart(partDoc) Then
            MessageBox.Show("See reegel ei tööta lehtmetalli detailidega.", "Mõõdud")
            Exit Sub
        End If
        
        CollectPartData(partDoc, partDocs, partNames, thicknessAxes, widthAxes, lengthAxes, customAxisDescs, selectedFlags)
        Logger.Info("Mõõdud: Processing single part - " & partDoc.DisplayName)
        
    ElseIf doc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
        ' Assembly document - process selected parts or all parts if none selected
        Dim asmDoc As AssemblyDocument = CType(doc, AssemblyDocument)
        Dim sel As SelectSet = asmDoc.SelectSet
        Dim useAllParts As Boolean = (sel Is Nothing OrElse sel.Count = 0)

        ' Collect unique part occurrences
        Dim processedDefs As New HashSet(Of Object)
        
        If useAllParts Then
            ' No selection - collect all parts from assembly
            CollectAllPartsFromAssembly(asmDoc.ComponentDefinition.Occurrences, processedDefs, _
                                        partDocs, partNames, thicknessAxes, widthAxes, lengthAxes, customAxisDescs, selectedFlags)
            Logger.Info("Mõõdud: No selection - using all " & partDocs.Count & " part(s) from assembly")
        Else
            ' Use selection
            For Each selObj As Object In sel
                If TypeOf selObj Is ComponentOccurrence Then
                    Dim occ As ComponentOccurrence = CType(selObj, ComponentOccurrence)
                    If occ.DefinitionDocumentType = DocumentTypeEnum.kPartDocumentObject Then
                        ' Avoid duplicates from same definition
                        If Not processedDefs.Contains(occ.Definition) Then
                            processedDefs.Add(occ.Definition)
                            Try
                                Dim partDoc As PartDocument = CType(occ.Definition.Document, PartDocument)
                                ' Skip sheet metal parts
                                If Not IsSheetMetalPart(partDoc) Then
                                    CollectPartData(partDoc, partDocs, partNames, thicknessAxes, widthAxes, lengthAxes, customAxisDescs, selectedFlags)
                                End If
                            Catch
                            End Try
                        End If
                    End If
                End If
            Next
            Logger.Info("Mõõdud: Processing " & partDocs.Count & " part(s) from assembly selection")
        End If

        If partDocs.Count = 0 Then
            MessageBox.Show("Koostus ei leitud sobivaid detaile." & vbCrLf & _
                            "(Lehtmetalli detailid on välja jäetud.)", "Mõõdud")
            Exit Sub
        End If
    Else
        MessageBox.Show("See reegel töötab ainult detaili (.ipt) või koostu (.iam) dokumentidega.", "Mõõdud")
        Exit Sub
    End If

    ' Dialog loop for face picking
    Dim dlgResult As DialogResult
    Dim pickRowIndex As Integer = -1

    Do
        dlgResult = ShowBatchDialog(app, partDocs, partNames, thicknessAxes, widthAxes, lengthAxes, _
                                    customAxisDescs, selectedFlags, pickRowIndex)

        If dlgResult = DialogResult.Retry AndAlso pickRowIndex >= 0 AndAlso pickRowIndex < partDocs.Count Then
            ' User clicked "Vali pind" - do face pick
            Dim partDoc As PartDocument = partDocs(pickRowIndex)
            Logger.Info("Mõõdud: Picking face for '" & partNames(pickRowIndex) & "'")

            Try
                Dim planeDesc As String = ""
                Dim pickedVector As String = BoundingBoxStockLib.PickPlaneForThickness(app, planeDesc, True)
                
                If pickedVector <> "" Then
                    thicknessAxes(pickRowIndex) = pickedVector
                    customAxisDescs(pickRowIndex) = planeDesc
                    
                    ' Compute perpendicular vectors for width/length
                    Dim tx As Double = 0, ty As Double = 0, tz As Double = 0
                    Dim wx As Double = 0, wy As Double = 0, wz As Double = 0
                    Dim lx As Double = 0, ly As Double = 0, lz As Double = 0
                    BoundingBoxStockLib.ParseVectorComponents(pickedVector, tx, ty, tz)
                    BoundingBoxStockLib.ComputePerpendicularVectors(tx, ty, tz, wx, wy, wz, lx, ly, lz)
                    
                    ' Measure extents to determine which is width vs length
                    Dim widthExtent As Double = BoundingBoxStockLib.GetOrientedExtent(partDoc, wx, wy, wz)
                    Dim lengthExtent As Double = BoundingBoxStockLib.GetOrientedExtent(partDoc, lx, ly, lz)
                    
                    If lengthExtent >= widthExtent Then
                        lengthAxes(pickRowIndex) = BoundingBoxStockLib.VectorToString(lx, ly, lz)
                        widthAxes(pickRowIndex) = BoundingBoxStockLib.VectorToString(wx, wy, wz)
                    Else
                        lengthAxes(pickRowIndex) = BoundingBoxStockLib.VectorToString(wx, wy, wz)
                        widthAxes(pickRowIndex) = BoundingBoxStockLib.VectorToString(lx, ly, lz)
                    End If
                    
                    Logger.Info("Mõõdud: Applied custom axis for '" & partNames(pickRowIndex) & "' - " & planeDesc)
                End If
            Catch
                ' User cancelled pick
            End Try

            pickRowIndex = -1
            Continue Do
        End If

        Exit Do
    Loop

    If dlgResult <> DialogResult.OK Then
        Logger.Info("Mõõdud: Cancelled by user")
        Exit Sub
    End If

    ' Apply rules to selected parts
    Dim processedCount As Integer = 0
    For i As Integer = 0 To partDocs.Count - 1
        If selectedFlags(i) Then
            Dim partDoc As PartDocument = partDocs(i)
            BoundingBoxStockLib.CreateOrUpdateRule(partDoc, thicknessAxes(i), widthAxes(i), lengthAxes(i), iLogicVb.Automation)
            processedCount += 1
            Logger.Info("Mõõdud: Updated '" & partNames(i) & "' - T:" & thicknessAxes(i) & " W:" & widthAxes(i) & " L:" & lengthAxes(i))
        End If
    Next

    Logger.Info("Mõõdud: Completed - processed " & processedCount & " part(s)")
End Sub

' ============================================================================
' Collect part data and auto-detect axes
' ============================================================================
Sub CollectPartData(ByVal partDoc As PartDocument, _
                    ByVal partDocs As List(Of PartDocument), _
                    ByVal partNames As List(Of String), _
                    ByVal thicknessAxes As List(Of String), _
                    ByVal widthAxes As List(Of String), _
                    ByVal lengthAxes As List(Of String), _
                    ByVal customAxisDescs As List(Of String), _
                    ByVal selectedFlags As List(Of Boolean))
    
    partDocs.Add(partDoc)
    
    ' Build display name: filename + description
    Dim fileName As String = System.IO.Path.GetFileNameWithoutExtension(partDoc.FullFileName)
    Dim desc As String = GetPartDescription(partDoc)
    Dim displayName As String = fileName
    If desc <> "" Then displayName &= " - " & desc
    partNames.Add(displayName)
    
    ' Try to read existing axis configuration from iProperties
    Dim thicknessAxis As String = BoundingBoxStockLib.GetCustomPropertyValue(partDoc, "BB_ThicknessAxis", "")
    Dim widthAxis As String = BoundingBoxStockLib.GetCustomPropertyValue(partDoc, "BB_WidthAxis", "")
    Dim lengthAxis As String = ""
    Dim customAxisDesc As String = ""
    
    ' Get bounding box sizes for axis-aligned detection
    Dim xSize As Double = 0, ySize As Double = 0, zSize As Double = 0
    BoundingBoxStockLib.GetBoundingBoxSizes(partDoc, xSize, ySize, zSize)
    
    If thicknessAxis = "" Then
        ' No existing config - auto-detect from geometry
        If Not BoundingBoxStockLib.AutoDetectAxesFromGeometry(partDoc, thicknessAxis, widthAxis, lengthAxis) Then
            ' Fall back to axis-aligned detection
            BoundingBoxStockLib.AutoDetectAxes(xSize, ySize, zSize, thicknessAxis, widthAxis, lengthAxis)
        End If
        
        ' Set description if auto-detected a vector
        If BoundingBoxStockLib.IsVectorFormat(thicknessAxis) Then
            Dim tx As Double = 0, ty As Double = 0, tz As Double = 0
            BoundingBoxStockLib.ParseVectorComponents(thicknessAxis, tx, ty, tz)
            customAxisDesc = "Auto (" & BoundingBoxStockLib.FormatVectorDesc(tx, ty, tz) & ")"
        End If
    ElseIf BoundingBoxStockLib.IsVectorFormat(thicknessAxis) Then
        ' Vector format stored
        Dim tx As Double = 0, ty As Double = 0, tz As Double = 0
        If BoundingBoxStockLib.ParseVectorComponents(thicknessAxis, tx, ty, tz) Then
            customAxisDesc = "Custom (" & BoundingBoxStockLib.FormatVectorDesc(tx, ty, tz) & ")"
            
            ' Compute width/length if not stored
            If widthAxis = "" OrElse Not BoundingBoxStockLib.IsVectorFormat(widthAxis) Then
                Dim wx As Double = 0, wy As Double = 0, wz As Double = 0
                Dim lx As Double = 0, ly As Double = 0, lz As Double = 0
                BoundingBoxStockLib.ComputePerpendicularVectors(tx, ty, tz, wx, wy, wz, lx, ly, lz)
                
                Dim widthExtent As Double = BoundingBoxStockLib.GetOrientedExtent(partDoc, wx, wy, wz)
                Dim lengthExtent As Double = BoundingBoxStockLib.GetOrientedExtent(partDoc, lx, ly, lz)
                
                If lengthExtent >= widthExtent Then
                    lengthAxis = BoundingBoxStockLib.VectorToString(lx, ly, lz)
                    widthAxis = BoundingBoxStockLib.VectorToString(wx, wy, wz)
                Else
                    lengthAxis = BoundingBoxStockLib.VectorToString(wx, wy, wz)
                    widthAxis = BoundingBoxStockLib.VectorToString(lx, ly, lz)
                End If
            Else
                ' Width stored - compute length as cross product
                Dim wx As Double = 0, wy As Double = 0, wz As Double = 0
                BoundingBoxStockLib.ParseVectorComponents(widthAxis, wx, wy, wz)
                Dim lx As Double = ty * wz - tz * wy
                Dim ly As Double = tz * wx - tx * wz
                Dim lz As Double = tx * wy - ty * wx
                lengthAxis = BoundingBoxStockLib.VectorToString(lx, ly, lz)
            End If
        Else
            thicknessAxis = "Z"
            BoundingBoxStockLib.AssignWidthLength(thicknessAxis, xSize, ySize, zSize, widthAxis, lengthAxis)
        End If
    Else
        ' Simple axis format (X/Y/Z)
        If widthAxis = "" Then
            BoundingBoxStockLib.AssignWidthLength(thicknessAxis, xSize, ySize, zSize, widthAxis, lengthAxis)
        Else
            lengthAxis = BoundingBoxStockLib.GetRemainingAxis(thicknessAxis, widthAxis)
        End If
    End If
    
    thicknessAxes.Add(thicknessAxis)
    widthAxes.Add(widthAxis)
    lengthAxes.Add(lengthAxis)
    customAxisDescs.Add(customAxisDesc)
    selectedFlags.Add(True)
End Sub

' ============================================================================
' Recursively collect all parts from assembly occurrences
' ============================================================================
Sub CollectAllPartsFromAssembly(ByVal occurrences As ComponentOccurrences, _
                                 ByVal processedDefs As HashSet(Of Object), _
                                 ByVal partDocs As List(Of PartDocument), _
                                 ByVal partNames As List(Of String), _
                                 ByVal thicknessAxes As List(Of String), _
                                 ByVal widthAxes As List(Of String), _
                                 ByVal lengthAxes As List(Of String), _
                                 ByVal customAxisDescs As List(Of String), _
                                 ByVal selectedFlags As List(Of Boolean))
    
    For Each occ As ComponentOccurrence In occurrences
        Try
            If occ.DefinitionDocumentType = DocumentTypeEnum.kPartDocumentObject Then
                ' Part occurrence - collect if not already processed
                If Not processedDefs.Contains(occ.Definition) Then
                    processedDefs.Add(occ.Definition)
                    Dim partDoc As PartDocument = CType(occ.Definition.Document, PartDocument)
                    ' Skip sheet metal parts
                    If Not IsSheetMetalPart(partDoc) Then
                        CollectPartData(partDoc, partDocs, partNames, thicknessAxes, widthAxes, lengthAxes, customAxisDescs, selectedFlags)
                    End If
                End If
            ElseIf occ.DefinitionDocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                ' Sub-assembly - recurse into it
                Dim subAsmDef As AssemblyComponentDefinition = CType(occ.Definition, AssemblyComponentDefinition)
                CollectAllPartsFromAssembly(subAsmDef.Occurrences, processedDefs, _
                                           partDocs, partNames, thicknessAxes, widthAxes, lengthAxes, customAxisDescs, selectedFlags)
            End If
        Catch
        End Try
    Next
End Sub

' ============================================================================
' Get part description from iProperties
' ============================================================================
Function GetPartDescription(ByVal partDoc As PartDocument) As String
    Try
        Dim designProps As PropertySet = partDoc.PropertySets.Item("Design Tracking Properties")
        Dim desc As String = CStr(designProps.Item("Description").Value)
        If desc IsNot Nothing AndAlso desc.Trim() <> "" Then
            Return desc.Trim()
        End If
    Catch
    End Try
    
    Try
        Dim summaryInfo As PropertySet = partDoc.PropertySets.Item("Inventor Summary Information")
        Dim subj As String = CStr(summaryInfo.Item("Subject").Value)
        If subj IsNot Nothing AndAlso subj.Trim() <> "" Then
            Return subj.Trim()
        End If
    Catch
    End Try
    
    Return ""
End Function

' ============================================================================
' Check if part is a sheet metal part
' ============================================================================
Function IsSheetMetalPart(ByVal partDoc As PartDocument) As Boolean
    Try
        ' Sheet metal SubType GUID
        Const SHEET_METAL_SUBTYPE As String = "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}"
        Return partDoc.SubType = SHEET_METAL_SUBTYPE
    Catch
        Return False
    End Try
End Function

' ============================================================================
' Show batch dialog with DataGridView
' ============================================================================
Function ShowBatchDialog(ByVal app As Inventor.Application, _
                         ByVal partDocs As List(Of PartDocument), _
                         ByVal partNames As List(Of String), _
                         ByVal thicknessAxes As List(Of String), _
                         ByVal widthAxes As List(Of String), _
                         ByVal lengthAxes As List(Of String), _
                         ByVal customAxisDescs As List(Of String), _
                         ByVal selectedFlags As List(Of Boolean), _
                         ByRef pickRowIndex As Integer) As DialogResult
    
    pickRowIndex = -1
    
    Dim frm As New System.Windows.Forms.Form()
    frm.Text = "Mõõdud - Gabariitmõõtude seadistamine"
    frm.Width = 900
    frm.Height = If(partDocs.Count = 1, 220, 450)
    frm.StartPosition = FormStartPosition.CenterScreen
    frm.FormBorderStyle = FormBorderStyle.Sizable
    frm.MinimizeBox = True
    frm.MaximizeBox = True
    
    Dim currentY As Integer = 10
    
    ' Header label
    Dim lblHeader As New System.Windows.Forms.Label()
    lblHeader.Text = "Detailid (" & partDocs.Count & "):"
    lblHeader.Left = 10
    lblHeader.Top = currentY
    lblHeader.Width = 200
    frm.Controls.Add(lblHeader)
    
    currentY += 20
    
    ' DataGridView
    Dim dgv As New System.Windows.Forms.DataGridView()
    dgv.Name = "dgvParts"
    dgv.Left = 10
    dgv.Top = currentY
    dgv.Width = 860
    dgv.Height = If(partDocs.Count = 1, 60, 280)
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
    
    ' Column: Part Name
    Dim colName As New DataGridViewTextBoxColumn()
    colName.Name = "colName"
    colName.HeaderText = "Detail"
    colName.Width = 250
    colName.ReadOnly = True
    dgv.Columns.Add(colName)
    
    ' Column: Thickness
    Dim colT As New DataGridViewTextBoxColumn()
    colT.Name = "colT"
    colT.HeaderText = "T (mm)"
    colT.Width = 70
    colT.ReadOnly = True
    dgv.Columns.Add(colT)
    
    ' Column: Width
    Dim colW As New DataGridViewTextBoxColumn()
    colW.Name = "colW"
    colW.HeaderText = "W (mm)"
    colW.Width = 70
    colW.ReadOnly = True
    dgv.Columns.Add(colW)
    
    ' Column: Length
    Dim colL As New DataGridViewTextBoxColumn()
    colL.Name = "colL"
    colL.HeaderText = "L (mm)"
    colL.Width = 70
    colL.ReadOnly = True
    dgv.Columns.Add(colL)
    
    ' Column: Axis (ComboBox)
    Dim colAxis As New DataGridViewComboBoxColumn()
    colAxis.Name = "colAxis"
    colAxis.HeaderText = "Telg"
    colAxis.Width = 100
    colAxis.FlatStyle = FlatStyle.Flat
    colAxis.Items.Add("X")
    colAxis.Items.Add("Y")
    colAxis.Items.Add("Z")
    colAxis.Items.Add("Kohandatud")
    dgv.Columns.Add(colAxis)
    
    ' Column: Flip button
    Dim colFlip As New DataGridViewButtonColumn()
    colFlip.Name = "colFlip"
    colFlip.HeaderText = "W/L"
    colFlip.Text = "Vaheta"
    colFlip.UseColumnTextForButtonValue = True
    colFlip.Width = 70
    dgv.Columns.Add(colFlip)
    
    ' Column: Pick face button
    Dim colPick As New DataGridViewButtonColumn()
    colPick.Name = "colPick"
    colPick.HeaderText = "Pind"
    colPick.Text = "Vali pind"
    colPick.UseColumnTextForButtonValue = True
    colPick.Width = 80
    dgv.Columns.Add(colPick)
    
    ' Populate rows
    For i As Integer = 0 To partDocs.Count - 1
        Dim rowIndex As Integer = dgv.Rows.Add()
        dgv.Rows(rowIndex).Tag = i
        
        dgv.Rows(rowIndex).Cells("colSelected").Value = selectedFlags(i)
        dgv.Rows(rowIndex).Cells("colName").Value = partNames(i)
        
        ' Calculate display values
        UpdateRowDisplayValues(dgv.Rows(rowIndex), partDocs(i), thicknessAxes(i), widthAxes(i), lengthAxes(i))
        
        ' Set axis combo value
        If BoundingBoxStockLib.IsVectorFormat(thicknessAxes(i)) Then
            dgv.Rows(rowIndex).Cells("colAxis").Value = "Kohandatud"
        Else
            dgv.Rows(rowIndex).Cells("colAxis").Value = thicknessAxes(i)
        End If
    Next
    
    ' Store form tag as -1 (no pick requested yet)
    frm.Tag = -1
    
    ' Handle button clicks
    AddHandler dgv.CellContentClick, Sub(sender As Object, e As DataGridViewCellEventArgs)
        If e.RowIndex < 0 Then Exit Sub
        
        Dim idx As Integer = CInt(dgv.Rows(e.RowIndex).Tag)
        
        If e.ColumnIndex = dgv.Columns("colFlip").Index Then
            ' Flip width/length
            Dim tempAxis As String = widthAxes(idx)
            widthAxes(idx) = lengthAxes(idx)
            lengthAxes(idx) = tempAxis
            
            ' Update display
            UpdateRowDisplayValues(dgv.Rows(e.RowIndex), partDocs(idx), thicknessAxes(idx), widthAxes(idx), lengthAxes(idx))
            
        ElseIf e.ColumnIndex = dgv.Columns("colPick").Index Then
            ' Pick face - sync state and close form
            SyncGridToLists(dgv, selectedFlags)
            frm.Tag = idx
            frm.DialogResult = DialogResult.Retry
            frm.Close()
        End If
    End Sub
    
    ' Handle axis combo change
    AddHandler dgv.CellValueChanged, Sub(sender As Object, e As DataGridViewCellEventArgs)
        If e.RowIndex < 0 Then Exit Sub
        If e.ColumnIndex <> dgv.Columns("colAxis").Index Then Exit Sub
        
        Dim idx As Integer = CInt(dgv.Rows(e.RowIndex).Tag)
        Dim newAxisValue As Object = dgv.Rows(e.RowIndex).Cells("colAxis").Value
        If newAxisValue Is Nothing Then Exit Sub
        
        Dim newAxis As String = newAxisValue.ToString()
        
        If newAxis = "Kohandatud" Then
            ' Trigger face pick
            SyncGridToLists(dgv, selectedFlags)
            frm.Tag = idx
            frm.DialogResult = DialogResult.Retry
            frm.Close()
        ElseIf newAxis = "X" OrElse newAxis = "Y" OrElse newAxis = "Z" Then
            ' Recalculate axes for standard axis
            thicknessAxes(idx) = newAxis
            customAxisDescs(idx) = ""
            
            Dim xSize As Double = 0, ySize As Double = 0, zSize As Double = 0
            BoundingBoxStockLib.GetBoundingBoxSizes(partDocs(idx), xSize, ySize, zSize)
            
            Dim newWidth As String = ""
            Dim newLength As String = ""
            BoundingBoxStockLib.AssignWidthLength(newAxis, xSize, ySize, zSize, newWidth, newLength)
            widthAxes(idx) = newWidth
            lengthAxes(idx) = newLength
            
            ' Update display
            UpdateRowDisplayValues(dgv.Rows(e.RowIndex), partDocs(idx), thicknessAxes(idx), widthAxes(idx), lengthAxes(idx))
        End If
    End Sub
    
    ' Commit edit when cell loses focus (needed for combo box)
    AddHandler dgv.CurrentCellDirtyStateChanged, Sub(sender As Object, e As EventArgs)
        If dgv.IsCurrentCellDirty Then
            dgv.CommitEdit(DataGridViewDataErrorContexts.Commit)
        End If
    End Sub
    
    frm.Controls.Add(dgv)
    
    currentY += dgv.Height + 10
    
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
            row.Cells("colSelected").Value = True
        Next
    End Sub
    
    Dim btnSelectNone As New System.Windows.Forms.Button()
    btnSelectNone.Text = "Tühjenda"
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
    
    ' OK/Cancel buttons
    Dim btnOK As New System.Windows.Forms.Button()
    btnOK.Text = "Käivita"
    btnOK.Left = 700
    btnOK.Top = currentY
    btnOK.Width = 90
    btnOK.Height = 28
    btnOK.Anchor = AnchorStyles.Bottom Or AnchorStyles.Right
    btnOK.DialogResult = DialogResult.OK
    frm.Controls.Add(btnOK)
    frm.AcceptButton = btnOK
    
    Dim btnCancel As New System.Windows.Forms.Button()
    btnCancel.Text = "Tühista"
    btnCancel.Left = 795
    btnCancel.Top = currentY
    btnCancel.Width = 75
    btnCancel.Height = 28
    btnCancel.Anchor = AnchorStyles.Bottom Or AnchorStyles.Right
    btnCancel.DialogResult = DialogResult.Cancel
    frm.Controls.Add(btnCancel)
    frm.CancelButton = btnCancel
    
    ' Show dialog
    Dim result As DialogResult = frm.ShowDialog()
    
    ' Sync final state back to lists
    If result = DialogResult.OK Then
        SyncGridToLists(dgv, selectedFlags)
    End If
    
    ' Read pick index from form.Tag
    pickRowIndex = CInt(frm.Tag)
    
    frm.Dispose()
    Return result
End Function

' ============================================================================
' Update row display values (T/W/L)
' ============================================================================
Sub UpdateRowDisplayValues(ByVal row As DataGridViewRow, ByVal partDoc As PartDocument, _
                           ByVal thicknessAxis As String, ByVal widthAxis As String, ByVal lengthAxis As String)
    
    Dim uom As UnitsOfMeasure = partDoc.UnitsOfMeasure
    
    Dim thicknessValue As Double = 0
    Dim widthValue As Double = 0
    Dim lengthValue As Double = 0
    
    If BoundingBoxStockLib.IsVectorFormat(thicknessAxis) Then
        BoundingBoxStockLib.GetOrientedSizes(partDoc, thicknessAxis, widthAxis, lengthAxis, thicknessValue, widthValue, lengthValue)
    Else
        Dim xSize As Double = 0, ySize As Double = 0, zSize As Double = 0
        BoundingBoxStockLib.GetBoundingBoxSizes(partDoc, xSize, ySize, zSize)
        thicknessValue = BoundingBoxStockLib.GetAxisSize(thicknessAxis, xSize, ySize, zSize)
        widthValue = BoundingBoxStockLib.GetAxisSize(widthAxis, xSize, ySize, zSize)
        lengthValue = BoundingBoxStockLib.GetAxisSize(lengthAxis, xSize, ySize, zSize)
    End If
    
    ' Convert from cm to mm and format
    row.Cells("colT").Value = FormatMm(thicknessValue * 10)
    row.Cells("colW").Value = FormatMm(widthValue * 10)
    row.Cells("colL").Value = FormatMm(lengthValue * 10)
End Sub

' ============================================================================
' Format value in mm
' ============================================================================
Function FormatMm(ByVal valueMm As Double) As String
    Return valueMm.ToString("0.0", System.Globalization.CultureInfo.InvariantCulture)
End Function

' ============================================================================
' Sync grid state to lists
' ============================================================================
Sub SyncGridToLists(ByVal dgv As DataGridView, ByVal selectedFlags As List(Of Boolean))
    For Each row As DataGridViewRow In dgv.Rows
        Dim idx As Integer = CInt(row.Tag)
        If idx >= 0 AndAlso idx < selectedFlags.Count Then
            selectedFlags(idx) = CBool(row.Cells("colSelected").Value)
        End If
    Next
End Sub
