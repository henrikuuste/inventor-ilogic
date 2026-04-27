' Copyright (c) 2026 Henri Kuuste
' ============================================================================
' Määra materjalide välimus - Set material appearances
'
' Allows setting appearances for multiple materials at once using:
' - Wildcard pattern matching (e.g., "HR*", "RG*")
' - Multi-select from filtered list
' - Shows current appearance for each material
'
' Usage: Run from any open part document
' ============================================================================

AddVbFile "Lib/RuntimeLib.vb"
AddVbFile "Lib/UtilsLib.vb"

Imports System.Collections.Generic
Imports System.Windows.Forms
Imports System.Text.RegularExpressions
Imports Inventor

Sub Main()
    If Not AppRuntime.Initialize(ThisApplication) Then Return
    
    Dim app As Inventor.Application = ThisApplication
    UtilsLib.SetLogger(Logger)
    
    Dim doc As Document = app.ActiveDocument
    
    If doc Is Nothing Then
        UtilsLib.LogError("Määra materjalide välimus: No active document.")
        MessageBox.Show("Aktiivne dokument puudub.", "Määra materjalide välimus")
        Exit Sub
    End If
    
    If doc.DocumentType <> DocumentTypeEnum.kPartDocumentObject Then
        UtilsLib.LogError("Määra materjalide välimus: Active document is not a part")
        MessageBox.Show("Aktiivseks dokumendiks peab olema detail (.ipt).", "Määra materjalide välimus")
        Exit Sub
    End If
    
    Dim partDoc As PartDocument = CType(doc, PartDocument)
    
    UtilsLib.LogInfo("Määra materjalide välimus: Starting for " & partDoc.DisplayName)
    
    ' Main loop - dialog reopens after each apply
    Do
        ' Collect all materials with their current render styles
        Dim materialAppearances As New Dictionary(Of String, String)
        Dim materialNames As New List(Of String)
        
        For Each mat As Material In partDoc.Materials
            Dim appearanceName As String = GetMaterialAppearanceName(mat)
            materialAppearances.Add(mat.Name, appearanceName)
            materialNames.Add(mat.Name)
        Next
        materialNames.Sort()
        
        ' Collect document-local appearances from AppearanceAssets
        ' AppearanceAssets contains only appearances stored in the document
        Dim allAppearances As New List(Of String)
        Try
            For Each asset As Asset In partDoc.AppearanceAssets
                allAppearances.Add(asset.DisplayName)
            Next
        Catch
        End Try
        allAppearances.Sort()
        
        ' Show dialog and get results
        Dim dialogData As Object() = ShowMaterialAppearanceDialog(materialNames, materialAppearances, allAppearances)
        
        ' Exit loop if cancelled
        If dialogData Is Nothing Then
            UtilsLib.LogInfo("Määra materjalide välimus: Closed by user")
            Exit Do
        End If
        
        Dim selectedMaterials As List(Of String) = CType(dialogData(0), List(Of String))
        Dim selectedAppearance As String = CStr(dialogData(1))
        
        ' Validate selection
        If selectedMaterials.Count = 0 OrElse String.IsNullOrEmpty(selectedAppearance) Then
            Continue Do
        End If
        
        ' Get the render style
        Dim renderStyle As RenderStyle = Nothing
        Try
            renderStyle = partDoc.RenderStyles.Item(selectedAppearance)
        Catch
            UtilsLib.LogError("Määra materjalide välimus: Could not find render style '" & selectedAppearance & "'")
            Continue Do
        End Try
        
        ' Apply render style to selected materials
        Dim trans As Transaction = app.TransactionManager.StartTransaction(doc, "Määra materjalide välimus")
        
        Try
            Dim successCount As Integer = 0
            
            For Each matName As String In selectedMaterials
                Try
                    Dim mat As Material = partDoc.Materials.Item(matName)
                    mat.RenderStyle = renderStyle
                    successCount += 1
                    UtilsLib.LogInfo("Määra materjalide välimus: Set '" & selectedAppearance & "' on '" & matName & "'")
                Catch ex As Exception
                    UtilsLib.LogWarn("Määra materjalide välimus: Failed on '" & matName & "': " & ex.Message)
                End Try
            Next
            
            trans.End()
            
            UtilsLib.LogInfo("Määra materjalide välimus: Updated " & successCount & " of " & selectedMaterials.Count & " materials")
            
            If successCount > 0 Then
                partDoc.Update()
            End If
            
        Catch ex As Exception
            trans.Abort()
            UtilsLib.LogError("Määra materjalide välimus: Error - " & ex.Message)
        End Try
        
        ' Loop continues - dialog will reopen with refreshed data
    Loop
End Sub

' Get the render style name from a material
Private Function GetMaterialAppearanceName(mat As Material) As String
    Try
        If mat.RenderStyle IsNot Nothing Then
            Return mat.RenderStyle.Name
        End If
    Catch
    End Try
    Return ""
End Function

' Convert wildcard pattern to regex
Private Function WildcardToRegex(pattern As String) As String
    If String.IsNullOrEmpty(pattern) Then Return ".*"
    
    Dim escaped As String = Regex.Escape(pattern)
    escaped = escaped.Replace("\*", ".*")
    escaped = escaped.Replace("\?", ".")
    Return "^" & escaped & "$"
End Function

' Filter materials by wildcard pattern
Private Function FilterMaterials(materials As List(Of String), pattern As String) As List(Of String)
    If String.IsNullOrEmpty(pattern) Then Return materials
    
    Dim result As New List(Of String)
    Dim regexPattern As String = WildcardToRegex(pattern)
    
    Try
        Dim regex As New Regex(regexPattern, RegexOptions.IgnoreCase)
        For Each mat As String In materials
            If regex.IsMatch(mat) Then
                result.Add(mat)
            End If
        Next
    Catch
        Return materials
    End Try
    
    Return result
End Function

' Format material display text with current appearance
Private Function FormatMaterialDisplay(matName As String, matAppearances As Dictionary(Of String, String)) As String
    Dim appearance As String = ""
    If matAppearances.ContainsKey(matName) Then
        appearance = matAppearances(matName)
    End If
    
    If String.IsNullOrEmpty(appearance) Then
        Return matName & "  [välimus puudub]"
    Else
        Return matName & "  [" & appearance & "]"
    End If
End Function

' Extract material name from display text
Private Function ExtractMaterialName(displayText As String) As String
    Dim bracketPos As Integer = displayText.IndexOf("  [")
    If bracketPos > 0 Then
        Return displayText.Substring(0, bracketPos)
    End If
    Return displayText
End Function

' Show the material appearance dialog - returns Object() {selectedMaterials, selectedAppearance} or Nothing if cancelled
Private Function ShowMaterialAppearanceDialog(
        materialNames As List(Of String),
        materialAppearances As Dictionary(Of String, String),
        allAppearances As List(Of String)) As Object()
    
    Dim frm As New System.Windows.Forms.Form()
    frm.Text = "Määra materjalide välimus"
    frm.Width = 550
    frm.Height = 580
    frm.StartPosition = FormStartPosition.CenterScreen
    frm.FormBorderStyle = FormBorderStyle.FixedDialog
    frm.MaximizeBox = False
    frm.MinimizeBox = False
    
    ' Store data in form Tag for access in handlers: {materialNames, materialAppearances}
    frm.Tag = New Object() {materialNames, materialAppearances}
    
    Dim currentY As Integer = 15
    
    ' Pattern label
    Dim lblPattern As New System.Windows.Forms.Label()
    lblPattern.Text = "Muster (nt. HR*, RG*, *Steel*):"
    lblPattern.Left = 15
    lblPattern.Top = currentY
    lblPattern.Width = 510
    lblPattern.Height = 20
    frm.Controls.Add(lblPattern)
    currentY += 25
    
    ' Pattern textbox
    Dim txtPattern As New System.Windows.Forms.TextBox()
    txtPattern.Name = "txtPattern"
    txtPattern.Left = 15
    txtPattern.Top = currentY
    txtPattern.Width = 400
    txtPattern.Height = 23
    frm.Controls.Add(txtPattern)
    
    ' Filter button
    Dim btnFilter As New System.Windows.Forms.Button()
    btnFilter.Text = "Filtreeri"
    btnFilter.Left = 425
    btnFilter.Top = currentY
    btnFilter.Width = 100
    btnFilter.Height = 25
    frm.Controls.Add(btnFilter)
    currentY += 35
    
    ' Materials label
    Dim lblMaterials As New System.Windows.Forms.Label()
    lblMaterials.Text = "Materjalid [praegune välimus] - vali üks või mitu:"
    lblMaterials.Left = 15
    lblMaterials.Top = currentY
    lblMaterials.Width = 510
    lblMaterials.Height = 20
    frm.Controls.Add(lblMaterials)
    currentY += 25
    
    ' Materials listbox (multiselect)
    Dim lstMaterials As New System.Windows.Forms.ListBox()
    lstMaterials.Name = "lstMaterials"
    lstMaterials.Left = 15
    lstMaterials.Top = currentY
    lstMaterials.Width = 510
    lstMaterials.Height = 220
    lstMaterials.SelectionMode = SelectionMode.MultiExtended
    For Each matName As String In materialNames
        lstMaterials.Items.Add(FormatMaterialDisplay(matName, materialAppearances))
    Next
    frm.Controls.Add(lstMaterials)
    currentY += 230
    
    ' Select all / clear buttons
    Dim btnSelectAll As New System.Windows.Forms.Button()
    btnSelectAll.Text = "Vali kõik"
    btnSelectAll.Left = 15
    btnSelectAll.Top = currentY
    btnSelectAll.Width = 100
    btnSelectAll.Height = 25
    frm.Controls.Add(btnSelectAll)
    
    Dim btnClearSelection As New System.Windows.Forms.Button()
    btnClearSelection.Text = "Tühista valik"
    btnClearSelection.Left = 125
    btnClearSelection.Top = currentY
    btnClearSelection.Width = 100
    btnClearSelection.Height = 25
    frm.Controls.Add(btnClearSelection)
    currentY += 40
    
    ' Appearance label
    Dim lblAppearance As New System.Windows.Forms.Label()
    lblAppearance.Text = "Uus välimus:"
    lblAppearance.Left = 15
    lblAppearance.Top = currentY
    lblAppearance.Width = 510
    lblAppearance.Height = 20
    frm.Controls.Add(lblAppearance)
    currentY += 25
    
    ' Appearance combobox
    Dim cboAppearance As New System.Windows.Forms.ComboBox()
    cboAppearance.Name = "cboAppearance"
    cboAppearance.Left = 15
    cboAppearance.Top = currentY
    cboAppearance.Width = 510
    cboAppearance.Height = 23
    cboAppearance.DropDownStyle = ComboBoxStyle.DropDownList
    cboAppearance.Items.Add("")
    For Each appearanceName As String In allAppearances
        cboAppearance.Items.Add(appearanceName)
    Next
    cboAppearance.SelectedIndex = 0
    frm.Controls.Add(cboAppearance)
    currentY += 45
    
    ' Status label
    Dim lblStatus As New System.Windows.Forms.Label()
    lblStatus.Name = "lblStatus"
    lblStatus.Text = "Valitud: 0 materjali"
    lblStatus.Left = 15
    lblStatus.Top = currentY
    lblStatus.Width = 510
    lblStatus.Height = 20
    frm.Controls.Add(lblStatus)
    currentY += 35
    
    ' Buttons panel
    Dim btnApply As New System.Windows.Forms.Button()
    btnApply.Text = "Rakenda"
    btnApply.Left = 310
    btnApply.Top = currentY
    btnApply.Width = 100
    btnApply.Height = 30
    frm.Controls.Add(btnApply)
    
    Dim btnCancel As New System.Windows.Forms.Button()
    btnCancel.Text = "Tühista"
    btnCancel.Left = 420
    btnCancel.Top = currentY
    btnCancel.Width = 100
    btnCancel.Height = 30
    frm.Controls.Add(btnCancel)
    
    ' Event handlers - using local variables only (no ByRef in lambdas)
    AddHandler btnFilter.Click, Sub(s, e)
        Dim tagData As Object() = CType(frm.Tag, Object())
        Dim storedNames As List(Of String) = CType(tagData(0), List(Of String))
        Dim storedAppearances As Dictionary(Of String, String) = CType(tagData(1), Dictionary(Of String, String))
        Dim filtered As List(Of String) = FilterMaterials(storedNames, txtPattern.Text)
        lstMaterials.Items.Clear()
        For Each matName As String In filtered
            lstMaterials.Items.Add(FormatMaterialDisplay(matName, storedAppearances))
        Next
        lblStatus.Text = "Filtreeritud: " & filtered.Count & " materjali"
    End Sub
    
    AddHandler txtPattern.KeyDown, Sub(s, e)
        If e.KeyCode = Keys.Enter Then
            Dim tagData As Object() = CType(frm.Tag, Object())
            Dim storedNames As List(Of String) = CType(tagData(0), List(Of String))
            Dim storedAppearances As Dictionary(Of String, String) = CType(tagData(1), Dictionary(Of String, String))
            Dim filtered As List(Of String) = FilterMaterials(storedNames, txtPattern.Text)
            lstMaterials.Items.Clear()
            For Each matName As String In filtered
                lstMaterials.Items.Add(FormatMaterialDisplay(matName, storedAppearances))
            Next
            lblStatus.Text = "Filtreeritud: " & filtered.Count & " materjali"
            e.SuppressKeyPress = True
        End If
    End Sub
    
    AddHandler btnSelectAll.Click, Sub(s, e)
        For i As Integer = 0 To lstMaterials.Items.Count - 1
            lstMaterials.SetSelected(i, True)
        Next
        lblStatus.Text = "Valitud: " & lstMaterials.SelectedItems.Count & " materjali"
    End Sub
    
    AddHandler btnClearSelection.Click, Sub(s, e)
        lstMaterials.ClearSelected()
        lblStatus.Text = "Valitud: 0 materjali"
    End Sub
    
    AddHandler lstMaterials.SelectedIndexChanged, Sub(s, e)
        lblStatus.Text = "Valitud: " & lstMaterials.SelectedItems.Count & " materjali"
    End Sub
    
    AddHandler btnApply.Click, Sub(s, e)
        frm.DialogResult = DialogResult.OK
        frm.Close()
    End Sub
    
    AddHandler btnCancel.Click, Sub(s, e)
        frm.DialogResult = DialogResult.Cancel
        frm.Close()
    End Sub
    
    ' Show dialog
    Dim result As DialogResult = frm.ShowDialog()
    
    ' Read values AFTER dialog closes (avoiding ByRef in lambda)
    If result = DialogResult.OK Then
        Dim selectedMaterials As New List(Of String)
        
        ' Extract material names from selected display items
        For Each item As Object In lstMaterials.SelectedItems
            Dim displayText As String = CStr(item)
            Dim matName As String = ExtractMaterialName(displayText)
            selectedMaterials.Add(matName)
        Next
        
        Dim selectedAppearance As String = ""
        If cboAppearance.SelectedItem IsNot Nothing Then
            selectedAppearance = CStr(cboAppearance.SelectedItem)
        End If
        
        Return New Object() {selectedMaterials, selectedAppearance}
    End If
    
    Return Nothing
End Function
