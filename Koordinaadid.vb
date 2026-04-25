' ============================================================================
' Koordinaadid - UCS Creation at Bounding Box Positions
' 
' Creates a User Coordinate System (UCS) at any position on the bounding box
' of selected assembly components. Supports corners, face centers, edge centers,
' and the box center. Live preview updates as user changes options.
'
' Usage:
' 1. Select one or more occurrences in an assembly
' 2. Run this rule
' 3. Choose position (X/Y/Z: Min/Keskpunkt/Max)
' 4. Choose orientation (which global axes map to UCS axes)
' 5. Enter a name for the UCS
' 6. Click OK to create, or Cancel to abort
'
' ============================================================================

AddVbFile "Lib/UtilsLib.vb"

Sub Main()
    Dim app As Inventor.Application = ThisApplication
    Dim doc As Document = app.ActiveDocument
    
    ' Validate document type
    If doc Is Nothing Then
        Logger.Error("Koordinaadid: No active document")
        MessageBox.Show("Aktiivne dokument puudub.", "Koordinaadid")
        Exit Sub
    End If
    
    If doc.DocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
        Logger.Error("Koordinaadid: Not an assembly document")
        MessageBox.Show("See reegel töötab ainult koostudokumentides (.iam).", "Koordinaadid")
        Exit Sub
    End If
    
    Dim asmDoc As AssemblyDocument = CType(doc, AssemblyDocument)
    Dim asmDef As AssemblyComponentDefinition = asmDoc.ComponentDefinition
    
    ' Get selected occurrences
    Dim selectedOccs As New System.Collections.Generic.List(Of ComponentOccurrence)
    
    For Each selObj As Object In asmDoc.SelectSet
        If TypeOf selObj Is ComponentOccurrence Then
            selectedOccs.Add(CType(selObj, ComponentOccurrence))
        End If
    Next
    
    If selectedOccs.Count = 0 Then
        Logger.Error("Koordinaadid: No occurrences selected")
        MessageBox.Show("Palun vali vähemalt üks komponent.", "Koordinaadid")
        Exit Sub
    End If
    
    Logger.Info("Koordinaadid: " & selectedOccs.Count.ToString() & " occurrence(s) selected")
    
    ' Calculate combined bounding box
    Dim minX As Double = Double.MaxValue
    Dim minY As Double = Double.MaxValue
    Dim minZ As Double = Double.MaxValue
    Dim maxX As Double = Double.MinValue
    Dim maxY As Double = Double.MinValue
    Dim maxZ As Double = Double.MinValue
    
    For Each occ As ComponentOccurrence In selectedOccs
        Dim occBox As Box = occ.RangeBox
        If occBox.MinPoint.X < minX Then minX = occBox.MinPoint.X
        If occBox.MinPoint.Y < minY Then minY = occBox.MinPoint.Y
        If occBox.MinPoint.Z < minZ Then minZ = occBox.MinPoint.Z
        If occBox.MaxPoint.X > maxX Then maxX = occBox.MaxPoint.X
        If occBox.MaxPoint.Y > maxY Then maxY = occBox.MaxPoint.Y
        If occBox.MaxPoint.Z > maxZ Then maxZ = occBox.MaxPoint.Z
    Next
    
    Logger.Info("Koordinaadid: Bounding box calculated - " & _
                "X: " & (minX * 10).ToString("0.0") & " to " & (maxX * 10).ToString("0.0") & " mm, " & _
                "Y: " & (minY * 10).ToString("0.0") & " to " & (maxY * 10).ToString("0.0") & " mm, " & _
                "Z: " & (minZ * 10).ToString("0.0") & " to " & (maxZ * 10).ToString("0.0") & " mm")
    
    ' Run the UCS creation dialog
    RunUcsDialog(app, asmDoc, selectedOccs.Count, minX, minY, minZ, maxX, maxY, maxZ)
End Sub

' ============================================================================
' State class to hold shared data for event handlers
' ============================================================================

Class UcsState
    Public App As Inventor.Application
    Public AsmDoc As AssemblyDocument
    Public PreviewUcs As UserCoordinateSystem
    Public MinX As Double
    Public MinY As Double
    Public MinZ As Double
    Public MaxX As Double
    Public MaxY As Double
    Public MaxZ As Double
    Public Confirmed As Boolean
End Class

' ============================================================================
' Main Dialog
' ============================================================================

Sub RunUcsDialog(app As Inventor.Application, asmDoc As AssemblyDocument, _
                 occCount As Integer, _
                 minX As Double, minY As Double, minZ As Double, _
                 maxX As Double, maxY As Double, maxZ As Double)
    
    Dim asmDef As AssemblyComponentDefinition = asmDoc.ComponentDefinition
    Dim tg As TransientGeometry = app.TransientGeometry
    
    ' Create state object
    Dim state As New UcsState()
    state.App = app
    state.AsmDoc = asmDoc
    state.MinX = minX
    state.MinY = minY
    state.MinZ = minZ
    state.MaxX = maxX
    state.MaxY = maxY
    state.MaxZ = maxZ
    state.Confirmed = False
    state.PreviewUcs = Nothing
    
    ' Generate unique default name
    Dim baseName As String = "UCS"
    Dim ucsName As String = GenerateUniqueName(asmDef, baseName)
    
    ' Create preview UCS at default position (Min, Min, Min with default orientation)
    Dim defaultOrigin As Point = tg.CreatePoint(minX, minY, minZ)
    Dim xAxis As Vector = tg.CreateVector(1, 0, 0)
    Dim yAxis As Vector = tg.CreateVector(0, 1, 0)
    Dim zAxis As Vector = tg.CreateVector(0, 0, 1)
    
    state.PreviewUcs = CreateUcs(app, asmDef, "_Preview_UCS", defaultOrigin, xAxis, yAxis, zAxis)
    
    If state.PreviewUcs Is Nothing Then
        Logger.Error("Koordinaadid: Failed to create preview UCS")
        MessageBox.Show("UCS-i loomine ebaõnnestus.", "Koordinaadid")
        Exit Sub
    End If
    
    app.ActiveView.Update()
    
    ' Create form
    Dim frm As New System.Windows.Forms.Form()
    frm.Text = "Koordinaadid - Loo UCS"
    frm.Width = 380
    frm.Height = 420
    frm.StartPosition = FormStartPosition.Manual
    frm.Left = 100
    frm.Top = 100
    frm.FormBorderStyle = FormBorderStyle.FixedToolWindow
    frm.MaximizeBox = False
    frm.MinimizeBox = False
    frm.TopMost = True
    frm.Tag = state
    
    Dim yPos As Integer = 15
    Dim labelWidth As Integer = 100
    Dim comboWidth As Integer = 90
    
    ' --- Selection count ---
    Dim lblCount As New System.Windows.Forms.Label()
    lblCount.Text = "Valitud elemente: " & occCount.ToString()
    lblCount.Left = 15
    lblCount.Top = yPos
    lblCount.Width = 300
    frm.Controls.Add(lblCount)
    yPos += 30
    
    ' --- Position section ---
    Dim lblPosSection As New System.Windows.Forms.Label()
    lblPosSection.Text = "--- Asukoht ---"
    lblPosSection.Left = 15
    lblPosSection.Top = yPos
    lblPosSection.Width = 300
    frm.Controls.Add(lblPosSection)
    yPos += 25
    
    ' Position dropdowns (X, Y, Z)
    Dim posOptions() As String = {"Min", "Keskpunkt", "Max"}
    
    Dim lblPosX As New System.Windows.Forms.Label()
    lblPosX.Text = "X:"
    lblPosX.Left = 15
    lblPosX.Top = yPos + 3
    lblPosX.Width = 20
    frm.Controls.Add(lblPosX)
    
    Dim cboPosX As New System.Windows.Forms.ComboBox()
    cboPosX.Name = "cboPosX"
    cboPosX.Left = 35
    cboPosX.Top = yPos
    cboPosX.Width = comboWidth
    cboPosX.DropDownStyle = ComboBoxStyle.DropDownList
    cboPosX.Items.AddRange(posOptions)
    cboPosX.SelectedIndex = 0
    frm.Controls.Add(cboPosX)
    
    Dim lblPosY As New System.Windows.Forms.Label()
    lblPosY.Text = "Y:"
    lblPosY.Left = 135
    lblPosY.Top = yPos + 3
    lblPosY.Width = 20
    frm.Controls.Add(lblPosY)
    
    Dim cboPosY As New System.Windows.Forms.ComboBox()
    cboPosY.Name = "cboPosY"
    cboPosY.Left = 155
    cboPosY.Top = yPos
    cboPosY.Width = comboWidth
    cboPosY.DropDownStyle = ComboBoxStyle.DropDownList
    cboPosY.Items.AddRange(posOptions)
    cboPosY.SelectedIndex = 0
    frm.Controls.Add(cboPosY)
    
    Dim lblPosZ As New System.Windows.Forms.Label()
    lblPosZ.Text = "Z:"
    lblPosZ.Left = 255
    lblPosZ.Top = yPos + 3
    lblPosZ.Width = 20
    frm.Controls.Add(lblPosZ)
    
    Dim cboPosZ As New System.Windows.Forms.ComboBox()
    cboPosZ.Name = "cboPosZ"
    cboPosZ.Left = 275
    cboPosZ.Top = yPos
    cboPosZ.Width = comboWidth
    cboPosZ.DropDownStyle = ComboBoxStyle.DropDownList
    cboPosZ.Items.AddRange(posOptions)
    cboPosZ.SelectedIndex = 0
    frm.Controls.Add(cboPosZ)
    
    yPos += 30
    
    ' --- Offset section ---
    Dim lblOffsetSection As New System.Windows.Forms.Label()
    lblOffsetSection.Text = "--- Nihe (mm) ---"
    lblOffsetSection.Left = 15
    lblOffsetSection.Top = yPos
    lblOffsetSection.Width = 300
    frm.Controls.Add(lblOffsetSection)
    yPos += 25
    
    Dim lblOffX As New System.Windows.Forms.Label()
    lblOffX.Text = "X:"
    lblOffX.Left = 15
    lblOffX.Top = yPos + 3
    lblOffX.Width = 20
    frm.Controls.Add(lblOffX)
    
    Dim txtOffX As New System.Windows.Forms.TextBox()
    txtOffX.Name = "txtOffX"
    txtOffX.Left = 35
    txtOffX.Top = yPos
    txtOffX.Width = 60
    txtOffX.Text = "0"
    frm.Controls.Add(txtOffX)
    
    Dim lblOffY As New System.Windows.Forms.Label()
    lblOffY.Text = "Y:"
    lblOffY.Left = 110
    lblOffY.Top = yPos + 3
    lblOffY.Width = 20
    frm.Controls.Add(lblOffY)
    
    Dim txtOffY As New System.Windows.Forms.TextBox()
    txtOffY.Name = "txtOffY"
    txtOffY.Left = 130
    txtOffY.Top = yPos
    txtOffY.Width = 60
    txtOffY.Text = "0"
    frm.Controls.Add(txtOffY)
    
    Dim lblOffZ As New System.Windows.Forms.Label()
    lblOffZ.Text = "Z:"
    lblOffZ.Left = 205
    lblOffZ.Top = yPos + 3
    lblOffZ.Width = 20
    frm.Controls.Add(lblOffZ)
    
    Dim txtOffZ As New System.Windows.Forms.TextBox()
    txtOffZ.Name = "txtOffZ"
    txtOffZ.Left = 225
    txtOffZ.Top = yPos
    txtOffZ.Width = 60
    txtOffZ.Text = "0"
    frm.Controls.Add(txtOffZ)
    
    ' Update on offset change
    AddHandler txtOffX.TextChanged, AddressOf OnOffsetChanged
    AddHandler txtOffY.TextChanged, AddressOf OnOffsetChanged
    AddHandler txtOffZ.TextChanged, AddressOf OnOffsetChanged
    txtOffX.Tag = frm
    txtOffY.Tag = frm
    txtOffZ.Tag = frm
    
    yPos += 35
    
    ' --- Orientation section ---
    Dim lblOrientSection As New System.Windows.Forms.Label()
    lblOrientSection.Text = "--- Orientatsioon ---"
    lblOrientSection.Left = 15
    lblOrientSection.Top = yPos
    lblOrientSection.Width = 300
    frm.Controls.Add(lblOrientSection)
    yPos += 25
    
    Dim orientOptions() As String = {"+X", "-X", "+Y", "-Y", "+Z", "-Z"}
    
    ' UCS X-axis
    Dim lblOrientX As New System.Windows.Forms.Label()
    lblOrientX.Text = "UCS X-telg:"
    lblOrientX.Left = 15
    lblOrientX.Top = yPos + 3
    lblOrientX.Width = 70
    frm.Controls.Add(lblOrientX)
    
    Dim cboOrientX As New System.Windows.Forms.ComboBox()
    cboOrientX.Name = "cboOrientX"
    cboOrientX.Left = 90
    cboOrientX.Top = yPos
    cboOrientX.Width = 70
    cboOrientX.DropDownStyle = ComboBoxStyle.DropDownList
    cboOrientX.Items.AddRange(orientOptions)
    cboOrientX.SelectedIndex = 0
    frm.Controls.Add(cboOrientX)
    
    yPos += 30
    
    ' UCS Y-axis
    Dim lblOrientY As New System.Windows.Forms.Label()
    lblOrientY.Text = "UCS Y-telg:"
    lblOrientY.Left = 15
    lblOrientY.Top = yPos + 3
    lblOrientY.Width = 70
    frm.Controls.Add(lblOrientY)
    
    Dim cboOrientY As New System.Windows.Forms.ComboBox()
    cboOrientY.Name = "cboOrientY"
    cboOrientY.Left = 90
    cboOrientY.Top = yPos
    cboOrientY.Width = 70
    cboOrientY.DropDownStyle = ComboBoxStyle.DropDownList
    cboOrientY.Items.AddRange(orientOptions)
    cboOrientY.SelectedIndex = 2
    frm.Controls.Add(cboOrientY)
    
    yPos += 30
    
    ' UCS Z-axis (auto-calculated, display only)
    Dim lblOrientZ As New System.Windows.Forms.Label()
    lblOrientZ.Text = "UCS Z-telg:"
    lblOrientZ.Left = 15
    lblOrientZ.Top = yPos + 3
    lblOrientZ.Width = 70
    frm.Controls.Add(lblOrientZ)
    
    Dim txtOrientZ As New System.Windows.Forms.TextBox()
    txtOrientZ.Name = "txtOrientZ"
    txtOrientZ.Left = 90
    txtOrientZ.Top = yPos
    txtOrientZ.Width = 70
    txtOrientZ.ReadOnly = True
    txtOrientZ.Text = "+Z"
    frm.Controls.Add(txtOrientZ)
    
    Dim lblOrientZNote As New System.Windows.Forms.Label()
    lblOrientZNote.Text = "(automaatne)"
    lblOrientZNote.Left = 165
    lblOrientZNote.Top = yPos + 3
    lblOrientZNote.Width = 80
    frm.Controls.Add(lblOrientZNote)
    
    yPos += 35
    
    ' --- Name ---
    Dim lblName As New System.Windows.Forms.Label()
    lblName.Text = "Nimi:"
    lblName.Left = 15
    lblName.Top = yPos + 3
    lblName.Width = 50
    frm.Controls.Add(lblName)
    
    Dim txtName As New System.Windows.Forms.TextBox()
    txtName.Name = "txtName"
    txtName.Left = 70
    txtName.Top = yPos
    txtName.Width = 200
    txtName.Text = ucsName
    frm.Controls.Add(txtName)
    
    yPos += 40
    
    ' --- Buttons ---
    Dim btnOK As New System.Windows.Forms.Button()
    btnOK.Name = "btnOK"
    btnOK.Text = "OK"
    btnOK.Left = 100
    btnOK.Top = yPos
    btnOK.Width = 80
    btnOK.Height = 30
    btnOK.Tag = frm
    frm.Controls.Add(btnOK)
    
    Dim btnCancel As New System.Windows.Forms.Button()
    btnCancel.Text = "Tühista"
    btnCancel.Left = 190
    btnCancel.Top = yPos
    btnCancel.Width = 80
    btnCancel.Height = 30
    btnCancel.DialogResult = DialogResult.Cancel
    frm.CancelButton = btnCancel
    frm.Controls.Add(btnCancel)
    
    ' Button click handlers
    AddHandler btnOK.Click, AddressOf OnOkClicked
    AddHandler btnCancel.Click, AddressOf OnCancelClicked
    btnCancel.Tag = frm
    
    ' --- Event handlers for live preview ---
    AddHandler cboPosX.SelectedIndexChanged, AddressOf OnSelectionChanged
    AddHandler cboPosY.SelectedIndexChanged, AddressOf OnSelectionChanged
    AddHandler cboPosZ.SelectedIndexChanged, AddressOf OnSelectionChanged
    AddHandler cboOrientX.SelectedIndexChanged, AddressOf OnOrientationChanged
    AddHandler cboOrientY.SelectedIndexChanged, AddressOf OnOrientationChanged
    
    ' Store combo references in Tag for access from handlers
    cboPosX.Tag = frm
    cboPosY.Tag = frm
    cboPosZ.Tag = frm
    cboOrientX.Tag = frm
    cboOrientY.Tag = frm
    
    ' Handle form closing to cleanup if cancelled
    AddHandler frm.FormClosing, AddressOf OnFormClosing
    
    ' Show form as modeless (allows viewport interaction)
    frm.Show()
    
    ' Process messages while form is open
    Do While frm.Visible
        System.Windows.Forms.Application.DoEvents()
        System.Threading.Thread.Sleep(10)
    Loop
    
    ' Result is handled in OnOkClicked and OnFormClosing
End Sub

' ============================================================================
' Event Handlers
' ============================================================================

Sub OnOkClicked(sender As Object, e As EventArgs)
    Dim btn As System.Windows.Forms.Button = CType(sender, System.Windows.Forms.Button)
    Dim frm As System.Windows.Forms.Form = CType(btn.Tag, System.Windows.Forms.Form)
    Dim state As UcsState = CType(frm.Tag, UcsState)
    
    ' Mark as confirmed BEFORE form closes
    state.Confirmed = True
    
    ' Get final name and rename UCS
    Dim txtName As System.Windows.Forms.TextBox = CType(frm.Controls("txtName"), System.Windows.Forms.TextBox)
    Dim finalName As String = txtName.Text.Trim()
    If finalName = "" Then finalName = "UCS_1"
    
    ' Rename the preview UCS to final name
    If state.PreviewUcs IsNot Nothing Then
        Try
            state.PreviewUcs.Name = finalName
            Logger.Info("Koordinaadid: UCS '" & finalName & "' created successfully")
        Catch ex As Exception
            Logger.Warn("Koordinaadid: Could not rename UCS - " & ex.Message)
        End Try
    End If
    
    ' Close the form
    frm.DialogResult = DialogResult.OK
    frm.Close()
End Sub

Sub OnCancelClicked(sender As Object, e As EventArgs)
    Dim btn As System.Windows.Forms.Button = CType(sender, System.Windows.Forms.Button)
    Dim frm As System.Windows.Forms.Form = CType(btn.Tag, System.Windows.Forms.Form)
    frm.Close()
End Sub

Sub OnSelectionChanged(sender As Object, e As EventArgs)
    Dim cbo As System.Windows.Forms.ComboBox = CType(sender, System.Windows.Forms.ComboBox)
    Dim frm As System.Windows.Forms.Form = CType(cbo.Tag, System.Windows.Forms.Form)
    UpdatePreviewUcs(frm)
End Sub

Sub OnOrientationChanged(sender As Object, e As EventArgs)
    Dim cbo As System.Windows.Forms.ComboBox = CType(sender, System.Windows.Forms.ComboBox)
    Dim frm As System.Windows.Forms.Form = CType(cbo.Tag, System.Windows.Forms.Form)
    
    ' Update Z-axis display based on X and Y selections
    UpdateZAxisDisplay(frm)
    
    ' Update the UCS
    UpdatePreviewUcs(frm)
End Sub

Sub OnOffsetChanged(sender As Object, e As EventArgs)
    Dim txt As System.Windows.Forms.TextBox = CType(sender, System.Windows.Forms.TextBox)
    Dim frm As System.Windows.Forms.Form = CType(txt.Tag, System.Windows.Forms.Form)
    UpdatePreviewUcs(frm)
End Sub

Sub OnFormClosing(sender As Object, e As FormClosingEventArgs)
    Dim frm As System.Windows.Forms.Form = CType(sender, System.Windows.Forms.Form)
    Dim state As UcsState = CType(frm.Tag, UcsState)
    
    ' If not confirmed, delete the preview UCS
    If Not state.Confirmed AndAlso state.PreviewUcs IsNot Nothing Then
        Try
            state.PreviewUcs.Delete()
            state.App.ActiveView.Update()
            Logger.Info("Koordinaadid: Preview UCS deleted (cancelled)")
        Catch ex As Exception
            Logger.Warn("Koordinaadid: Could not delete preview UCS - " & ex.Message)
        End Try
    End If
End Sub

' ============================================================================
' UCS Update Logic
' ============================================================================

Sub UpdatePreviewUcs(frm As System.Windows.Forms.Form)
    Dim state As UcsState = CType(frm.Tag, UcsState)
    If state Is Nothing OrElse state.PreviewUcs Is Nothing Then Exit Sub
    
    ' Get position selections
    Dim cboPosX As System.Windows.Forms.ComboBox = CType(frm.Controls("cboPosX"), System.Windows.Forms.ComboBox)
    Dim cboPosY As System.Windows.Forms.ComboBox = CType(frm.Controls("cboPosY"), System.Windows.Forms.ComboBox)
    Dim cboPosZ As System.Windows.Forms.ComboBox = CType(frm.Controls("cboPosZ"), System.Windows.Forms.ComboBox)
    
    ' Get offset values (in mm, convert to cm for internal units)
    Dim txtOffX As System.Windows.Forms.TextBox = CType(frm.Controls("txtOffX"), System.Windows.Forms.TextBox)
    Dim txtOffY As System.Windows.Forms.TextBox = CType(frm.Controls("txtOffY"), System.Windows.Forms.TextBox)
    Dim txtOffZ As System.Windows.Forms.TextBox = CType(frm.Controls("txtOffZ"), System.Windows.Forms.TextBox)
    
    Dim offX As Double = 0
    Dim offY As Double = 0
    Dim offZ As Double = 0
    Double.TryParse(txtOffX.Text, offX)
    Double.TryParse(txtOffY.Text, offY)
    Double.TryParse(txtOffZ.Text, offZ)
    
    ' Convert mm to cm (internal units)
    offX = offX / 10.0
    offY = offY / 10.0
    offZ = offZ / 10.0
    
    ' Get orientation selections
    Dim cboOrientX As System.Windows.Forms.ComboBox = CType(frm.Controls("cboOrientX"), System.Windows.Forms.ComboBox)
    Dim cboOrientY As System.Windows.Forms.ComboBox = CType(frm.Controls("cboOrientY"), System.Windows.Forms.ComboBox)
    
    ' Calculate base position
    Dim posX As Double = GetAxisValue(CStr(cboPosX.SelectedItem), state.MinX, state.MaxX)
    Dim posY As Double = GetAxisValue(CStr(cboPosY.SelectedItem), state.MinY, state.MaxY)
    Dim posZ As Double = GetAxisValue(CStr(cboPosZ.SelectedItem), state.MinZ, state.MaxZ)
    
    ' Apply offsets
    posX = posX + offX
    posY = posY + offY
    posZ = posZ + offZ
    
    ' Get orientation vectors
    Dim tg As TransientGeometry = state.App.TransientGeometry
    Dim origin As Point = tg.CreatePoint(posX, posY, posZ)
    
    Dim xAxis As Vector = GetDirectionVector(tg, CStr(cboOrientX.SelectedItem))
    Dim yAxis As Vector = GetDirectionVector(tg, CStr(cboOrientY.SelectedItem))
    
    ' Calculate Z as cross product of X and Y
    Dim zAxis As Vector = xAxis.CrossProduct(yAxis)
    
    ' Check if X and Y are parallel (invalid)
    If zAxis.Length < 0.001 Then
        ' Invalid orientation - X and Y are parallel, don't update
        Exit Sub
    End If
    
    zAxis.Normalize()
    
    ' Build transformation matrix
    Dim m As Matrix = tg.CreateMatrix()
    m.SetCoordinateSystem(origin, xAxis, yAxis, zAxis)
    
    ' Update UCS transformation
    Try
        state.PreviewUcs.Transformation = m
        state.App.ActiveView.Update()
    Catch ex As Exception
        Logger.Warn("Koordinaadid: Could not update UCS - " & ex.Message)
    End Try
End Sub

Sub UpdateZAxisDisplay(frm As System.Windows.Forms.Form)
    Dim cboOrientX As System.Windows.Forms.ComboBox = CType(frm.Controls("cboOrientX"), System.Windows.Forms.ComboBox)
    Dim cboOrientY As System.Windows.Forms.ComboBox = CType(frm.Controls("cboOrientY"), System.Windows.Forms.ComboBox)
    Dim txtOrientZ As System.Windows.Forms.TextBox = CType(frm.Controls("txtOrientZ"), System.Windows.Forms.TextBox)
    
    Dim state As UcsState = CType(frm.Tag, UcsState)
    Dim tg As TransientGeometry = state.App.TransientGeometry
    
    Dim xAxis As Vector = GetDirectionVector(tg, CStr(cboOrientX.SelectedItem))
    Dim yAxis As Vector = GetDirectionVector(tg, CStr(cboOrientY.SelectedItem))
    Dim zAxis As Vector = xAxis.CrossProduct(yAxis)
    
    If zAxis.Length < 0.001 Then
        txtOrientZ.Text = "(viga)"
    Else
        zAxis.Normalize()
        txtOrientZ.Text = VectorToAxisName(zAxis)
    End If
End Sub

' ============================================================================
' Helper Functions
' ============================================================================

Function GetAxisValue(selection As String, minVal As Double, maxVal As Double) As Double
    Select Case selection
        Case "Min"
            Return minVal
        Case "Keskpunkt"
            Return (minVal + maxVal) / 2
        Case "Max"
            Return maxVal
        Case Else
            Return minVal
    End Select
End Function

Function GetDirectionVector(tg As TransientGeometry, axisName As String) As Vector
    Select Case axisName
        Case "+X"
            Return tg.CreateVector(1, 0, 0)
        Case "-X"
            Return tg.CreateVector(-1, 0, 0)
        Case "+Y"
            Return tg.CreateVector(0, 1, 0)
        Case "-Y"
            Return tg.CreateVector(0, -1, 0)
        Case "+Z"
            Return tg.CreateVector(0, 0, 1)
        Case "-Z"
            Return tg.CreateVector(0, 0, -1)
        Case Else
            Return tg.CreateVector(1, 0, 0)
    End Select
End Function

Function VectorToAxisName(v As Vector) As String
    Dim threshold As Double = 0.9
    
    If v.X > threshold Then Return "+X"
    If v.X < -threshold Then Return "-X"
    If v.Y > threshold Then Return "+Y"
    If v.Y < -threshold Then Return "-Y"
    If v.Z > threshold Then Return "+Z"
    If v.Z < -threshold Then Return "-Z"
    
    ' Not aligned with principal axis
    Return "(" & v.X.ToString("0.00") & ", " & v.Y.ToString("0.00") & ", " & v.Z.ToString("0.00") & ")"
End Function

Function CreateUcs(app As Inventor.Application, asmDef As AssemblyComponentDefinition, _
                   ucsName As String, origin As Point, xAxis As Vector, yAxis As Vector, zAxis As Vector) As UserCoordinateSystem
    Try
        Dim tg As TransientGeometry = app.TransientGeometry
        
        ' Build transformation matrix
        Dim m As Matrix = tg.CreateMatrix()
        m.SetCoordinateSystem(origin, xAxis, yAxis, zAxis)
        
        ' Create UCS definition and add
        Dim ucsDef As UserCoordinateSystemDefinition = asmDef.UserCoordinateSystems.CreateDefinition
        ucsDef.Transformation = m
        
        Dim ucs As UserCoordinateSystem = asmDef.UserCoordinateSystems.Add(ucsDef)
        ucs.Name = ucsName
        
        Return ucs
    Catch ex As Exception
        Logger.Error("Koordinaadid: CreateUcs failed - " & ex.Message)
        Return Nothing
    End Try
End Function

Function GenerateUniqueName(asmDef As AssemblyComponentDefinition, baseName As String) As String
    Dim index As Integer = 1
    Dim candidateName As String = baseName & "_" & index.ToString()
    
    ' Check existing UCS names
    Dim existingNames As New System.Collections.Generic.HashSet(Of String)
    For Each ucs As UserCoordinateSystem In asmDef.UserCoordinateSystems
        existingNames.Add(ucs.Name)
    Next
    
    While existingNames.Contains(candidateName)
        index += 1
        candidateName = baseName & "_" & index.ToString()
    End While
    
    Return candidateName
End Function
