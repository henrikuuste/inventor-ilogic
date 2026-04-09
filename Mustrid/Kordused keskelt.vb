' ============================================================================
' Kordused Keskelt - Center-Based Occurrence Pattern (Extended)
' 
' Creates a parametric pattern that distributes instances evenly across a span.
' The span is defined by two boundary planes/faces/points/vertices.
'
' Features:
' - Uniform or symmetric-from-center distribution
' - Include/exclude instances at span boundaries
' - Start/end offset to reduce effective span
' - Start/end alignment options (center, inward, outward)
' - Zero instance handling when span is too small
' - Support for non-planar boundaries (points, vertices, edges)
' - Non-parallel plane handling with explicit axis selection
' - Update mode: re-run on existing pattern to update parameters
'
' Usage:
' 1. Select an occurrence to pattern (or have it pre-selected)
' 2. Pick start and end boundaries (planes, faces, points, vertices)
' 3. If boundaries are non-parallel, select an axis
' 4. Enter max spacing (mm) or select existing parameter
' 5. Configure offsets, alignment, and distribution options
' 6. Click OK to create/update the pattern
'
' ESC cancels any selection operation.
'
' ============================================================================

AddVbFile "Lib/UtilsLib.vb"
AddVbFile "Lib/GeoLib.vb"
AddVbFile "Lib/WorkFeatureLib.vb"
AddVbFile "Lib/PatternLib.vb"
AddVbFile "Lib/DocumentUpdateLib.vb"
AddVbFile "Lib/CenterPatternLib.vb"

Sub Main()
    Dim app As Inventor.Application = ThisApplication
    Dim doc As Document = app.ActiveDocument
    Dim iLogicAuto As Object = iLogicVb.Automation
    
    ' Validate document
    If doc Is Nothing Then
        Logger.Error("Kordused keskelt: No active document")
        MessageBox.Show("Aktiivne dokument puudub.", "Kordused keskelt")
        Exit Sub
    End If
    
    If doc.DocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
        Logger.Error("Kordused keskelt: Not an assembly document")
        MessageBox.Show("See reegel töötab ainult koostudokumentides (.iam).", "Kordused keskelt")
        Exit Sub
    End If
    
    Dim asmDoc As AssemblyDocument = CType(doc, AssemblyDocument)
    
    ' Run the main picker loop
    RunPatternSetup(app, asmDoc, iLogicAuto)
End Sub

' ============================================================================
' Main Setup Loop
' ============================================================================

Sub RunPatternSetup(app As Inventor.Application, asmDoc As AssemblyDocument, iLogicAuto As Object)
    ' State variables - basic
    Dim seedOcc As ComponentOccurrence = Nothing
    Dim baseName As String = ""
    Dim startGeometry As Object = Nothing
    Dim endGeometry As Object = Nothing
    Dim explicitAxis As Object = Nothing
    Dim maxSpacingInput As String = "500"
    Dim mode As String = CenterPatternLib.MODE_UNIFORM
    Dim includeEnds As Boolean = False
    
    ' State variables - extended
    Dim startOffsetMm As Double = 0
    Dim endOffsetMm As Double = 0
    Dim startAlignment As String = CenterPatternLib.ALIGN_CENTER
    Dim endAlignment As String = CenterPatternLib.ALIGN_CENTER
    Dim allowZeroInstances As Boolean = False
    Dim isUpdateMode As Boolean = False
    
    ' Check for pre-selected occurrence or pattern
    If asmDoc.SelectSet.Count = 1 Then
        Dim selectedObj As Object = asmDoc.SelectSet.Item(1)
        
        If TypeOf selectedObj Is ComponentOccurrence Then
            seedOcc = CType(selectedObj, ComponentOccurrence)
            baseName = ExtractBaseName(seedOcc.Name)
            
            ' Check if this is an existing pattern (update mode)
            Dim existingSeed As ComponentOccurrence = CenterPatternLib.DetectExistingPattern(asmDoc, seedOcc)
            If existingSeed IsNot Nothing Then
                isUpdateMode = True
                seedOcc = existingSeed
                ' Load existing configuration
                LoadExistingConfig(seedOcc, baseName, maxSpacingInput, mode, includeEnds, _
                                   startOffsetMm, endOffsetMm, startAlignment, endAlignment, allowZeroInstances)
                ' Load existing work features as geometry references
                LoadExistingGeometry(asmDoc, seedOcc, startGeometry, endGeometry, explicitAxis)
            End If
            
        ElseIf TypeOf selectedObj Is RectangularOccurrencePattern Then
            ' User selected the pattern itself - find the seed occurrence
            Dim pattern As RectangularOccurrencePattern = CType(selectedObj, RectangularOccurrencePattern)
            Dim patternSeed As ComponentOccurrence = FindPatternSeedOccurrence(asmDoc, pattern)
            
            If patternSeed IsNot Nothing AndAlso CenterPatternLib.HasPatternConfig(patternSeed) Then
                isUpdateMode = True
                seedOcc = patternSeed
                baseName = ExtractBaseName(seedOcc.Name)
                ' Load existing configuration
                LoadExistingConfig(seedOcc, baseName, maxSpacingInput, mode, includeEnds, _
                                   startOffsetMm, endOffsetMm, startAlignment, endAlignment, allowZeroInstances)
                ' Load existing work features as geometry references
                LoadExistingGeometry(asmDoc, seedOcc, startGeometry, endGeometry, explicitAxis)
            End If
        End If
    End If
    
    ' Main dialog loop
    Dim keepGoing As Boolean = True
    Do While keepGoing
        Dim action As String = ""
        Dim result As DialogResult = ShowSetupForm(app, asmDoc, _
            seedOcc, baseName, startGeometry, endGeometry, explicitAxis, _
            maxSpacingInput, mode, includeEnds, _
            startOffsetMm, endOffsetMm, startAlignment, endAlignment, _
            allowZeroInstances, isUpdateMode, action)
        
        Select Case result
            Case DialogResult.Cancel
                Exit Do
                
            Case DialogResult.Abort
                ' Handle delete action
                If action = "DELETE" AndAlso seedOcc IsNot Nothing Then
                    DeletePattern(app, asmDoc, iLogicAuto, seedOcc)
                    Exit Do
                End If
                
            Case DialogResult.OK
                ' Create or update the pattern
                If isUpdateMode Then
                    ' In update mode, rebuild the pattern with new settings
                    If ValidateInputs(seedOcc, startGeometry, endGeometry, maxSpacingInput) Then
                        RebuildPattern(app, asmDoc, iLogicAuto, seedOcc, baseName, _
                                       startGeometry, endGeometry, explicitAxis, maxSpacingInput, _
                                       mode, includeEnds, startOffsetMm, endOffsetMm, _
                                       startAlignment, endAlignment, allowZeroInstances)
                        Exit Do
                    End If
                ElseIf ValidateInputs(seedOcc, startGeometry, endGeometry, maxSpacingInput) Then
                    CreatePatternEx(app, asmDoc, iLogicAuto, seedOcc, baseName, _
                                    startGeometry, endGeometry, explicitAxis, maxSpacingInput, _
                                    mode, includeEnds, startOffsetMm, endOffsetMm, _
                                    startAlignment, endAlignment, allowZeroInstances)
                    Exit Do
                End If
                
            Case DialogResult.Retry
                ' Handle picker actions
                Select Case action
                    Case "PICK_OCC"
                        Dim picked As ComponentOccurrence = PickOccurrence(app)
                        If picked IsNot Nothing Then
                            seedOcc = picked
                            baseName = ExtractBaseName(seedOcc.Name)
                            
                            ' Check for existing pattern
                            Dim existingSeed As ComponentOccurrence = CenterPatternLib.DetectExistingPattern(asmDoc, picked)
                            If existingSeed IsNot Nothing Then
                                isUpdateMode = True
                                seedOcc = existingSeed
                                LoadExistingConfig(seedOcc, baseName, maxSpacingInput, mode, includeEnds, _
                                                   startOffsetMm, endOffsetMm, startAlignment, endAlignment, allowZeroInstances)
                                ' Load existing work features as geometry references
                                LoadExistingGeometry(asmDoc, seedOcc, startGeometry, endGeometry, explicitAxis)
                            Else
                                isUpdateMode = False
                                ' Clear geometry when switching to new occurrence
                                startGeometry = Nothing
                                endGeometry = Nothing
                                explicitAxis = Nothing
                            End If
                        End If
                        
                    Case "PICK_START"
                        Dim picked As Object = PickBoundaryGeometry(app, "Vali alguspunkt/pind - ESC tühistamiseks")
                        If picked IsNot Nothing Then
                            startGeometry = picked
                        End If
                        
                    Case "PICK_END"
                        Dim picked As Object = PickBoundaryGeometry(app, "Vali lõpupunkt/pind - ESC tühistamiseks")
                        If picked IsNot Nothing Then
                            endGeometry = picked
                            ' Check if axis is needed
                            If NeedsExplicitAxis(startGeometry, endGeometry) Then
                                MessageBox.Show("Valitud pinnad ei ole paralleelsed. Palun vali mustri telg.", "Kordused keskelt")
                            End If
                        End If
                        
                    Case "PICK_AXIS"
                        Dim picked As Object = PickAxis(app, "Vali mustri telg (serv või telg) - ESC tühistamiseks")
                        If picked IsNot Nothing Then
                            explicitAxis = picked
                        End If
                End Select
        End Select
    Loop
End Sub

' ============================================================================
' Setup Form
' ============================================================================

Function ShowSetupForm(app As Inventor.Application, asmDoc As AssemblyDocument, _
                       ByRef seedOcc As ComponentOccurrence, _
                       ByRef baseName As String, _
                       ByRef startGeometry As Object, _
                       ByRef endGeometry As Object, _
                       ByRef explicitAxis As Object, _
                       ByRef maxSpacingInput As String, _
                       ByRef mode As String, _
                       ByRef includeEnds As Boolean, _
                       ByRef startOffsetMm As Double, _
                       ByRef endOffsetMm As Double, _
                       ByRef startAlignment As String, _
                       ByRef endAlignment As String, _
                       ByRef allowZeroInstances As Boolean, _
                       ByRef isUpdateMode As Boolean, _
                       ByRef action As String) As DialogResult
    
    action = ""
    
    ' Create form
    Dim frm As New System.Windows.Forms.Form()
    frm.Text = If(isUpdateMode, "Uuenda mustrit", "Kordused Keskelt")
    frm.Width = 480
    frm.Height = 580
    frm.StartPosition = FormStartPosition.CenterScreen
    frm.FormBorderStyle = FormBorderStyle.FixedDialog
    frm.MaximizeBox = False
    frm.MinimizeBox = False
    
    Dim yPos As Integer = 15
    Dim labelWidth As Integer = 120
    Dim controlLeft As Integer = 135
    Dim controlWidth As Integer = 200
    Dim btnWidth As Integer = 70
    Dim rowHeight As Integer = 30
    
    ' --- Update mode indicator ---
    If isUpdateMode Then
        Dim lblUpdate As New System.Windows.Forms.Label()
        lblUpdate.Text = "*** Olemasoleva mustri uuendamine ***"
        lblUpdate.Left = 15
        lblUpdate.Top = yPos
        lblUpdate.Width = 400
        frm.Controls.Add(lblUpdate)
        yPos += rowHeight
    End If
    
    ' --- Occurrence selection ---
    Dim lblOcc As New System.Windows.Forms.Label()
    lblOcc.Text = "Element:"
    lblOcc.Left = 15
    lblOcc.Top = yPos + 3
    lblOcc.Width = labelWidth
    frm.Controls.Add(lblOcc)
    
    Dim txtOcc As New System.Windows.Forms.TextBox()
    txtOcc.Name = "txtOcc"
    txtOcc.Left = controlLeft
    txtOcc.Top = yPos
    txtOcc.Width = controlWidth
    txtOcc.ReadOnly = True
    txtOcc.Text = If(baseName <> "", baseName, "(vali element)")
    ' In update mode, occurrence is fixed (can't change which part to pattern)
    txtOcc.Enabled = Not isUpdateMode
    frm.Controls.Add(txtOcc)
    
    Dim btnPickOcc As New System.Windows.Forms.Button()
    btnPickOcc.Text = "Vali..."
    btnPickOcc.Left = controlLeft + controlWidth + 10
    btnPickOcc.Top = yPos - 2
    btnPickOcc.Width = btnWidth
    btnPickOcc.Height = 26
    btnPickOcc.Enabled = Not isUpdateMode
    frm.Controls.Add(btnPickOcc)
    
    yPos += rowHeight + 5
    
    ' --- Start boundary ---
    Dim lblStart As New System.Windows.Forms.Label()
    lblStart.Text = "Alguspind:"
    lblStart.Left = 15
    lblStart.Top = yPos + 3
    lblStart.Width = labelWidth
    frm.Controls.Add(lblStart)
    
    Dim txtStart As New System.Windows.Forms.TextBox()
    txtStart.Name = "txtStart"
    txtStart.Left = controlLeft
    txtStart.Top = yPos
    txtStart.Width = controlWidth
    txtStart.ReadOnly = True
    txtStart.Text = If(startGeometry IsNot Nothing, UtilsLib.GetObjectDisplayName(startGeometry), "(vali pind/punkt)")
    ' Enable in update mode - user can change boundaries
    frm.Controls.Add(txtStart)
    
    Dim btnPickStart As New System.Windows.Forms.Button()
    btnPickStart.Text = "Vali..."
    btnPickStart.Left = controlLeft + controlWidth + 10
    btnPickStart.Top = yPos - 2
    btnPickStart.Width = btnWidth
    btnPickStart.Height = 26
    ' Enable in update mode - user can change boundaries
    frm.Controls.Add(btnPickStart)
    
    yPos += rowHeight + 5
    
    ' --- End boundary ---
    Dim lblEnd As New System.Windows.Forms.Label()
    lblEnd.Text = "Lõpupind:"
    lblEnd.Left = 15
    lblEnd.Top = yPos + 3
    lblEnd.Width = labelWidth
    frm.Controls.Add(lblEnd)
    
    Dim txtEnd As New System.Windows.Forms.TextBox()
    txtEnd.Name = "txtEnd"
    txtEnd.Left = controlLeft
    txtEnd.Top = yPos
    txtEnd.Width = controlWidth
    txtEnd.ReadOnly = True
    txtEnd.Text = If(endGeometry IsNot Nothing, UtilsLib.GetObjectDisplayName(endGeometry), "(vali pind/punkt)")
    ' Enable in update mode - user can change boundaries
    frm.Controls.Add(txtEnd)
    
    Dim btnPickEnd As New System.Windows.Forms.Button()
    btnPickEnd.Text = "Vali..."
    btnPickEnd.Left = controlLeft + controlWidth + 10
    btnPickEnd.Top = yPos - 2
    btnPickEnd.Width = btnWidth
    btnPickEnd.Height = 26
    ' Enable in update mode - user can change boundaries
    frm.Controls.Add(btnPickEnd)
    
    yPos += rowHeight + 5
    
    ' --- Explicit Axis (for non-parallel boundaries) ---
    Dim lblAxis As New System.Windows.Forms.Label()
    lblAxis.Text = "Telg (valikuline):"
    lblAxis.Left = 15
    lblAxis.Top = yPos + 3
    lblAxis.Width = labelWidth
    frm.Controls.Add(lblAxis)
    
    Dim txtAxis As New System.Windows.Forms.TextBox()
    txtAxis.Name = "txtAxis"
    txtAxis.Left = controlLeft
    txtAxis.Top = yPos
    txtAxis.Width = controlWidth
    txtAxis.ReadOnly = True
    txtAxis.Text = If(explicitAxis IsNot Nothing, UtilsLib.GetObjectDisplayName(explicitAxis), "(automaatne)")
    ' Enable in update mode - user can change axis
    frm.Controls.Add(txtAxis)
    
    Dim btnPickAxis As New System.Windows.Forms.Button()
    btnPickAxis.Text = "Vali..."
    btnPickAxis.Left = controlLeft + controlWidth + 10
    btnPickAxis.Top = yPos - 2
    btnPickAxis.Width = btnWidth
    btnPickAxis.Height = 26
    ' Enable in update mode - user can change axis
    frm.Controls.Add(btnPickAxis)
    
    yPos += rowHeight + 10
    
    ' --- Max spacing ---
    Dim lblSpacing As New System.Windows.Forms.Label()
    lblSpacing.Text = "Max vahe (mm):"
    lblSpacing.Left = 15
    lblSpacing.Top = yPos + 3
    lblSpacing.Width = labelWidth
    frm.Controls.Add(lblSpacing)
    
    Dim txtSpacing As New System.Windows.Forms.TextBox()
    txtSpacing.Name = "txtSpacing"
    txtSpacing.Left = controlLeft
    txtSpacing.Top = yPos
    txtSpacing.Width = 80
    txtSpacing.Text = maxSpacingInput
    frm.Controls.Add(txtSpacing)
    
    ' Parameter dropdown
    Dim cboParams As New System.Windows.Forms.ComboBox()
    cboParams.Name = "cboParams"
    cboParams.Left = controlLeft + 90
    cboParams.Top = yPos
    cboParams.Width = 150
    cboParams.DropDownStyle = ComboBoxStyle.DropDownList
    cboParams.Items.Add("(parameeter)")
    
    Dim paramNames As String() = CenterPatternLib.GetUserParameterNames(asmDoc)
    For Each pName As String In paramNames
        cboParams.Items.Add(pName)
    Next
    cboParams.SelectedIndex = 0
    frm.Controls.Add(cboParams)
    
    yPos += rowHeight + 10
    
    ' --- Start offset ---
    Dim lblStartOff As New System.Windows.Forms.Label()
    lblStartOff.Text = "Alguse nihe (mm):"
    lblStartOff.Left = 15
    lblStartOff.Top = yPos + 3
    lblStartOff.Width = labelWidth
    frm.Controls.Add(lblStartOff)
    
    Dim txtStartOff As New System.Windows.Forms.TextBox()
    txtStartOff.Name = "txtStartOff"
    txtStartOff.Left = controlLeft
    txtStartOff.Top = yPos
    txtStartOff.Width = 60
    txtStartOff.Text = startOffsetMm.ToString("0")
    frm.Controls.Add(txtStartOff)
    
    ' Start alignment combo
    Dim lblStartAlign As New System.Windows.Forms.Label()
    lblStartAlign.Text = "Joondus:"
    lblStartAlign.Left = controlLeft + 70
    lblStartAlign.Top = yPos + 3
    lblStartAlign.Width = 50
    frm.Controls.Add(lblStartAlign)
    
    Dim cboStartAlign As New System.Windows.Forms.ComboBox()
    cboStartAlign.Name = "cboStartAlign"
    cboStartAlign.Left = controlLeft + 125
    cboStartAlign.Top = yPos
    cboStartAlign.Width = 100
    cboStartAlign.DropDownStyle = ComboBoxStyle.DropDownList
    cboStartAlign.Items.Add("Keskel")
    cboStartAlign.Items.Add("Sissepoole")
    cboStartAlign.Items.Add("Väljapoole")
    cboStartAlign.SelectedIndex = GetAlignmentIndex(startAlignment)
    frm.Controls.Add(cboStartAlign)
    
    yPos += rowHeight + 5
    
    ' --- End offset ---
    Dim lblEndOff As New System.Windows.Forms.Label()
    lblEndOff.Text = "Lõpu nihe (mm):"
    lblEndOff.Left = 15
    lblEndOff.Top = yPos + 3
    lblEndOff.Width = labelWidth
    frm.Controls.Add(lblEndOff)
    
    Dim txtEndOff As New System.Windows.Forms.TextBox()
    txtEndOff.Name = "txtEndOff"
    txtEndOff.Left = controlLeft
    txtEndOff.Top = yPos
    txtEndOff.Width = 60
    txtEndOff.Text = endOffsetMm.ToString("0")
    frm.Controls.Add(txtEndOff)
    
    ' End alignment combo
    Dim lblEndAlign As New System.Windows.Forms.Label()
    lblEndAlign.Text = "Joondus:"
    lblEndAlign.Left = controlLeft + 70
    lblEndAlign.Top = yPos + 3
    lblEndAlign.Width = 50
    frm.Controls.Add(lblEndAlign)
    
    Dim cboEndAlign As New System.Windows.Forms.ComboBox()
    cboEndAlign.Name = "cboEndAlign"
    cboEndAlign.Left = controlLeft + 125
    cboEndAlign.Top = yPos
    cboEndAlign.Width = 100
    cboEndAlign.DropDownStyle = ComboBoxStyle.DropDownList
    cboEndAlign.Items.Add("Keskel")
    cboEndAlign.Items.Add("Sissepoole")
    cboEndAlign.Items.Add("Väljapoole")
    cboEndAlign.SelectedIndex = GetAlignmentIndex(endAlignment)
    frm.Controls.Add(cboEndAlign)
    
    yPos += rowHeight + 10
    
    ' --- Distribution mode ---
    Dim lblMode As New System.Windows.Forms.Label()
    lblMode.Text = "Jaotus:"
    lblMode.Left = 15
    lblMode.Top = yPos + 3
    lblMode.Width = labelWidth
    frm.Controls.Add(lblMode)
    
    Dim cboMode As New System.Windows.Forms.ComboBox()
    cboMode.Name = "cboMode"
    cboMode.Left = controlLeft
    cboMode.Top = yPos
    cboMode.Width = 150
    cboMode.DropDownStyle = ComboBoxStyle.DropDownList
    cboMode.Items.Add("Ühtlane")
    cboMode.Items.Add("Sümmeetriline keskelt")
    cboMode.SelectedIndex = If(mode = CenterPatternLib.MODE_SYMMETRIC, 1, 0)
    frm.Controls.Add(cboMode)
    
    yPos += rowHeight + 5
    
    ' --- Include ends checkbox ---
    Dim chkEnds As New System.Windows.Forms.CheckBox()
    chkEnds.Name = "chkEnds"
    chkEnds.Text = "Elemendid otstes"
    chkEnds.Left = controlLeft
    chkEnds.Top = yPos
    chkEnds.Width = 130
    chkEnds.Checked = includeEnds
    frm.Controls.Add(chkEnds)
    
    ' --- Zero instance checkbox ---
    Dim chkZero As New System.Windows.Forms.CheckBox()
    chkZero.Name = "chkZero"
    chkZero.Text = "0 kordust, kui ulatus < vahe"
    chkZero.Left = controlLeft + 140
    chkZero.Top = yPos
    chkZero.Width = 190
    chkZero.Checked = allowZeroInstances
    frm.Controls.Add(chkZero)
    
    yPos += rowHeight + 20
    
    ' --- Buttons ---
    Dim btnOK As New System.Windows.Forms.Button()
    btnOK.Text = If(isUpdateMode, "Uuenda", "Loo muster")
    btnOK.Left = 120
    btnOK.Top = yPos
    btnOK.Width = 100
    btnOK.Height = 32
    btnOK.DialogResult = DialogResult.OK
    frm.AcceptButton = btnOK
    frm.Controls.Add(btnOK)
    
    Dim btnDelete As New System.Windows.Forms.Button()
    btnDelete.Text = "Kustuta"
    btnDelete.Left = 230
    btnDelete.Top = yPos
    btnDelete.Width = 80
    btnDelete.Height = 32
    btnDelete.Enabled = isUpdateMode ' Only enabled for existing patterns
    frm.Controls.Add(btnDelete)
    
    Dim btnCancel As New System.Windows.Forms.Button()
    btnCancel.Text = "Tühista"
    btnCancel.Left = 320
    btnCancel.Top = yPos
    btnCancel.Width = 80
    btnCancel.Height = 32
    btnCancel.DialogResult = DialogResult.Cancel
    frm.CancelButton = btnCancel
    frm.Controls.Add(btnCancel)
    
    ' --- Event handlers ---
    frm.Tag = ""
    
    AddHandler btnPickOcc.Click, Sub(s, e)
        frm.Tag = "PICK_OCC"
        frm.DialogResult = DialogResult.Retry
    End Sub
    
    AddHandler btnPickStart.Click, Sub(s, e)
        frm.Tag = "PICK_START"
        frm.DialogResult = DialogResult.Retry
    End Sub
    
    AddHandler btnPickEnd.Click, Sub(s, e)
        frm.Tag = "PICK_END"
        frm.DialogResult = DialogResult.Retry
    End Sub
    
    AddHandler btnPickAxis.Click, Sub(s, e)
        frm.Tag = "PICK_AXIS"
        frm.DialogResult = DialogResult.Retry
    End Sub
    
    AddHandler cboParams.SelectedIndexChanged, Sub(s, e)
        If cboParams.SelectedIndex > 0 Then
            txtSpacing.Text = cboParams.SelectedItem.ToString()
        End If
    End Sub
    
    AddHandler btnDelete.Click, Sub(s, e)
        ' Confirm deletion
        If MessageBox.Show("Kas oled kindel, et soovid mustri kustutada?" & vbCrLf & _
                           "See taastab esialgse oleku.", _
                           "Kustuta muster", _
                           MessageBoxButtons.YesNo, _
                           MessageBoxIcon.Warning) = DialogResult.Yes Then
            frm.Tag = "DELETE"
            frm.DialogResult = DialogResult.Abort
        End If
    End Sub
    
    ' Show form
    Dim result As DialogResult = frm.ShowDialog()
    
    ' Read values back
    action = CStr(frm.Tag)
    maxSpacingInput = txtSpacing.Text.Trim()
    mode = If(cboMode.SelectedIndex = 1, CenterPatternLib.MODE_SYMMETRIC, CenterPatternLib.MODE_UNIFORM)
    includeEnds = chkEnds.Checked
    allowZeroInstances = chkZero.Checked
    
    ' Parse offsets
    Double.TryParse(txtStartOff.Text, startOffsetMm)
    Double.TryParse(txtEndOff.Text, endOffsetMm)
    
    ' Get alignment from combos
    startAlignment = GetAlignmentFromIndex(cboStartAlign.SelectedIndex)
    endAlignment = GetAlignmentFromIndex(cboEndAlign.SelectedIndex)
    
    Return result
End Function

' ============================================================================
' Pickers
' ============================================================================

Function PickOccurrence(app As Inventor.Application) As ComponentOccurrence
    Try
        Dim picked As Object = app.CommandManager.Pick( _
            SelectionFilterEnum.kAssemblyOccurrenceFilter, _
            "Vali element mida korrata - ESC tühistamiseks")
        
        If TypeOf picked Is ComponentOccurrence Then
            Return CType(picked, ComponentOccurrence)
        End If
    Catch
    End Try
    Return Nothing
End Function

Function PickPlanarGeometry(app As Inventor.Application, prompt As String) As Object
    Try
        ' Use kAllPlanarEntities to allow both work planes and planar faces
        Dim picked As Object = app.CommandManager.Pick( _
            SelectionFilterEnum.kAllPlanarEntities, prompt)
        Return picked
    Catch
    End Try
    Return Nothing
End Function

Function PickBoundaryGeometry(app As Inventor.Application, prompt As String) As Object
    ' Try planar entities first
    Try
        Dim picked As Object = app.CommandManager.Pick( _
            SelectionFilterEnum.kAllPlanarEntities, prompt)
        If picked IsNot Nothing Then Return picked
    Catch
    End Try
    
    ' Try points/vertices
    Try
        Dim picked As Object = app.CommandManager.Pick( _
            SelectionFilterEnum.kAllPointEntities, prompt)
        If picked IsNot Nothing Then Return picked
    Catch
    End Try
    
    Return Nothing
End Function

Function PickAxis(app As Inventor.Application, prompt As String) As Object
    Try
        ' Try linear entities (edges, axes)
        Dim picked As Object = app.CommandManager.Pick( _
            SelectionFilterEnum.kAllLinearEntities, prompt)
        Return picked
    Catch
    End Try
    Return Nothing
End Function

Function NeedsExplicitAxis(startGeometry As Object, endGeometry As Object) As Boolean
    ' Check if both geometries are planar and parallel
    If startGeometry Is Nothing OrElse endGeometry Is Nothing Then Return False
    
    ' Check geometry types
    Dim startType As Integer = WorkFeatureLib.GetGeometryType(startGeometry)
    Dim endType As Integer = WorkFeatureLib.GetGeometryType(endGeometry)
    
    ' If both are planar, check if parallel
    If startType = WorkFeatureLib.GEOM_TYPE_PLANE AndAlso endType = WorkFeatureLib.GEOM_TYPE_PLANE Then
        Dim n1 As UnitVector = GeoLib.GetPlaneNormal(startGeometry)
        Dim n2 As UnitVector = GeoLib.GetPlaneNormal(endGeometry)
        
        If n1 IsNot Nothing AndAlso n2 IsNot Nothing Then
            Return Not GeoLib.AreVectorsParallel(n1, n2, 0.95)
        End If
    End If
    
    ' If mixed geometry types (planar + point), axis can be determined automatically
    If (startType = WorkFeatureLib.GEOM_TYPE_PLANE AndAlso endType <> WorkFeatureLib.GEOM_TYPE_PLANE) OrElse _
       (startType <> WorkFeatureLib.GEOM_TYPE_PLANE AndAlso endType = WorkFeatureLib.GEOM_TYPE_PLANE) Then
        Return False
    End If
    
    ' Two points - axis can be determined automatically
    If (startType = WorkFeatureLib.GEOM_TYPE_POINT OrElse startType = WorkFeatureLib.GEOM_TYPE_VERTEX) AndAlso _
       (endType = WorkFeatureLib.GEOM_TYPE_POINT OrElse endType = WorkFeatureLib.GEOM_TYPE_VERTEX) Then
        Return False
    End If
    
    Return False
End Function

' ============================================================================
' Alignment Helpers
' ============================================================================

Function GetAlignmentIndex(alignment As String) As Integer
    Select Case alignment
        Case CenterPatternLib.ALIGN_CENTER
            Return 0
        Case CenterPatternLib.ALIGN_INWARD
            Return 1
        Case CenterPatternLib.ALIGN_OUTWARD
            Return 2
        Case Else
            Return 0
    End Select
End Function

Function GetAlignmentFromIndex(index As Integer) As String
    Select Case index
        Case 0
            Return CenterPatternLib.ALIGN_CENTER
        Case 1
            Return CenterPatternLib.ALIGN_INWARD
        Case 2
            Return CenterPatternLib.ALIGN_OUTWARD
        Case Else
            Return CenterPatternLib.ALIGN_CENTER
    End Select
End Function

' ============================================================================
' Configuration Loading
' ============================================================================

Sub LoadExistingConfig(occ As ComponentOccurrence, _
                       ByRef baseName As String, _
                       ByRef maxSpacingInput As String, _
                       ByRef mode As String, _
                       ByRef includeEnds As Boolean, _
                       ByRef startOffsetMm As Double, _
                       ByRef endOffsetMm As Double, _
                       ByRef startAlignment As String, _
                       ByRef endAlignment As String, _
                       ByRef allowZeroInstances As Boolean)
    
    Dim startPlane As String = ""
    Dim endPlane As String = ""
    Dim axisName As String = ""
    Dim maxSpacing As String = ""
    Dim startOffset As String = ""
    Dim endOffset As String = ""
    
    If CenterPatternLib.LoadPatternConfigEx(occ, baseName, startPlane, endPlane, axisName, _
                                             maxSpacing, mode, includeEnds, startOffset, endOffset, _
                                             startAlignment, endAlignment, allowZeroInstances) Then
        maxSpacingInput = maxSpacing
        Double.TryParse(startOffset, startOffsetMm)
        Double.TryParse(endOffset, endOffsetMm)
    ElseIf CenterPatternLib.LoadPatternConfig(occ, baseName, startPlane, endPlane, maxSpacing, mode, includeEnds) Then
        maxSpacingInput = maxSpacing
    End If
End Sub

Sub LoadExistingGeometry(asmDoc As AssemblyDocument, _
                         occ As ComponentOccurrence, _
                         ByRef startGeometry As Object, _
                         ByRef endGeometry As Object, _
                         ByRef explicitAxis As Object)
    
    ' Load plane/axis names from config
    Dim baseName As String = ""
    Dim startPlaneName As String = ""
    Dim endPlaneName As String = ""
    Dim axisName As String = ""
    Dim maxSpacing As String = ""
    Dim mode As String = ""
    Dim includeEnds As Boolean = False
    Dim startOffset As String = ""
    Dim endOffset As String = ""
    Dim startAlign As String = ""
    Dim endAlign As String = ""
    Dim allowZero As Boolean = False
    
    If Not CenterPatternLib.LoadPatternConfigEx(occ, baseName, startPlaneName, endPlaneName, axisName, _
                                                 maxSpacing, mode, includeEnds, startOffset, endOffset, _
                                                 startAlign, endAlign, allowZero) Then
        Exit Sub
    End If
    
    Dim asmDef As AssemblyComponentDefinition = asmDoc.ComponentDefinition
    
    ' Find existing work planes to use as geometry references
    If Not String.IsNullOrEmpty(startPlaneName) Then
        startGeometry = WorkFeatureLib.FindWorkPlaneByName(asmDef, startPlaneName)
    End If
    
    If Not String.IsNullOrEmpty(endPlaneName) Then
        endGeometry = WorkFeatureLib.FindWorkPlaneByName(asmDef, endPlaneName)
    End If
    
    If Not String.IsNullOrEmpty(axisName) Then
        explicitAxis = WorkFeatureLib.FindWorkAxisByName(asmDef, axisName)
    End If
End Sub

' ============================================================================
' Validation
' ============================================================================

Function ValidateInputs(seedOcc As ComponentOccurrence, _
                        startGeometry As Object, _
                        endGeometry As Object, _
                        maxSpacingInput As String) As Boolean
    
    If seedOcc Is Nothing Then
        MessageBox.Show("Palun vali element mida korrata.", "Kordused keskelt")
        Return False
    End If
    
    If startGeometry Is Nothing Then
        MessageBox.Show("Palun vali alguspind.", "Kordused keskelt")
        Return False
    End If
    
    If endGeometry Is Nothing Then
        MessageBox.Show("Palun vali lõpupind.", "Kordused keskelt")
        Return False
    End If
    
    If maxSpacingInput = "" Then
        MessageBox.Show("Palun sisesta max vahe väärtus.", "Kordused keskelt")
        Return False
    End If
    
    ' Try to parse as number or check if it's a parameter name
    Dim numValue As Double
    If Not Double.TryParse(maxSpacingInput, numValue) Then
        ' Not a number - check if it's a valid parameter name
        ' We'll validate this later in CreatePattern
    ElseIf numValue <= 0 Then
        MessageBox.Show("Max vahe peab olema positiivne arv.", "Kordused keskelt")
        Return False
    End If
    
    Return True
End Function

' ============================================================================
' Pattern Creation
' ============================================================================

Sub CreatePattern(app As Inventor.Application, _
                  asmDoc As AssemblyDocument, _
                  iLogicAuto As Object, _
                  seedOcc As ComponentOccurrence, _
                  baseName As String, _
                  startGeometry As Object, _
                  endGeometry As Object, _
                  maxSpacingInput As String, _
                  mode As String, _
                  includeEnds As Boolean)
    
    ' Resolve max spacing value
    Dim maxSpacingMm As Double = 0
    Dim numValue As Double
    
    If Double.TryParse(maxSpacingInput, numValue) Then
        maxSpacingMm = numValue
    Else
        ' It's a parameter name - get value
        Dim paramValueCm As Double = CenterPatternLib.GetParameterValue(asmDoc, maxSpacingInput)
        If paramValueCm > 0 Then
            maxSpacingMm = paramValueCm * 10.0
        Else
            MessageBox.Show("Parameetrit '" & maxSpacingInput & "' ei leitud või on väärtus 0.", "Kordused keskelt")
            Exit Sub
        End If
    End If
    
    ' Create pattern using library
    Dim logs As New System.Collections.Generic.List(Of String)
    Dim success As Boolean = CenterPatternLib.CreateCenterPattern( _
        app, asmDoc, iLogicAuto, seedOcc, _
        startGeometry, endGeometry, _
        maxSpacingMm, mode, includeEnds, baseName, logs)
    
    ' Output logs
    For Each logMsg As String In logs
        Logger.Info(logMsg)
    Next
    
    If success Then
        Logger.Info("Kordused keskelt: Muster '" & baseName & "' loodud edukalt")
    Else
        Logger.Error("Kordused keskelt: Mustri loomine ebaõnnestus")
        MessageBox.Show("Mustri loomine ebaõnnestus. Vaata logi akent.", "Kordused keskelt")
    End If
End Sub

Sub CreatePatternEx(app As Inventor.Application, _
                    asmDoc As AssemblyDocument, _
                    iLogicAuto As Object, _
                    seedOcc As ComponentOccurrence, _
                    baseName As String, _
                    startGeometry As Object, _
                    endGeometry As Object, _
                    explicitAxis As Object, _
                    maxSpacingInput As String, _
                    mode As String, _
                    includeEnds As Boolean, _
                    startOffsetMm As Double, _
                    endOffsetMm As Double, _
                    startAlignment As String, _
                    endAlignment As String, _
                    allowZeroInstances As Boolean)
    
    ' Create extended pattern using library
    ' maxSpacingInput can be a number, parameter name, or formula - library handles parsing
    Dim logs As New System.Collections.Generic.List(Of String)
    Dim success As Boolean = CenterPatternLib.CreateCenterPatternEx( _
        app, asmDoc, iLogicAuto, seedOcc, _
        startGeometry, endGeometry, explicitAxis, _
        maxSpacingInput, mode, includeEnds, baseName, _
        startOffsetMm, endOffsetMm, _
        startAlignment, endAlignment, _
        allowZeroInstances, logs)
    
    ' Output logs
    For Each logMsg As String In logs
        Logger.Info(logMsg)
    Next
    
    If success Then
        Logger.Info("Kordused keskelt: Muster '" & baseName & "' loodud edukalt")
    Else
        Logger.Error("Kordused keskelt: Mustri loomine ebaõnnestus")
        MessageBox.Show("Mustri loomine ebaõnnestus. Vaata logi akent.", "Kordused keskelt")
    End If
End Sub

Sub RebuildPattern(app As Inventor.Application, _
                   asmDoc As AssemblyDocument, _
                   iLogicAuto As Object, _
                   seedOcc As ComponentOccurrence, _
                   baseName As String, _
                   startGeometry As Object, _
                   endGeometry As Object, _
                   explicitAxis As Object, _
                   maxSpacingInput As String, _
                   mode As String, _
                   includeEnds As Boolean, _
                   startOffsetMm As Double, _
                   endOffsetMm As Double, _
                   startAlignment As String, _
                   endAlignment As String, _
                   allowZeroInstances As Boolean)
    
    Dim logs As New System.Collections.Generic.List(Of String)
    Dim success As Boolean = CenterPatternLib.RebuildCenterPattern( _
        app, asmDoc, iLogicAuto, seedOcc, _
        startGeometry, endGeometry, explicitAxis, _
        maxSpacingInput, mode, includeEnds, baseName, _
        startOffsetMm, endOffsetMm, _
        startAlignment, endAlignment, _
        allowZeroInstances, logs)
    
    ' Output logs
    For Each logMsg As String In logs
        Logger.Info(logMsg)
    Next
    
    If success Then
        Logger.Info("Kordused keskelt: Muster '" & baseName & "' uuendatud edukalt")
    Else
        Logger.Error("Kordused keskelt: Mustri uuendamine ebaõnnestus")
        MessageBox.Show("Mustri uuendamine ebaõnnestus. Vaata logi akent.", "Kordused keskelt")
    End If
End Sub

Sub DeletePattern(app As Inventor.Application, asmDoc As AssemblyDocument, _
                  iLogicAuto As Object, seedOcc As ComponentOccurrence)
    Dim logs As New System.Collections.Generic.List(Of String)
    Dim success As Boolean = CenterPatternLib.DeleteCenterPattern(asmDoc, iLogicAuto, seedOcc, keepWorkFeatures:=False, logs:=logs)
    
    ' Output logs
    For Each logMsg As String In logs
        Logger.Info(logMsg)
    Next
    
    If success Then
        Logger.Info("Kordused keskelt: Muster kustutatud edukalt")
        MessageBox.Show("Muster kustutatud edukalt.", "Kordused keskelt")
    Else
        Logger.Error("Kordused keskelt: Mustri kustutamine ebaõnnestus")
        MessageBox.Show("Mustri kustutamine ebaõnnestus. Vaata logi akent.", "Kordused keskelt")
    End If
End Sub

' ============================================================================
' Utilities
' ============================================================================

Function ExtractBaseName(occName As String) As String
    Dim colonPos As Integer = occName.LastIndexOf(":")
    If colonPos > 0 Then
        Return occName.Substring(0, colonPos)
    End If
    Return occName
End Function

Function FindPatternSeedOccurrence(asmDoc As AssemblyDocument, pattern As RectangularOccurrencePattern) As ComponentOccurrence
    ' Get pattern name and derive seed occurrence name
    Dim patternName As String = pattern.Name
    
    ' Pattern name is typically "BaseName_Muster", seed occurrence is "BaseName:1"
    If patternName.EndsWith("_Muster") Then
        Dim baseName As String = patternName.Substring(0, patternName.Length - 7)
        Dim seedName As String = baseName & ":1"
        
        ' Find occurrence by name
        For Each occ As ComponentOccurrence In asmDoc.ComponentDefinition.Occurrences
            If occ.Name = seedName Then
                Return occ
            End If
        Next
        
        ' If not found by exact name, look for any occurrence with pattern config and matching base name
        For Each occ As ComponentOccurrence In asmDoc.ComponentDefinition.Occurrences
            If CenterPatternLib.HasPatternConfig(occ) Then
                Dim occBaseName As String = ExtractBaseName(occ.Name)
                If occBaseName = baseName Then
                    Return occ
                End If
            End If
        Next
    End If
    
    Return Nothing
End Function
