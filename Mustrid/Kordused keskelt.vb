' ============================================================================
' Kordused Keskelt - Center-Based Occurrence Pattern
' 
' Creates a parametric pattern that distributes instances evenly across a span.
' The span is defined by two boundary planes/faces, and instances are placed
' based on maximum spacing with options for uniform or symmetric distribution.
'
' Features:
' - Uniform or symmetric-from-center distribution
' - Include/exclude instances at span boundaries
' - Constraint-based seed positioning (updates with geometry changes)
' - Parametric spacing via assembly parameters
'
' Usage:
' 1. Select an occurrence to pattern (or have it pre-selected)
' 2. Pick start and end boundaries (planes or faces)
' 3. Enter max spacing (mm) or select existing parameter
' 4. Choose distribution mode and end options
' 5. Click OK to create the pattern
'
' ESC cancels any selection operation.
'
' ============================================================================

AddVbFile "Lib/UtilsLib.vb"
AddVbFile "Lib/GeoLib.vb"
AddVbFile "Lib/WorkFeatureLib.vb"
AddVbFile "Lib/PatternLib.vb"
AddVbFile "Lib/CenterPatternLib.vb"

Sub Main()
    Dim app As Inventor.Application = ThisApplication
    Dim doc As Document = app.ActiveDocument
    
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
    RunPatternSetup(app, asmDoc)
End Sub

' ============================================================================
' Main Setup Loop
' ============================================================================

Sub RunPatternSetup(app As Inventor.Application, asmDoc As AssemblyDocument)
    ' State variables
    Dim seedOcc As ComponentOccurrence = Nothing
    Dim baseName As String = ""
    Dim startGeometry As Object = Nothing
    Dim endGeometry As Object = Nothing
    Dim maxSpacingInput As String = "500"
    Dim mode As String = CenterPatternLib.MODE_UNIFORM
    Dim includeEnds As Boolean = False
    
    ' Check for pre-selected occurrence
    If asmDoc.SelectSet.Count = 1 Then
        If TypeOf asmDoc.SelectSet.Item(1) Is ComponentOccurrence Then
            seedOcc = CType(asmDoc.SelectSet.Item(1), ComponentOccurrence)
            baseName = ExtractBaseName(seedOcc.Name)
        End If
    End If
    
    ' Main dialog loop
    Dim keepGoing As Boolean = True
    Do While keepGoing
        Dim action As String = ""
        Dim result As DialogResult = ShowSetupForm(app, asmDoc, _
            seedOcc, baseName, startGeometry, endGeometry, _
            maxSpacingInput, mode, includeEnds, action)
        
        Select Case result
            Case DialogResult.Cancel
                Exit Do
                
            Case DialogResult.OK
                ' Create the pattern
                If ValidateInputs(seedOcc, startGeometry, endGeometry, maxSpacingInput) Then
                    CreatePattern(app, asmDoc, seedOcc, baseName, _
                                  startGeometry, endGeometry, maxSpacingInput, mode, includeEnds)
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
                        End If
                        
                    Case "PICK_START"
                        Dim picked As Object = PickPlanarGeometry(app, "Vali alguspind (tasand või pind) - ESC tühistamiseks")
                        If picked IsNot Nothing Then
                            startGeometry = picked
                        End If
                        
                    Case "PICK_END"
                        Dim picked As Object = PickPlanarGeometry(app, "Vali lõpupind (tasand või pind) - ESC tühistamiseks")
                        If picked IsNot Nothing Then
                            endGeometry = picked
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
                       ByRef maxSpacingInput As String, _
                       ByRef mode As String, _
                       ByRef includeEnds As Boolean, _
                       ByRef action As String) As DialogResult
    
    action = ""
    
    ' Create form
    Dim frm As New System.Windows.Forms.Form()
    frm.Text = "Kordused Keskelt"
    frm.Width = 450
    frm.Height = 380
    frm.StartPosition = FormStartPosition.CenterScreen
    frm.FormBorderStyle = FormBorderStyle.FixedDialog
    frm.MaximizeBox = False
    frm.MinimizeBox = False
    
    Dim yPos As Integer = 15
    Dim labelWidth As Integer = 120
    Dim controlLeft As Integer = 130
    Dim controlWidth As Integer = 200
    Dim btnWidth As Integer = 70
    Dim rowHeight As Integer = 32
    
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
    frm.Controls.Add(txtOcc)
    
    Dim btnPickOcc As New System.Windows.Forms.Button()
    btnPickOcc.Text = "Vali..."
    btnPickOcc.Left = controlLeft + controlWidth + 10
    btnPickOcc.Top = yPos - 2
    btnPickOcc.Width = btnWidth
    btnPickOcc.Height = 26
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
    txtStart.Text = If(startGeometry IsNot Nothing, UtilsLib.GetObjectDisplayName(startGeometry), "(vali pind)")
    frm.Controls.Add(txtStart)
    
    Dim btnPickStart As New System.Windows.Forms.Button()
    btnPickStart.Text = "Vali..."
    btnPickStart.Left = controlLeft + controlWidth + 10
    btnPickStart.Top = yPos - 2
    btnPickStart.Width = btnWidth
    btnPickStart.Height = 26
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
    txtEnd.Text = If(endGeometry IsNot Nothing, UtilsLib.GetObjectDisplayName(endGeometry), "(vali pind)")
    frm.Controls.Add(txtEnd)
    
    Dim btnPickEnd As New System.Windows.Forms.Button()
    btnPickEnd.Text = "Vali..."
    btnPickEnd.Left = controlLeft + controlWidth + 10
    btnPickEnd.Top = yPos - 2
    btnPickEnd.Width = btnWidth
    btnPickEnd.Height = 26
    frm.Controls.Add(btnPickEnd)
    
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
    txtSpacing.Width = 100
    txtSpacing.Text = maxSpacingInput
    frm.Controls.Add(txtSpacing)
    
    ' Parameter dropdown
    Dim cboParams As New System.Windows.Forms.ComboBox()
    cboParams.Name = "cboParams"
    cboParams.Left = controlLeft + 110
    cboParams.Top = yPos
    cboParams.Width = 140
    cboParams.DropDownStyle = ComboBoxStyle.DropDownList
    cboParams.Items.Add("(või vali parameeter)")
    
    Dim paramNames As String() = CenterPatternLib.GetUserParameterNames(asmDoc)
    For Each pName As String In paramNames
        cboParams.Items.Add(pName)
    Next
    cboParams.SelectedIndex = 0
    frm.Controls.Add(cboParams)
    
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
    
    yPos += rowHeight + 10
    
    ' --- Include ends checkbox ---
    Dim chkEnds As New System.Windows.Forms.CheckBox()
    chkEnds.Name = "chkEnds"
    chkEnds.Text = "Elemendid otstes (algus- ja lõpupinnal)"
    chkEnds.Left = controlLeft
    chkEnds.Top = yPos
    chkEnds.Width = 250
    chkEnds.Checked = includeEnds
    frm.Controls.Add(chkEnds)
    
    yPos += rowHeight + 15
    
    ' --- Buttons ---
    Dim btnOK As New System.Windows.Forms.Button()
    btnOK.Text = "Loo muster"
    btnOK.Left = 200
    btnOK.Top = yPos
    btnOK.Width = 100
    btnOK.Height = 32
    btnOK.DialogResult = DialogResult.OK
    frm.AcceptButton = btnOK
    frm.Controls.Add(btnOK)
    
    Dim btnCancel As New System.Windows.Forms.Button()
    btnCancel.Text = "Tühista"
    btnCancel.Left = 310
    btnCancel.Top = yPos
    btnCancel.Width = 80
    btnCancel.Height = 32
    btnCancel.DialogResult = DialogResult.Cancel
    frm.CancelButton = btnCancel
    frm.Controls.Add(btnCancel)
    
    ' --- Event handlers ---
    ' Use Tag to pass action back
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
    
    AddHandler cboParams.SelectedIndexChanged, Sub(s, e)
        If cboParams.SelectedIndex > 0 Then
            txtSpacing.Text = cboParams.SelectedItem.ToString()
        End If
    End Sub
    
    ' Show form
    Dim result As DialogResult = frm.ShowDialog()
    
    ' Read values back
    action = CStr(frm.Tag)
    maxSpacingInput = txtSpacing.Text.Trim()
    mode = If(cboMode.SelectedIndex = 1, CenterPatternLib.MODE_SYMMETRIC, CenterPatternLib.MODE_UNIFORM)
    includeEnds = chkEnds.Checked
    
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
        app, asmDoc, seedOcc, _
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
