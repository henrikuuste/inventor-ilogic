' ============================================================================
' RotateOriginAxes - Reorient Part via Derived Transformation
' 
' Creates a new part file with a rotation transformation applied to reorient
' the model relative to origin axes. Uses DerivedPartTransformDef API.
'
' Includes by default:
' - All solid bodies
' - All surfaces  
' - All sketches
' - All parameters (model and user)
'
' Usage: Open a part file (.ipt), run this rule, select rotation preset,
'        and specify output filename.
'
' ============================================================================

Sub Main()
    Dim app As Inventor.Application = ThisApplication
    Dim doc As Document = app.ActiveDocument
    
    Logger.Info("RotateOriginAxes: Starting...")
    
    ' Validate document type
    If doc Is Nothing Then
        Logger.Error("RotateOriginAxes: No active document.")
        MessageBox.Show("Aktiivne dokument puudub.", "Pööra telgi")
        Exit Sub
    End If
    
    If doc.DocumentType <> DocumentTypeEnum.kPartDocumentObject Then
        Logger.Error("RotateOriginAxes: Active document is not a part file.")
        MessageBox.Show("See reegel töötab ainult detaili failidega (.ipt).", "Pööra telgi")
        Exit Sub
    End If
    
    Dim sourceDoc As PartDocument = CType(doc, PartDocument)
    
    ' Check if source document is saved
    If String.IsNullOrEmpty(sourceDoc.FullDocumentName) OrElse Not System.IO.File.Exists(sourceDoc.FullDocumentName) Then
        Logger.Error("RotateOriginAxes: Source document must be saved first.")
        MessageBox.Show("Dokument peab olema esmalt salvestatud.", "Pööra telgi")
        Exit Sub
    End If
    
    Dim sourceFilePath As String = sourceDoc.FullDocumentName
    Dim sourceFolder As String = System.IO.Path.GetDirectoryName(sourceFilePath)
    Dim sourceFileName As String = System.IO.Path.GetFileNameWithoutExtension(sourceFilePath)
    
    Logger.Info("RotateOriginAxes: Source file: " & sourceFilePath)
    
    ' Show rotation selection dialog
    Dim selectedPresetIndex As Integer = -1
    Dim outputFileName As String = sourceFileName & "_Rotated.ipt"
    
    Dim dialogResult As DialogResult = ShowRotationDialog(selectedPresetIndex, outputFileName)
    
    If dialogResult <> DialogResult.OK Then
        Logger.Info("RotateOriginAxes: User cancelled.")
        Exit Sub
    End If
    
    If selectedPresetIndex < 0 Then
        Logger.Error("RotateOriginAxes: No rotation preset selected.")
        MessageBox.Show("Pööret ei valitud.", "Pööra telgi")
        Exit Sub
    End If
    
    ' Get rotation parameters for selected preset
    Dim axisX As Double = 0, axisY As Double = 0, axisZ As Double = 0
    Dim angleDegrees As Double = 0
    Dim presetName As String = ""
    
    GetRotationPreset(selectedPresetIndex, axisX, axisY, axisZ, angleDegrees, presetName)
    
    Logger.Info("RotateOriginAxes: Selected preset: " & presetName)
    Logger.Info("RotateOriginAxes: Rotation axis: (" & axisX & ", " & axisY & ", " & axisZ & ")")
    Logger.Info("RotateOriginAxes: Rotation angle: " & angleDegrees & " degrees")
    
    ' Build output path
    Dim outputFilePath As String = System.IO.Path.Combine(sourceFolder, outputFileName)
    If Not outputFileName.EndsWith(".ipt", StringComparison.OrdinalIgnoreCase) Then
        outputFilePath = outputFilePath & ".ipt"
    End If
    
    Logger.Info("RotateOriginAxes: Output file: " & outputFilePath)
    
    ' Check if output file already exists
    If System.IO.File.Exists(outputFilePath) Then
        Dim overwriteResult As DialogResult = MessageBox.Show( _
            "Fail '" & outputFileName & "' on juba olemas." & vbCrLf & vbCrLf & _
            "Kas soovid selle üle kirjutada?", _
            "Pööra telgi", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
        
        If overwriteResult <> DialogResult.Yes Then
            Logger.Info("RotateOriginAxes: User declined overwrite.")
            Exit Sub
        End If
    End If
    
    ' Create the derived part with transformation
    Try
        CreateDerivedPartWithRotation(app, sourceDoc, outputFilePath, axisX, axisY, axisZ, angleDegrees)
        
        Logger.Info("RotateOriginAxes: Successfully created rotated part: " & outputFilePath)
        MessageBox.Show("Pööratud detail loodud:" & vbCrLf & vbCrLf & outputFilePath, "Pööra telgi")
        
    Catch ex As Exception
        Logger.Error("RotateOriginAxes: Failed to create derived part: " & ex.Message)
        MessageBox.Show("Viga pööratud detaili loomisel:" & vbCrLf & vbCrLf & ex.Message, "Pööra telgi")
    End Try
    
    Logger.Info("RotateOriginAxes: Completed.")
End Sub

' ============================================================================
' ROTATION PRESETS
' ============================================================================

Function GetPresetCount() As Integer
    Return 11
End Function

Sub GetRotationPreset(index As Integer, ByRef axisX As Double, ByRef axisY As Double, ByRef axisZ As Double, _
                      ByRef angleDegrees As Double, ByRef presetName As String)
    Select Case index
        Case 0  ' Cyclic forward: X->Y, Y->Z, Z->X
            axisX = 1 : axisY = 1 : axisZ = 1
            angleDegrees = 120
            presetName = "X → Y, Y → Z, Z → X (tsükliline)"
            
        Case 1  ' Cyclic reverse: X->Z, Z->Y, Y->X
            axisX = 1 : axisY = 1 : axisZ = 1
            angleDegrees = -120
            presetName = "X → Z, Z → Y, Y → X (vastupidine tsükliline)"
            
        Case 2  ' Swap X and Y
            axisX = 1 : axisY = 1 : axisZ = 0
            angleDegrees = 180
            presetName = "Vaheta X ja Y"
            
        Case 3  ' Swap Y and Z
            axisX = 0 : axisY = 1 : axisZ = 1
            angleDegrees = 180
            presetName = "Vaheta Y ja Z"
            
        Case 4  ' Swap X and Z
            axisX = 1 : axisY = 0 : axisZ = 1
            angleDegrees = 180
            presetName = "Vaheta X ja Z"
            
        Case 5  ' 90 deg around X
            axisX = 1 : axisY = 0 : axisZ = 0
            angleDegrees = 90
            presetName = "90° ümber X-telje"
            
        Case 6  ' 90 deg around Y
            axisX = 0 : axisY = 1 : axisZ = 0
            angleDegrees = 90
            presetName = "90° ümber Y-telje"
            
        Case 7  ' 90 deg around Z
            axisX = 0 : axisY = 0 : axisZ = 1
            angleDegrees = 90
            presetName = "90° ümber Z-telje"
            
        Case 8  ' -90 deg around X
            axisX = 1 : axisY = 0 : axisZ = 0
            angleDegrees = -90
            presetName = "-90° ümber X-telje"
            
        Case 9  ' -90 deg around Y
            axisX = 0 : axisY = 1 : axisZ = 0
            angleDegrees = -90
            presetName = "-90° ümber Y-telje"
            
        Case 10  ' -90 deg around Z
            axisX = 0 : axisY = 0 : axisZ = 1
            angleDegrees = -90
            presetName = "-90° ümber Z-telje"
            
        Case Else
            axisX = 0 : axisY = 0 : axisZ = 1
            angleDegrees = 0
            presetName = "Tundmatu"
    End Select
End Sub

Function GetPresetDescription(index As Integer) As String
    Select Case index
        Case 0
            Return "X-telg saab Y-teljeks" & vbCrLf & _
                   "Y-telg saab Z-teljeks" & vbCrLf & _
                   "Z-telg saab X-teljeks"
        Case 1
            Return "X-telg saab Z-teljeks" & vbCrLf & _
                   "Z-telg saab Y-teljeks" & vbCrLf & _
                   "Y-telg saab X-teljeks"
        Case 2
            Return "X-telg saab Y-teljeks" & vbCrLf & _
                   "Y-telg saab X-teljeks" & vbCrLf & _
                   "Z-telg jääb samaks"
        Case 3
            Return "X-telg jääb samaks" & vbCrLf & _
                   "Y-telg saab Z-teljeks" & vbCrLf & _
                   "Z-telg saab Y-teljeks"
        Case 4
            Return "X-telg saab Z-teljeks" & vbCrLf & _
                   "Y-telg jääb samaks" & vbCrLf & _
                   "Z-telg saab X-teljeks"
        Case 5
            Return "X-telg jääb samaks" & vbCrLf & _
                   "Y-telg saab Z-teljeks" & vbCrLf & _
                   "Z-telg saab -Y-teljeks"
        Case 6
            Return "X-telg saab -Z-teljeks" & vbCrLf & _
                   "Y-telg jääb samaks" & vbCrLf & _
                   "Z-telg saab X-teljeks"
        Case 7
            Return "X-telg saab Y-teljeks" & vbCrLf & _
                   "Y-telg saab -X-teljeks" & vbCrLf & _
                   "Z-telg jääb samaks"
        Case 8
            Return "X-telg jääb samaks" & vbCrLf & _
                   "Y-telg saab -Z-teljeks" & vbCrLf & _
                   "Z-telg saab Y-teljeks"
        Case 9
            Return "X-telg saab Z-teljeks" & vbCrLf & _
                   "Y-telg jääb samaks" & vbCrLf & _
                   "Z-telg saab -X-teljeks"
        Case 10
            Return "X-telg saab -Y-teljeks" & vbCrLf & _
                   "Y-telg saab X-teljeks" & vbCrLf & _
                   "Z-telg jääb samaks"
        Case Else
            Return ""
    End Select
End Function

' ============================================================================
' USER INTERFACE
' ============================================================================

Function ShowRotationDialog(ByRef selectedPresetIndex As Integer, ByRef outputFileName As String) As DialogResult
    ' Create form
    Dim frm As New System.Windows.Forms.Form()
    frm.Text = "Pööra detaili telgi"
    frm.Width = 420
    frm.Height = 340
    frm.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
    frm.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
    frm.MaximizeBox = False
    frm.MinimizeBox = False
    
    Dim yPos As Integer = 15
    
    ' Label for rotation selection
    Dim lblRotation As New System.Windows.Forms.Label()
    lblRotation.Text = "Vali telje pööre:"
    lblRotation.Left = 15
    lblRotation.Top = yPos
    lblRotation.Width = 380
    lblRotation.Height = 20
    frm.Controls.Add(lblRotation)
    yPos += 25
    
    ' ComboBox for rotation presets
    Dim cboRotation As New System.Windows.Forms.ComboBox()
    cboRotation.Left = 15
    cboRotation.Top = yPos
    cboRotation.Width = 370
    cboRotation.Height = 25
    cboRotation.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
    cboRotation.Name = "cboRotation"
    
    ' Populate presets
    For i As Integer = 0 To GetPresetCount() - 1
        Dim axisX As Double = 0, axisY As Double = 0, axisZ As Double = 0
        Dim angleDegrees As Double = 0
        Dim presetName As String = ""
        GetRotationPreset(i, axisX, axisY, axisZ, angleDegrees, presetName)
        cboRotation.Items.Add(presetName)
    Next
    
    cboRotation.SelectedIndex = 0
    frm.Controls.Add(cboRotation)
    yPos += 35
    
    ' Label for description
    Dim lblDescTitle As New System.Windows.Forms.Label()
    lblDescTitle.Text = "Kirjeldus:"
    lblDescTitle.Left = 15
    lblDescTitle.Top = yPos
    lblDescTitle.Width = 380
    lblDescTitle.Height = 20
    frm.Controls.Add(lblDescTitle)
    yPos += 22
    
    ' TextBox for description (read-only, multiline)
    Dim txtDescription As New System.Windows.Forms.TextBox()
    txtDescription.Left = 15
    txtDescription.Top = yPos
    txtDescription.Width = 370
    txtDescription.Height = 60
    txtDescription.Multiline = True
    txtDescription.ReadOnly = True
    txtDescription.Name = "txtDescription"
    txtDescription.Text = GetPresetDescription(0)
    frm.Controls.Add(txtDescription)
    yPos += 70
    
    ' Update description when selection changes
    AddHandler cboRotation.SelectedIndexChanged, Sub(s, e)
        Dim descBox As System.Windows.Forms.TextBox = CType(frm.Controls("txtDescription"), System.Windows.Forms.TextBox)
        Dim combo As System.Windows.Forms.ComboBox = CType(s, System.Windows.Forms.ComboBox)
        descBox.Text = GetPresetDescription(combo.SelectedIndex)
    End Sub
    
    ' Label for output filename
    Dim lblOutput As New System.Windows.Forms.Label()
    lblOutput.Text = "Väljundfail:"
    lblOutput.Left = 15
    lblOutput.Top = yPos
    lblOutput.Width = 380
    lblOutput.Height = 20
    frm.Controls.Add(lblOutput)
    yPos += 22
    
    ' TextBox for output filename
    Dim txtOutput As New System.Windows.Forms.TextBox()
    txtOutput.Left = 15
    txtOutput.Top = yPos
    txtOutput.Width = 370
    txtOutput.Height = 23
    txtOutput.Name = "txtOutput"
    txtOutput.Text = outputFileName
    frm.Controls.Add(txtOutput)
    yPos += 40
    
    ' OK button
    Dim btnOK As New System.Windows.Forms.Button()
    btnOK.Text = "OK"
    btnOK.Left = 200
    btnOK.Top = yPos
    btnOK.Width = 85
    btnOK.Height = 28
    btnOK.DialogResult = System.Windows.Forms.DialogResult.OK
    frm.Controls.Add(btnOK)
    frm.AcceptButton = btnOK
    
    ' Cancel button
    Dim btnCancel As New System.Windows.Forms.Button()
    btnCancel.Text = "Tühista"
    btnCancel.Left = 295
    btnCancel.Top = yPos
    btnCancel.Width = 85
    btnCancel.Height = 28
    btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
    frm.Controls.Add(btnCancel)
    frm.CancelButton = btnCancel
    
    ' Show dialog
    Dim result As DialogResult = frm.ShowDialog()
    
    ' Get values from form
    If result = DialogResult.OK Then
        selectedPresetIndex = CType(frm.Controls("cboRotation"), System.Windows.Forms.ComboBox).SelectedIndex
        outputFileName = CType(frm.Controls("txtOutput"), System.Windows.Forms.TextBox).Text.Trim()
    End If
    
    frm.Dispose()
    Return result
End Function

' ============================================================================
' DERIVED PART CREATION
' ============================================================================

Sub CreateDerivedPartWithRotation(app As Inventor.Application, sourceDoc As PartDocument, _
                                   outputFilePath As String, _
                                   axisX As Double, axisY As Double, axisZ As Double, _
                                   angleDegrees As Double)
    
    Logger.Info("RotateOriginAxes: Creating derived part with rotation...")
    
    ' Find part template
    Dim templatePath As String = FindPartTemplate(app)
    
    ' Create new part document
    Dim newDoc As PartDocument = Nothing
    
    If String.IsNullOrEmpty(templatePath) Then
        newDoc = app.Documents.Add(DocumentTypeEnum.kPartDocumentObject)
        Logger.Info("RotateOriginAxes: Created new part from default template")
    Else
        newDoc = app.Documents.Add(DocumentTypeEnum.kPartDocumentObject, templatePath, True)
        Logger.Info("RotateOriginAxes: Created new part from template: " & templatePath)
    End If
    
    Dim newCompDef As PartComponentDefinition = newDoc.ComponentDefinition
    Dim dpcs As DerivedPartComponents = newCompDef.ReferenceComponents.DerivedPartComponents
    
    ' Build rotation matrix
    Dim tg As TransientGeometry = app.TransientGeometry
    Dim rotMatrix As Matrix = tg.CreateMatrix()
    
    ' Normalize axis vector
    Dim axisLen As Double = Math.Sqrt(axisX * axisX + axisY * axisY + axisZ * axisZ)
    If axisLen > 0.0001 Then
        axisX = axisX / axisLen
        axisY = axisY / axisLen
        axisZ = axisZ / axisLen
    End If
    
    ' Create axis vector and origin point
    Dim axisVec As Vector = tg.CreateVector(axisX, axisY, axisZ)
    Dim origin As Point = tg.CreatePoint(0, 0, 0)
    
    ' Convert angle to radians
    Dim angleRadians As Double = angleDegrees * Math.PI / 180.0
    
    ' Set rotation
    Logger.Info("RotateOriginAxes: Setting rotation: " & angleDegrees & " deg around (" & _
                axisX.ToString("0.###") & ", " & axisY.ToString("0.###") & ", " & axisZ.ToString("0.###") & ")")
    rotMatrix.SetToRotation(angleRadians, axisVec, origin)
    
    ' Use DerivedPartTransformDef for transformation (includes bodies by default)
    ' NOTE: This only works for solid bodies. Sketches are NOT transformed.
    Logger.Info("RotateOriginAxes: Creating DerivedPartTransformDef...")
    Dim derivedDef As DerivedPartTransformDef = dpcs.CreateTransformDef(sourceDoc.FullDocumentName)
    
    ' Apply transformation
    derivedDef.SetTransformation(rotMatrix)
    Logger.Info("RotateOriginAxes: Transformation applied to derived definition")
    
    ' Add derived component - this derives bodies with the transformation applied
    Logger.Info("RotateOriginAxes: Adding derived component...")
    Dim derivedComp As DerivedPartComponent = dpcs.Add(derivedDef)
    Logger.Info("RotateOriginAxes: Derived component added")
    
    ' Update document
    newDoc.Update()
    
    ' Check what was derived
    Dim bodyCount As Integer = 0
    For Each body As SurfaceBody In newCompDef.SurfaceBodies
        bodyCount += 1
    Next
    
    Logger.Info("RotateOriginAxes: Body count in derived part: " & bodyCount)
    
    If bodyCount = 0 Then
        ' No bodies derived - the source part likely has no solid bodies
        ' DerivedPartTransformDef only transforms bodies, not sketches
        Logger.Warn("RotateOriginAxes: No bodies in derived part.")
        Logger.Warn("RotateOriginAxes: Source part may have no solid bodies, only sketches.")
        Logger.Warn("RotateOriginAxes: DerivedPartTransformDef only transforms bodies, NOT sketches.")
        
        ' Close without saving
        newDoc.Close(True)
        
        ' Show error to user
        MessageBox.Show( _
            "Lähtedokumendil pole tahkkehasid - ainult eskiisid." & vbCrLf & vbCrLf & _
            "DerivedPartTransformDef saab pöörata ainult tahkkehasid. " & _
            "Eskiise ei saa Inventor API kaudu automaatselt pöörata." & vbCrLf & vbCrLf & _
            "Võimalikud lahendused:" & vbCrLf & _
            "1. Loo esmalt tahkkeha (ekstrusioon vms), siis käivita see tööriist" & vbCrLf & _
            "2. Määratle eskiisid käsitsi ümber uutele tasapindadele", _
            "Pööra telgi - Viga", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Exit Sub
    End If
    
    ' Save new file
    Logger.Info("RotateOriginAxes: Saving to: " & outputFilePath)
    app.SilentOperation = True
    Try
        newDoc.SaveAs(outputFilePath, False)
    Finally
        app.SilentOperation = False
    End Try
    
    Logger.Info("RotateOriginAxes: File saved successfully")
    
    ' Close the new document (user can open it manually if desired)
    newDoc.Close(False)
    Logger.Info("RotateOriginAxes: New document closed")
End Sub

Function FindPartTemplate(app As Inventor.Application) As String
    Dim templateNames() As String = {"Part.ipt", "Standard.ipt", "Metric\Standard (mm).ipt"}
    
    Try
        Dim templatesPath As String = app.DesignProjectManager.ActiveDesignProject.TemplatesPath
        
        For Each templateName As String In templateNames
            Dim fullPath As String = System.IO.Path.Combine(templatesPath, templateName)
            If System.IO.File.Exists(fullPath) Then
                Return fullPath
            End If
        Next
        
        Dim enUSPath As String = System.IO.Path.Combine(templatesPath, "en-US")
        If System.IO.Directory.Exists(enUSPath) Then
            For Each templateName As String In templateNames
                Dim fullPath As String = System.IO.Path.Combine(enUSPath, templateName)
                If System.IO.File.Exists(fullPath) Then
                    Return fullPath
                End If
            Next
        End If
        
    Catch ex As Exception
        Logger.Warn("RotateOriginAxes: Error finding template: " & ex.Message)
    End Try
    
    Return ""
End Function
