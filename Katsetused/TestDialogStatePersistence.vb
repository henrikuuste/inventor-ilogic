' ============================================================================
' TestDialogStatePersistence - Test dialog close/reopen with state preservation
' 
' Tests:
' - Can we close dialog, do a CommandManager.Pick, and reopen?
' - Does form.Tag work for storing complex state?
' - Is state correctly restored after pick?
'
' Usage: Open any part document, then run this rule.
'        Click [Vali pind] to test the pick workflow.
' ============================================================================

Sub Main()
    Dim app As Inventor.Application = ThisApplication
    Dim doc As Document = app.ActiveDocument
    
    Logger.Info("TestDialogStatePersistence: Starting dialog state tests...")
    
    ' Validate document
    If doc Is Nothing Then
        Logger.Error("TestDialogStatePersistence: No active document.")
        MessageBox.Show("Aktiivne dokument puudub.", "TestDialogStatePersistence")
        Exit Sub
    End If
    
    If doc.DocumentType <> DocumentTypeEnum.kPartDocumentObject Then
        Logger.Error("TestDialogStatePersistence: Active document is not a part file.")
        MessageBox.Show("See reegel töötab ainult detaili failidega (.ipt).", "TestDialogStatePersistence")
        Exit Sub
    End If
    
    ' Run the dialog loop
    Dim textValue As String = "Algväärtus"
    Dim pickedInfo As String = ""
    Dim pickCount As Integer = 0
    
    Do
        Dim action As String = ""
        Dim result As DialogResult = ShowTestDialog(textValue, pickedInfo, pickCount, action)
        
        If result = DialogResult.Cancel Then
            Logger.Info("TestDialogStatePersistence: Dialog cancelled")
            Exit Do
        End If
        
        If result = DialogResult.OK Then
            Logger.Info("TestDialogStatePersistence: Dialog OK - Text value: '" & textValue & "'")
            Logger.Info("TestDialogStatePersistence: Picked info: '" & pickedInfo & "'")
            Logger.Info("TestDialogStatePersistence: Pick count: " & pickCount)
            Exit Do
        End If
        
        If action = "PICK" Then
            Logger.Info("TestDialogStatePersistence: Performing pick (dialog closed)...")
            
            ' Dialog is now closed - we can pick
            Dim pickedFace As Face = Nothing
            Try
                pickedFace = app.CommandManager.Pick(SelectionFilterEnum.kPartFacePlanarFilter, _
                    "Vali pind - ESC tühistamiseks")
            Catch
                Logger.Info("TestDialogStatePersistence: Pick cancelled")
            End Try
            
            If pickedFace IsNot Nothing Then
                pickCount += 1
                
                ' Get face normal info
                Dim nx As Double = 0, ny As Double = 0, nz As Double = 0
                GetFaceNormal(pickedFace, nx, ny, nz)
                
                pickedInfo = "Pind " & pickCount & ": normal=(" & _
                             FormatNumber(nx, 3) & ", " & _
                             FormatNumber(ny, 3) & ", " & _
                             FormatNumber(nz, 3) & ")"
                
                Logger.Info("TestDialogStatePersistence: Picked face - " & pickedInfo)
            End If
            
            ' Loop will reopen dialog with preserved state
        End If
    Loop
    
    Logger.Info("TestDialogStatePersistence: Test completed!")
    Logger.Info("TestDialogStatePersistence: Final text value: '" & textValue & "'")
    Logger.Info("TestDialogStatePersistence: Total picks: " & pickCount)
End Sub

Function ShowTestDialog(ByRef textValue As String, ByVal pickedInfo As String, ByVal pickCount As Integer, ByRef action As String) As DialogResult
    action = ""
    
    ' Create form
    Dim frm As New System.Windows.Forms.Form()
    frm.Text = "Test Dialog State Persistence"
    frm.Width = 450
    frm.Height = 300
    frm.StartPosition = FormStartPosition.CenterScreen
    frm.FormBorderStyle = FormBorderStyle.FixedDialog
    frm.MaximizeBox = False
    frm.MinimizeBox = False
    
    ' Instructions label
    Dim lblInstructions As New System.Windows.Forms.Label()
    lblInstructions.Text = "See test kontrollib, kas dialoogi olek säilib" & vbCrLf & _
                           "pärast pinna valimist ja dialoogi taasavamist."
    lblInstructions.Left = 15
    lblInstructions.Top = 15
    lblInstructions.Width = 400
    lblInstructions.Height = 40
    frm.Controls.Add(lblInstructions)
    
    ' Text field label
    Dim lblText As New System.Windows.Forms.Label()
    lblText.Text = "Teksti väli (peaks säilima):"
    lblText.Left = 15
    lblText.Top = 65
    lblText.Width = 150
    frm.Controls.Add(lblText)
    
    ' Text field
    Dim txtValue As New System.Windows.Forms.TextBox()
    txtValue.Name = "txtValue"
    txtValue.Text = textValue
    txtValue.Left = 170
    txtValue.Top = 62
    txtValue.Width = 200
    frm.Controls.Add(txtValue)
    
    ' Picked info label
    Dim lblPicked As New System.Windows.Forms.Label()
    lblPicked.Text = "Viimane valik:"
    lblPicked.Left = 15
    lblPicked.Top = 100
    lblPicked.Width = 100
    frm.Controls.Add(lblPicked)
    
    ' Picked info value
    Dim lblPickedValue As New System.Windows.Forms.Label()
    lblPickedValue.Text = If(String.IsNullOrEmpty(pickedInfo), "(pole valitud)", pickedInfo)
    lblPickedValue.Left = 120
    lblPickedValue.Top = 100
    lblPickedValue.Width = 300
    frm.Controls.Add(lblPickedValue)
    
    ' Pick count label
    Dim lblCount As New System.Windows.Forms.Label()
    lblCount.Text = "Valikute arv: " & pickCount
    lblCount.Left = 15
    lblCount.Top = 130
    lblCount.Width = 200
    frm.Controls.Add(lblCount)
    
    ' Pick button
    Dim btnPick As New System.Windows.Forms.Button()
    btnPick.Text = "Vali pind"
    btnPick.Left = 15
    btnPick.Top = 170
    btnPick.Width = 120
    btnPick.Height = 30
    frm.Controls.Add(btnPick)
    
    ' OK button
    Dim btnOK As New System.Windows.Forms.Button()
    btnOK.Text = "OK"
    btnOK.Left = 240
    btnOK.Top = 220
    btnOK.Width = 80
    btnOK.Height = 30
    btnOK.DialogResult = DialogResult.OK
    frm.Controls.Add(btnOK)
    
    ' Cancel button
    Dim btnCancel As New System.Windows.Forms.Button()
    btnCancel.Text = "Tühista"
    btnCancel.Left = 330
    btnCancel.Top = 220
    btnCancel.Width = 80
    btnCancel.Height = 30
    btnCancel.DialogResult = DialogResult.Cancel
    frm.Controls.Add(btnCancel)
    
    frm.AcceptButton = btnOK
    frm.CancelButton = btnCancel
    
    ' Store action in form.Tag (will be set by button handlers)
    frm.Tag = ""
    
    ' Pick button handler - close dialog with special result
    ' Note: Cannot use ByRef params in lambda, so we use Tag and read text after close
    AddHandler btnPick.Click, Sub(sender, e)
        frm.Tag = "PICK"
        frm.DialogResult = DialogResult.Retry  ' Use Retry as "pick" signal
    End Sub
    
    ' Show dialog
    Dim result As DialogResult = frm.ShowDialog()
    
    ' Get action from Tag
    action = CStr(frm.Tag)
    
    ' Save text value after dialog closes (works for all cases)
    textValue = txtValue.Text
    
    Return result
End Function

Function GetFaceNormal(face As Face, ByRef nx As Double, ByRef ny As Double, ByRef nz As Double) As Boolean
    Try
        Dim geom As Object = face.Geometry
        If TypeOf geom Is Plane Then
            Dim plane As Plane = CType(geom, Plane)
            Dim normal As UnitVector = plane.Normal
            nx = normal.X
            ny = normal.Y
            nz = normal.Z
            Return True
        End If
    Catch
    End Try
    Return False
End Function
