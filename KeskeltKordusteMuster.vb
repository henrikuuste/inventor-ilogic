' ============================================================================
' KeskeltKordusteMuster - Support Pattern Rule Generator
' 
' Run this iLogic rule to interactively create a support pattern rule.
' Shows a form to pick a component, displays the name, and creates
' a document-local rule named <baseName>Pattern on OK.
' ============================================================================

Sub Main()
    Dim app As Inventor.Application = ThisApplication
    Dim doc As Document = app.ActiveDocument

    If doc Is Nothing Then
        MessageBox.Show("No active document.", "Support Pattern Generator")
        Exit Sub
    End If

    If doc.DocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
        MessageBox.Show("This rule only works in assembly documents (.iam).", "Support Pattern Generator")
        Exit Sub
    End If

    Dim asmDoc As AssemblyDocument = CType(doc, AssemblyDocument)

    ' ---- Show the picker form (loops until OK or Cancel) ----
    Dim baseName As String = RunPickerLoop(app)

    If baseName = "" Then
        Exit Sub
    End If

    ' ---- Build the rule content ----
    Dim ruleName As String = baseName & "Pattern"
    Dim ruleText As String = _
        "AddVbFile ""Lib/SupportPatternLibrary.vb""" & vbCrLf & _
        "Dim _dep1 = " & baseName & "AlaSuurus" & vbCrLf & _
        "Dim _dep2 = " & baseName & "MaxVahe" & vbCrLf & _
        "SupportPatternLibrary.ApplySupports(ThisApplication, ThisDoc.Document, """ & baseName & """)"

    ' ---- Create or overwrite the iLogic rule ----
    Dim iLogicAuto As Object = Nothing
    Try
        iLogicAuto = iLogicVb.Automation
    Catch ex As Exception
        MessageBox.Show("Failed to access iLogic Automation: " & ex.Message, "Support Pattern Generator")
        Exit Sub
    End Try

    Dim existingRule As Object = Nothing
    Try
        existingRule = iLogicAuto.GetRule(asmDoc, ruleName)
    Catch
        existingRule = Nothing
    End Try

    Try
        If existingRule IsNot Nothing Then
            ' Update existing rule
            existingRule.Text = ruleText
            MessageBox.Show( _
                "Updated existing rule: " & ruleName & vbCrLf & vbCrLf & _
                "Rule content:" & vbCrLf & ruleText, _
                "Support Pattern Generator")
        Else
            ' Create new rule
            iLogicAuto.AddRule(asmDoc, ruleName, ruleText)
            MessageBox.Show( _
                "Created new rule: " & ruleName & vbCrLf & vbCrLf & _
                "Rule content:" & vbCrLf & ruleText, _
                "Support Pattern Generator")
        End If
    Catch ex As Exception
        MessageBox.Show("Failed to create/update rule '" & ruleName & "': " & ex.Message, "Support Pattern Generator")
    End Try

End Sub

Function RunPickerLoop(ByVal app As Inventor.Application) As String
    Dim baseName As String = ""
    Dim keepGoing As Boolean = True

    Do While keepGoing
        ' Show form and get user action
        Dim action As String = ""
        Dim result As DialogResult = ShowPickerForm(baseName, action)

        If result = DialogResult.Cancel Then
            ' User cancelled
            Return ""
        ElseIf result = DialogResult.OK Then
            ' User confirmed selection
            If baseName <> "" Then
                Return baseName
            End If
        ElseIf action = "PICK" Then
            ' User clicked Pick - form closed, now do the actual pick
            Dim pickedName As String = DoComponentPick(app)
            If pickedName <> "" Then
                baseName = pickedName
            End If
            ' Loop continues, form will reopen with the picked name
        End If
    Loop

    Return ""
End Function

Function ShowPickerForm(ByRef baseName As String, ByRef action As String) As DialogResult
    action = ""

    ' Create the form
    Dim frm As New System.Windows.Forms.Form()
    frm.Text = "Keskelt Korduste Muster"
    frm.Width = 360
    frm.Height = 180
    frm.StartPosition = FormStartPosition.CenterScreen
    frm.FormBorderStyle = FormBorderStyle.FixedDialog
    frm.MaximizeBox = False
    frm.MinimizeBox = False

    ' Label for instructions
    Dim lblInstruction As New System.Windows.Forms.Label()
    lblInstruction.Text = "Valitud element:"
    lblInstruction.Left = 15
    lblInstruction.Top = 20
    lblInstruction.AutoSize = True
    frm.Controls.Add(lblInstruction)

    ' TextBox to display selected component name
    Dim txtSelected As New System.Windows.Forms.TextBox()
    txtSelected.Name = "txtSelected"
    txtSelected.Left = 15
    txtSelected.Top = 45
    txtSelected.Width = 220
    txtSelected.Height = 24
    txtSelected.ReadOnly = True
    If baseName <> "" Then
        txtSelected.Text = baseName
    Else
        txtSelected.Text = "(pole valitud)"
    End If
    frm.Controls.Add(txtSelected)

    ' Pick button - uses Retry result to signal pick action
    Dim btnPick As New System.Windows.Forms.Button()
    btnPick.Name = "btnPick"
    btnPick.Text = "Vali..."
    btnPick.Left = 245
    btnPick.Top = 43
    btnPick.Width = 80
    btnPick.Height = 28
    btnPick.DialogResult = DialogResult.Retry ' Special result for pick
    frm.Controls.Add(btnPick)

    ' OK button
    Dim btnOK As New System.Windows.Forms.Button()
    btnOK.Name = "btnOK"
    btnOK.Text = "OK"
    btnOK.Left = 160
    btnOK.Top = 95
    btnOK.Width = 80
    btnOK.Height = 30
    btnOK.Enabled = (baseName <> "")
    btnOK.DialogResult = DialogResult.OK
    frm.AcceptButton = btnOK
    frm.Controls.Add(btnOK)

    ' Cancel button
    Dim btnCancel As New System.Windows.Forms.Button()
    btnCancel.Text = "Tühista"
    btnCancel.Left = 245
    btnCancel.Top = 95
    btnCancel.Width = 80
    btnCancel.Height = 30
    btnCancel.DialogResult = DialogResult.Cancel
    frm.CancelButton = btnCancel
    frm.Controls.Add(btnCancel)

    ' Show the form as modal dialog
    Dim result As DialogResult = frm.ShowDialog()

    ' Check which button was pressed
    If result = DialogResult.Retry Then
        action = "PICK"
    End If

    Return result
End Function

Function DoComponentPick(ByVal app As Inventor.Application) As String
    Dim selFilter As SelectionFilterEnum = SelectionFilterEnum.kAssemblyOccurrenceFilter
    Dim selectedObj As Object = Nothing

    Try
        selectedObj = app.CommandManager.Pick( _
            selFilter, _
            "Vali element mida korrata (e.g., Põõn:1):")
    Catch
        ' User cancelled or selection failed
        Return ""
    End Try

    If selectedObj Is Nothing Then
        Return ""
    End If

    If Not TypeOf selectedObj Is ComponentOccurrence Then
        Return ""
    End If

    Dim occ As ComponentOccurrence = CType(selectedObj, ComponentOccurrence)
    Dim occName As String = occ.Name

    ' Extract base name by stripping the :# suffix
    Dim colonPos As Integer = occName.LastIndexOf(":")
    If colonPos > 0 Then
        Return occName.Substring(0, colonPos)
    Else
        Return occName
    End If
End Function
