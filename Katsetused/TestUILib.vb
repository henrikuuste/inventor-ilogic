' TestUILib.vb - Test script for UILib functionality
' Run this to verify Phase 1 of the Unified UI Library implementation
'
' Tests:
' 1. Form creation (standard, large, compact)
' 2. Layout panels (content, buttons)
' 3. Control factory methods
' 4. Fluid layout behavior (resize to test)
' 5. Non-modal form with viewport interaction

AddVbFile "Lib/UILib.vb"

Sub Main()
    Dim app As Inventor.Application = ThisApplication
    
    Logger.Info("TestUILib: Starting UILib verification tests...")
    
    ' Test 1: Basic form creation
    TestBasicForm()
    
    ' Test 2: Large form with DataGridView
    TestLargeForm()
    
    ' Test 3: Compact form
    TestCompactForm()
    
    ' Test 4: Non-modal form (main test - verify viewport interaction)
    TestNonModalForm(app)
    
    Logger.Info("TestUILib: All tests completed.")
End Sub

Sub TestBasicForm()
    Logger.Info("TestUILib: Test 1 - Basic form with layout...")
    
    Dim frm As System.Windows.Forms.Form = UILib.CreateForm("Test: Basic Form")
    Dim content As System.Windows.Forms.TableLayoutPanel = UILib.CreateContentPanel()
    Dim buttons As System.Windows.Forms.FlowLayoutPanel = UILib.CreateButtonPanel()
    
    ' Add various controls
    UILib.AddRow(content, "Name:", UILib.CreateTextBox("Sample text"))
    UILib.AddRow(content, "Count:", UILib.CreateNumericUpDown(0, 100, 10))
    UILib.AddRow(content, "Option:", UILib.CreateComboBox(New String() {"Option A", "Option B", "Option C"}))
    UILib.AddFullWidthRow(content, UILib.CreateCheckBox("Enable feature", True))
    
    UILib.AddSectionHeader(content, "Additional Settings")
    UILib.AddRow(content, "Value:", UILib.CreateNumericUpDown(0, 1000, 50, 2))
    
    ' Buttons
    Dim btnOK As System.Windows.Forms.Button = UILib.CreateButton("OK")
    btnOK.DialogResult = System.Windows.Forms.DialogResult.OK
    Dim btnCancel As System.Windows.Forms.Button = UILib.CreateButton("Cancel")
    btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
    
    buttons.Controls.Add(btnCancel)
    buttons.Controls.Add(btnOK)
    
    frm.Controls.Add(content)
    frm.Controls.Add(buttons)
    frm.AcceptButton = btnOK
    frm.CancelButton = btnCancel
    
    ' Finalize to set minimum size based on content
    UILib.FinalizeForm(frm)
    
    Logger.Info("TestUILib: Showing basic form - try resizing to test fluid layout...")
    Dim result As System.Windows.Forms.DialogResult = frm.ShowDialog()
    Logger.Info("TestUILib: Basic form closed with: " & result.ToString())
End Sub

Sub TestLargeForm()
    Logger.Info("TestUILib: Test 2 - Large form with ListBox...")
    
    Dim frm As System.Windows.Forms.Form = UILib.CreateLargeForm("Test: Large Form with List")
    Dim content As System.Windows.Forms.TableLayoutPanel = UILib.CreateContentPanel()
    Dim buttons As System.Windows.Forms.FlowLayoutPanel = UILib.CreateButtonPanel()
    
    ' Header row
    UILib.AddRow(content, "Filter:", UILib.CreateTextBox())
    
    ' Fill row with ListBox - should expand vertically
    Dim lst As System.Windows.Forms.ListBox = UILib.CreateListBox()
    For i As Integer = 1 To 50
        lst.Items.Add("Item " & i.ToString("D3") & " - Sample list entry")
    Next
    UILib.AddFillRow(content, lst)
    
    ' Status row
    UILib.AddRow(content, "Selected:", UILib.CreateLabel("0 items"))
    
    ' Buttons
    Dim btnSelectAll As System.Windows.Forms.Button = UILib.CreateButton("Select All", 100)
    Dim btnOK As System.Windows.Forms.Button = UILib.CreateButton("OK")
    btnOK.DialogResult = System.Windows.Forms.DialogResult.OK
    Dim btnCancel As System.Windows.Forms.Button = UILib.CreateButton("Cancel")
    btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
    
    buttons.Controls.Add(btnCancel)
    buttons.Controls.Add(btnOK)
    buttons.Controls.Add(btnSelectAll)
    
    frm.Controls.Add(content)
    frm.Controls.Add(buttons)
    frm.AcceptButton = btnOK
    frm.CancelButton = btnCancel
    
    ' Finalize to set minimum size based on content
    UILib.FinalizeForm(frm)
    
    Logger.Info("TestUILib: Showing large form - resize to verify ListBox fills available space...")
    Dim result As System.Windows.Forms.DialogResult = frm.ShowDialog()
    Logger.Info("TestUILib: Large form closed with: " & result.ToString())
End Sub

Sub TestCompactForm()
    Logger.Info("TestUILib: Test 3 - Compact form (fixed size)...")
    
    Dim frm As System.Windows.Forms.Form = UILib.CreateCompactForm("Test: Compact Form", 300, 150)
    Dim content As System.Windows.Forms.TableLayoutPanel = UILib.CreateContentPanel()
    Dim buttons As System.Windows.Forms.FlowLayoutPanel = UILib.CreateButtonPanel()
    
    UILib.AddFullWidthRow(content, UILib.CreateLabel("This is a compact, fixed-size dialog."))
    UILib.AddFullWidthRow(content, UILib.CreateCheckBox("Confirm action"))
    
    Dim btnOK As System.Windows.Forms.Button = UILib.CreateButton("OK")
    btnOK.DialogResult = System.Windows.Forms.DialogResult.OK
    Dim btnCancel As System.Windows.Forms.Button = UILib.CreateButton("Cancel")
    btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
    
    buttons.Controls.Add(btnCancel)
    buttons.Controls.Add(btnOK)
    
    frm.Controls.Add(content)
    frm.Controls.Add(buttons)
    frm.AcceptButton = btnOK
    frm.CancelButton = btnCancel
    
    Logger.Info("TestUILib: Showing compact form - note it cannot be resized...")
    Dim result As System.Windows.Forms.DialogResult = frm.ShowDialog()
    Logger.Info("TestUILib: Compact form closed with: " & result.ToString())
End Sub

Sub TestNonModalForm(app As Inventor.Application)
    Logger.Info("TestUILib: Test 4 - Non-modal form with viewport interaction...")
    Logger.Info("TestUILib: IMPORTANT: While the form is open, try to:")
    Logger.Info("TestUILib:   - Orbit the view (middle mouse button)")
    Logger.Info("TestUILib:   - Pan the view (Shift + middle mouse)")
    Logger.Info("TestUILib:   - Zoom in/out (scroll wheel)")
    Logger.Info("TestUILib:   - Select objects in the viewport")
    
    Dim frm As System.Windows.Forms.Form = UILib.CreateForm("Test: Non-Modal Form (interact with viewport!)")
    Dim content As System.Windows.Forms.TableLayoutPanel = UILib.CreateContentPanel()
    Dim buttons As System.Windows.Forms.FlowLayoutPanel = UILib.CreateButtonPanel()
    
    ' Info text
    Dim lblInfo As System.Windows.Forms.Label = UILib.CreateLabel( _
        "This form is NON-MODAL." & vbCrLf & vbCrLf & _
        "While this window is open, you should be able to:" & vbCrLf & _
        "• Orbit the viewport (middle mouse)" & vbCrLf & _
        "• Pan the view (Shift + middle mouse)" & vbCrLf & _
        "• Zoom with scroll wheel" & vbCrLf & _
        "• Select objects in the model" & vbCrLf & vbCrLf & _
        "The form stays on top but doesn't block Inventor.")
    lblInfo.AutoSize = True
    UILib.AddFullWidthRow(content, lblInfo)
    
    UILib.AddSectionHeader(content, "Test Controls")
    UILib.AddRow(content, "Value:", UILib.CreateNumericUpDown(0, 100, 50))
    UILib.AddRow(content, "Text:", UILib.CreateTextBox("Edit while viewing model"))
    
    ' Close button
    Dim btnClose As System.Windows.Forms.Button = UILib.CreateButton("Close Test")
    AddHandler btnClose.Click, Sub(s, e)
        frm.DialogResult = System.Windows.Forms.DialogResult.OK
        frm.Close()
    End Sub
    
    buttons.Controls.Add(btnClose)
    
    frm.Controls.Add(content)
    frm.Controls.Add(buttons)
    
    ' Finalize to set minimum size based on content
    UILib.FinalizeForm(frm)
    
    ' This is the key test - ShowNonModal should allow viewport interaction
    UILib.ShowNonModal(frm)
    
    Logger.Info("TestUILib: Non-modal form closed. Result: " & frm.DialogResult.ToString())
    Logger.Info("TestUILib: If you could orbit/pan/zoom while the form was open, the test PASSED.")
End Sub
