' TestStringsLib.vb - Test script for StringsLib functionality
' Run this to verify Phase 2 of the Unified UI Library implementation
'
' Tests:
' 1. Document guard messages
' 2. Button labels
' 3. Picker prompts
' 4. Format helpers
' 5. Combined UILib + StringsLib usage

AddVbFile "Lib/UILib.vb"
AddVbFile "Lib/StringsLib.vb"

Sub Main()
    Logger.Info("TestStringsLib: Starting StringsLib verification tests...")
    
    ' Test 1: Show all document guard messages
    TestDocumentGuards()
    
    ' Test 2: Show dialog with Estonian buttons and labels
    TestEstonianDialog()
    
    ' Test 3: Show picker prompt formats
    TestPickerPrompts()
    
    Logger.Info("TestStringsLib: All tests completed.")
End Sub

Sub TestDocumentGuards()
    Logger.Info("TestStringsLib: Test 1 - Document guard messages...")
    
    Dim frm As System.Windows.Forms.Form = UILib.CreateForm("Test: Document Guards", 500, 350)
    Dim content As System.Windows.Forms.TableLayoutPanel = UILib.CreateContentPanel()
    Dim buttons As System.Windows.Forms.FlowLayoutPanel = UILib.CreateButtonPanel()
    
    UILib.AddSectionHeader(content, "Document Guard Messages")
    
    ' Show all guard messages
    UILib.AddRow(content, "No document:", UILib.CreateLabel(StringsLib.MSG_NO_ACTIVE_DOCUMENT))
    UILib.AddRow(content, "Requires assembly:", UILib.CreateLabel(StringsLib.MSG_REQUIRES_ASSEMBLY))
    UILib.AddRow(content, "Requires part:", UILib.CreateLabel(StringsLib.MSG_REQUIRES_PART))
    UILib.AddRow(content, "Requires drawing:", UILib.CreateLabel(StringsLib.MSG_REQUIRES_DRAWING))
    UILib.AddRow(content, "Assembly or part:", UILib.CreateLabel(StringsLib.MSG_REQUIRES_ASSEMBLY_OR_PART))
    
    UILib.AddSectionHeader(content, "Common Messages")
    UILib.AddRow(content, "No selection:", UILib.CreateLabel(StringsLib.MSG_NO_SELECTION))
    UILib.AddRow(content, "Cancelled:", UILib.CreateLabel(StringsLib.MSG_OPERATION_CANCELLED))
    UILib.AddRow(content, "Complete:", UILib.CreateLabel(StringsLib.MSG_OPERATION_COMPLETE))
    UILib.AddRow(content, "Failed:", UILib.CreateLabel(StringsLib.MSG_OPERATION_FAILED))
    
    Dim btnClose As System.Windows.Forms.Button = UILib.CreateButton(StringsLib.BTN_CLOSE)
    btnClose.DialogResult = System.Windows.Forms.DialogResult.OK
    buttons.Controls.Add(btnClose)
    
    frm.Controls.Add(content)
    frm.Controls.Add(buttons)
    frm.AcceptButton = btnClose
    
    UILib.FinalizeForm(frm)
    frm.ShowDialog()
End Sub

Sub TestEstonianDialog()
    Logger.Info("TestStringsLib: Test 2 - Estonian dialog with all button types...")
    
    Dim frm As System.Windows.Forms.Form = UILib.CreateForm(StringsLib.FormatDialogTitle("Test", "Estonian Buttons"), 450, 400)
    Dim content As System.Windows.Forms.TableLayoutPanel = UILib.CreateContentPanel()
    Dim buttons As System.Windows.Forms.FlowLayoutPanel = UILib.CreateButtonPanel()
    
    UILib.AddSectionHeader(content, "Form Labels")
    UILib.AddRow(content, StringsLib.LBL_NAME, UILib.CreateTextBox("Näidis nimi"))
    UILib.AddRow(content, StringsLib.LBL_DESCRIPTION, UILib.CreateTextBox("Kirjelduse tekst"))
    UILib.AddRow(content, StringsLib.LBL_VALUE, UILib.CreateNumericUpDown(0, 100, 50))
    UILib.AddRow(content, StringsLib.LBL_COUNT, UILib.CreateNumericUpDown(1, 10, 1))
    UILib.AddRow(content, StringsLib.LBL_MATERIAL, UILib.CreateComboBox(New String() {"Puit", "Metall", "Plast"}))
    
    UILib.AddSectionHeader(content, "Dimensions")
    UILib.AddRow(content, StringsLib.LBL_WIDTH, UILib.CreateNumericUpDown(0, 1000, 100, 1))
    UILib.AddRow(content, StringsLib.LBL_HEIGHT, UILib.CreateNumericUpDown(0, 1000, 50, 1))
    UILib.AddRow(content, StringsLib.LBL_THICKNESS, UILib.CreateNumericUpDown(0, 100, 5, 1))
    
    ' Multiple button types
    Dim btnOK As System.Windows.Forms.Button = UILib.CreateButton(StringsLib.BTN_OK)
    btnOK.DialogResult = System.Windows.Forms.DialogResult.OK
    Dim btnCancel As System.Windows.Forms.Button = UILib.CreateButton(StringsLib.BTN_CANCEL)
    btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
    Dim btnApply As System.Windows.Forms.Button = UILib.CreateButton(StringsLib.BTN_APPLY)
    
    buttons.Controls.Add(btnCancel)
    buttons.Controls.Add(btnApply)
    buttons.Controls.Add(btnOK)
    
    frm.Controls.Add(content)
    frm.Controls.Add(buttons)
    frm.AcceptButton = btnOK
    frm.CancelButton = btnCancel
    
    UILib.FinalizeForm(frm)
    
    Dim result As System.Windows.Forms.DialogResult = frm.ShowDialog()
    Logger.Info("TestStringsLib: Dialog closed with: " & result.ToString())
End Sub

Sub TestPickerPrompts()
    Logger.Info("TestStringsLib: Test 3 - Picker prompt formats...")
    
    Dim frm As System.Windows.Forms.Form = UILib.CreateForm("Test: Picker Prompts", 500, 400)
    Dim content As System.Windows.Forms.TableLayoutPanel = UILib.CreateContentPanel()
    Dim buttons As System.Windows.Forms.FlowLayoutPanel = UILib.CreateButtonPanel()
    
    UILib.AddSectionHeader(content, "Standard Picker Prompts")
    UILib.AddRow(content, "Point:", UILib.CreateLabel(StringsLib.PICK_POINT))
    UILib.AddRow(content, "Axis:", UILib.CreateLabel(StringsLib.PICK_AXIS))
    UILib.AddRow(content, "Plane:", UILib.CreateLabel(StringsLib.PICK_PLANE))
    UILib.AddRow(content, "Face:", UILib.CreateLabel(StringsLib.PICK_FACE))
    UILib.AddRow(content, "Edge:", UILib.CreateLabel(StringsLib.PICK_EDGE))
    UILib.AddRow(content, "Component:", UILib.CreateLabel(StringsLib.PICK_COMPONENT))
    UILib.AddRow(content, "Occurrence:", UILib.CreateLabel(StringsLib.PICK_OCCURRENCE))
    
    UILib.AddSectionHeader(content, "Custom Prompts (FormatPickPrompt)")
    UILib.AddRow(content, "Start point:", UILib.CreateLabel(StringsLib.FormatPickPrompt("Vali alguspunkt")))
    UILib.AddRow(content, "End point:", UILib.CreateLabel(StringsLib.FormatPickPrompt("Vali lõpp-punkt")))
    UILib.AddRow(content, "A-side face:", UILib.CreateLabel(StringsLib.FormatPickPrompt("Vali A-külje pind")))
    
    UILib.AddSectionHeader(content, "Format Helpers")
    UILib.AddRow(content, "FormatCount:", UILib.CreateLabel(StringsLib.FormatCount(5, "elementi")))
    UILib.AddRow(content, "FormatSelectedCount:", UILib.CreateLabel(StringsLib.FormatSelectedCount(3)))
    
    Dim btnClose As System.Windows.Forms.Button = UILib.CreateButton(StringsLib.BTN_CLOSE)
    btnClose.DialogResult = System.Windows.Forms.DialogResult.OK
    buttons.Controls.Add(btnClose)
    
    frm.Controls.Add(content)
    frm.Controls.Add(buttons)
    frm.AcceptButton = btnClose
    
    UILib.FinalizeForm(frm)
    frm.ShowDialog()
End Sub
