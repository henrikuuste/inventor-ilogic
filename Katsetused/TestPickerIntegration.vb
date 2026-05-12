' TestPickerIntegration.vb - Test script for picker integration
' Phase 4 verification: Picker operations with non-modal forms
'
' Tests:
' 1. PickWithForm - single pick while form visible
' 2. MultiPickWithForm - multiple picks until ESC
' 3. PickWithFormByType - type-based filter selection
' 4. UtilsLib Estonian prompts (PickPointET, etc.)

AddVbFile "Lib/UILib.vb"
AddVbFile "Lib/StringsLib.vb"
AddVbFile "Lib/UtilsLib.vb"
AddVbFile "Lib/ViewportHelperLib.vb"

Sub Main()
    Dim app As Inventor.Application = ThisApplication
    Dim doc As Document = app.ActiveDocument
    
    ' Initialize logging
    UtilsLib.SetLogger(Logger)
    
    Logger.Info("=== Testing Picker Integration (Phase 4) ===")
    
    ' Create main test form
    Dim frm As System.Windows.Forms.Form = UILib.CreateForm(StringsLib.FormatDialogTitle("Picker Test"))
    Dim content As System.Windows.Forms.TableLayoutPanel = UILib.CreateContentPanel()
    frm.Controls.Add(content)
    
    ' Status label
    Dim lblStatus As System.Windows.Forms.Label = UILib.CreateLabel("Valitud: (midagi pole)")
    lblStatus.Dock = System.Windows.Forms.DockStyle.Fill
    UILib.AddFullWidthRow(content, lblStatus)
    
    ' List for multi-pick results
    Dim lstPicked As System.Windows.Forms.ListBox = UILib.CreateListBox()
    lstPicked.Height = 100
    lstPicked.Dock = System.Windows.Forms.DockStyle.Fill
    UILib.AddFullWidthRow(content, lstPicked)
    
    ' Buttons panel
    Dim buttons As System.Windows.Forms.FlowLayoutPanel = UILib.CreateButtonPanel()
    
    ' Single pick buttons
    Dim btnPickPoint As System.Windows.Forms.Button = UILib.CreateButton("Vali Punkt")
    AddHandler btnPickPoint.Click, Sub()
        Logger.Info("Starting single point pick...")
        Dim picked As Object = UILib.PickWithFormByType(app, frm, "point", StringsLib.PICK_POINT)
        If picked IsNot Nothing Then
            Dim name As String = UtilsLib.GetObjectDisplayName(picked)
            lblStatus.Text = "Valitud: " & name
            Logger.Info("Picked: " & name)
        Else
            lblStatus.Text = "Valitud: (tühistatud)"
            Logger.Info("Pick cancelled")
        End If
    End Sub
    buttons.Controls.Add(btnPickPoint)
    
    Dim btnPickPlane As System.Windows.Forms.Button = UILib.CreateButton("Vali Tasand")
    AddHandler btnPickPlane.Click, Sub()
        Logger.Info("Starting single plane pick...")
        Dim picked As Object = UILib.PickWithFormByType(app, frm, "plane", StringsLib.PICK_PLANE)
        If picked IsNot Nothing Then
            Dim name As String = UtilsLib.GetObjectDisplayName(picked)
            lblStatus.Text = "Valitud: " & name
            Logger.Info("Picked: " & name)
        Else
            lblStatus.Text = "Valitud: (tühistatud)"
            Logger.Info("Pick cancelled")
        End If
    End Sub
    buttons.Controls.Add(btnPickPlane)
    
    Dim btnPickFace As System.Windows.Forms.Button = UILib.CreateButton("Vali Pind")
    AddHandler btnPickFace.Click, Sub()
        Logger.Info("Starting single face pick...")
        Dim picked As Object = UILib.PickWithFormByType(app, frm, "face", StringsLib.PICK_FACE)
        If picked IsNot Nothing Then
            Dim name As String = UtilsLib.GetObjectDisplayName(picked)
            lblStatus.Text = "Valitud: " & name
            Logger.Info("Picked: " & name)
        Else
            lblStatus.Text = "Valitud: (tühistatud)"
            Logger.Info("Pick cancelled")
        End If
    End Sub
    buttons.Controls.Add(btnPickFace)
    
    ' Multi-pick button
    Dim btnMultiPick As System.Windows.Forms.Button = UILib.CreateButton("Mitu Elementi")
    AddHandler btnMultiPick.Click, Sub()
        Logger.Info("Starting multi-pick (ESC to finish)...")
        lstPicked.Items.Clear()
        
        ' Use kAllEntitiesFilter (0) for multi-pick demo
        Dim picked As List(Of Object) = UILib.MultiPickWithForm(app, frm, 0, StringsLib.FormatPickPrompt("Vali elemendid"))
        
        For Each obj As Object In picked
            Dim name As String = UtilsLib.GetObjectDisplayName(obj)
            lstPicked.Items.Add(name)
            Logger.Info("Multi-picked: " & name)
        Next
        
        lblStatus.Text = StringsLib.FormatSelectedCount(picked.Count)
        Logger.Info("Multi-pick complete: " & picked.Count & " items")
    End Sub
    buttons.Controls.Add(btnMultiPick)
    
    ' Clear button
    Dim btnClear As System.Windows.Forms.Button = UILib.CreateButton(StringsLib.BTN_CLEAR)
    AddHandler btnClear.Click, Sub()
        lstPicked.Items.Clear()
        lblStatus.Text = "Valitud: (midagi pole)"
        Logger.Info("Cleared results")
    End Sub
    buttons.Controls.Add(btnClear)
    
    ' Test UtilsLib Estonian pickers
    Dim btnUtilsPick As System.Windows.Forms.Button = UILib.CreateButton("UtilsLib Test")
    AddHandler btnUtilsPick.Click, Sub()
        Logger.Info("Testing UtilsLib Estonian pickers...")
        
        ' Test PickPointET
        Logger.Info("PickPointET - click point or ESC...")
        Dim pt As Object = UtilsLib.PickPointET(app)
        If pt IsNot Nothing Then
            lblStatus.Text = "UtilsLib: " & UtilsLib.GetObjectDisplayName(pt)
            Logger.Info("PickPointET result: " & UtilsLib.GetObjectDisplayName(pt))
        Else
            lblStatus.Text = "UtilsLib: (tühistatud)"
            Logger.Info("PickPointET cancelled")
        End If
    End Sub
    buttons.Controls.Add(btnUtilsPick)
    
    ' Close button
    Dim btnClose As System.Windows.Forms.Button = UILib.CreateButton(StringsLib.BTN_CLOSE)
    AddHandler btnClose.Click, Sub()
        frm.Close()
        Logger.Info("Test form closed")
    End Sub
    buttons.Controls.Add(btnClose)
    
    UILib.AddFullWidthRow(content, buttons)
    
    ' Finalize and show
    UILib.FinalizeForm(frm)
    
    Logger.Info("")
    Logger.Info("=== Picker Integration Test Instructions ===")
    Logger.Info("1. Click 'Vali Punkt' - pick a work point, result shows in status")
    Logger.Info("2. Click 'Vali Tasand' - pick a work plane or planar face")
    Logger.Info("3. Click 'Vali Pind' - pick a planar face")
    Logger.Info("4. Click 'Mitu Elementi' - pick multiple items, ESC to finish")
    Logger.Info("5. Click 'UtilsLib Test' - tests Estonian prompt pickers")
    Logger.Info("6. All picks should work while form is visible")
    Logger.Info("7. ESC should cancel picks gracefully")
    Logger.Info("")
    
    UILib.ShowNonModal(frm)
    
    Logger.Info("=== Picker Integration Tests Complete ===")
End Sub
