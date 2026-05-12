' UILib.vb - Unified UI Management Library
' Provides consistent form creation, non-modal management, and fluid layout helpers
' Layouts adapt to window size - no DPI scaling, content reflows naturally

Public Module UILib
    
    ' ============================================================
    ' CONSTANTS - Default sizes for ~1500px screens
    ' ============================================================
    
    ' Form default sizes (user can resize)
    Public Const FORM_WIDTH_SMALL As Integer = 350
    Public Const FORM_WIDTH_MEDIUM As Integer = 450
    Public Const FORM_WIDTH_LARGE As Integer = 700
    Public Const FORM_HEIGHT_SMALL As Integer = 200
    Public Const FORM_HEIGHT_MEDIUM As Integer = 350
    Public Const FORM_HEIGHT_LARGE As Integer = 500
    
    ' Minimum sizes to prevent unusable layouts
    Public Const FORM_MIN_WIDTH As Integer = 300
    Public Const FORM_MIN_HEIGHT As Integer = 150
    
    ' Control sizing (buttons only - other controls fill available space)
    Public Const BUTTON_WIDTH As Integer = 90
    Public Const BUTTON_HEIGHT As Integer = 28
    
    ' Spacing
    Public Const PADDING As Integer = 10
    Public Const SPACING As Integer = 6
    
    ' ============================================================
    ' NON-MODAL FORM MANAGEMENT
    ' ============================================================
    
    ' Active non-modal forms for cleanup
    Private m_ActiveForms As New List(Of System.Windows.Forms.Form)
    
    ''' <summary>
    ''' Shows a form as non-modal and runs a message loop until closed.
    ''' Allows full Inventor interaction including CommandManager.Pick.
    ''' </summary>
    Public Sub ShowNonModal(frm As System.Windows.Forms.Form)
        frm.TopMost = True
        m_ActiveForms.Add(frm)
        
        AddHandler frm.FormClosed, Sub(s, e)
            m_ActiveForms.Remove(frm)
        End Sub
        
        frm.Show()
        
        Do While frm.Visible
            System.Windows.Forms.Application.DoEvents()
            System.Threading.Thread.Sleep(10)
        Loop
    End Sub
    
    ''' <summary>
    ''' Shows a form as non-modal and returns immediately.
    ''' Caller must manage the message loop or use callbacks.
    ''' </summary>
    Public Sub ShowNonModalAsync(frm As System.Windows.Forms.Form)
        frm.TopMost = True
        m_ActiveForms.Add(frm)
        
        AddHandler frm.FormClosed, Sub(s, e)
            m_ActiveForms.Remove(frm)
        End Sub
        
        frm.Show()
    End Sub
    
    ''' <summary>
    ''' Pumps message queue once. Call repeatedly in picker loops.
    ''' </summary>
    Public Sub PumpMessages()
        System.Windows.Forms.Application.DoEvents()
        System.Threading.Thread.Sleep(5)
    End Sub
    
    ''' <summary>
    ''' Closes all active non-modal forms (cleanup on rule exit).
    ''' </summary>
    Public Sub CloseAllForms()
        For Each frm In m_ActiveForms.ToArray()
            If frm.Visible Then frm.Close()
        Next
        m_ActiveForms.Clear()
    End Sub
    
    ' ============================================================
    ' FORM FACTORY
    ' ============================================================
    
    ''' <summary>
    ''' Creates a standard resizable tool window form.
    ''' Content adapts when user resizes the window.
    ''' Call FinalizeForm() after adding all controls to set proper minimum size.
    ''' </summary>
    Public Function CreateForm(title As String, Optional width As Integer = FORM_WIDTH_MEDIUM, Optional height As Integer = FORM_HEIGHT_MEDIUM) As System.Windows.Forms.Form
        Dim frm As New System.Windows.Forms.Form()
        
        frm.Text = title
        frm.Width = width
        frm.Height = height
        frm.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        frm.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable
        frm.MaximizeBox = False
        frm.MinimizeBox = True
        frm.Padding = New System.Windows.Forms.Padding(PADDING)
        
        Return frm
    End Function
    
    ''' <summary>
    ''' Creates a larger resizable form for dialogs with grids/lists.
    ''' Call FinalizeForm() after adding all controls to set proper minimum size.
    ''' </summary>
    Public Function CreateLargeForm(title As String, Optional width As Integer = FORM_WIDTH_LARGE, Optional height As Integer = FORM_HEIGHT_LARGE) As System.Windows.Forms.Form
        Dim frm As New System.Windows.Forms.Form()
        
        frm.Text = title
        frm.Width = width
        frm.Height = height
        frm.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        frm.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable
        frm.MaximizeBox = True
        frm.MinimizeBox = True
        frm.Padding = New System.Windows.Forms.Padding(PADDING)
        
        Return frm
    End Function
    
    ''' <summary>
    ''' Sets minimum size for a form.
    ''' Avoids System.Drawing.Size which may not work in iLogic.
    ''' Uses closure to capture min values (does not consume Form.Tag).
    ''' </summary>
    Public Sub SetMinimumSize(frm As System.Windows.Forms.Form, minWidth As Integer, minHeight As Integer)
        ' Closure captures minWidth/minHeight - this is safe (not ByRef)
        
        ' Enforce on resize (continuous check)
        AddHandler frm.Resize, Sub(s, e)
            If frm.WindowState = System.Windows.Forms.FormWindowState.Normal Then
                Dim needsAdjust As Boolean = False
                Dim newWidth As Integer = frm.Width
                Dim newHeight As Integer = frm.Height
                
                If frm.Width < minWidth Then
                    newWidth = minWidth
                    needsAdjust = True
                End If
                If frm.Height < minHeight Then
                    newHeight = minHeight
                    needsAdjust = True
                End If
                
                If needsAdjust Then
                    frm.Width = newWidth
                    frm.Height = newHeight
                End If
            End If
        End Sub
        
        ' Also enforce when resize ends (for smoother UX)
        AddHandler frm.ResizeEnd, Sub(s, e)
            If frm.Width < minWidth Then frm.Width = minWidth
            If frm.Height < minHeight Then frm.Height = minHeight
        End Sub
    End Sub
    
    ''' <summary>
    ''' Finalizes the form after all controls are added.
    ''' Calculates minimum height based on actual content.
    ''' Call this AFTER adding all controls to the form.
    ''' </summary>
    Public Sub FinalizeForm(frm As System.Windows.Forms.Form)
        ' Calculate preferred height based on content
        Dim contentHeight As Integer = 0
        
        For Each ctrl As System.Windows.Forms.Control In frm.Controls
            ' For TableLayoutPanel, calculate based on row heights
            If TypeOf ctrl Is System.Windows.Forms.TableLayoutPanel Then
                Dim tlp As System.Windows.Forms.TableLayoutPanel = CType(ctrl, System.Windows.Forms.TableLayoutPanel)
                Dim tlpHeight As Integer = CalculateTableLayoutHeight(tlp)
                If tlpHeight > contentHeight Then contentHeight = tlpHeight
            End If
            
            ' For FlowLayoutPanel (buttons), add its height
            If TypeOf ctrl Is System.Windows.Forms.FlowLayoutPanel Then
                Dim flp As System.Windows.Forms.FlowLayoutPanel = CType(ctrl, System.Windows.Forms.FlowLayoutPanel)
                ' Estimate: button height + padding
                Dim flpHeight As Integer = BUTTON_HEIGHT + SPACING * 2 + flp.Padding.Top + flp.Padding.Bottom
                contentHeight += flpHeight
            End If
            
            ' For other controls, use bounds
            If ctrl.Dock = System.Windows.Forms.DockStyle.None Then
                Dim ctrlBottom As Integer = ctrl.Top + ctrl.Height + ctrl.Margin.Bottom
                If ctrlBottom > contentHeight Then
                    contentHeight = ctrlBottom
                End If
            End If
        Next
        
        ' Add padding and chrome (title bar ~30px, borders ~8px)
        Dim minHeight As Integer = contentHeight + frm.Padding.Top + frm.Padding.Bottom + 40
        If minHeight < FORM_MIN_HEIGHT Then minHeight = FORM_MIN_HEIGHT
        
        ' Ensure form is at least as tall as content
        If frm.Height < minHeight Then
            frm.Height = minHeight
        End If
        
        ' Update minimum size enforcement
        SetMinimumSize(frm, FORM_MIN_WIDTH, minHeight)
    End Sub
    
    ''' <summary>
    ''' Calculates the preferred height for a TableLayoutPanel based on its rows.
    ''' </summary>
    Private Function CalculateTableLayoutHeight(tlp As System.Windows.Forms.TableLayoutPanel) As Integer
        Dim height As Integer = tlp.Padding.Top + tlp.Padding.Bottom
        
        ' Track heights per row to avoid double-counting
        Dim rowHeights As New Dictionary(Of Integer, Integer)
        
        For Each ctrl As System.Windows.Forms.Control In tlp.Controls
            Dim row As Integer = tlp.GetRow(ctrl)
            If row >= 0 Then
                Dim ctrlHeight As Integer = ctrl.Height + ctrl.Margin.Top + ctrl.Margin.Bottom
                
                ' Keep the tallest control in each row
                If Not rowHeights.ContainsKey(row) OrElse rowHeights(row) < ctrlHeight Then
                    rowHeights(row) = ctrlHeight
                End If
            End If
        Next
        
        ' Sum all row heights
        For Each rowHeight In rowHeights.Values
            height += rowHeight
        Next
        
        Return height
    End Function
    
    ''' <summary>
    ''' Creates a fixed-size compact dialog (no resize).
    ''' Use sparingly - prefer resizable forms.
    ''' </summary>
    Public Function CreateCompactForm(title As String, Optional width As Integer = FORM_WIDTH_SMALL, Optional height As Integer = FORM_HEIGHT_SMALL) As System.Windows.Forms.Form
        Dim frm As New System.Windows.Forms.Form()
        
        frm.Text = title
        frm.Width = width
        frm.Height = height
        frm.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        frm.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        frm.MaximizeBox = False
        frm.MinimizeBox = False
        frm.Padding = New System.Windows.Forms.Padding(PADDING)
        
        Return frm
    End Function
    
    ' ============================================================
    ' LAYOUT HELPERS - Fluid layouts that adapt to window size
    ' ============================================================
    
    ''' <summary>
    ''' Creates a TableLayoutPanel for form content with fluid column layout.
    ''' Column 0: Labels (auto-width), Column 1: Controls (fills remaining space)
    ''' </summary>
    Public Function CreateContentPanel() As System.Windows.Forms.TableLayoutPanel
        Dim panel As New System.Windows.Forms.TableLayoutPanel()
        
        panel.Dock = System.Windows.Forms.DockStyle.Fill
        panel.AutoSize = False  ' Fill available space
        panel.ColumnCount = 2
        panel.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.AutoSize))
        panel.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100))
        panel.Padding = New System.Windows.Forms.Padding(0)
        
        Return panel
    End Function
    
    ''' <summary>
    ''' Creates a TableLayoutPanel with three columns: label, control, button.
    ''' Useful for rows with pick buttons.
    ''' </summary>
    Public Function CreateContentPanelWithButtons() As System.Windows.Forms.TableLayoutPanel
        Dim panel As New System.Windows.Forms.TableLayoutPanel()
        
        panel.Dock = System.Windows.Forms.DockStyle.Fill
        panel.AutoSize = False
        panel.ColumnCount = 3
        panel.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.AutoSize))
        panel.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100))
        panel.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.AutoSize))
        panel.Padding = New System.Windows.Forms.Padding(0)
        
        Return panel
    End Function
    
    ''' <summary>
    ''' Creates a FlowLayoutPanel for button bars (right-aligned, wraps on narrow windows).
    ''' </summary>
    Public Function CreateButtonPanel() As System.Windows.Forms.FlowLayoutPanel
        Dim panel As New System.Windows.Forms.FlowLayoutPanel()
        
        panel.Dock = System.Windows.Forms.DockStyle.Bottom
        panel.FlowDirection = System.Windows.Forms.FlowDirection.RightToLeft
        panel.WrapContents = True  ' Buttons wrap to next line if window too narrow
        panel.AutoSize = True
        panel.Padding = New System.Windows.Forms.Padding(0, SPACING, 0, 0)
        
        Return panel
    End Function
    
    ''' <summary>
    ''' Adds a label+control row to a TableLayoutPanel.
    ''' Control stretches to fill available width.
    ''' </summary>
    Public Sub AddRow(panel As System.Windows.Forms.TableLayoutPanel, labelText As String, control As System.Windows.Forms.Control, Optional controlName As String = Nothing)
        Dim rowIndex As Integer = panel.RowCount
        panel.RowCount += 1
        panel.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
        
        Dim lbl As New System.Windows.Forms.Label()
        lbl.Text = labelText
        lbl.AutoSize = True
        lbl.Anchor = System.Windows.Forms.AnchorStyles.Left
        lbl.Margin = New System.Windows.Forms.Padding(0, 6, SPACING, 0)
        
        control.Anchor = System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right
        control.Margin = New System.Windows.Forms.Padding(0, 3, 0, 3)
        If controlName IsNot Nothing Then control.Name = controlName
        
        panel.Controls.Add(lbl, 0, rowIndex)
        panel.Controls.Add(control, 1, rowIndex)
    End Sub
    
    ''' <summary>
    ''' Adds a label+control+button row (for pick buttons).
    ''' </summary>
    Public Sub AddRowWithButton(panel As System.Windows.Forms.TableLayoutPanel, labelText As String, control As System.Windows.Forms.Control, button As System.Windows.Forms.Button, Optional controlName As String = Nothing)
        Dim rowIndex As Integer = panel.RowCount
        panel.RowCount += 1
        panel.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
        
        Dim lbl As New System.Windows.Forms.Label()
        lbl.Text = labelText
        lbl.AutoSize = True
        lbl.Anchor = System.Windows.Forms.AnchorStyles.Left
        lbl.Margin = New System.Windows.Forms.Padding(0, 6, SPACING, 0)
        
        control.Anchor = System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right
        control.Margin = New System.Windows.Forms.Padding(0, 3, SPACING, 3)
        If controlName IsNot Nothing Then control.Name = controlName
        
        button.Margin = New System.Windows.Forms.Padding(0, 3, 0, 3)
        
        panel.Controls.Add(lbl, 0, rowIndex)
        panel.Controls.Add(control, 1, rowIndex)
        panel.Controls.Add(button, 2, rowIndex)
    End Sub
    
    ''' <summary>
    ''' Adds a full-width control row (spans all columns).
    ''' Use for section headers, checkboxes, or wide controls.
    ''' </summary>
    Public Sub AddFullWidthRow(panel As System.Windows.Forms.TableLayoutPanel, control As System.Windows.Forms.Control, Optional controlName As String = Nothing)
        Dim rowIndex As Integer = panel.RowCount
        panel.RowCount += 1
        panel.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
        
        control.Anchor = System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right
        control.Margin = New System.Windows.Forms.Padding(0, 3, 0, 3)
        If controlName IsNot Nothing Then control.Name = controlName
        
        panel.Controls.Add(control, 0, rowIndex)
        panel.SetColumnSpan(control, panel.ColumnCount)
    End Sub
    
    ''' <summary>
    ''' Adds a fill row for controls that should expand vertically (ListBox, DataGridView).
    ''' </summary>
    Public Sub AddFillRow(panel As System.Windows.Forms.TableLayoutPanel, control As System.Windows.Forms.Control, Optional controlName As String = Nothing)
        Dim rowIndex As Integer = panel.RowCount
        panel.RowCount += 1
        panel.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100))
        
        control.Dock = System.Windows.Forms.DockStyle.Fill
        control.Margin = New System.Windows.Forms.Padding(0, 3, 0, 3)
        If controlName IsNot Nothing Then control.Name = controlName
        
        panel.Controls.Add(control, 0, rowIndex)
        panel.SetColumnSpan(control, panel.ColumnCount)
    End Sub
    
    ''' <summary>
    ''' Adds a section header label (visual separator, full width).
    ''' </summary>
    Public Sub AddSectionHeader(panel As System.Windows.Forms.TableLayoutPanel, text As String)
        Dim lbl As New System.Windows.Forms.Label()
        lbl.Text = "— " & text & " —"
        lbl.AutoSize = True
        lbl.Margin = New System.Windows.Forms.Padding(0, SPACING * 2, 0, SPACING)
        
        AddFullWidthRow(panel, lbl)
    End Sub
    
    ' ============================================================
    ' CONTROL FACTORY
    ' ============================================================
    
    Public Function CreateLabel(text As String) As System.Windows.Forms.Label
        Dim lbl As New System.Windows.Forms.Label()
        lbl.Text = text
        lbl.AutoSize = True
        Return lbl
    End Function
    
    Public Function CreateTextBox(Optional text As String = "") As System.Windows.Forms.TextBox
        Dim txt As New System.Windows.Forms.TextBox()
        txt.Text = text
        Return txt
    End Function
    
    Public Function CreateMultilineTextBox(Optional text As String = "", Optional minLines As Integer = 3) As System.Windows.Forms.TextBox
        Dim txt As New System.Windows.Forms.TextBox()
        txt.Text = text
        txt.Multiline = True
        txt.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        ' Use approximate line height (avoid System.Drawing.Font reference)
        txt.Height = 16 * minLines + 8
        Return txt
    End Function
    
    Public Function CreateComboBox(items() As String, Optional selectedIndex As Integer = 0) As System.Windows.Forms.ComboBox
        Dim cbo As New System.Windows.Forms.ComboBox()
        cbo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        cbo.Items.AddRange(items)
        If items.Length > selectedIndex Then cbo.SelectedIndex = selectedIndex
        Return cbo
    End Function
    
    Public Function CreateCheckBox(text As String, Optional checked As Boolean = False) As System.Windows.Forms.CheckBox
        Dim chk As New System.Windows.Forms.CheckBox()
        chk.Text = text
        chk.Checked = checked
        chk.AutoSize = True
        Return chk
    End Function
    
    Public Function CreateNumericUpDown(min As Decimal, max As Decimal, value As Decimal, Optional decimalPlaces As Integer = 0) As System.Windows.Forms.NumericUpDown
        Dim nud As New System.Windows.Forms.NumericUpDown()
        nud.Minimum = min
        nud.Maximum = max
        nud.Value = value
        nud.DecimalPlaces = decimalPlaces
        Return nud
    End Function
    
    Public Function CreateButton(text As String, Optional width As Integer = BUTTON_WIDTH) As System.Windows.Forms.Button
        Dim btn As New System.Windows.Forms.Button()
        btn.Text = text
        btn.Width = width
        btn.Height = BUTTON_HEIGHT
        btn.AutoSize = False
        Return btn
    End Function
    
    Public Function CreateSmallButton(text As String) As System.Windows.Forms.Button
        Dim btn As New System.Windows.Forms.Button()
        btn.Text = text
        btn.AutoSize = True
        btn.Padding = New System.Windows.Forms.Padding(4, 2, 4, 2)
        Return btn
    End Function
    
    Public Function CreateListBox() As System.Windows.Forms.ListBox
        Dim lst As New System.Windows.Forms.ListBox()
        lst.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended
        lst.IntegralHeight = False  ' Allow partial row display for fluid sizing
        Return lst
    End Function
    
    Public Function CreateDataGridView() As System.Windows.Forms.DataGridView
        Dim dgv As New System.Windows.Forms.DataGridView()
        dgv.AllowUserToAddRows = False
        dgv.AllowUserToDeleteRows = False
        dgv.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        dgv.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        dgv.RowHeadersVisible = False
        Return dgv
    End Function
    
    ' ============================================================
    ' PICKER INTEGRATION (Non-Modal Compatible)
    ' ============================================================
    
    ''' <summary>
    ''' Performs a pick operation while keeping the form visible.
    ''' Returns Nothing on cancel/ESC.
    ''' app should be Inventor.Application, filter should be SelectionFilterEnum value.
    ''' </summary>
    Public Function PickWithForm(app As Object, frm As System.Windows.Forms.Form, filter As Object, prompt As String) As Object
        ' Temporarily allow Inventor to receive focus
        Dim wasTopMost As Boolean = frm.TopMost
        frm.TopMost = False
        
        Dim result As Object = Nothing
        
        Try
            ' Pump messages to keep form responsive
            PumpMessages()
            
            ' Perform pick
            result = app.CommandManager.Pick(filter, prompt)
        Catch
            ' User cancelled with ESC
            result = Nothing
        End Try
        
        ' Restore form state
        frm.TopMost = wasTopMost
        frm.Focus()
        
        Return result
    End Function
    
    ''' <summary>
    ''' Performs a pick operation with pre-defined filter constants.
    ''' filterType: "point", "axis", "plane", "face", "edge", "occurrence"
    ''' </summary>
    Public Function PickWithFormByType(app As Object, frm As System.Windows.Forms.Form, filterType As String, prompt As String) As Object
        Dim filter As Integer
        
        Select Case filterType.ToLower()
            Case "point"
                filter = 262144  ' kWorkPointFilter
            Case "axis"
                filter = 131072  ' kWorkAxisFilter
            Case "plane"
                filter = 65536   ' kWorkPlaneFilter
            Case "face"
                filter = 32776   ' kPartFacePlanarFilter
            Case "edge"
                filter = 1       ' kPartEdgeFilter
            Case "occurrence"
                filter = 16      ' kAssemblyOccurrenceFilter
            Case "planar"
                filter = 32768   ' kAllPlanarEntities
            Case "linear"
                filter = 16384   ' kAllLinearEntities
            Case Else
                filter = 0       ' kAllEntitiesFilter
        End Select
        
        Return PickWithForm(app, frm, filter, prompt)
    End Function
    
    ''' <summary>
    ''' Performs multiple pick operations until ESC is pressed.
    ''' Returns list of picked objects (empty list if cancelled immediately).
    ''' </summary>
    Public Function MultiPickWithForm(app As Object, frm As System.Windows.Forms.Form, filter As Object, prompt As String) As List(Of Object)
        Dim results As New List(Of Object)
        Dim wasTopMost As Boolean = frm.TopMost
        frm.TopMost = False
        
        Dim keepPicking As Boolean = True
        Do While keepPicking
            PumpMessages()
            
            Try
                Dim obj As Object = app.CommandManager.Pick(filter, prompt)
                If obj IsNot Nothing Then
                    results.Add(obj)
                End If
            Catch
                ' ESC pressed - done picking
                keepPicking = False
            End Try
        Loop
        
        frm.TopMost = wasTopMost
        frm.Focus()
        
        Return results
    End Function
    
End Module
