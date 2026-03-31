' ============================================================================
' BoundingBoxStockLib - Shared Library for Bounding Box Stock Calculations
' 
' This module contains shared functions used by both:
' - BoundingBoxStock.vb (standalone part processing)
' - BoundingBoxStockBatch.vb (batch processing from assembly)
'
' Usage: AddVbFile "BoundingBoxStockLib.vb"
' ============================================================================

Imports Inventor
Imports System.Windows.Forms

Public Module BoundingBoxStockLib

    ' ============================================================================
    ' Main processing function - called by both standalone and batch scripts
    ' Returns: "OK" = success, "SKIP" = user skipped, "CANCEL" = user cancelled
    ' iLogicAuto: Pass iLogicVb.Automation from the calling script
    ' ============================================================================
    Public Function ProcessPartDocument(ByVal app As Inventor.Application, ByVal partDoc As PartDocument, _
                                        ByVal formTitle As String, ByVal showSkipButton As Boolean, _
                                        ByVal iLogicAuto As Object) As String
        ' Get bounding box dimensions
        Dim xSize As Double = 0
        Dim ySize As Double = 0
        Dim zSize As Double = 0
        GetBoundingBoxSizes(partDoc, xSize, ySize, zSize)

        ' Run the UI loop
        Dim axisConfig As String = RunConfigLoop(app, partDoc, xSize, ySize, zSize, formTitle, showSkipButton)

        If axisConfig = "" Then
            Return "CANCEL" ' User cancelled
        ElseIf axisConfig = "SKIP" Then
            Return "SKIP" ' User skipped this part
        End If

        ' Parse the axis config (format: "T:Z,W:X,L:Y")
        Dim thicknessAxis As String = ""
        Dim widthAxis As String = ""
        Dim lengthAxis As String = ""
        ParseAxisConfig(axisConfig, thicknessAxis, widthAxis, lengthAxis)

        ' Create/update the iProperties and document rule
        CreateOrUpdateRule(partDoc, thicknessAxis, widthAxis, lengthAxis, iLogicAuto)

        Return "OK"
    End Function

    Public Sub GetBoundingBoxSizes(ByVal partDoc As PartDocument, ByRef xSize As Double, ByRef ySize As Double, ByRef zSize As Double)
        Dim rangebox As Box = partDoc.ComponentDefinition.RangeBox
        xSize = rangebox.MaxPoint.X - rangebox.MinPoint.X
        ySize = rangebox.MaxPoint.Y - rangebox.MinPoint.Y
        zSize = rangebox.MaxPoint.Z - rangebox.MinPoint.Z
    End Sub

    Public Function RunConfigLoop(ByVal app As Inventor.Application, ByVal partDoc As PartDocument, _
                                  ByVal xSize As Double, ByVal ySize As Double, ByVal zSize As Double, _
                                  ByVal formTitle As String, ByVal showSkipButton As Boolean) As String
        ' Try to read existing axis configuration from iProperties
        Dim thicknessAxis As String = GetCustomPropertyValue(partDoc, "BB_ThicknessAxis", "")
        Dim widthAxis As String = GetCustomPropertyValue(partDoc, "BB_WidthAxis", "")
        Dim lengthAxis As String = ""
        Dim customAxisDesc As String = "" ' Description when custom axis is picked

        ' If no existing config, use defaults: Z = Thickness, then Width = longer of X/Y
        If thicknessAxis = "" Then
            thicknessAxis = "Z"
            AssignWidthLength(thicknessAxis, xSize, ySize, zSize, widthAxis, lengthAxis)
        Else
            ' Length is the remaining axis (determined by thickness and width)
            lengthAxis = GetRemainingAxis(thicknessAxis, widthAxis)
        End If

        Dim keepGoing As Boolean = True
        Do While keepGoing
            Dim action As String = ""
            Dim result As DialogResult = ShowConfigForm(app, partDoc, xSize, ySize, zSize, _
                                                         thicknessAxis, widthAxis, lengthAxis, _
                                                         customAxisDesc, formTitle, showSkipButton, action)

            If result = DialogResult.Cancel Then
                Return ""
            ElseIf result = DialogResult.OK Then
                Return "T:" & thicknessAxis & ",W:" & widthAxis & ",L:" & lengthAxis
            ElseIf result = DialogResult.No Then
                ' Skip button pressed (using No for skip)
                Return "SKIP"
            ElseIf action = "PICK_THICKNESS" Then
                Dim pickedAxis As String = PickAxisWithDescription(app, customAxisDesc)
                If pickedAxis <> "" Then
                    thicknessAxis = pickedAxis
                    AssignWidthLength(thicknessAxis, xSize, ySize, zSize, widthAxis, lengthAxis)
                End If
            ElseIf action = "SELECT_AXIS" Then
                customAxisDesc = "" ' Clear custom when selecting standard axis
                AssignWidthLength(thicknessAxis, xSize, ySize, zSize, widthAxis, lengthAxis)
            ElseIf action = "FLIP" Then
                Dim temp As String = widthAxis
                widthAxis = lengthAxis
                lengthAxis = temp
            End If
        Loop

        Return ""
    End Function

    Public Sub AssignWidthLength(ByVal thicknessAxis As String, ByVal xSize As Double, ByVal ySize As Double, ByVal zSize As Double, _
                                 ByRef widthAxis As String, ByRef lengthAxis As String)
        Dim axis1 As String = ""
        Dim axis2 As String = ""
        Dim size1 As Double = 0
        Dim size2 As Double = 0

        If thicknessAxis = "X" Then
            axis1 = "Y" : size1 = ySize
            axis2 = "Z" : size2 = zSize
        ElseIf thicknessAxis = "Y" Then
            axis1 = "X" : size1 = xSize
            axis2 = "Z" : size2 = zSize
        Else ' Z
            axis1 = "X" : size1 = xSize
            axis2 = "Y" : size2 = ySize
        End If

        If size1 >= size2 Then
            widthAxis = axis1
            lengthAxis = axis2
        Else
            widthAxis = axis2
            lengthAxis = axis1
        End If
    End Sub

    Public Function ShowConfigForm(ByVal app As Inventor.Application, ByVal partDoc As PartDocument, _
                                   ByVal xSize As Double, ByVal ySize As Double, ByVal zSize As Double, _
                                   ByRef thicknessAxis As String, ByVal widthAxis As String, ByVal lengthAxis As String, _
                                   ByVal customAxisDesc As String, ByVal formTitle As String, ByVal showSkipButton As Boolean, _
                                   ByRef action As String) As DialogResult
        action = ""
        Dim uom As UnitsOfMeasure = partDoc.UnitsOfMeasure

        Dim thicknessValue As Double = GetAxisSize(thicknessAxis, xSize, ySize, zSize)
        Dim widthValue As Double = GetAxisSize(widthAxis, xSize, ySize, zSize)
        Dim lengthValue As Double = GetAxisSize(lengthAxis, xSize, ySize, zSize)

        Dim thicknessStr As String = uom.GetStringFromValue(thicknessValue, uom.LengthUnits)
        Dim widthStr As String = uom.GetStringFromValue(widthValue, uom.LengthUnits)
        Dim lengthStr As String = uom.GetStringFromValue(lengthValue, uom.LengthUnits)

        Dim isCustomPick As Boolean = (customAxisDesc <> "")

        Dim frm As New System.Windows.Forms.Form()
        If formTitle <> "" Then
            frm.Text = formTitle
        Else
            frm.Text = "Bounding Box Stock Sizes"
        End If
        frm.Width = 420
        frm.Height = 320
        frm.StartPosition = FormStartPosition.CenterScreen
        frm.FormBorderStyle = FormBorderStyle.FixedDialog
        frm.MaximizeBox = False
        frm.MinimizeBox = False

        Dim lblThicknessAxis As New System.Windows.Forms.Label()
        lblThicknessAxis.Text = "Thickness Axis:"
        lblThicknessAxis.Left = 20
        lblThicknessAxis.Top = 20
        lblThicknessAxis.Width = 100
        frm.Controls.Add(lblThicknessAxis)

        Dim cboAxis As New System.Windows.Forms.ComboBox()
        cboAxis.Name = "cboAxis"
        cboAxis.Left = 125
        cboAxis.Top = 17
        cboAxis.Width = 80
        cboAxis.DropDownStyle = ComboBoxStyle.DropDownList
        cboAxis.Items.Add("X")
        cboAxis.Items.Add("Y")
        cboAxis.Items.Add("Z")
        cboAxis.Items.Add("Custom")
        If isCustomPick Then
            cboAxis.SelectedItem = "Custom"
        Else
            cboAxis.SelectedItem = thicknessAxis
        End If
        frm.Controls.Add(cboAxis)

        Dim txtCustomAxis As New System.Windows.Forms.TextBox()
        txtCustomAxis.Name = "txtCustomAxis"
        txtCustomAxis.Left = 215
        txtCustomAxis.Top = 17
        txtCustomAxis.Width = 170
        txtCustomAxis.ReadOnly = True
        If isCustomPick Then
            txtCustomAxis.Text = customAxisDesc
        Else
            txtCustomAxis.Text = ""
        End If
        frm.Controls.Add(txtCustomAxis)

        Dim lblThickness As New System.Windows.Forms.Label()
        lblThickness.Text = "Thickness:"
        lblThickness.Left = 20
        lblThickness.Top = 60
        lblThickness.Width = 100
        frm.Controls.Add(lblThickness)

        Dim txtThickness As New System.Windows.Forms.TextBox()
        txtThickness.Text = thicknessStr
        txtThickness.Left = 125
        txtThickness.Top = 57
        txtThickness.Width = 100
        txtThickness.ReadOnly = True
        frm.Controls.Add(txtThickness)

        Dim lblWidth As New System.Windows.Forms.Label()
        lblWidth.Text = "Width (" & widthAxis & " axis):"
        lblWidth.Left = 20
        lblWidth.Top = 100
        lblWidth.Width = 100
        frm.Controls.Add(lblWidth)

        Dim txtWidth As New System.Windows.Forms.TextBox()
        txtWidth.Text = widthStr
        txtWidth.Left = 125
        txtWidth.Top = 97
        txtWidth.Width = 100
        txtWidth.ReadOnly = True
        frm.Controls.Add(txtWidth)

        Dim lblLength As New System.Windows.Forms.Label()
        lblLength.Text = "Length (" & lengthAxis & " axis):"
        lblLength.Left = 20
        lblLength.Top = 140
        lblLength.Width = 100
        frm.Controls.Add(lblLength)

        Dim txtLength As New System.Windows.Forms.TextBox()
        txtLength.Text = lengthStr
        txtLength.Left = 125
        txtLength.Top = 137
        txtLength.Width = 100
        txtLength.ReadOnly = True
        frm.Controls.Add(txtLength)

        Dim btnFlip As New System.Windows.Forms.Button()
        btnFlip.Text = "Flip Width/Length"
        btnFlip.Left = 235
        btnFlip.Top = 117
        btnFlip.Width = 120
        btnFlip.Height = 28
        btnFlip.DialogResult = DialogResult.Ignore
        frm.Controls.Add(btnFlip)

        If showSkipButton Then
            Dim btnSkip As New System.Windows.Forms.Button()
            btnSkip.Text = "Skip"
            btnSkip.Left = 120
            btnSkip.Top = 230
            btnSkip.Width = 80
            btnSkip.Height = 30
            btnSkip.DialogResult = DialogResult.No
            frm.Controls.Add(btnSkip)
        End If

        Dim btnOK As New System.Windows.Forms.Button()
        btnOK.Text = "OK"
        btnOK.Left = 210
        btnOK.Top = 230
        btnOK.Width = 80
        btnOK.Height = 30
        btnOK.DialogResult = DialogResult.OK
        frm.AcceptButton = btnOK
        frm.Controls.Add(btnOK)

        Dim btnCancel As New System.Windows.Forms.Button()
        btnCancel.Text = "Cancel"
        btnCancel.Left = 300
        btnCancel.Top = 230
        btnCancel.Width = 80
        btnCancel.Height = 30
        btnCancel.DialogResult = DialogResult.Cancel
        frm.CancelButton = btnCancel
        frm.Controls.Add(btnCancel)

        cboAxis.Tag = frm
        AddHandler cboAxis.SelectedIndexChanged, AddressOf OnAxisComboChanged

        Dim result As DialogResult = frm.ShowDialog()
        Dim selectedAxis As String = CStr(cboAxis.SelectedItem)

        If result = DialogResult.Ignore Then
            action = "FLIP"
        ElseIf result = DialogResult.Yes Then
            thicknessAxis = selectedAxis
            action = "SELECT_AXIS"
        ElseIf result = DialogResult.Retry Then
            action = "PICK_THICKNESS"
        End If

        Return result
    End Function

    Public Sub OnAxisComboChanged(sender As Object, e As EventArgs)
        Dim cbo As System.Windows.Forms.ComboBox = CType(sender, System.Windows.Forms.ComboBox)
        Dim frm As System.Windows.Forms.Form = CType(cbo.Tag, System.Windows.Forms.Form)
        Dim selected As String = CStr(cbo.SelectedItem)

        If selected = "Custom" Then
            frm.DialogResult = DialogResult.Retry
            frm.Close()
        ElseIf selected = "X" OrElse selected = "Y" OrElse selected = "Z" Then
            frm.DialogResult = DialogResult.Yes
            frm.Close()
        End If
    End Sub

    Public Function GetAxisSize(ByVal axis As String, ByVal xSize As Double, ByVal ySize As Double, ByVal zSize As Double) As Double
        If axis = "X" Then Return xSize
        If axis = "Y" Then Return ySize
        Return zSize
    End Function

    Public Function PickAxisWithDescription(ByVal app As Inventor.Application, ByRef axisDesc As String) As String
        axisDesc = ""
        
        Dim selFilter As SelectionFilterEnum = SelectionFilterEnum.kAllLinearEntities
        Dim selectedObj As Object = Nothing

        Try
            selectedObj = app.CommandManager.Pick( _
                selFilter, _
                "Select an edge or work axis to define the thickness direction:")
        Catch
            Return ""
        End Try

        If selectedObj Is Nothing Then
            Return ""
        End If

        Dim dirVector As Vector = Nothing
        Dim objName As String = ""

        If TypeOf selectedObj Is Edge Then
            Dim edge As Edge = CType(selectedObj, Edge)
            dirVector = GetEdgeDirection(app, edge)
            objName = "Edge"
        ElseIf TypeOf selectedObj Is WorkAxis Then
            Dim workAxis As WorkAxis = CType(selectedObj, WorkAxis)
            dirVector = GetWorkAxisDirection(app, workAxis)
            Try
                objName = workAxis.Name
            Catch
                objName = "Work Axis"
            End Try
        Else
            MessageBox.Show("Please select a linear edge or work axis.", "Bounding Box Stock")
            Return ""
        End If

        If dirVector Is Nothing Then
            Return ""
        End If

        Dim principalAxis As String = GetPrincipalAxis(dirVector)
        axisDesc = principalAxis & " axis (from " & objName & ")"
        
        Return principalAxis
    End Function

    Public Function GetEdgeDirection(ByVal app As Inventor.Application, ByVal edge As Edge) As Vector
        Try
            Dim geom As Object = edge.Geometry
            If TypeOf geom Is Line Then
                Dim line As Line = CType(geom, Line)
                Dim dir As UnitVector = line.Direction
                Dim vec As Vector = app.TransientGeometry.CreateVector(dir.X, dir.Y, dir.Z)
                Return vec
            ElseIf TypeOf geom Is LineSegment Then
                Dim lineSeg As LineSegment = CType(geom, LineSegment)
                Dim dir As UnitVector = lineSeg.Direction
                Dim vec As Vector = app.TransientGeometry.CreateVector(dir.X, dir.Y, dir.Z)
                Return vec
            End If
        Catch
        End Try
        Return Nothing
    End Function

    Public Function GetWorkAxisDirection(ByVal app As Inventor.Application, ByVal workAxis As WorkAxis) As Vector
        Try
            Dim line As Line = workAxis.Line
            Dim dir As UnitVector = line.Direction
            Dim vec As Vector = app.TransientGeometry.CreateVector(dir.X, dir.Y, dir.Z)
            Return vec
        Catch
        End Try
        Return Nothing
    End Function

    Public Function GetPrincipalAxis(ByVal vec As Vector) As String
        Dim absX As Double = Math.Abs(vec.X)
        Dim absY As Double = Math.Abs(vec.Y)
        Dim absZ As Double = Math.Abs(vec.Z)

        If absX >= absY AndAlso absX >= absZ Then
            Return "X"
        ElseIf absY >= absX AndAlso absY >= absZ Then
            Return "Y"
        Else
            Return "Z"
        End If
    End Function

    Public Function GetRemainingAxis(ByVal axis1 As String, ByVal axis2 As String) As String
        ' Returns the third axis given two axes
        If (axis1 = "X" AndAlso axis2 = "Y") OrElse (axis1 = "Y" AndAlso axis2 = "X") Then
            Return "Z"
        ElseIf (axis1 = "X" AndAlso axis2 = "Z") OrElse (axis1 = "Z" AndAlso axis2 = "X") Then
            Return "Y"
        Else
            Return "X"
        End If
    End Function

    Public Sub ParseAxisConfig(ByVal config As String, ByRef thicknessAxis As String, ByRef widthAxis As String, ByRef lengthAxis As String)
        Dim parts() As String = config.Split(","c)
        For Each part As String In parts
            Dim kv() As String = part.Split(":"c)
            If kv.Length = 2 Then
                If kv(0) = "T" Then thicknessAxis = kv(1)
                If kv(0) = "W" Then widthAxis = kv(1)
                If kv(0) = "L" Then lengthAxis = kv(1)
            End If
        Next
    End Sub

    Public Sub CreateOrUpdateRule(ByVal partDoc As PartDocument, ByVal thicknessAxis As String, ByVal widthAxis As String, ByVal lengthAxis As String, ByVal iLogicAuto As Object)
        ' Only store thickness and width - length is determined by the other two
        SetCustomProperty(partDoc, "BB_ThicknessAxis", thicknessAxis)
        SetCustomProperty(partDoc, "BB_WidthAxis", widthAxis)

        Dim ruleName As String = "BoundingBoxStockUpdate"
        Dim ruleText As String = BuildRuleText()

        Dim existingRule As Object = Nothing

        Try
            existingRule = iLogicAuto.GetRule(partDoc, ruleName)
        Catch
            existingRule = Nothing
        End Try

        If existingRule IsNot Nothing Then
            existingRule.Text = ruleText
        Else
            iLogicAuto.AddRule(partDoc, ruleName, ruleText)
        End If

        iLogicAuto.RunRule(partDoc, ruleName)
    End Sub

    Public Sub SetCustomProperty(ByVal doc As Document, ByVal propName As String, ByVal propValue As String)
        Dim propSet As PropertySet = doc.PropertySets.Item("Inventor User Defined Properties")
        Try
            propSet.Item(propName).Value = propValue
        Catch
            propSet.Add(propValue, propName)
        End Try
    End Sub

    Public Function GetCustomPropertyValue(ByVal doc As Document, ByVal propName As String, ByVal defaultValue As String) As String
        Try
            Dim propSet As PropertySet = doc.PropertySets.Item("Inventor User Defined Properties")
            Return CStr(propSet.Item(propName).Value)
        Catch
            Return defaultValue
        End Try
    End Function

    Public Function BuildRuleText() As String
        Dim sb As New System.Text.StringBuilder()
        
        sb.AppendLine("' Auto-generated rule: Updates Width, Length, Thickness iProperties from bounding box")
        sb.AppendLine("' Override by creating parameters: WidthOverride, LengthOverride, ThicknessOverride")
        sb.AppendLine("")
        sb.AppendLine("Sub Main()")
        sb.AppendLine("    Dim partDoc As PartDocument = CType(ThisDoc.Document, PartDocument)")
        sb.AppendLine("")
        sb.AppendLine("    ' Get bounding box")
        sb.AppendLine("    Dim rangebox As Box = partDoc.ComponentDefinition.RangeBox")
        sb.AppendLine("    Dim xSize As Double = rangebox.MaxPoint.X - rangebox.MinPoint.X")
        sb.AppendLine("    Dim ySize As Double = rangebox.MaxPoint.Y - rangebox.MinPoint.Y")
        sb.AppendLine("    Dim zSize As Double = rangebox.MaxPoint.Z - rangebox.MinPoint.Z")
        sb.AppendLine("")
        sb.AppendLine("    ' Read axis configuration from iProperties (length is derived from the other two)")
        sb.AppendLine("    Dim thicknessAxis As String = GetCustomProp(partDoc, ""BB_ThicknessAxis"", ""Z"")")
        sb.AppendLine("    Dim widthAxis As String = GetCustomProp(partDoc, ""BB_WidthAxis"", ""X"")")
        sb.AppendLine("    Dim lengthAxis As String = GetRemainingAxis(thicknessAxis, widthAxis)")
        sb.AppendLine("")
        sb.AppendLine("    ' Calculate values based on axis mapping")
        sb.AppendLine("    Dim thicknessVal As Double = GetAxisValue(thicknessAxis, xSize, ySize, zSize)")
        sb.AppendLine("    Dim widthVal As Double = GetAxisValue(widthAxis, xSize, ySize, zSize)")
        sb.AppendLine("    Dim lengthVal As Double = GetAxisValue(lengthAxis, xSize, ySize, zSize)")
        sb.AppendLine("")
        sb.AppendLine("    ' Check for overrides")
        sb.AppendLine("    thicknessVal = GetOverrideOrValue(partDoc, ""ThicknessOverride"", thicknessVal)")
        sb.AppendLine("    widthVal = GetOverrideOrValue(partDoc, ""WidthOverride"", widthVal)")
        sb.AppendLine("    lengthVal = GetOverrideOrValue(partDoc, ""LengthOverride"", lengthVal)")
        sb.AppendLine("")
        sb.AppendLine("    ' Update iProperties (always in mm with 1 decimal place for consistency)")
        sb.AppendLine("    ' Internal values are in cm, convert to mm (* 10)")
        sb.AppendLine("    Dim thicknessStr As String = FormatDimensionMm(thicknessVal * 10)")
        sb.AppendLine("    Dim widthStr As String = FormatDimensionMm(widthVal * 10)")
        sb.AppendLine("    Dim lengthStr As String = FormatDimensionMm(lengthVal * 10)")
        sb.AppendLine("")
        sb.AppendLine("    SetCustomProp(partDoc, ""Thickness"", thicknessStr)")
        sb.AppendLine("    SetCustomProp(partDoc, ""Width"", widthStr)")
        sb.AppendLine("    SetCustomProp(partDoc, ""Length"", lengthStr)")
        sb.AppendLine("End Sub")
        sb.AppendLine("")
        sb.AppendLine("Function GetAxisValue(axis As String, xSize As Double, ySize As Double, zSize As Double) As Double")
        sb.AppendLine("    If axis = ""X"" Then Return xSize")
        sb.AppendLine("    If axis = ""Y"" Then Return ySize")
        sb.AppendLine("    Return zSize")
        sb.AppendLine("End Function")
        sb.AppendLine("")
        sb.AppendLine("Function GetRemainingAxis(axis1 As String, axis2 As String) As String")
        sb.AppendLine("    If (axis1 = ""X"" AndAlso axis2 = ""Y"") OrElse (axis1 = ""Y"" AndAlso axis2 = ""X"") Then Return ""Z""")
        sb.AppendLine("    If (axis1 = ""X"" AndAlso axis2 = ""Z"") OrElse (axis1 = ""Z"" AndAlso axis2 = ""X"") Then Return ""Y""")
        sb.AppendLine("    Return ""X""")
        sb.AppendLine("End Function")
        sb.AppendLine("")
        sb.AppendLine("Function GetCustomProp(doc As Document, propName As String, defaultVal As String) As String")
        sb.AppendLine("    Try")
        sb.AppendLine("        Return CStr(doc.PropertySets.Item(""Inventor User Defined Properties"").Item(propName).Value)")
        sb.AppendLine("    Catch")
        sb.AppendLine("        Return defaultVal")
        sb.AppendLine("    End Try")
        sb.AppendLine("End Function")
        sb.AppendLine("")
        sb.AppendLine("Function GetOverrideOrValue(doc As PartDocument, paramName As String, calcValue As Double) As Double")
        sb.AppendLine("    Try")
        sb.AppendLine("        Dim param As Parameter = doc.ComponentDefinition.Parameters.Item(paramName)")
        sb.AppendLine("        If param IsNot Nothing Then")
        sb.AppendLine("            Return param.Value")
        sb.AppendLine("        End If")
        sb.AppendLine("    Catch")
        sb.AppendLine("    End Try")
        sb.AppendLine("    Return calcValue")
        sb.AppendLine("End Function")
        sb.AppendLine("")
        sb.AppendLine("Sub SetCustomProp(doc As Document, propName As String, propValue As String)")
        sb.AppendLine("    Dim propSet As PropertySet = doc.PropertySets.Item(""Inventor User Defined Properties"")")
        sb.AppendLine("    Try")
        sb.AppendLine("        propSet.Item(propName).Value = propValue")
        sb.AppendLine("    Catch")
        sb.AppendLine("        propSet.Add(propValue, propName)")
        sb.AppendLine("    End Try")
        sb.AppendLine("End Sub")
        sb.AppendLine("")
        sb.AppendLine("Function FormatDimensionMm(valueMm As Double) As String")
        sb.AppendLine("    Return valueMm.ToString(""0.0"", System.Globalization.CultureInfo.InvariantCulture) & "" mm""")
        sb.AppendLine("End Function")

        Return sb.ToString()
    End Function

End Module

